import os
import curses
import zipfile
import fnmatch
import shutil
import subprocess
import argparse
import threading
import queue
import time
import types
# import winsound  # Standard on Windows
from collections import Counter, deque
from datetime import datetime

# Paths to DOS emulator and PKunzip — assumed to be in the same folder as this script
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
MSDOS_EXE   = os.path.join(SCRIPT_DIR, 'msdos.exe')
PKUNZIP_EXE = os.path.join(SCRIPT_DIR, 'PKUNZIP.EXE')
PKUNPAK_EXE  = os.path.join(SCRIPT_DIR, 'PKUNPAK.EXE')

#def play_success_sound():
    # Plays a three-note rising chime
#    winsound.Beep(440, 200)
#    winsound.Beep(554, 200)
#    winsound.Beep(659, 400)


# ---------------------------------------------------------------------------
# Curses TUI
# ---------------------------------------------------------------------------

class CursesUI:
    """
    Full-screen terminal dashboard using curses.

    Layout (top → bottom):
        Title bar
        ─────────────────────────────
        Status  : <current operation>
        Current : <filename>
        Progress: [████░░░░░░] n/total (pct%)
        ─────────────────────────────
        Archives   n   Successful  n   Queued    n
        Legacy ZIP n   Zip Errors  n   Files     n
        ARC        n   ARC Errors  n
        [--ext only] Skipped   n
        [--ext only] Ext matches → .ext: n  …
        ─────────────────────────────
        Recent Activity:        │ Live Extraction:
          … scrolling log …     │   … file ops stream …
    """

    COLOR_HEADER  = 1
    COLOR_SUCCESS = 2
    COLOR_ERROR   = 3
    COLOR_WARNING = 4
    COLOR_NORMAL  = 5
    COLOR_DIM     = 6
    COLOR_LIVE    = 7  # plain white (no bold) — softer than COLOR_NORMAL

    def __init__(self, stdscr):
        self.stdscr = stdscr
        self._status       = "Initialising…"
        self._current_file = ""
        self._prog_cur     = 0
        self._prog_total   = 0
        self._log_lines    = deque(maxlen=200)  # list of (text, color_pair_int)
        self._live_queue   = queue.Queue()          # thread-safe feed from reader threads
        self._live_lines   = deque(maxlen=200)      # display buffer for live window
        self._eta_window   = deque(maxlen=10)       # rolling last-10 archive durations (seconds)
        self._live_log     = None                   # optional file handle for --log mode
        self._process_all  = True        # False when --ext flag is active
        self._target_exts  = ()          # populated by set_mode when --ext is active
        self._ext_counts   = {}          # ext → count, populated only in --ext mode
        self.stats = dict(
            processed=0, skipped=0, successful=0,
            queued=0,
            reduce=0, crc_errors=0, files=0,
            arc=0, arc_crc=0, zips=0,
        )

        curses.start_color()
        curses.use_default_colors()
        curses.init_pair(self.COLOR_HEADER,  curses.COLOR_CYAN,    -1)
        curses.init_pair(self.COLOR_SUCCESS, curses.COLOR_GREEN,   -1)
        curses.init_pair(self.COLOR_ERROR,   curses.COLOR_RED,     -1)
        curses.init_pair(self.COLOR_WARNING, curses.COLOR_YELLOW,  -1)
        curses.init_pair(self.COLOR_NORMAL,  curses.COLOR_WHITE,   -1)
        curses.init_pair(self.COLOR_DIM,     curses.COLOR_BLACK,   -1)  # bold black = dark grey
        curses.init_pair(self.COLOR_LIVE,    curses.COLOR_WHITE,   -1)  # plain white, no bold

        curses.curs_set(0)
        self.stdscr.nodelay(True)
        self.draw()

    # ── public API ──────────────────────────────────────────────────────────

    def record_eta(self, elapsed_seconds: float):
        """Record one archive's elapsed time into the rolling window.
        Does not call draw() — the next natural draw() will pick up the updated window."""
        if elapsed_seconds > 0:
            self._eta_window.append(elapsed_seconds)

    @staticmethod
    def _fmt_eta(seconds: float) -> str:
        """Format a duration in seconds as a human-readable ETA string."""
        s = int(seconds)
        if s < 60:
            return f"{s}s"
        m, s = divmod(s, 60)
        if m < 60:
            return f"{m}m {s:02d}s"
        h, m = divmod(m, 60)
        return f"{h}h {m:02d}m"

    def set_live_log(self, file_handle):
        """Attach an open file handle to mirror all live and activity output."""
        self._live_log = file_handle

    def set_status(self, msg: str):
        self._status = msg
        self.draw()

    def set_current_file(self, fname: str):
        self._current_file = fname
        self.draw()

    def set_progress(self, current: int, total: int):
        self._prog_cur   = current
        self._prog_total = total
        self.draw()

    def update_stats(self, **kwargs):
        self.stats.update(kwargs)
        self.draw()

    def set_mode(self, process_all: bool, target_exts: tuple = ()):
        """Call once at startup to record whether --ext filtering is active."""
        self._process_all  = process_all
        self._target_exts  = target_exts
        self.draw()

    def update_ext_counts(self, counts: dict):
        """Refresh the per-extension match tally (only meaningful in --ext mode)."""
        self._ext_counts = dict(counts)
        self.draw()

    def log(self, msg: str, level: str = "normal"):
        color_map = {
            "success": self.COLOR_SUCCESS,
            "error":   self.COLOR_ERROR,
            "warning": self.COLOR_WARNING,
            "header":  self.COLOR_HEADER,
            "normal":  self.COLOR_NORMAL,
            "dim":     self.COLOR_DIM,
        }
        pair = color_map.get(level, self.COLOR_NORMAL)
        self._log_lines.append((msg, pair))
        if self._live_log:
            self._live_log.write(f"[{level.upper():<7}] {msg}\n")
            self._live_log.flush()
        self.draw()

    def push_live(self, msg: str, level: str = "live"):
        """Thread-safe: enqueue a line for the live extraction panel (no draw call)."""
        color_map = {
            "success": self.COLOR_SUCCESS,
            "error":   self.COLOR_ERROR,
            "warning": self.COLOR_WARNING,
            "normal":  self.COLOR_NORMAL,
            "dim":     self.COLOR_DIM,
            "live":    self.COLOR_LIVE,
        }
        self._live_queue.put((msg, color_map.get(level, self.COLOR_LIVE)))
        if self._live_log:
            self._live_log.write(f"  {msg}\n")
            self._live_log.flush()

    def wait_for_key(self, prompt: str = "  Press any key to exit…"):
        h, w = self.stdscr.getmaxyx()
        self.stdscr.nodelay(False)
        try:
            self.stdscr.attron(curses.color_pair(self.COLOR_WARNING) | curses.A_BOLD)
            self.stdscr.addstr(h - 1, 0, prompt[: w - 1])
            self.stdscr.attroff(curses.color_pair(self.COLOR_WARNING) | curses.A_BOLD)
        except curses.error:
            pass
        self.stdscr.refresh()
        self.stdscr.getch()

    # ── internal drawing ─────────────────────────────────────────────────────

    def _safe_addstr(self, row: int, col: int, text: str, attr: int = 0, max_w: int = None):
        h, w = self.stdscr.getmaxyx()
        if row < 0 or row >= h - 1:
            return
        max_len = (max_w if max_w is not None else w) - col - 1
        if max_len <= 0:
            return
        try:
            self.stdscr.addstr(row, col, text[:max_len], attr)
        except curses.error:
            pass

    def _hline(self, row: int):
        h, w = self.stdscr.getmaxyx()
        if 0 <= row < h - 1:
            try:
                self.stdscr.hline(row, 0, curses.ACS_HLINE, w - 1)
            except curses.error:
                pass

    def draw(self):
        try:
            h, w = self.stdscr.getmaxyx()
            self.stdscr.erase()
            row = 0
            col_split = w // 2
            s = self.stats

            # ── Title bar ──────────────────────────────────────────────
            title = "  Archive Extractor — ZIP / ARC  "
            attr  = curses.color_pair(self.COLOR_HEADER) | curses.A_BOLD | curses.A_REVERSE
            padded = title.center(w - 1)
            self._safe_addstr(row, 0, padded, attr)
            row += 1
            self._hline(row); row += 1

            # ── Status / current file ──────────────────────────────────
            self._safe_addstr(row, 0, f" Status  : {self._status}",
                              curses.color_pair(self.COLOR_NORMAL))
            row += 1
            if self._process_all:
                mode_text = " Mode    : ALL files"
                mode_attr = curses.color_pair(self.COLOR_DIM) | curses.A_BOLD
            else:
                exts_str  = "  ".join(e.lstrip('.').upper() for e in self._target_exts)
                mode_text = f" EXT     : {exts_str}"
                mode_attr = curses.color_pair(self.COLOR_WARNING) | curses.A_BOLD
            self._safe_addstr(row, 0, mode_text, mode_attr)
            row += 1
            cf = self._current_file or "—"
            self._safe_addstr(row, 0, f" Current : {cf}",
                              curses.color_pair(self.COLOR_DIM) | curses.A_BOLD)
            row += 1

            # ── Progress bar ───────────────────────────────────────────
            if self._prog_total > 0:
                pct     = self._prog_cur / self._prog_total
                bar_w   = max(10, w - 32)
                filled  = int(bar_w * pct)
                bar     = "█" * filled + "░" * (bar_w - filled)
                prog_ln = f" [{bar}] {self._prog_cur}/{self._prog_total} ({pct:.0%})"
                bar_colour = (curses.color_pair(self.COLOR_SUCCESS)
                              if pct >= 1.0 else
                              curses.color_pair(self.COLOR_WARNING)
                              if pct < 0.5 else
                              curses.color_pair(self.COLOR_SUCCESS))
                self._safe_addstr(row, 0, prog_ln, bar_colour)
                row += 1

                # ── ETA line ───────────────────────────────────────────
                remaining = max(0, s["queued"] - s["processed"] - s["skipped"])
                if self._eta_window and remaining > 0:
                    avg      = sum(self._eta_window) / len(self._eta_window)
                    eta_secs = avg * remaining
                    samples  = len(self._eta_window)
                    eta_str  = self._fmt_eta(eta_secs)
                    eta_ln   = f" ETA ~{eta_str}  (avg {avg:.1f}s/archive, {samples} sample{'s' if samples != 1 else ''})"
                    self._safe_addstr(row, 0, eta_ln,
                                      curses.color_pair(self.COLOR_NORMAL) | curses.A_BOLD)
                elif remaining == 0:
                    self._safe_addstr(row, 0, " ETA —  complete",
                                      curses.color_pair(self.COLOR_DIM) | curses.A_BOLD)
                else:
                    self._safe_addstr(row, 0, " ETA —  calculating…",
                                      curses.color_pair(self.COLOR_DIM) | curses.A_BOLD)
                row += 1
            else:
                self._safe_addstr(row, 0, " [waiting for task…]",
                                  curses.color_pair(self.COLOR_DIM) | curses.A_BOLD)
                row += 1
            self._hline(row); row += 1

            # ── Stats grid ─────────────────────────────────────────────
            def stat_line(pairs):
                col_w = (w - 2) // 3
                parts = []
                for label, val in pairs:
                    cell = f"{label}: {val}"
                    parts.append(f"{cell:<{col_w}}")
                return "  " + "".join(parts)

            self._safe_addstr(row, 0,
                stat_line([("Archives",   s["processed"]),
                           ("Successful", s["successful"]),
                           ("Queued",     max(0, s["queued"] - s["processed"] - s["skipped"]))]),
                curses.color_pair(self.COLOR_NORMAL))
            row += 1
            self._safe_addstr(row, 0,
                stat_line([("ZIP",        s["zips"]),
                           ("Legacy ZIP", s["reduce"]),
                           ("Zip Errors", s["crc_errors"])]),
                curses.color_pair(self.COLOR_WARNING))
            row += 1
            self._safe_addstr(row, 0,
                stat_line([("ARC",        s["arc"]),
                           ("ARC Errors", s["arc_crc"]),
                           ("Files",      s["files"])]),
                curses.color_pair(self.COLOR_NORMAL))
            row += 1

            # ── Skipped row (--ext mode only) ──────────────────────────
            if not self._process_all:
                self._safe_addstr(row, 0,
                    stat_line([("Skipped",    s["skipped"])]),
                    curses.color_pair(self.COLOR_WARNING))
                row += 1

            # ── Per-extension match counts (--ext mode only) ───────────
            if not self._process_all and self._ext_counts:
                # Sort by count descending, render as many as fit on one row
                sorted_exts = sorted(self._ext_counts.items(),
                                     key=lambda kv: kv[1], reverse=True)
                tokens = [f"{ext}: {cnt}" for ext, cnt in sorted_exts]
                # Pack tokens into rows of width w
                indent = "  Ext matches → "
                line   = indent
                first  = True
                for tok in tokens:
                    segment = tok if first else "   " + tok
                    if len(line) + len(segment) >= w - 1 and not first:
                        self._safe_addstr(row, 0, line,
                                          curses.color_pair(self.COLOR_SUCCESS))
                        row += 1
                        line = " " * len(indent) + tok
                        first = False
                        continue
                    line  += segment
                    first  = False
                if line.strip():
                    self._safe_addstr(row, 0, line,
                                      curses.color_pair(self.COLOR_SUCCESS))
                    row += 1
            self._hline(row)
            # T-junction where the horizontal separator meets the vertical panel divider
            try:
                self.stdscr.addch(row, col_split, curses.ACS_TTEE,
                                  curses.color_pair(self.COLOR_DIM) | curses.A_BOLD)
            except curses.error:
                pass
            row += 1
            try:
                while True:
                    self._live_lines.append(self._live_queue.get_nowait())
            except queue.Empty:
                pass

            # ── Two-panel log area ──────────────────────────────────────

            # Panel headers
            self._safe_addstr(row, 0, " Recent Activity:",
                              curses.color_pair(self.COLOR_HEADER) | curses.A_BOLD,
                              max_w=col_split)
            self._safe_addstr(row, col_split + 1, " Live Extraction:",
                              curses.color_pair(self.COLOR_HEADER) | curses.A_BOLD)
            row += 1

            # Vertical divider for remaining rows
            for r in range(row, h - 1):
                try:
                    self.stdscr.addch(r, col_split, curses.ACS_VLINE,
                                      curses.color_pair(self.COLOR_DIM))
                except curses.error:
                    pass

            # Render both panels side by side
            log_rows_available = h - row - 1
            visible_log  = list(self._log_lines)[-log_rows_available:] if log_rows_available > 0 else []
            visible_live = list(self._live_lines)[-log_rows_available:] if log_rows_available > 0 else []

            for i in range(log_rows_available):
                r = row + i
                if r >= h - 1:
                    break
                if i < len(visible_log):
                    msg, pair = visible_log[i]
                    self._safe_addstr(r, 0, f"  {msg}",
                                      curses.color_pair(pair), max_w=col_split)
                if i < len(visible_live):
                    msg, pair = visible_live[i]
                    self._safe_addstr(r, col_split + 1, f" {msg}",
                                      curses.color_pair(pair))

            self.stdscr.refresh()
        except curses.error:
            pass


# ---------------------------------------------------------------------------
# Helper utilities
# ---------------------------------------------------------------------------

def get_sort_folder(matches, wildcard_pattern, target_exts, sorted_root):
    """
    Determine which subfolder inside _sorted this ZIP should go into.
    - Single extension match   -> _sorted_zip/.ima/  (etc.)
    - Wildcard-only match      -> _sorted_zip/_wildcard/
    - Multiple extensions      -> _sorted_zip/mixed/
    - Wildcard + ext(s)        -> _sorted_zip/mixed/
    """
    exts_found = set()
    has_wildcard = False

    for name in matches:
        low = name.lower()
        if fnmatch.fnmatch(low, wildcard_pattern):
            has_wildcard = True
        if low.endswith(target_exts):
            _, ext = os.path.splitext(low)
            if ext:
                exts_found.add(ext)

    # Decide folder name
    if has_wildcard and not exts_found:
        folder_name = "_wildcard"
    elif not has_wildcard and len(exts_found) == 1:
        folder_name = exts_found.pop()          # e.g. ".ima"
    else:
        folder_name = "mixed"                   # multiple exts, or wildcard + ext

    return os.path.join(sorted_root, folder_name)


def read_zip_comment_raw(zip_path):
    """
    Read the ZIP comment directly from raw bytes using the EOCD record.
    Returns the raw comment bytes, or empty bytes if not found.
    """
    try:
        with open(zip_path, 'rb') as f:
            data = f.read()
        eocd_sig = b'\x50\x4B\x05\x06'
        pos = data.rfind(eocd_sig)  # rfind — use last occurrence
        if pos == -1:
            return b''
        comment_length = int.from_bytes(data[pos+20:pos+22], 'little')
        if comment_length == 0:
            return b''
        return data[pos+22:pos+22+comment_length]
    except Exception:
        return b''


def force_remove(func, path, _):
    """Read-only safe removal helper for shutil.rmtree."""
    import stat
    os.chmod(path, stat.S_IWRITE)
    func(path)


def process_nested_zips(folder, error_log, depth=0, max_depth=10, processed=None,
                        log_buffer=None, reduce_buffer=None):
    """
    Recursively find and extract any ZIP files inside an already-extracted folder.
    Each nested ZIP gets its own child folder, metadata file, and the ZIP is moved in.
    """
    if depth >= max_depth:
        return

    if processed is None:
        processed = set()

    # Snapshot all ZIP files in the entire folder tree before any extraction
    zip_files = []
    for dirpath, dirnames, filenames in os.walk(folder):
        for fname in filenames:
            if fname.lower().endswith('.zip'):
                zip_files.append(os.path.join(dirpath, fname))

    for nested_zip_path_raw in zip_files:
        nested_zip_path = os.path.abspath(nested_zip_path_raw)
        item = os.path.basename(nested_zip_path)
        nested_folder_name = os.path.splitext(item)[0]
        nested_dest = os.path.join(os.path.dirname(nested_zip_path), nested_folder_name)

        # Skip if already processed or no longer exists
        if nested_zip_path in processed or not os.path.isfile(nested_zip_path):
            continue

        processed.add(nested_zip_path)

        # Collision guard
        if os.path.exists(nested_dest):
            counter = 1
            while os.path.exists(f"{nested_dest}_{counter}-Dupe"):
                counter += 1
            nested_dest = f"{nested_dest}_{counter}-Dupe"

        _safe_makedirs(nested_dest)

        # Pre-detect Reduce compression (methods 2-5) by checking ZIP entries before extraction
        is_reduce = False
        try:
            with zipfile.ZipFile(nested_zip_path, 'r') as z:
                if any(info.compress_type in (2, 3, 4, 5) for info in z.infolist()):
                    is_reduce = True
        except Exception:
            pass

        if is_reduce:
            # Force into Reduce branch without trying 7-Zip
            result_returncode = 2
            sevenzip_output = "ERROR: Unsupported Method"
        else:
            # Try 7-Zip first
            result = subprocess.run(
                ['7z', 'x', nested_zip_path, f'-o{nested_dest}', '-y'],
                stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True
            )
            sevenzip_output = result.stdout + result.stderr
            result_returncode = result.returncode

        if result_returncode == 0 and "ERROR: Unsupported Method" not in sevenzip_output:
            # Write metadata
            comment_bytes = read_zip_comment_raw(nested_zip_path)
            disk_label = read_disk_label(nested_zip_path)
            meta_path = os.path.join(nested_dest, f"__{item}__metadata.nfo")
            with open(meta_path, 'wb') as m:
                m.write(f"Source ZIP: {item}\n".encode('utf-8'))
                m.write(f"Disk Label: {disk_label}\n".encode('utf-8'))
                m.write(b"Archive Comment: ")
                m.write(comment_bytes if comment_bytes else b"None")
                m.write(b"\n")
            # Recurse BEFORE moving ZIP in so snapshot won't find it
            process_nested_zips(nested_dest, error_log, depth + 1, max_depth, processed,
                                 log_buffer, reduce_buffer)
            shutil.move(nested_zip_path, os.path.join(nested_dest, item))

            # Log to success buffer
            if log_buffer is not None:
                indent = "  " * (depth + 1)
                log_buffer.append(f"{indent}[Nested] ZIP OK: {item}\n")
                log_buffer.append(f"{indent}  Location: {nested_dest}\n")

        elif "ERROR: Unsupported Method" in sevenzip_output:
            # Reduce nested ZIP — use pkunzip
            _safe_rmtree(nested_dest)
            _safe_makedirs(nested_dest)
            moved_zip = os.path.join(nested_dest, item)
            shutil.move(nested_zip_path, moved_zip)
            processed.add(os.path.abspath(moved_zip))
            subprocess.run(
                [MSDOS_EXE, '-d', PKUNZIP_EXE, '-e', '-o', '-d', item],
                cwd=nested_dest,
                stdout=subprocess.DEVNULL, stderr=subprocess.PIPE, text=True
            )
            comment_bytes = read_zip_comment_raw(moved_zip)
            disk_label = read_disk_label(moved_zip)
            meta_path = os.path.join(nested_dest, f"__{item}__metadata.nfo")
            with open(meta_path, 'wb') as m:
                m.write(f"Source ZIP: {item}\n".encode('utf-8'))
                m.write(f"Disk Label: {disk_label}\n".encode('utf-8'))
                m.write(b"Archive Comment: ")
                m.write(comment_bytes if comment_bytes else b"None")
                m.write(b"\n")
            process_nested_zips(nested_dest, error_log, depth + 1, max_depth, processed,
                                 log_buffer, reduce_buffer)

            # Log to Reduce buffer
            if reduce_buffer is not None:
                indent = "  " * (depth + 1)
                reduce_buffer.append(f"{indent}[Nested] Legacy ZIP: {item}\n")
                reduce_buffer.append(f"{indent}  Location: {nested_dest}\n")

        else:
            # Failed — clean up empty folder and log it
            _safe_rmtree(nested_dest)
            with open(error_log, 'a', encoding='utf-8', errors='replace') as f_err:
                f_err.write(f"NESTED ZIP ERROR (depth {depth+1}) {item}: {result.stderr}\n")


def read_disk_label(zip_path):
    """
    Read the disk label from a ZIP file using the byte pattern method.
    Returns the label string, or 'None Detected' if not found.
    """
    try:
        with open(zip_path, 'rb') as f:
            data = f.read()
        pattern = b'\x50\x4B\x01\x02\x0B\x00\x0B'
        pos = data.find(pattern)
        if pos != -1:
            label_length_pos = pos + 28
            if label_length_pos < len(data):
                label_length = data[label_length_pos]
                label_start = label_length_pos + 1 + 17
                if label_start + label_length <= len(data):
                    label = data[label_start:label_start + label_length].decode('utf-8', errors='replace').rstrip('\x00')
                    if label:
                        return label
    except Exception:
        pass
    return 'None Detected'


def read_arc_comment(arc_path):
    """
    Read the comment from an ARC file using the end-of-archive marker (0x1A).
    Returns raw comment bytes, or empty bytes if not found.
    """
    try:
        with open(arc_path, 'rb') as f:
            data = f.read()

        marker = data.rfind(b'\x1a')
        if marker == -1:
            return b''

        comment_bytes = data[marker + 1:]

        if len(comment_bytes) < 4:
            return b''

        field_len = comment_bytes[3]
        raw = comment_bytes[0 : 3 + field_len]

        # Truncate at PK signature if present (ZIP appended to ARC)
        pk_pos = raw.find(b'\x50\x4b')
        if pk_pos != -1:
            raw = raw[:pk_pos]

        # Strip space and null padding from both ends
        result = raw.strip(b'\x20\x00')
        return result if result else b''

    except Exception:
        return b''


def get_arc_members(arc_path):
    """
    Read member filenames from an ARC file by scanning headers.
    Returns a list of member filename strings.
    """
    members = []
    try:
        with open(arc_path, 'rb') as f:
            data = f.read()
        pos = 0
        while pos < len(data):
            if data[pos] != 0x1a:
                break
            header_type = data[pos + 1] if pos + 1 < len(data) else 0
            if header_type == 0:
                break
            name_end = data.find(b'\x00', pos + 2, pos + 17)
            if name_end == -1:
                break
            name = data[pos + 2:name_end].decode('ascii', errors='replace').lower()
            members.append(name)
            size = int.from_bytes(data[pos + 15:pos + 19], 'little') if pos + 19 <= len(data) else 0
            pos += 29 + size
    except Exception:
        pass
    return members


# ---------------------------------------------------------------------------
# Output file counter
# ---------------------------------------------------------------------------

def count_output_files():
    """Count all files currently present in both sorted output directories."""
    count = 0
    for root_dir in ('_sorted_zip', '_sorted_arc'):
        p = os.path.abspath(root_dir)
        if os.path.isdir(p):
            for _, _, fs in os.walk(p):
                count += len(fs)
    return count


# ---------------------------------------------------------------------------
# Robust directory creation (Windows NTFS timing workaround)
# ---------------------------------------------------------------------------

def _safe_makedirs(path: str, retries: int = 5, delay: float = 0.15):
    """
    os.makedirs with retry logic for Windows NTFS timing issues.
    After a rmtree, Windows may not immediately release the path — retrying
    with a short sleep resolves the transient WinError 3 / WinError 5.
    """
    for attempt in range(retries):
        try:
            os.makedirs(path, exist_ok=True)
            return
        except OSError:
            if attempt < retries - 1:
                time.sleep(delay)
            else:
                raise


def _safe_rmtree(path: str):
    """
    Remove a directory tree only if it exists.
    Prevents WinError 2 / WinError 3 from force_remove being called
    on a path that was already removed or never created.
    """
    if os.path.exists(path):
        shutil.rmtree(path, onerror=force_remove)




def _stream_to_live(stdout_pipe, ui, level="live"):
    """
    Read lines from a subprocess stdout pipe and push each into the UI live
    window queue.  Designed to run in a daemon thread — only touches the
    thread-safe queue, never calls draw() directly.
    """
    try:
        for raw in stdout_pipe:
            line = raw.rstrip('\r\n')
            if line.strip():
                ui.push_live(line, level)
    finally:
        try:
            stdout_pipe.close()
        except Exception:
            pass


def _collect_stderr(stderr_pipe, buf: list):
    """
    Read stderr from a subprocess into buf[0] on a daemon thread,
    preventing pipe-buffer deadlocks when the process writes a lot to stderr.
    """
    try:
        buf[0] = stderr_pipe.read()
    finally:
        try:
            stderr_pipe.close()
        except Exception:
            pass




def process_arcs(report_file, error_file, arc_buffer, arc_crc_buffer,
                 total_arc, total_arc_crc, process_all=True, target_exts=(),
                 extension_counts=None, total_matches=0,
                 total_skipped=0, total_successful=0, total_processed=0,
                 ui: CursesUI = None):
    """
    Scan recursively for .arc files, extract with PKXARC via msdos.exe,
    write metadata, and sort into _sorted_arc/Arc/<name>/.
    Returns updated (total_arc, total_arc_crc) counts.
    """
    arc_root     = os.path.abspath('_sorted_arc')
    arc_folder   = os.path.join(arc_root, 'Arc')
    arc_err_root = os.path.join(arc_root, 'CRC-Errors')

    if ui:
        ui.set_status("Pre-scanning for ARC files…")
    sorted_root  = os.path.abspath('_sorted_zip')
    arc_tasks = []
    for root, dirs, files in os.walk('.'):
        # Don't walk into any of our output folders
        dirs[:] = [d for d in dirs if not os.path.abspath(os.path.join(root, d)).startswith(
                   (sorted_root, arc_root))]
        for file in files:
            if file.lower().endswith('.arc'):
                arc_tasks.append((root, file))

    if not arc_tasks:
        if ui:
            ui.log("No ARC files found.", "dim")
        return total_arc, total_arc_crc, total_matches, total_skipped, total_successful, total_processed

    # Apply extension filter if --ext flag is set
    if not process_all and target_exts:
        filtered = []
        for root, file in arc_tasks:
            arc_path = os.path.normpath(os.path.join(root, file))
            try:
                with open(arc_path, 'rb') as f:
                    data = f.read()
                # Scan ARC member headers for matching extensions
                pos = 0
                matched = False
                while pos < len(data):
                    if data[pos] != 0x1a:
                        break
                    header_type = data[pos + 1] if pos + 1 < len(data) else 0
                    if header_type == 0:
                        break
                    name_end = data.find(b'\x00', pos + 2, pos + 17)
                    if name_end == -1:
                        break
                    name = data[pos + 2:name_end].decode('ascii', errors='replace').lower()
                    if any(name.endswith(ext) for ext in target_exts):
                        matched = True
                        break
                    size = int.from_bytes(data[pos + 15:pos + 19], 'little') if pos + 19 <= len(data) else 0
                    pos += 29 + size
                if matched:
                    filtered.append((root, file))
            except Exception:
                pass
        arc_tasks_all = arc_tasks
        arc_tasks = filtered
        total_skipped += len(arc_tasks_all) - len(arc_tasks)
        if not arc_tasks:
            if ui:
                ui.log("No ARC files found matching target extensions.", "warning")
            return total_arc, total_arc_crc, total_matches, total_skipped, total_successful, total_processed

    dirs_to_delete = set()
    cwd = os.path.abspath('.')

    def queue_arc_deletion(path):
        current = os.path.abspath(path)
        while current != cwd:
            dirs_to_delete.add(current)
            current = os.path.dirname(current)

    if ui:
        ui.set_status("Processing ARCs…")
        ui.set_progress(0, len(arc_tasks))

    with open(error_file, 'a', encoding='utf-8', errors='replace') as f_err:
        for idx, (root, file) in enumerate(arc_tasks):
            if ui:
                ui.set_current_file(file)
                ui.set_progress(idx + 1, len(arc_tasks))

            arc_path = os.path.normpath(os.path.join(root, file))
            folder_name = os.path.splitext(file)[0]

            # File may have been moved by a previous iteration
            if not os.path.isfile(arc_path):
                if ui:
                    ui.log(f"Skipped (already moved): {file}", "dim")
                continue
            total_processed += 1
            if ui:
                ui.update_stats(processed=total_processed)
            _archive_start = time.monotonic()
            members = get_arc_members(arc_path)
            sort_folder = get_sort_folder(members, "*.?@?", target_exts, arc_root)
            dest_folder = os.path.join(sort_folder, folder_name)

            # Collision guard
            if os.path.exists(dest_folder):
                counter = 1
                while os.path.exists(f"{dest_folder}_{counter}-Dupe"):
                    counter += 1
                dest_folder = f"{dest_folder}_{counter}-Dupe"

            _safe_makedirs(dest_folder)

            try:
                # Move ARC into dest folder first, then extract in place
                shutil.move(arc_path, os.path.join(dest_folder, file))
                moved_arc_path = os.path.join(dest_folder, file)

                # Extract using msdos.exe + pkunpak from within the dest folder
                # Read char by char so we catch the y/n prompt which has no newline
                killed = False
                pkunpak_output = ''
                _live_line     = ''
                proc = subprocess.Popen(
                    [MSDOS_EXE, PKUNPAK_EXE, file],
                    cwd=dest_folder,
                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    text=True, errors='replace'
                )
                try:
                    while True:
                        ch = proc.stdout.read(1)
                        if not ch:
                            break
                        pkunpak_output += ch
                        _live_line     += ch
                        if ch == '\n':
                            if ui and _live_line.strip():
                                ui.push_live(_live_line.rstrip(), "live")
                                ui.draw()
                            _live_line = ''
                        if pkunpak_output.endswith('overwrite (y/n)?'):
                            if ui and _live_line.strip():
                                ui.push_live(_live_line.rstrip(), "warning")
                                ui.draw()
                            _live_line = ''
                            proc.kill()
                            proc.wait()
                            killed = True
                            break
                    if not killed:
                        proc.wait()
                except Exception:
                    proc.kill()
                    proc.wait()
                    killed = True

                # Check for CRC error — covers prompt kill, bad returncode, and known error strings
                is_crc_error = (killed or
                                proc.returncode != 0 or
                                'fails CRC check' in pkunpak_output or
                                'error in archive' in pkunpak_output)

                # Save metadata using raw byte comment reader
                comment_bytes = read_arc_comment(moved_arc_path)
                meta_path = os.path.join(dest_folder, f"__{file}__metadata.nfo")
                with open(meta_path, 'wb') as m:
                    m.write(f"Source ARC: {file}\n".encode('utf-8'))
                    m.write(b"Archive Comment: ")
                    m.write(comment_bytes if comment_bytes else b"None")
                    m.write(b"\n")

                queue_arc_deletion(root)

                if is_crc_error:
                    # Move to CRC-Errors sorted by extension
                    crc_ext_folder = os.path.basename(get_sort_folder(members, "*.?@?", target_exts, arc_err_root))
                    err_dest = os.path.join(arc_err_root, crc_ext_folder, folder_name)
                    _safe_makedirs(err_dest)
                    for item in os.listdir(dest_folder):
                        shutil.move(os.path.join(dest_folder, item), os.path.join(err_dest, item))
                    _safe_rmtree(dest_folder)
                    # Clean up empty sort folder if nothing else landed there
                    if os.path.exists(sort_folder) and not os.listdir(sort_folder):
                        shutil.rmtree(sort_folder, onerror=force_remove)
                    total_arc_crc += 1
                    arc_crc_buffer.append(f"ARC CRC Error: {file}\n")
                    arc_crc_buffer.append(f"  Location: {err_dest}\n")
                    arc_crc_buffer.append("-" * 60 + "\n")
                    if ui:
                        ui.log(f"ARC CRC Error: {file}", "error")
                        ui.update_stats(arc_crc=total_arc_crc,
                                        processed=total_processed)
                        ui.record_eta(time.monotonic() - _archive_start)
                else:
                    total_arc += 1
                    total_successful += 1
                    arc_buffer.append(f"ARC OK: Extracted {file}\n")
                    arc_buffer.append(f"  Location: {dest_folder}\n")
                    arc_buffer.append("-" * 60 + "\n")
                    if extension_counts is not None:
                        for member in members:
                            _, ext = os.path.splitext(member.lower())
                            if ext and member.lower().endswith(target_exts):
                                extension_counts[ext] += 1
                                total_matches += 1
                    if ui:
                        ui.log(f"ARC OK: {file}", "success")
                        ui.update_stats(arc=total_arc, successful=total_successful,
                                        files=count_output_files(), processed=total_processed)
                        if not process_all:
                            ui.update_ext_counts(extension_counts)
                        ui.record_eta(time.monotonic() - _archive_start)

            except FileNotFoundError as e:
                _safe_rmtree(dest_folder)
                f_err.write(f"FILE UNAVAILABLE (skipped) {file}: {str(e)}\n")
                if ui:
                    ui.log(f"Unavailable (skipped): {file}", "warning")

            except Exception as e:
                _safe_rmtree(dest_folder)
                f_err.write(f"ARC ERROR processing {file}: {str(e)}\n")
                if ui:
                    ui.log(f"ARC error: {file} — {e}", "error")
                    ui.record_eta(time.monotonic() - _archive_start)
    # Clean up original directories
    for d in sorted(dirs_to_delete, key=lambda p: p.count(os.sep), reverse=True):
        if os.path.exists(d) and not os.listdir(d):
            shutil.rmtree(d, onerror=force_remove)

    return total_arc, total_arc_crc, total_matches, total_skipped, total_successful, total_processed


# ---------------------------------------------------------------------------
# Main archive processing
# ---------------------------------------------------------------------------

def process_archives(report_file, error_file, process_all=False,
                     ui: CursesUI = None):
    # The image extensions to search for INSIDE the zips
    target_exts = ('.ima', '.flp', '.dd', '.raw', '.td0', '.fdd',
                   '.vfd', '.sdi', '.cp2', '.dmg', '.pdi', '.ana',
                   '.imd', '.ddi', '.dsk', '.img', '.sqz')
    wildcard_pattern = "*.?@?"
    total_matches    = 0
    total_processed  = 0
    total_skipped    = 0
    total_successful = 0
    total_reduce     = 0
    total_crc_errors = 0
    total_zips       = 0
    extension_counts = Counter()

    start_time = datetime.now()

    if not process_all:
        if ui:
            ui.log("--ext flag set: processing only ZIPs matching target extensions.", "warning")

    # Central root folder for all sorted output
    sorted_root  = os.path.abspath('_sorted_zip')

    if ui:
        ui.set_mode(process_all, target_exts)
        ui.set_status("Pre-scanning for ZIP files…")

    zip_tasks = []
    for root, dirs, files in os.walk('.'):
        # Don't walk into _sorted or any of its subdirectories
        dirs[:] = [d for d in dirs
                   if not os.path.abspath(os.path.join(root, d)).startswith(sorted_root)]
        for file in files:
            if file.lower().endswith('.zip'):
                zip_tasks.append((root, file))

    if not zip_tasks:
        if ui:
            ui.log("No ZIP files found.", "warning")

    # Quick pre-scan for ARC files to establish the total Queued count early
    _arc_root_prescan = os.path.abspath('_sorted_arc')
    _arc_prescan_count = 0
    for _root, _dirs, _files in os.walk('.'):
        _dirs[:] = [d for d in _dirs
                    if not os.path.abspath(os.path.join(_root, d)).startswith(
                       (sorted_root, _arc_root_prescan))]
        for _f in _files:
            if _f.lower().endswith('.arc'):
                _arc_prescan_count += 1
    if ui:
        ui.update_stats(queued=len(zip_tasks) + _arc_prescan_count)

    # Buffer for per-ZIP log lines; summary will be written to file first
    log_buffer     = []
    reduce_buffer = []
    crc_buffer     = []
    dirs_to_delete = set()  # original dirs to clean up after all ZIPs are processed

    # Helper to queue a directory and all its parents (up to cwd) for deletion
    cwd = os.path.abspath('.')

    def queue_for_deletion(path):
        current = os.path.abspath(path)
        while current != cwd:
            dirs_to_delete.add(current)
            current = os.path.dirname(current)

    if ui:
        ui.set_status("Processing Archives…")
        ui.set_progress(0, len(zip_tasks))

    with open(error_file, 'w', encoding='utf-8', errors='replace') as f_err:

        for idx, (root, file) in enumerate(zip_tasks):
            if ui:
                ui.set_current_file(file)
                ui.set_progress(idx + 1, len(zip_tasks))

            zip_path = os.path.normpath(os.path.join(root, file))

            # File may have been moved by process_nested_zips on a previous iteration
            if not os.path.isfile(zip_path):
                if ui:
                    ui.log(f"Skipped (already moved): {file}", "dim")
                continue

            try:
                # 1. Peek inside with zipfile to check for target extensions
                with zipfile.ZipFile(zip_path, 'r') as z:
                    all_contents = z.namelist()

                    if process_all:
                        matches = all_contents  # treat everything as a match
                    else:
                        matches = [n for n in all_contents
                                   if n.lower().endswith(target_exts)
                                   or fnmatch.fnmatch(n.lower(), wildcard_pattern)]
                        if not matches:
                            total_skipped += 1
                            if ui:
                                ui.update_stats(skipped=total_skipped)
                            continue

                    # Capture Metadata (Comment and Potential Disk Label)
                    disk_label = "None Detected"

                total_processed += 1
                if ui:
                    ui.update_stats(processed=total_processed)
                _archive_start = time.monotonic()
                disk_label = read_disk_label(zip_path)

                # --- Processing Logic ---

                # Determine sort destination
                sort_folder = get_sort_folder(matches, wildcard_pattern, target_exts, sorted_root)
                folder_name = os.path.splitext(file)[0]
                dest_folder = os.path.join(sort_folder, folder_name)

                # Collision guard — if folder already exists, append a counter suffix
                if os.path.exists(dest_folder):
                    counter = 1
                    while os.path.exists(f"{dest_folder}_{counter}-Dupe"):
                        counter += 1
                    dest_folder = f"{dest_folder}_{counter}-Dupe"

                _safe_makedirs(dest_folder)

                # 2. Extract using 7-Zip (handles Deflate64, LZMA, etc.)
                if ui:
                    proc_7z = subprocess.Popen(
                        ['7z', 'x', zip_path, f'-o{dest_folder}', '-y'],
                        stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                        text=True, errors='replace'
                    )
                    _stderr_buf = ['']
                    t_live = threading.Thread(
                        target=_stream_to_live, args=(proc_7z.stdout, ui), daemon=True
                    )
                    t_err = threading.Thread(
                        target=_collect_stderr, args=(proc_7z.stderr, _stderr_buf), daemon=True
                    )
                    t_live.start()
                    t_err.start()
                    while proc_7z.poll() is None:
                        ui.draw()
                        time.sleep(0.05)
                    ui.draw()
                    t_live.join()
                    t_err.join()
                    result = types.SimpleNamespace(
                        returncode=proc_7z.returncode, stderr=_stderr_buf[0]
                    )
                else:
                    result = subprocess.run(
                        ['7z', 'x', zip_path, f'-o{dest_folder}', '-y'],
                        stdout=subprocess.DEVNULL, stderr=subprocess.PIPE, text=True
                    )

                if result.returncode == 0:
                    # 3. Save Metadata to a file inside the new folder
                    meta_path = os.path.join(dest_folder, f"__{file}__metadata.nfo")
                    comment_bytes = read_zip_comment_raw(zip_path)
                    with open(meta_path, 'wb') as m:
                        m.write(f"Source ZIP: {file}\n".encode('utf-8'))
                        m.write(f"Disk Label: {disk_label}\n".encode('utf-8'))
                        m.write(b"Archive Comment: ")
                        if comment_bytes:
                            m.write(comment_bytes)
                        else:
                            m.write(b"None")
                        m.write(b"\n")

                    # 4. Move original ZIP into the destination folder
                    shutil.move(zip_path, os.path.join(dest_folder, file))

                    # Queue original directory and all parents for deletion
                    queue_for_deletion(root)

                    # 5. Buffer parent log entry and update stats
                    total_successful += 1
                    total_zips       += 1
                    log_buffer.append(f"ZIP OK: Extracted and Moved {file}\n")
                    log_buffer.append(f"  Location: {dest_folder} --->\n")
                    for m_item in matches:
                        log_buffer.append(f"  [OK] {m_item}\n")
                        total_matches += 1
                        _, ext = os.path.splitext(m_item.lower())
                        extension_counts[ext if ext else "wildcard"] += 1

                    if ui:
                        ui.log(f"ZIP OK: {file}", "success")
                        ui.update_stats(
                    # Process any nested ZIPs — entries appended to buffers under parent
                            processed=total_processed, successful=total_successful,
                            files=count_output_files(), zips=total_zips,
                        )
                        if not process_all:
                            ui.update_ext_counts(extension_counts)
                        ui.record_eta(time.monotonic() - _archive_start)

                    parent_processed = {os.path.abspath(os.path.join(dest_folder, file))}
                    process_nested_zips(dest_folder, error_file, processed=parent_processed,
                                        log_buffer=log_buffer, reduce_buffer=reduce_buffer)
                    log_buffer.append("-" * 60 + "\n")

                elif "ERROR: Unsupported Method" in result.stderr:
                    # Reduce — delete the messy 7-Zip dest folder
                    _safe_rmtree(dest_folder)
                    # Clean up empty parent extension folder if nothing else is in it
                    if os.path.exists(sort_folder) and not os.listdir(sort_folder):
                        shutil.rmtree(sort_folder, onerror=force_remove)

                    # Extract into normal _sorted extension folder alongside everything else
                    r_dest = os.path.join(sort_folder, folder_name)
                    _safe_makedirs(r_dest)
                    r_zip_path = os.path.join(r_dest, file)
                    shutil.move(zip_path, r_zip_path)

                    # Queue original directory and all parents for deletion
                    queue_for_deletion(root)

                    # Extract using msdos.exe + pkunzip from within the dest folder
                    if ui:
                        proc_pkz = subprocess.Popen(
                            [MSDOS_EXE, '-d', PKUNZIP_EXE, '-e', '-o', '-d', file],
                            cwd=r_dest,
                            stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                            text=True, errors='replace'
                        )
                        _pkz_err_buf = ['']
                        t_pkz = threading.Thread(
                            target=_stream_to_live, args=(proc_pkz.stdout, ui, "warning"), daemon=True
                        )
                        t_pkz_err = threading.Thread(
                            target=_collect_stderr, args=(proc_pkz.stderr, _pkz_err_buf), daemon=True
                        )
                        t_pkz.start()
                        t_pkz_err.start()
                        while proc_pkz.poll() is None:
                            ui.draw()
                            time.sleep(0.05)
                        ui.draw()
                        t_pkz.join()
                        t_pkz_err.join()
                        r_result = types.SimpleNamespace(returncode=proc_pkz.returncode)
                    else:
                        r_result = subprocess.run(
                            [MSDOS_EXE, '-d', PKUNZIP_EXE, '-e', '-o', '-d', file],
                            cwd=r_dest,
                            stdout=subprocess.DEVNULL, stderr=subprocess.PIPE, text=True
                        )

                    total_reduce += 1
                    reduce_buffer.append(f"Legacy ZIP OK: {file}\n")
                    reduce_buffer.append(f"  Location: {r_dest}\n")

                    # Save metadata using raw byte comment reader
                    r_comment = read_zip_comment_raw(r_zip_path)
                    meta_path = os.path.join(r_dest, f"__{file}__metadata.nfo")
                    with open(meta_path, 'wb') as m:
                        m.write(f"Source ZIP: {file}\n".encode('utf-8'))
                        m.write(f"Disk Label: {disk_label}\n".encode('utf-8'))
                        m.write(b"Archive Comment: ")
                        if r_comment:
                            m.write(r_comment)
                        else:
                            m.write(b"None")
                        m.write(b"\n")

                    if ui:
                        ui.log(f"Legacy ZIP OK: {file}", "warning")
                        ui.update_stats(
                            processed=total_processed, reduce=total_reduce,
                        )
                        ui.record_eta(time.monotonic() - _archive_start)

                    if r_result.returncode != 0:
                        # pkunzip failed — move ZIP and metadata to CRC-Errors/__Reduce
                        r_crc_dest = os.path.join(sorted_root, 'CRC-Errors', '__Reduce', folder_name)
                        _safe_makedirs(r_crc_dest)
                        shutil.move(r_zip_path, os.path.join(r_crc_dest, file))
                        shutil.move(meta_path, os.path.join(r_crc_dest, os.path.basename(meta_path)))
                        # Clean up the now-empty dest subfolder
                        _safe_rmtree(r_dest)
                        if os.path.exists(sort_folder) and not os.listdir(sort_folder):
                            shutil.rmtree(sort_folder, onerror=force_remove)
                        reduce_buffer.append(f"  WARNING: pkunzip failed (code {r_result.returncode}) — moved to CRC-Errors/__Reduce\n")
                        total_crc_errors += 1
                        crc_buffer.append(f"CRC Error (Legacy ZIP): {file}\n")
                        crc_buffer.append(f"  Location: {r_crc_dest}\n")
                        crc_buffer.append("-" * 60 + "\n")
                        if ui:
                            ui.log(f"CRC Error (Legacy ZIP): {file}", "error")
                            ui.update_stats(crc_errors=total_crc_errors)
                    else:
                        # Process nested ZIPs — entries appended to buffers under parent
                        r_parent_processed = {os.path.abspath(r_zip_path)}
                        process_nested_zips(r_dest, error_file, processed=r_parent_processed,
                                            log_buffer=log_buffer, reduce_buffer=reduce_buffer)
                    reduce_buffer.append("-" * 60 + "\n")

                elif "ERROR: Data Error" in result.stderr:
                    # CRC Error — clean up dest folder, move ZIP to CRC-Errors/<ext>/
                    _safe_rmtree(dest_folder)
                    # Clean up empty parent extension folder if nothing else is in it
                    if os.path.exists(sort_folder) and not os.listdir(sort_folder):
                        shutil.rmtree(sort_folder, onerror=force_remove)

                    # Queue original directory and all parents for deletion
                    queue_for_deletion(root)

                    crc_ext_folder = os.path.basename(get_sort_folder(matches, wildcard_pattern, target_exts, sorted_root))
                    crc_dest = os.path.join(sorted_root, 'CRC-Errors', crc_ext_folder, folder_name)
                    _safe_makedirs(crc_dest)
                    shutil.move(zip_path, os.path.join(crc_dest, file))

                    # Save metadata
                    crc_comment = read_zip_comment_raw(os.path.join(crc_dest, file))
                    meta_path = os.path.join(crc_dest, f"__{file}__metadata.nfo")
                    with open(meta_path, 'wb') as m:
                        m.write(f"Source ZIP: {file}\n".encode('utf-8'))
                        m.write(f"Disk Label: {disk_label}\n".encode('utf-8'))
                        m.write(b"Archive Comment: ")
                        m.write(crc_comment if crc_comment else b"None")
                        m.write(b"\n")

                    total_crc_errors += 1
                    crc_buffer.append(f"CRC Error: {file}\n")
                    crc_buffer.append(f"  Location: {crc_dest}\n")
                    crc_buffer.append("-" * 60 + "\n")
                    if ui:
                        ui.log(f"CRC Error (Data): {file}", "error")
                        ui.update_stats(crc_errors=total_crc_errors,
                                        processed=total_processed)
                        ui.record_eta(time.monotonic() - _archive_start)

                elif "ERROR: CRC Failed" in result.stderr:
                    # CRC Error — clean up dest folder, move ZIP to CRC-Errors/<ext>/
                    _safe_rmtree(dest_folder)
                    # Clean up empty parent extension folder if nothing else is in it
                    if os.path.exists(sort_folder) and not os.listdir(sort_folder):
                        shutil.rmtree(sort_folder, onerror=force_remove)

                    # Queue original directory and all parents for deletion
                    queue_for_deletion(root)

                    crc_ext_folder = os.path.basename(get_sort_folder(matches, wildcard_pattern, target_exts, sorted_root))
                    crc_dest = os.path.join(sorted_root, 'CRC-Errors', crc_ext_folder, folder_name)
                    _safe_makedirs(crc_dest)
                    shutil.move(zip_path, os.path.join(crc_dest, file))

                    # Save metadata
                    crc_comment = read_zip_comment_raw(os.path.join(crc_dest, file))
                    meta_path = os.path.join(crc_dest, f"__{file}__metadata.nfo")
                    with open(meta_path, 'wb') as m:
                        m.write(f"Source ZIP: {file}\n".encode('utf-8'))
                        m.write(f"Disk Label: {disk_label}\n".encode('utf-8'))
                        m.write(b"Archive Comment: ")
                        m.write(crc_comment if crc_comment else b"None")
                        m.write(b"\n")

                    total_crc_errors += 1
                    crc_buffer.append(f"CRC Error: {file}\n")
                    crc_buffer.append(f"  Location: {crc_dest}\n")
                    crc_buffer.append("-" * 60 + "\n")
                    if ui:
                        ui.log(f"CRC Error (Failed): {file}", "error")
                        ui.update_stats(crc_errors=total_crc_errors,
                                        processed=total_processed)
                        ui.record_eta(time.monotonic() - _archive_start)

                else:
                    # Unknown 7-Zip error — clean up empty dest folder and parent if empty
                    _safe_rmtree(dest_folder)
                    # Clean up empty parent extension folder if nothing else is in it
                    if os.path.exists(sort_folder) and not os.listdir(sort_folder):
                        shutil.rmtree(sort_folder, onerror=force_remove)

                    # Queue original directory and all parents for deletion
                    queue_for_deletion(root)
                    f_err.write(f"7-ZIP ERROR for {file}: {result.stderr}\n")
                    if ui:
                        ui.log(f"7-Zip error: {file}", "error")
                        ui.record_eta(time.monotonic() - _archive_start)

            except FileNotFoundError as e:
                # File disappeared mid-run — most likely quarantined by AV software
                # or moved externally. Log clearly and continue without counting as error.
                f_err.write(f"FILE UNAVAILABLE (skipped) {file}: {str(e)}\n")
                if ui:
                    ui.log(f"Unavailable (skipped): {file}", "warning")

            except Exception as e:
                f_err.write(f"PYTHON ERROR processing {file}: {str(e)}\n")
                if ui:
                    ui.log(f"Python error: {file} — {e}", "error")
    # Process ARC files before deleting source directories
    # ── ARC pass ────────────────────────────────────────────────────────────
    arc_buffer     = []
    arc_crc_buffer = []
    total_arc      = 0
    total_arc_crc  = 0
    total_arc, total_arc_crc, arc_matches, total_skipped, total_successful, total_processed = process_arcs(
        report_file, error_file, arc_buffer, arc_crc_buffer, total_arc, total_arc_crc,
        process_all=process_all, target_exts=target_exts,
        extension_counts=extension_counts, total_matches=total_matches,
        total_skipped=total_skipped, total_successful=total_successful,
        total_processed=total_processed, ui=ui,
    )
    total_matches += arc_matches

    if ui:
        ui.set_status("Cleaning up empty source directories…")
        ui.set_current_file("")
    # Delete original directories now that all ZIPs and ARCs have been processed
    # Sort by depth deepest first so children are removed before parents
    # Only delete if the directory is empty — non-archive files are left untouched
    for d in sorted(dirs_to_delete, key=lambda p: p.count(os.sep), reverse=True):
        if os.path.exists(d) and not os.listdir(d):
            shutil.rmtree(d, onerror=force_remove)

    # Calculate duration
    end_time = datetime.now()
    duration = end_time - start_time

    if ui:
        ui.set_status("Writing report…")
    # Write report: summary first, then buffered per-ZIP log
    with open(report_file, 'w', encoding='utf-8', errors='replace') as f_out:

        f_out.write("EXTRACTION SUMMARY REPORT".center(60) + "\n")
        f_out.write("=" * 60 + "\n")
        f_out.write(f"Date/Time          : {start_time.strftime('%Y-%m-%d %H:%M:%S')}\n")
        f_out.write(f"Duration           : {duration}\n")
        f_out.write(f"Total Processed    : {total_processed}\n")
        if not process_all:
            f_out.write(f"Skipped            : {total_skipped}\n")
        f_out.write(f"Successful         : {total_successful}\n")
        f_out.write(f"  ZIP              : {total_zips}\n")
        f_out.write(f"  Legacy ZIP       : {total_reduce}\n")
        f_out.write(f"  Arc              : {total_arc}\n")
        f_out.write(f"CRC Errors         : {total_crc_errors}\n")
        f_out.write(f"Arc CRC Errors     : {total_arc_crc}\n")
        f_out.write(f"Target Files Found : {'-' * 3 if process_all else total_matches}\n")
        if not process_all:
            f_out.write("\n  By Extension:\n")
            for ext, count in sorted(extension_counts.items()):
                f_out.write(f"{ext}     : {count}\n")
        f_out.write("-" * 60 + "\n\n")

        # Arc section
        if arc_buffer:
            f_out.write("=" * 60 + "\n")
            f_out.write("ARC".center(60) + "\n")
            f_out.write("=" * 60 + "\n")
            f_out.writelines(arc_buffer)
            f_out.write("\n")

        # Arc CRC Errors section
        if arc_crc_buffer:
            f_out.write("=" * 60 + "\n")
            f_out.write("ARC CRC ERRORS".center(60) + "\n")
            f_out.write("=" * 60 + "\n")
            f_out.writelines(arc_crc_buffer)
            f_out.write("\n")

        # Legacy ZIP section
        if reduce_buffer:
            f_out.write("=" * 60 + "\n")
            f_out.write("LEGACY ZIP".center(60) + "\n")
            f_out.write("=" * 60 + "\n")
            f_out.writelines(reduce_buffer)
            f_out.write("\n")

        # CRC Errors section
        if crc_buffer:
            f_out.write("=" * 60 + "\n")
            f_out.write("FAILED".center(60) + "\n")
            f_out.write("=" * 60 + "\n")
            f_out.writelines(crc_buffer)
            f_out.write("\n")

        # Per-archive detail log
        f_out.write("=" * 60 + "\n")
        f_out.write("SUCCESS".center(60) + "\n")
        f_out.write("=" * 60 + "\n")
        f_out.writelines(log_buffer)

    if ui:
        ui.update_stats(
            processed=total_processed, skipped=total_skipped,
            successful=total_successful, reduce=total_reduce,
            crc_errors=total_crc_errors, files=count_output_files(),
            arc=total_arc, arc_crc=total_arc_crc, zips=total_zips,
        )
        ui.set_status(f"Done! Duration: {duration}")
        ui.set_current_file("")
#        ui.log(f"Report saved → {report_file}", "success")
        ui.log(f"Report saved → {report_file}", "normal")

# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Process ZIP archives.")
    parser.add_argument('--ext', action='store_true',
                        help="Only process ZIPs containing files matching target_exts or the wildcard pattern")
    parser.add_argument('--log', metavar='FILE', default=None,
                        help="Mirror live extraction output to a text file in real time")
    args = parser.parse_args()

    # Check required executables before entering curses (so errors print normally)
    missing = [exe for exe in [MSDOS_EXE, PKUNZIP_EXE, PKUNPAK_EXE] if not os.path.isfile(exe)]
    if missing:
        for m in missing:
            print(f"ERROR: Required executable not found: {m}")
        print("Please ensure msdos.exe, PKUNZIP.EXE and PKUNPAK.EXE are in the same folder as this script.")
        exit(1)

    folder_name = os.path.basename(os.path.dirname(os.path.abspath(__file__)))
    date_str    = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    rep         = f"{date_str}_{folder_name}_Report.txt"
    err         = 'error_log.txt'

    def _curses_main(stdscr):
        ui = CursesUI(stdscr)
        if args.log:
            live_log_handle = open(args.log, 'w', encoding='utf-8', errors='replace')
            live_log_handle.write(f"Live Log — {date_str}\n")
            live_log_handle.write("=" * 60 + "\n")
            live_log_handle.flush()
            ui.set_live_log(live_log_handle)
        else:
            live_log_handle = None
        try:
            process_archives(rep, err, process_all=not args.ext, ui=ui)
            ui.wait_for_key("  Done — press any key to exit and open the report…")
        finally:
            if live_log_handle:
                live_log_handle.write("=" * 60 + "\n")
                live_log_handle.write("Log complete.\n")
                live_log_handle.close()

    curses.wrapper(_curses_main)

    # Restore terminal is complete at this point; open the report normally
    print(f"\nScan complete. Opening {rep} now…")
    # AUTO-OPEN the report file
    os.startfile(rep)
#    play_success_sound()