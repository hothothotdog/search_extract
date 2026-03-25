import os
import zipfile
import fnmatch
import shutil
import subprocess
import argparse
# import winsound  # Standard on Windows
from collections import Counter
from datetime import datetime
from tqdm import tqdm  # Ensure you run 'pip install tqdm' first

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

        os.makedirs(nested_dest, exist_ok=True)

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
                log_buffer.append(f"{indent}[Nested] SUCCESS: {item}\n")
                log_buffer.append(f"{indent}  Location: {nested_dest}\n")

        elif "ERROR: Unsupported Method" in sevenzip_output:
            # Reduce nested ZIP — use pkunzip
            shutil.rmtree(nested_dest, onerror=force_remove)
            os.makedirs(nested_dest, exist_ok=True)
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
                reduce_buffer.append(f"{indent}[Nested] Reduce: {item}\n")
                reduce_buffer.append(f"{indent}  Location: {nested_dest}\n")

        else:
            # Failed — clean up empty folder and log it
            shutil.rmtree(nested_dest, onerror=force_remove)
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

def process_arcs(report_file, error_file, arc_buffer, arc_crc_buffer,
                 total_arc, total_arc_crc, process_all=True, target_exts=(),
                 extension_counts=None, total_matches=0,
                 total_skipped=0, total_successful=0, total_processed=0):
    """
    Scan recursively for .arc files, extract with PKXARC via msdos.exe,
    write metadata, and sort into _sorted_arc/Arc/<name>/.
    Returns updated (total_arc, total_arc_crc) counts.
    """
    arc_root     = os.path.abspath('_sorted_arc')
    arc_folder   = os.path.join(arc_root, 'Arc')
    arc_err_root = os.path.join(arc_root, 'CRC-Errors')

    print("Pre-scanning for ARC files...")
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
        print("No ARC files found.")
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
            print("No ARC files found matching target extensions.")
            return total_arc, total_arc_crc, total_matches, total_skipped, total_successful, total_processed

    dirs_to_delete = set()
    cwd = os.path.abspath('.')

    def queue_arc_deletion(path):
        current = os.path.abspath(path)
        while current != cwd:
            dirs_to_delete.add(current)
            current = os.path.dirname(current)

    with open(error_file, 'a', encoding='utf-8', errors='replace') as f_err:
        for root, file in tqdm(arc_tasks, desc="Processing ARCs", unit="arc"):
            arc_path = os.path.normpath(os.path.join(root, file))
            folder_name = os.path.splitext(file)[0]
            total_processed += 1

            # Determine sort folder from ARC member extensions
            members = get_arc_members(arc_path)
            sort_folder = get_sort_folder(members, "*.?@?", target_exts, arc_root)
            dest_folder = os.path.join(sort_folder, folder_name)

            # Collision guard
            if os.path.exists(dest_folder):
                counter = 1
                while os.path.exists(f"{dest_folder}_{counter}-Dupe"):
                    counter += 1
                dest_folder = f"{dest_folder}_{counter}-Dupe"

            os.makedirs(dest_folder, exist_ok=True)

            try:
                # Move ARC into dest folder first, then extract in place
                shutil.move(arc_path, os.path.join(dest_folder, file))
                moved_arc_path = os.path.join(dest_folder, file)

                # Extract using msdos.exe + PKXARC from within the dest folder
                result = subprocess.run(
                    [MSDOS_EXE, PKUNPAK_EXE, file],
                    cwd=dest_folder,
                    stdout=subprocess.DEVNULL, stderr=subprocess.PIPE, text=True
                )

                # Check for CRC error in output
                is_crc_error = result.returncode != 0

                # Write metadata
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
                    os.makedirs(err_dest, exist_ok=True)
                    for item in os.listdir(dest_folder):
                        shutil.move(os.path.join(dest_folder, item), os.path.join(err_dest, item))
                    shutil.rmtree(dest_folder, onerror=force_remove)
                    # Clean up empty sort folder if nothing else landed there
                    if os.path.exists(sort_folder) and not os.listdir(sort_folder):
                        shutil.rmtree(sort_folder, onerror=force_remove)
                    total_arc_crc += 1
                    arc_crc_buffer.append(f"ARC CRC Error: {file}\n")
                    arc_crc_buffer.append(f"  Location: {err_dest}\n")
                    arc_crc_buffer.append("-" * 60 + "\n")
                else:
                    total_arc += 1
                    total_successful += 1
                    arc_buffer.append(f"SUCCESS: Extracted {file}\n")
                    arc_buffer.append(f"  Location: {dest_folder}\n")
                    arc_buffer.append("-" * 60 + "\n")
                    if extension_counts is not None:
                        for member in members:
                            _, ext = os.path.splitext(member.lower())
                            if ext and member.lower().endswith(target_exts):
                                extension_counts[ext] += 1
                                total_matches += 1

            except Exception as e:
                shutil.rmtree(dest_folder, onerror=force_remove)
                f_err.write(f"ARC ERROR processing {file}: {str(e)}\n")

    # Clean up original directories
    for d in sorted(dirs_to_delete, key=lambda p: p.count(os.sep), reverse=True):
        if os.path.exists(d) and not os.listdir(d):
            shutil.rmtree(d, onerror=force_remove)

    return total_arc, total_arc_crc, total_matches, total_skipped, total_successful, total_processed


def process_archives(report_file, error_file, process_all=False):
    # The image extensions to search for INSIDE the zips
    target_exts = ('.ima', '.flp', '.dd', '.raw', '.td0', '.fdd',
                   '.vfd', '.sdi', '.cp2', '.dmg', '.pdi', '.ana',
                   '.imd', '.ddi', '.dsk', '.img', '.sqz')
    wildcard_pattern = "*.?@?"
    total_matches = 0
    total_processed = 0
    total_skipped = 0
    total_successful = 0
    total_reduce = 0
    total_crc_errors = 0
    extension_counts = Counter()

    start_time = datetime.now()

    if not process_all:
        print("--ext flag set: processing only ZIPs matching target extensions.")

    # Central root folder for all sorted output
    sorted_root  = os.path.abspath('_sorted_zip')

    print("Pre-scanning for ZIP files...")
    zip_tasks = []
    for root, dirs, files in os.walk('.'):
        # Don't walk into _sorted or any of its subdirectories
        dirs[:] = [d for d in dirs
                   if not os.path.abspath(os.path.join(root, d)).startswith(sorted_root)]
        for file in files:
            if file.lower().endswith('.zip'):
                zip_tasks.append((root, file))

    if not zip_tasks:
        print("No ZIP files found.")

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

    with open(error_file, 'w', encoding='utf-8', errors='replace') as f_err:

        for root, file in tqdm(zip_tasks, desc="Processing Archives", unit="zip"):
            zip_path = os.path.normpath(os.path.join(root, file))

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
                            continue

                    # Capture Metadata (Comment and Potential Disk Label)
                    disk_label = "None Detected"

                total_processed += 1

                # Try to find disk label using byte pattern
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

                os.makedirs(dest_folder, exist_ok=True)

                # 2. Extract using 7-Zip (handles Deflate64, LZMA, etc.)
                result = subprocess.run(
                    ['7z', 'x', zip_path, f'-o{dest_folder}', '-y'],
                    # remove comma above '-y'], and Comment below to hide 7z from terminal
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
                    log_buffer.append(f"SUCCESS: Extracted and Moved {file}\n")
                    log_buffer.append(f"  Location: {dest_folder} --->\n")
                    for m_item in matches:
                        log_buffer.append(f"  [Match] {m_item}\n")
                        total_matches += 1
                        _, ext = os.path.splitext(m_item.lower())
                        extension_counts[ext if ext else "wildcard"] += 1

                    # Process any nested ZIPs — entries appended to buffers under parent
                    parent_processed = {os.path.abspath(os.path.join(dest_folder, file))}
                    process_nested_zips(dest_folder, error_file, processed=parent_processed,
                                        log_buffer=log_buffer, reduce_buffer=reduce_buffer)
                    log_buffer.append("-" * 60 + "\n")

                elif "ERROR: Unsupported Method" in result.stderr:
                    # Reduce — delete the messy 7-Zip dest folder
                    shutil.rmtree(dest_folder, onerror=force_remove)
                    # Clean up empty parent extension folder if nothing else is in it
                    if os.path.exists(sort_folder) and not os.listdir(sort_folder):
                        shutil.rmtree(sort_folder, onerror=force_remove)

                    # Extract into normal _sorted extension folder alongside everything else
                    r_dest = os.path.join(sort_folder, folder_name)
                    os.makedirs(r_dest, exist_ok=True)
                    r_zip_path = os.path.join(r_dest, file)
                    shutil.move(zip_path, r_zip_path)

                    # Queue original directory and all parents for deletion
                    queue_for_deletion(root)

                    # Extract using msdos.exe + pkunzip from within the dest folder
                    r_result = subprocess.run(
                        [MSDOS_EXE, '-d', PKUNZIP_EXE, '-e', '-o', '-d', file],
                        cwd=r_dest,
                    # remove comma above dest, and Comment below to hide pkunzip from terminal
                    stdout=subprocess.DEVNULL, stderr=subprocess.PIPE, text=True
                    )

                    total_reduce += 1
                    reduce_buffer.append(f"Reduce: {file}\n")
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

                    if r_result.returncode != 0:
                        # pkunzip failed — move ZIP and metadata to CRC-Errors/__Reduce
                        r_crc_dest = os.path.join(sorted_root, 'CRC-Errors', '__Reduce', folder_name)
                        os.makedirs(r_crc_dest, exist_ok=True)
                        shutil.move(r_zip_path, os.path.join(r_crc_dest, file))
                        shutil.move(meta_path, os.path.join(r_crc_dest, os.path.basename(meta_path)))
                        # Clean up the now-empty dest subfolder
                        shutil.rmtree(r_dest, onerror=force_remove)
                        if os.path.exists(sort_folder) and not os.listdir(sort_folder):
                            shutil.rmtree(sort_folder, onerror=force_remove)
                        reduce_buffer.append(f"  WARNING: pkunzip failed (code {r_result.returncode}) — moved to CRC-Errors/__Reduce\n")
                        total_crc_errors += 1
                        crc_buffer.append(f"CRC Error (Reduce): {file}\n")
                        crc_buffer.append(f"  Location: {r_crc_dest}\n")
                        crc_buffer.append("-" * 60 + "\n")
                    else:
                        # Process nested ZIPs — entries appended to buffers under parent
                        r_parent_processed = {os.path.abspath(r_zip_path)}
                        process_nested_zips(r_dest, error_file, processed=r_parent_processed,
                                            log_buffer=log_buffer, reduce_buffer=reduce_buffer)
                    reduce_buffer.append("-" * 60 + "\n")

                elif "ERROR: Data Error" in result.stderr:
                    # CRC Error — clean up dest folder, move ZIP to CRC-Errors/<ext>/
                    shutil.rmtree(dest_folder, onerror=force_remove)
                    # Clean up empty parent extension folder if nothing else is in it
                    if os.path.exists(sort_folder) and not os.listdir(sort_folder):
                        shutil.rmtree(sort_folder, onerror=force_remove)

                    # Queue original directory and all parents for deletion
                    queue_for_deletion(root)

                    crc_ext_folder = os.path.basename(get_sort_folder(matches, wildcard_pattern, target_exts, sorted_root))
                    crc_dest = os.path.join(sorted_root, 'CRC-Errors', crc_ext_folder, folder_name)
                    os.makedirs(crc_dest, exist_ok=True)
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

                elif "ERROR: CRC Failed" in result.stderr:
                    # CRC Error — clean up dest folder, move ZIP to CRC-Errors/<ext>/
                    shutil.rmtree(dest_folder, onerror=force_remove)
                    # Clean up empty parent extension folder if nothing else is in it
                    if os.path.exists(sort_folder) and not os.listdir(sort_folder):
                        shutil.rmtree(sort_folder, onerror=force_remove)

                    # Queue original directory and all parents for deletion
                    queue_for_deletion(root)

                    crc_ext_folder = os.path.basename(get_sort_folder(matches, wildcard_pattern, target_exts, sorted_root))
                    crc_dest = os.path.join(sorted_root, 'CRC-Errors', crc_ext_folder, folder_name)
                    os.makedirs(crc_dest, exist_ok=True)
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

                else:
                    # Unknown 7-Zip error — clean up empty dest folder and parent if empty
                    shutil.rmtree(dest_folder, onerror=force_remove)
                    # Clean up empty parent extension folder if nothing else is in it
                    if os.path.exists(sort_folder) and not os.listdir(sort_folder):
                        shutil.rmtree(sort_folder, onerror=force_remove)

                    # Queue original directory and all parents for deletion
                    queue_for_deletion(root)

                    f_err.write(f"7-ZIP ERROR for {file}: {result.stderr}\n")

            except Exception as e:
                f_err.write(f"PYTHON ERROR processing {file}: {str(e)}\n")

    # Process ARC files before deleting source directories
    arc_buffer     = []
    arc_crc_buffer = []
    total_arc      = 0
    total_arc_crc  = 0
    total_arc, total_arc_crc, arc_matches, total_skipped, total_successful, total_processed = process_arcs(
        report_file, error_file, arc_buffer, arc_crc_buffer, total_arc, total_arc_crc,
        process_all=process_all, target_exts=target_exts,
        extension_counts=extension_counts, total_matches=total_matches,
        total_skipped=total_skipped, total_successful=total_successful,
        total_processed=total_processed
    )
    total_matches += arc_matches

    # Delete original directories now that all ZIPs and ARCs have been processed
    # Sort by depth deepest first so children are removed before parents
    # Only delete if the directory is empty — non-archive files are left untouched
    for d in sorted(dirs_to_delete, key=lambda p: p.count(os.sep), reverse=True):
        if os.path.exists(d) and not os.listdir(d):
            shutil.rmtree(d, onerror=force_remove)

    # Calculate duration
    end_time = datetime.now()
    duration = end_time - start_time

    # Write report: summary first, then buffered per-ZIP log
    with open(report_file, 'w', encoding='utf-8', errors='replace') as f_out:

        f_out.write("          EXTRACTION SUMMARY REPORT\n")
        f_out.write("=" * 60 + "\n")
        f_out.write(f"Date/Time          : {start_time.strftime('%Y-%m-%d %H:%M:%S')}\n")
        f_out.write(f"Duration           : {duration}\n")
        f_out.write(f"Total Processed    : {total_processed}\n")
        f_out.write(f"Skipped            : {total_skipped}\n")
        f_out.write(f"Successful         : {total_successful}\n")
        f_out.write(f"Reduce             : {total_reduce}\n")
        f_out.write(f"CRC Errors         : {total_crc_errors}\n")
        f_out.write(f"Arc                : {total_arc}\n")
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
            f_out.write("               ARC\n")
            f_out.write("=" * 60 + "\n")
            f_out.writelines(arc_buffer)
            f_out.write("\n")

        # Arc CRC Errors section
        if arc_crc_buffer:
            f_out.write("=" * 60 + "\n")
            f_out.write("           ARC CRC ERRORS\n")
            f_out.write("=" * 60 + "\n")
            f_out.writelines(arc_crc_buffer)
            f_out.write("\n")

        # Reduce section
        if reduce_buffer:
            f_out.write("=" * 60 + "\n")
            f_out.write("               REDUCE\n")
            f_out.write("=" * 60 + "\n")
            f_out.writelines(reduce_buffer)
            f_out.write("\n")

        # CRC Errors section
        if crc_buffer:
            f_out.write("=" * 60 + "\n")
            f_out.write("             CRC ERRORS\n")
            f_out.write("=" * 60 + "\n")
            f_out.writelines(crc_buffer)
            f_out.write("\n")

        # Per-ZIP detail log
        f_out.write("=" * 60 + "\n")
        f_out.write("             SUCCESS\n")
        f_out.write("=" * 60 + "\n")

        f_out.writelines(log_buffer)

    # AUTO-OPEN the report file
    os.startfile(report_file)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Process ZIP archives.")
    parser.add_argument('--ext', action='store_true',
                        help="Only process ZIPs containing files matching target_exts or the wildcard pattern")
    args = parser.parse_args()

    # Check required executables are present before starting
    missing = [exe for exe in [MSDOS_EXE, PKUNZIP_EXE, PKUNPAK_EXE] if not os.path.isfile(exe)]
    if missing:
        for m in missing:
            print(f"ERROR: Required executable not found: {m}")
        print("Please ensure msdos.exe, PKUNZIP.EXE and PKUNPAK.EXE are in the same folder as this script.")
        exit(1)

    folder_name = os.path.basename(os.path.dirname(os.path.abspath(__file__)))
    date_str = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    rep = f"{date_str}_{folder_name}_Report.txt"
    err = 'error_log.txt'
    process_archives(rep, err, process_all=not args.ext)
    print(f"\nScan complete. Opening {rep} now...")
#    play_success_sound()
