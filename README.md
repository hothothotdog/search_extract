# search_extract.py

A Windows command-line tool that recursively scans a directory tree for **ZIP** and **ARC** archives, extracts them using the best available extractor, sorts the output by content type, and produces a full extraction report — all inside a live curses terminal dashboard.

It is designed to handle archives from the DOS/early PC era, including formats and compression methods that modern tools struggle with, such as PKzip Reduce compression which was a precursor to the more modern zip and the original ARC archives created in the 80's.
This script will work equally well on a single archive.

## Some context
After obtaining a dump of 1980-1990 archives I originally looked for a way to mass extract and document them but, found none truly suitable.

Inside all these archives there were multiple disk image types, nested archives within archives, ancient .arc and .zip files that were a pain to work on in en-mass with DOSBox. They all needed sorting too and many had CRC errors that i wanted to deal with quickly, log and put aside separately.

---

## Requirements

| Python 3.8+ | Runtime |
### Python packages
```
pip install windows-curses
```

### External executables
| [7-Zip](https://www.7-zip.org/) (`7z` on PATH) | Primary ZIP extractor (Deflate, Deflate64, LZMA, etc.) |

| File          | Purpose                                                       |
| ------------- | ------------------------------------------------------------- |
| `msdos.exe`   | DOS emulator for running legacy tools                         |
| `PKUNZIP.EXE` | Legacy extractor for **Reduce**-compressed ZIPs (methods 2–5) |
| `PKUNPAK.EXE` | Legacy extractor for **ARC** archives                         |

`msdos.exe`, `PKUNZIP.EXE`, and `PKUNPAK.EXE` are checked at startup. The script exits with a clear error message if any are missing before curses takes over the terminal.

## Obtaining Prerequisites

Each executable requires a few extraction steps from its source archive. All extraction steps below use 7-Zip unless otherwise noted.

Alternatively, you can skip the extraction steps and download the `prerequisites.7z` from the releases page.
---

### PKARC / PKPAK v3.61 (1988)

**Provides:** `PKUNPAK.EXE`

|                   |                                                               |
| ----------------- | ------------------------------------------------------------- |
| **Download file** | `pkarc-v3-61_35dd_ima_en.zip`                                 |
| **Topic**         | http://www.win3x.org/win3board/viewtopic.php?t=28463&view=min |
| **Download**      | http://www.win3x.org/win3board/ext/win3x/download.php?id=4316 |

**Extraction steps:**

1. Extract `pkarc-v3-61_35dd_ima_en.zip` with 7-Zip
2. Extract `PK361.EXE` from `disk.ima` with 7-Zip
3. Extract `PK361.EXE` via terminal: `msdos.exe PK361.EXE`
4. Copy `PKUNPAK.EXE` to the script folder

---

### PKZIP v2.04g (1993)

**Provides:** `PKUNZIP.EXE`

|                   |                                                                   |
| ----------------- | ----------------------------------------------------------------- |
| **Download file** | `pkzip_204g_dd35_dos_en.zip`                                      |
| **Topic**         | http://www.win3x.org/win3board/viewtopic.php?t=4336&hilit=pkunzip |
| **Download**      | http://www.win3x.org/win3board/ext/win3x/download.php?id=4168     |

**Extraction steps:**

1. Extract `pkzip_204g_dd35_dos_en.zip` with 7-Zip
2. Extract `PKUNZIP.EXE` from `disk.ima` with 7-Zip
3. Copy `PKUNZIP.EXE` to the script folder

---

### MS-DOS Player (i8086) for Win32-x64

**Provides:** `msdos.exe`

There are many versions to choose from inside the download. I went with the simplest x86 for speed.

|                   |                                                      |
| ----------------- | ---------------------------------------------------- |
| **Download file** | `msdos.7z`                                           |
| **Topic**         | https://takeda-toshiya.my.coocan.jp/msdos/index.html |
| **Download**      | https://takeda-toshiya.my.coocan.jp/msdos/msdos.7z   |

**Extraction steps:**

1. Extract `msdos.7z` with 7-Zip
2. Navigate to `\msdos\binary\i86_x86\`
3. Copy `msdos.exe` to the script folder

---
### Useful extras

### NFOPad v1.81 (2022)

To view comments as intended from archives back in the day this does an excellent job.
It's a free, small and fast text editor that supports ASCII art with the extension `.nfo`
|                   |                                                      |
| ----------------- | ---------------------------------------------------- |
| **Download file** | `nfopad181.exe`                                      |
| **Main page**     | https://truehumandesign.se/s_nfopad.php              |
| **Download**      | https://truehumandesign.se/dl.php?file=nfopad181.exe |

---

## Usage

```
# Process every ZIP and ARC found under the current directory
python search_extract.py

# log file captures both panels in real time with
python search_extract.py --log <filename>


# Only process archives that contain disk image files (see Target Extensions below)
python search_extract.py --ext

# Search for target extensions and log live output to file
python search_extract.py --ext --log live.txt

```

### `--ext` flag

When passed, only archives whose contents match the **target extension list** or the wildcard pattern `*.?@?` are extracted. Archives with no matching content are counted as **Skipped** and left untouched. Without the flag every archive is processed unconditionally and the Skipped counter is hidden.

---

## Target Extensions

The following file extensions are recognised as disk image files when `--ext` is active:

`.ima` `.flp` `.dd` `.raw` `.td0` `.fdd` `.vfd` `.sdi` `.cp2` `.dmg` `.pdi` `.ana` `.imd` `.ddi` `.dsk` `.img` `.sqz`

The wildcard pattern `*.?@?` additionally catches non-standard disk image naming conventions.
```

## Edit the script to add or remove extensions

Find `def process_archives`
```
The snippet below shows what needs editing.

```
    # The image extensions to search for INSIDE the archives
    target_exts = ('.ima', '.flp', '.dd', '.raw', '.td0', '.fdd',
                   '.vfd', '.sdi', '.cp2', '.dmg', '.pdi', '.ana',
                   '.imd', '.ddi', '.dsk', '.img', '.sqz')
    wildcard_pattern = "*.?@?"
```


## Output Structure

All output lands alongside the script. Source directories are cleaned up (deleted if empty) after processing completes.

```
(working directory)/
│
├── _sorted_zip/                 ← extracted ZIP contents
│   ├── .img/                    ← single-extension match
│   ├── .ima/
│   ├── mixed/                   ← multiple extensions in one archive
│   ├── _wildcard/               ← wildcard-only match
│   └── CRC-Errors/              ← failed ZIPs, sorted by extension
│       ├── .img/
│       └── __Reduce/            ← Reduce ZIPs that pkunzip failed on
│
├── _sorted_arc/                 ← extracted ARC contents
│   ├── Arc/                     ← successfully extracted ARCs
│   └── CRC-Errors/              ← ARCs that failed CRC check
│
├── YYYY-MM-DD_HH-MM-SS_<folder>_Report.txt   ← extraction report
└── error_log.txt                              ← Python / 7-Zip errors
```


## Metadata

Every processed archive produces a `__<filename>__metadata.nfo` file. Metadata is written in raw binary mode to preserve original encoding.

### ZIP metadata format

```
Source ZIP: disk001.zip
Disk Label: DISK 1
Archive Comment: <raw bytes of original comment>
```

### ARC metadata format

```
Source ARC: 0UTILS.ARC
Archive Comment: <raw bytes of original comment>
```

---

## Nested ZIP support
After a ZIP is extracted, its output folder is recursively scanned for further ZIPs up to a depth of 10. Metadata is saved for nested archives.

---

## Source files and folders
Source directories are only deleted if they are **empty** after processing — any non-archive files present are left in place and the directory is kept.

---

## Terminal Dashboard

The script runs inside a full-screen curses TUI.


![](<Clipboard-20260407-01.png>)



### `--ext` mode adds two extra rows in the stats block



![](<Clipboard-20260407.png>)

```

### Stat descriptions

| Stat | Description |
|---|---|
| **Archives** | Total archives extracted so far (ZIPs + ARCs) |
| **Successful** | Extractions that completed without error |
| **Queued** | Remaining archives still to be processed — counts down to zero in real time |
| **Legacy** | ZIPs using legacy Reduce compression, routed through pkunzip |
| **Zip Errors** | ZIPs that failed CRC verification |
| **Files** | Live count of all files present in `_sorted_zip` and `_sorted_arc` |
| **ARC** | ARC archives successfully extracted |
| **ARC Errors** | ARCs that failed CRC check |
| **Skipped** | *(--ext mode only)* Archives with no matching content, left untouched |
```
### Panel descriptions

**Recent Activity** (left) — one line per archive processed, colour-coded: green for success, yellow for Reduce, red for CRC errors.

**Live Extraction** (right) — Terminal output streamed directly as each file is written to disk.

## Report File

A timestamped report is written to the working directory on completion and opened automatically:

YYYY-MM-DD_HH-MM-SS_<parent-folder-name>_Report.txt
```

The report is structured as:


          EXTRACTION SUMMARY REPORT
============================================================
Date/Time          : 2025-03-14 09:41:22
Duration           : 0:04:37.112894
Total Processed    : 380
Successful         : 371
Reduce             : 5
CRC Errors         : 4
Arc                : 12
Arc CRC Errors     : 1
Target Files Found : ---
------------------------------------------------------------

============================================================
               ARC
============================================================
SUCCESS: Extracted DISKSET1.ARC
  Location: _sorted_arc\Arc\DISKSET1
...

============================================================
             FAILED
============================================================
CRC Error: CORRUPT.ZIP
  Location: _sorted_zip\CRC-Errors\.img\CORRUPT
...

============================================================
             SUCCESS
============================================================
SUCCESS: Extracted and Moved DISK001.ZIP
  Location: _sorted_zip\.img\DISK001 --->
  [OK] DISK.IMG
...

```
# Error Handling Summary


## Error Log

`error_log.txt` captures errors that don't fit the known categories:

- Unknown 7-Zip errors
- Python exceptions during processing
- Nested ZIP extraction failures (with depth level)
- ARC processing exceptions
---

## Archive-Level Errors

### CRC Errors — ZIP Using 7-zip
The ZIP and its metadata are moved to `_sorted_zip/CRC-Errors/<ext>/`.

### CRC Errors — Legacy ZIP (Reduce)
The ZIP and its metadata are moved to `_sorted_zip/CRC-Errors/__Reduce/`.

### CRC Errors — ARC
The ARC, its metadata and any files that could be extracted are moved to `_sorted_arc/CRC-Errors/`.

### Unknown 7-Zip Errors
Output is written to `error_log.txt`.

---

## File System Errors

### WinError 2
Checks are made before a file not found error is encountered.
### WinError 3 and 5 
Timing problems with NTFS are worked around with retrys.
### Locked files
Read-only is cleared before retrying.
### File-already-moved check
Files that got removed whilst the script is running are skipped and logged silently.
### General `Exception` handler
Catches any unexpected Python error.

