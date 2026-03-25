# search_extract.py — Documentation

## Overview

`search_extract.py` is a Windows command-line tool for bulk processing legacy archive files (ZIP and ARC formats). It recursively scans a directory, extracts archives, sorts the results and seperates disk image files by file extension, captures metadata (disk labels and comments), and produces a detailed extraction report.

It is designed to handle archives from the DOS/early PC era, including formats and compression methods that modern tools struggle with, such as PKzip Reduce compression which was a precursor to the more modern zip and the original ARC archives created in the 80's.
That said, this script will work equally well on a single archive.

## Some context
After obtaining a dump of 1980-1990 archives I originally looked for a way to mass extract and document them but, found none truly suitable.

Inside all these archives there were multiple disk image types, nested archives within archives, ancient .arc and .zip files that were a pain to work on in en-mass with DOSBox. They all needed sorting too and many had CRC errors that i wanted to deal with quickly, log and put aside separately.

---

## Requirements
### Python


### Python packages

```
pip install tqdm
```

### External executables

First three must be placed in the **same folder as the script**:

| File          | Purpose                                                |
| ------------- | ------------------------------------------------------ |
| `msdos.exe`   | DOS emulator for running legacy tools                  |
| `PKUNZIP.EXE` | Legacy PKzip decompressor (for Reduce-compressed ZIPs) |
| `PKUNPAK.EXE` | Legacy ARC extractor                                   |
| `7z.exe`      | 7-Zip (must be on system PATH)                         |

---

## Obtaining Prerequisites

Each executable requires a few extraction steps from its source archive. All extraction steps below use 7-Zip unless otherwise noted.

Alternatively, you ca skip the extraction steps and download the `prerequisites.7z` from the releases page.

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
Place the script and prerequisite files in the top level folder that you want to work on.

```
python process_zips.py          # Normal mode — Process ALL archives regardless of contents
python process_zips.py --ext    # process archives containing target extensions
```

### `--ext` flag

When set, the Target extensions is enabled.
This mode is useful if you want to work on a specific data type.
Archives are only processed if they contain files matching these extensions (or the wildcard pattern `*.?@?`):

`.ima` `.flp` `.dd` `.raw` `.td0` `.fdd` `.vfd` `.sdi` `.cp2` `.dmg` `.pdi` `.ana` `.imd` `.ddi` `.dsk` `.img` `.sqz`

## Edit the script to add or remove extensions

Find `def process_archives`
The snippet below shows what needs editing.

```
    # The image extensions to search for INSIDE the archives
    target_exts = ('.ima', '.flp', '.dd', '.raw', '.td0', '.fdd',
                   '.vfd', '.sdi', '.cp2', '.dmg', '.pdi', '.ana',
                   '.imd', '.ddi', '.dsk', '.img', '.sqz')
    wildcard_pattern = "*.?@?"
```



---

## Output Structure

### ZIP output — `_sorted_zip/`

```
_sorted-zip/
├── .ima/                  ← Single-extension match
│   └── disk001/
│       ├── disk001.ima
│       ├── disk001.zip
│       └── __disk001.zip__metadata.nfo
├── mixed/                 ← Multiple extensions or wildcard + extension
├── _wildcard/             ← Wildcard-only matches (*.?@?)
├── zip_Reduce/            ← (removed — Reduce ZIPs now sort into normal folders)
└── CRC-Errors/
    ├── .ima/              ← CRC errors sorted by extension
    ├── mixed/
    └── __Reduce/          ← Reduce ZIPs that failed pkunzip extraction
```

### ARC output — `_sorted_arc/`

```
_sorted_arc/
├── Arc/
│   └── 0UTILS/
│       ├── 0UTILS.ARC
│       ├── __0UTILS.ARC__metadata.nfo
│       └── (extracted files)
└── CRC-Errors/
    └── (failed ARC extractions)
```

---

## Metadata Files

Every successfully processed archive produces a `__<filename>__metadata.nfo` file alongside the extracted contents. Metadata is written in raw binary mode to preserve original encoding.

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

Metadata is also written for CRC error archives and Reduce extraction failures, so comment data is preserved regardless of whether extraction succeeded.

---

## ZIP Processing Details



### Extraction flow

1. Peek inside the ZIP with Python's `zipfile` module to check for target extensions
2. Extract with **7-Zip** (`7z x`)
3. Write metadata file
4. Move original ZIP into the destination folder
5. Queue source directory for deletion

### Error handling

| 7-Zip output                | Action                                                                                                                      |
| --------------------------- | ----------------------------------------------------------------------------------------------------------------------------|
| `returncode 0`              | Success — sorted into `_sorted[]/<ext>`                                                                                     |
| `ERROR: Unsupported Method` | Reduce compression detected — script will extract with pkunzip via msdos.exe and sort into normal `_sorted[]/<ext>/` folder |
| `ERROR: Data Error`         | CRC error — move to `_sorted[]/CRC-Errors/<ext>/`                                                                           |
| `ERROR: CRC Failed`         | CRC error — move to `_sorted[]/CRC-Errors/<ext>/`                                                                           |
| Other error                 | Log to `error_log.txt`                                                                                                      |

### Reduce (legacy PKzip compression)

ZIPs using PKzip Reduce compression (methods 2–5) cannot be extracted by 7-Zip. When detected, they are extracted using `msdos.exe` + `PKunzip.exe -e -o -d`. If pkunzip also fails, the ZIP is moved to `_sorted-zip/CRC-Errors/__Reduce/`.

### Nested ZIPs

After extraction, the script recursively scans the destination folder for any ZIPs contained within. Each nested ZIP is:

- Extracted into its own subfolder alongside the parent's extracted contents
- Given its own metadata file
- Checked for Reduce compression before extraction (using `zipfile` to inspect compression method)
- Reported in the extraction report grouped under its parent entry

Recursion is capped at **10 levels deep** to guard against circular archives. A `processed` set tracks every ZIP path to prevent any archive being extracted more than once.

### Collision handling

If a destination folder already exists (e.g. two ZIPs with the same name in different source directories), the second is suffixed `_1-Dupe`, `_2-Dupe`, etc.

### Source directory cleanup

After all archives are processed, original source directories are deleted deepest-first. Read-only files and folders are handled automatically.
#### `All original files are moved when processed and none are deleted regardless. If a file is skipped then it will be kept in it's original location.`

---

## ARC Processing Details

ARC files are processed after ZIPs, before source directory cleanup.

### Extraction flow

1. Move ARC into destination folder `_sorted_arc/Arc/<name>/`
2. Extract with `msdos.exe PKUNPAK.EXE <filename>` from within the destination folder
3. Write metadata file
4. If `returncode != 0`, move to `_sorted_arc/CRC-Errors/`

### Comment extraction

ARC comments are read directly from raw bytes using the end-of-archive marker (`0x1A`). The comment structure after the marker is:

```
[0x00] end-of-archive null
[0x01] comment type byte
[0x1e] length byte
[0x58] first content byte
[...]  comment text, space-padded
```

If a PK signature (`\x50\x4B`) is found within the comment data (indicating a ZIP appended to the ARC), the comment is truncated at that point.

---

## Extraction Report

The report is written to the working folder, including name and date `2026-03-26_test_Report.txt` and opened automatically on completion.

### Summary section

```
          EXTRACTION SUMMARY REPORT
============================================================
Date/Time          : 2026-03-23 14:32:01
Duration           : 0:00:04.123456
Total Processed    : 47
Skipped            : 12          ← Counts archives that contained no matching extensions (Target mode only)
Successful         : 43
Reduce             : 2
CRC Errors         : 2
Arc                : 8
Arc CRC Errors     : 1
Target Files Found : 112         ← Shows --- in normal mode

  By Extension:                  ← Hidden in normal mode
.ima     : 54
.flp     : 31
------------------------------------------------------------
```

`Skipped` counts archives that contained no matching extensions (Target mode only). `Target Files Found` shows `---` in `normal` mode. `By Extension` is hidden in `normal` mode.

### Detail sections (in order)

1. **ARC** — successfully extracted ARC files
2. **ARC CRC ERRORS** — ARC files that failed extraction
3. **REDUCE** — ZIPs processed via 1980's pkunzip, including nested Reduce ZIPs grouped under their parent
4. **CRC ERRORS** — ZIPs with data errors
5. **SUCCESS** — successfully extracted ZIPs, with nested ZIPs grouped under their parent

---

## Error Log

`error_log.txt` captures errors that don't fit the known categories:

- Unknown 7-Zip errors
- Python exceptions during processing
- Nested ZIP extraction failures (with depth level)
- ARC processing exceptions

---

## Helper Functions

| Function                 | Purpose                                                                 |
| ------------------------ | ----------------------------------------------------------------------- |
| `get_sort_folder()`      | Determines `_sorted` subfolder from matched extensions                  |
| `read_zip_comment_raw()` | Reads ZIP comment bytes from EOCD record                                |
| `read_disk_label()`      | Reads disk label from ZIP central directory byte pattern                |
| `read_arc_comment()`     | Reads ARC comment bytes from end-of-archive marker                      |
| `force_remove()`         | `shutil.rmtree` error handler — strips read-only attribute before retry |
| `process_nested_zips()`  | Recursively extracts ZIPs found inside an already-extracted folder      |
| `process_arcs()`         | Scans and processes all ARC files                                       |
| `process_archives()`     | Main ZIP processing loop                                                |

---


## License and warranty
Do what you want with it.
I take no responsibility for usage.