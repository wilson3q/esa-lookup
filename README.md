# esa-lookup

Windows GUI app that mass-looks-up SAP data into an Excel file. Replaces the
per-cell notebook workflow with ALV grid exports + bulk Excel writes.

## Gen 4 (this branch): in-memory chaining, single write pass

Steps no longer write to Excel one-by-one. All SAP lookups run first and
chain **in memory** (e.g. TO step 3 keys off the OBJNR values step 2 just
fetched, not off column C in the sheet); the workbook is then written
**once**, at the end. Consequences:

- Nothing is written until the SAP phase is over, so a run can never leave
  a half-written column. **Stop** writes nothing at all. A *failure* writes
  the steps that already completed and leaves the rest of the columns as
  they were -- see [What happens when a step fails](#what-happens-when-a-step-fails).
- Large runs are faster: 1 bulk key-read + 1 bulk write total instead of
  a write+re-read round-trip between every step.
- Output columns, matching rules, and preserve/clear semantics are
  unchanged from the previous version.
- SAP steps are self-checking: the status bar is read after every query
  (SAP errors surface with SAP's own message), export row counts are
  verified against the grid (a truncated export aborts the run instead of
  silently under-matching), an empty result is a clean "all unmatched"
  outcome, and key lists larger than 2000 are split into multiple
  query+export rounds automatically. The multi-select dialog is also
  cleared before each paste, so values left over from a previous run can
  no longer leak into the filter.
- Filter values now reach SAP through a **temp text file** ('Import from
  Text File' in the multi-select dialog) instead of the OS clipboard --
  copying something to the clipboard while a run is in flight can no
  longer corrupt the filter. If the import dialog is not scriptable on a
  given SAP GUI version, the app logs a warning and automatically falls
  back to the old clipboard paste.
- **Dry run mode** (checkbox in the GUI, `DRY_RUN = True` in the
  notebook): runs every SAP lookup and reports match counts + sample rows
  in the log, but leaves the workbook completely untouched. Use it to
  sanity-check a new file or ALV layout before writing anything.

## Prerequisites

- Windows with **SAP GUI for Windows** installed (Java client is not supported)
- **SAP GUI Scripting enabled** on the client (Options -> Accessibility &
  Scripting -> Enable scripting) and permitted on the server side
  (Basis-controlled; already true if the reference notebook works)
- Access to transaction **`ZTBV`** for plant **`ESA1`** with read access to
  tables `LTAP`, `Z50CFG_ENG_CRNT`, `Z50CFG_ENG_VALD`
- **Microsoft Excel** installed (Excel COM is used to attach/read/write)
- Python 3.10+

## Install

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Run

1. Log into SAP GUI (any session; script attaches to the first one)
2. Open the target Excel file in Excel (OneDrive/SharePoint paths are fine)
3. Launch:
   ```powershell
   python esa_lookup.py
   ```
4. Browse to the Excel file, pick either the **TO Number process** or
   **Notification Number process**, click **Run** (tick **Dry run** first
   if you only want a match report without writing anything)

There is also a **Diagnose** button. It reads the ZTBV selection screens for
the chosen process and logs what it finds -- no Excel file needed, no query
run, nothing written. Use it when a run stops with *"Could not resolve the
... filter"*; see [Troubleshooting](#troubleshooting).

## What happens when a step fails

Each step writes its own block of columns, and steps run in order. If a step
fails, the steps that **already completed are written and saved**, and the
failed step plus everything after it leave their columns exactly as they
were. The log names both sets:

```
WRITTEN  step 1 (LTAP): columns M -- 3 matched / 0 unmatched
NOT RUN  step 2 (Z50CFG_ENG_CRNT): columns C, D, E, A, N, O, P left as they were
```

Re-running after a failure is safe: every step rewrites its own columns from
scratch, so a partially-filled sheet is completed rather than corrupted.

Pressing **Stop** is different -- an explicit cancel writes nothing at all.

## Workflows

### TO Number process (3 steps)

| # | Read from | SAP table         | SAP filter field   | Writes to Excel columns |
|---|-----------|-------------------|--------------------|--------------------------|
| 1 | Col K     | `LTAP`            | TO Number          | `ABLAD` -> M             |
| 2 | Col N + O | `Z50CFG_ENG_CRNT` | Reservation Number | `QMNUM` -> A, `OBJNR` -> C, `DISP_MATNR` -> D, `DISP_QTY` -> E |
| 3 | Col C     | `Z50CFG_ENG_VALD` | Object Number      | `Z_SECTION`/`Z_MODULE`/`DESCRIPT`/`SALES_ORDER`/`LID` -> F..J |

### Notification Number process (2 steps)

| # | Read from | SAP table         | SAP filter field   | Writes to Excel columns |
|---|-----------|-------------------|--------------------|--------------------------|
| 1 | Col A     | `Z50CFG_ENG_CRNT` | Notification (QMNUM) | `OBJNR` -> C, `DISP_MATNR` -> D, `DISP_QTY` -> E |
| 2 | Col C     | `Z50CFG_ENG_VALD` | Object Number      | `Z_SECTION`/`Z_MODULE`/`DESCRIPT`/`SALES_ORDER`/`LID` -> F..J |

## ALV layout requirement

The ZTBV export writes only the currently visible columns. Before running
the app for the first time, open each of the three tables in ZTBV, add
every field the app reads (see tables above), and **save as a layout**
(default variant). The app depends on those fields being present in the
exported file.

## Troubleshooting

- **"Cannot attach to SAP GUI"** - not logged in, or scripting disabled.
  Check Options -> Accessibility & Scripting.
- **"Cannot attach to Excel"** - open the target .xlsx in Excel first, or
  let the app open it (blank/read-only files trip OneDrive).
- **Export dialog IDs don't match** - the exact `wnd[1]` field names for
  the Save-As dialog vary slightly between SAP GUI versions. Record a
  single export via SAP GUI Script Recorder and adjust `EXPORT_SEQUENCES`
  in `sap_ops.py`.
- **"file import unavailable ... falling back to clipboard paste"** - the
  multi-select dialog's Import-from-Text-File path could not be scripted
  (older SAP GUI, or 'Show native Microsoft Windows dialogs' is On). The
  run still works via the clipboard; to fix the import path, check that
  native dialogs are Off, or record the import once with the Script
  Recorder and adjust `fill_multi_value_filter_from_file` in `sap_ops.py`.
- **"Could not resolve the `<FIELD>` filter on `<TABLE>`"** - ZTBV names its
  select-options generically (`S3`, `S15`, `S29`), so a filter that is not
  already recorded in `PUSH_BUTTONS` has to be identified by the label
  printed beside it on the selection screen. This error means zero or
  several labels matched, and the app refuses to guess -- pasting keys into
  the wrong filter returns confidently wrong data, which is worse than
  stopping. To fix it:
  1. Click **Diagnose**. It lists every `S<n>` on that screen with its
     label and tooltip, and dumps a full per-control inventory to the log
     file. Read off which `S<n>` is the field you need.
  2. Pin it **without a rebuild** by setting an environment variable, then
     re-run:
     ```powershell
     $env:ESA_LOOKUP_PUSH_Z50CFG_ENG_CRNT_TO_NUMBER = "S7"
     ```
     The name is `ESA_LOOKUP_PUSH_<TABLE>_<FIELD>`, uppercased, with any
     non-alphanumeric character replaced by `_`. The value is the bare
     `S<n>` (or a full control id).
  3. To make it permanent, add it to `PUSH_BUTTONS` in `sap_ops.py`:
     ```python
     ("Z50CFG_ENG_CRNT", "TO_NUMBER"): push_button_id("S7"),
     ```
  If every label in the Diagnose output reads `(no label found)`, send the
  log **file** -- it carries the raw control dump with the coordinates and
  tooltips needed to work out the mapping.
- **Fewer matches than expected** - your ALV layout is probably missing a
  key column; edit and re-save the layout, then re-run.
- **"WARNING: sent N unique key(s) to SAP but only M matched"** - shown
  in the log when SAP came back with less than 10% of what was requested.
  Usually means the query hit an error screen, the ALV filter rejected
  the values (wrong data type / leading-zero mismatch), or you're pointed
  at the wrong plant / table for this environment.

## Logs

Every run writes a persistent log file. The GUI shows a one-line summary
per event; the file gets the same events **plus** full exception
tracebacks, elapsed timings, environment info (Python / pandas / openpyxl
versions), and any SAP screen titles observed. If the GUI closed before
you could read it, check the file.

| Where                                                        | What          |
|--------------------------------------------------------------|---------------|
| `%LOCALAPPDATA%\esa-lookup\logs\esa-lookup-YYYYMMDD-HHMMSS.log` | One file per run |
| Line format                                                  | `[YYYY-MM-DD HH:MM:SS] LEVEL msg` |
| Retention                                                    | Most-recent 20 runs; older files auto-pruned on each launch |

The log file path is echoed into the GUI log at run start so you can
copy it easily. To open the folder quickly:

```powershell
explorer "$env:LOCALAPPDATA\esa-lookup\logs"
```

When reporting a bug or asking for help, attach the log file for the
failed run - it usually has the full traceback plus everything the app
was doing right before the failure.

## Build a standalone `.exe`

Ship the app to users who don't have Python installed.

### 1. Prereqs on the build machine

- **Python 3.10+ from [python.org](https://www.python.org/downloads/)**.
  During install, tick "Add python.exe to PATH". Confirm afterwards:
  ```powershell
  py --version
  ```
  The `python.exe` Windows offers by default from the Microsoft Store is a
  redirect shim and will not work - the build script rejects it.
- ~1 GB free on your local `C:\Users\<you>\AppData\Local` drive for the
  build virtualenv + PyInstaller work cache.
- Internet access for the first build (to download pywin32, pandas,
  openpyxl, pyinstaller wheels).

### 2. Run the build

```powershell
cd Z:\shared\esa-lookup
powershell -ExecutionPolicy Bypass -File .\build.ps1
```

First build: 3-8 minutes (downloads + freezes ~300 MB of wheels).
Subsequent builds: 30-90 seconds (venv is reused).

### 3. Ship the output

The finished binary lands at:

```
Z:\shared\esa-lookup\dist\esa-lookup.exe
```

Single file, ~80-120 MB, no console window. Copy it anywhere. The target
machine needs **SAP GUI for Windows** and **Microsoft Excel** installed,
but NOT Python.

### Where the build lives

The venv and PyInstaller's work directory are placed on the **local disk**,
not next to the source:

| Path | Purpose |
|------|---------|
| `%LOCALAPPDATA%\esa-lookup-build\.venv-build\` | Python venv (deps + PyInstaller) |
| `%LOCALAPPDATA%\esa-lookup-build\work\` | PyInstaller build cache + .spec file |
| `Z:\shared\esa-lookup\dist\esa-lookup.exe` | Final output (only file on the share) |

**Why**: `Z:` is a network / SMB share. pip installs files with an
atomic-rename step that SMB does not reliably support, so installing on
the share fails partway through with `[Errno 17] File exists` and leaves
a half-built venv. Keeping the venv on the local NTFS disk avoids the
whole class of failures.

The script auto-detects and deletes any legacy `Z:\shared\esa-lookup\.venv-build\`
left over from a previous version of `build.ps1`.

### Troubleshooting the build

| Symptom | Fix |
|---------|-----|
| `No real Python interpreter found` | Install Python 3.10+ from python.org and tick "Add to PATH". The Windows Store shim is rejected. |
| `[Errno 17] File exists` during pip install | You're still on the old on-share venv, or `%LOCALAPPDATA%\esa-lookup-build\` got corrupted. Delete that folder and re-run. |
| `PyInstaller finished but ... esa-lookup.exe is missing` | Read the PyInstaller output above the failure. Usually a missing hidden import - add `--hidden-import <name>` to the `$pyinstallerArgs` array in `build.ps1`. |
| Rebuilding after editing the .py files still ships old behavior | The `dist\` folder is wiped every build; if you see the old exe, another instance may be holding it open. Close any running `esa-lookup.exe` first. |
| Corporate AV quarantines the .exe on first launch | Common for fresh PyInstaller binaries. Sign the exe (`signtool`) before distribution, or ask IT for an AV exception on the download path. |
| Want to force a specific Python version (e.g. Python 3.12, not the newest) | Edit `build.ps1`: change the `Resolve-Python` function to return `'py -3.12'` instead of `'py'`. |
| Want to start completely fresh | Delete `%LOCALAPPDATA%\esa-lookup-build\` and `Z:\shared\esa-lookup\dist\`, then re-run. |

### Rebuilding after code changes

Just re-run `build.ps1`. It:

- Reuses the venv (unless it's missing or broken)
- Wipes `dist\` and the local `work\` cache
- Rebuilds fresh from the current source

You do **not** need to bump a version, edit a spec file, or clean up
anything manually between builds.

## Setup on end-user machines (running the `.exe`)

If you built the `.exe` and are distributing it, each user needs a
**one-time setup** on their Windows machine before the app will run
without popup dialogs interrupting the automation.

### 1. What must already be installed

| Required                             | Why                                                                |
|--------------------------------------|--------------------------------------------------------------------|
| Windows 10 or 11 (64-bit)            | The exe is a Windows binary                                        |
| **SAP GUI for Windows**              | The app scripts it via COM. Java SAP GUI is NOT supported.         |
| **Microsoft Excel** (desktop)        | The app scripts Excel via COM. Excel Online / Web will NOT work.   |

**Not needed on the target machine:** Python, pip, pywin32, admin
rights, .NET, or Visual C++ redistributables. Everything is either
bundled inside the `.exe` or already present on Windows 10 / 11.

### 2. Enable SAP GUI Scripting (client side)

Open SAP GUI, click the **Options** icon (gear, top-right of the main
window) → **Accessibility & Scripting** → **Scripting**.

Set these:

| Setting                                                       | Value    |
|---------------------------------------------------------------|----------|
| Enable scripting                                              | **On**   |
| Notify when a script attaches to SAP GUI                      | **Off**  |
| Notify when a script opens a connection                       | **Off**  |

### 3. Disable the file-write / file-read security warning (critical)

The ALV export step writes a spreadsheet file to a temp folder. By
default SAP GUI pops up a **"A script is trying to write / read a file"**
modal for every export - which freezes the automation until the user
clicks Allow. This must be turned off for unattended runs.

In the same **Scripting** panel (or a **Security** sub-tab in newer SAP
GUI versions), set:

| Setting                                                       | Value    |
|---------------------------------------------------------------|----------|
| Notify when a script attempts to **write** to files           | **Off**  |
| Notify when a script attempts to **read** files               | **Off**  |
| Show native Microsoft Windows dialogs                         | **Off**  |

Click **OK / Apply**. These settings persist for the Windows user
profile - one-time setup per user, per machine.

### 4. Server-side (SAP Basis / IT, once per SAP instance)

Server-side scripting must be enabled by the Basis team:

```
sapgui/user_scripting = TRUE
```

If your reference notebook already works, this is already set - no need
to re-ask.

### 5. Every-run prerequisites

Before double-clicking `esa-lookup.exe`, the user must:

1. Be **logged into an SAP GUI session** with access to transaction
   `ZTBV`, plant `ESA1`.
2. Have the **target Excel file open in Microsoft Excel** (or at least
   openable - not locked read-only by OneDrive sync).

That's it. The exe attaches to whatever SAP session and Excel instance
are currently running.

### First-launch notes

- The first time an end user runs the exe, PyInstaller's onefile
  bootstrapper self-extracts to `%TEMP%\_MEI*\` and takes a few seconds
  before the window appears. Subsequent launches are faster.
- Corporate antivirus sometimes false-flags fresh PyInstaller binaries.
  If IT quarantines it, sign the exe with `signtool` before distribution
  or request an AV exception on the download location.

## Tests

The test suite stubs the COM boundary (win32com / pythoncom), so it runs
on **any OS** -- no SAP GUI, no Excel, no Windows required:

```
pip install -r requirements-dev.txt
python -m pytest
```

Covered: key normalization + lookup building (`tests/test_keys.py`), the
Gen 4 in-memory column resolver (`tests/test_virtual_sheet.py`), reading
the `S<n>` -> field mapping off the selection screen
(`tests/test_selection_screen.py`), and both workflows end-to-end against a
fake SAP/Excel -- including partial-write-on-failure, chunking, 0-row
results, truncated exports, the clipboard fallback, and dry run
(`tests/test_workflow_e2e.py`).

Two notes for anyone extending these, both learned the hard way:

- `tests/conftest.py` monkeypatches `sap_ops.resolve_push_button`, so the
  end-to-end tests do **not** exercise filter resolution. That path is
  covered only by `tests/test_selection_screen.py`.
- The fakes in `test_selection_screen.py` carry both coordinate systems
  (`CharTop`/`CharLeft` and `Top`/`Left`), with each label's pixel `Top`
  deliberately offset from the button beside it. An earlier fake gave them
  identical `Top` values, which made a broken pairing look correct -- the
  suite stayed green while every filter on the customer's screen came back
  `(no label found)`.

## Files

- `esa_lookup.py` - tkinter GUI entry point
- `sap_ops.py`    - SAP GUI Scripting helpers
- `excel_ops.py`  - Excel COM helpers (bulk read/write)
- `pipeline.py`   - workflow orchestrator + key-normalization helpers
- `build.ps1`     - build a standalone `dist\esa-lookup.exe`
- `tests/`        - COM-stubbed pytest suite (runs anywhere)
