# esa-lookup (single-notebook branch)

One self-contained Jupyter notebook — `esa_lookup.ipynb` — that mass-looks-up
SAP data into the GTF SS Database workbook via transaction **ZTBV**, plant
**ESA1**. This branch keeps the ESA developer's original structure: one code
cell per step, run top to bottom, each step writes the workbook and saves
immediately. No other source files.

(The `main` branch holds a GUI/app variant of the same process, split into
Python modules. This branch is the notebook-only line.)

## The two processes

### TO Number process (3 cells)

| # | Reads | SAP table | Filter slot | Writes |
|---|-------|-----------|-------------|--------|
| 1 | Col K (TO Number) | `LTAP` | `S3` | `ABLAD` (Unloading Point) -> **M** |
| 2 | Col N (Reservation) + Col O (Item) | `Z50CFG_ENG_CRNT` | `S15` | `QMNUM` -> **A** (matched rows only), `OBJNR`/`DISP_MATNR`/`DISP_QTY` -> **C, D, E**, match key -> **P** |
| 3 | Col C (Object Number) | `Z50CFG_ENG_VALD` | `S2` | `Z_SECTION`/`Z_MODULE`/`DESCRIPT`/`SALES_ORDER` -> **F, G, H, I** (J cleared on non-match) |

Step 2 pastes only the Reservation Numbers into SAP and then matches on the
composite key `RSNUM|RSPOS` against `Column N|Column O`. It also fills a
`SAP_Debug_Table` sheet with the raw SAP rows so a mismatch can be inspected.

### Notification Number process (2 cells)

| # | Reads | SAP table | Filter slot | Writes |
|---|-------|-----------|-------------|--------|
| 1 | Col A (Notification) | `Z50CFG_ENG_CRNT` | `S29` | `OBJNR`/`DISP_MATNR`/`DISP_QTY` -> **C, D, E** |
| 2 | Col C (Object Number) | `Z50CFG_ENG_VALD` | `S2` | -> **F, G, H, I** (J cleared on non-match) |

## Column letters — read this once

The workbook once gained one extra column on the left (Notification Number in
Column A), so **older descriptions and scripts referred to columns one letter
to the LEFT** of the current file. The tables above and the notebook's
description cells match the current workbook and the code. When in doubt, the
`*_col = <number>` constants at the top of each code cell are the truth
(A=1 ... P=16).

## ZTBV filter slots (recorded)

ZTBV names its select-options generically; these are the recorded slots per
table. If SAP's screen ever changes, re-record the slot and update the one
`findById` line in the affected cell.

| Table | Field | Slot |
|-------|-------|------|
| `LTAP` | Transfer Order Number | `S3` |
| `Z50CFG_ENG_CRNT` | Number of Reservation/Depend | `S15` |
| `Z50CFG_ENG_CRNT` | Transfer Order Number (unused, for reference) | `S19` |
| `Z50CFG_ENG_CRNT` | Notification No | `S29` |
| `Z50CFG_ENG_VALD` | Object number | `S2` |

## Prerequisites

- Windows with **SAP GUI for Windows** (Java client not supported) and
  **Microsoft Excel** installed
- Python 3.10+ with Jupyter and `pip install -r requirements.txt` (pywin32
  only)
- SAP GUI Scripting enabled, and the scripting notifications turned OFF so
  no popup freezes a run (one-time, per user):
  Options -> Accessibility & Scripting -> Scripting:
  - Enable scripting: **On**
  - Notify when a script attaches / opens a connection: **Off**
  - Notify when a script attempts to read / write files: **Off**
  - Show native Microsoft Windows dialogs: **Off**
- Server side `sapgui/user_scripting = TRUE` (already true if these scripts
  have ever worked)

## How to run

1. Log into SAP GUI (any session with ZTBV / plant ESA1 access) and leave it
   open. The scripts attach to the first session.
2. Close the target workbook in Excel if you have it open — each cell opens
   the file itself (hidden), writes, saves, and closes it.
3. Open `esa_lookup.ipynb` in Jupyter.
4. Check the `excel_file_path` at the top of each code cell you plan to run.
5. Run the cells of your process **in order, one at a time** (TO: steps
   1 -> 2 -> 3; Notification: steps 1 -> 2). Each cell shows a completion
   popup with matched / not-matched counts. Re-running a cell is safe — every
   step rewrites its own output columns from scratch.

## Troubleshooting

- **A cell hangs with a SAP popup open** — a scripting notification is still
  enabled; see Prerequisites and turn them off.
- **"Excel file opened as Read-Only"** — the workbook is open somewhere else
  (or locked by OneDrive). Close it and re-run the cell.
- **Wrong or stale values were matched** — every code cell clears the SAP
  multi-select dialog (`btn[16]`, "Delete Entire Selection") before uploading
  this run's keys, so leftovers from a previous run cannot join the query. If
  results still look wrong, check the `SAP_Debug_Table` sheet written by TO
  step 2: it holds the raw SAP rows and the exact keys used on both sides.
- **Fewer matches than expected** — compare Column P (Excel match key) with
  the `SAP Match Key Used` column of `SAP_Debug_Table`; a leading-zero or
  formatting difference is usually visible immediately.
