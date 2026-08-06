"""Regenerate esa_lookup.ipynb from the .py source files.

Run this after editing pipeline.py, excel_ops.py, or sap_ops.py so the
self-contained notebook stays in sync with the source of truth.

Usage:
    py build_notebook.py

No dependencies beyond the stdlib.

What it does:
- Reads pipeline.py + excel_ops.py + sap_ops.py verbatim
- Strips their module docstrings and top-level import statements
  (consolidated into one imports cell in the notebook)
- Renames `attach` -> `excel_attach` / `sap_attach` and `save` ->
  `excel_save` so they can coexist in one module namespace
- Rewrites `excel_ops.X` / `sap_ops.X` call sites everywhere to bare
  identifiers so the code runs after inlining
- Assembles ONE collapsible "Engine" cell (imports + all three modules +
  a message-box event handler + `run_process()`), then one runnable cell
  per process: "# TO Number Process" and "# Notification Number Process"
  -- the shape the ESA operator's own notebook used

If any of the source .py files gain a NEW function or dataclass, it
automatically shows up in the regenerated notebook -- no builder edit
needed. If a function is RENAMED (specifically `attach` on either module,
or `save` on excel_ops), the rename maps in this file need to grow.
"""
from __future__ import annotations

import json
import re
from pathlib import Path


ROOT = Path(__file__).resolve().parent


def read_py(name: str) -> str:
    return (ROOT / name).read_text(encoding="utf-8")


def strip_header(src: str) -> str:
    """Remove the module docstring, `from __future__ ...`, and every
    top-level import statement. Keep everything after the last import.
    """
    lines = src.splitlines(keepends=True)
    i = 0
    # shebang / encoding cookies
    while i < len(lines) and (lines[i].startswith("#!") or "coding:" in lines[i]):
        i += 1
    # blank / comment lines before the docstring
    while i < len(lines) and (lines[i].strip() == "" or lines[i].strip().startswith("#")):
        i += 1
    # module docstring
    if i < len(lines):
        stripped = lines[i].lstrip()
        quote = '"""' if stripped.startswith('"""') else ("'''" if stripped.startswith("'''") else None)
        if quote:
            if lines[i].count(quote) >= 2:
                # single-line docstring
                i += 1
            else:
                i += 1
                while i < len(lines) and quote not in lines[i]:
                    i += 1
                if i < len(lines):
                    i += 1
    # imports and interstitial blanks / comments
    while i < len(lines):
        stripped = lines[i].strip()
        if stripped == "" or stripped.startswith("#"):
            i += 1
            continue
        if (stripped.startswith("from __future__")
                or stripped.startswith("import ")
                or stripped.startswith("from ")):
            i += 1
            continue
        break
    # trim leading blank lines from the remainder
    while i < len(lines) and lines[i].strip() == "":
        i += 1
    return "".join(lines[i:])


def _strip_module_refs(src: str) -> str:
    """Rewrite every excel_ops.X / sap_ops.X reference. Rename maps handle
    the two functions whose bare name would collide (`attach`) or shadow a
    builtin (`save`); everything else drops the module prefix."""
    src = re.sub(r'\bexcel_ops\.attach\b', 'excel_attach', src)
    src = re.sub(r'\bexcel_ops\.save\b', 'excel_save', src)
    src = re.sub(r'\bsap_ops\.attach\b', 'sap_attach', src)
    src = re.sub(r'\bexcel_ops\.', '', src)
    src = re.sub(r'\bsap_ops\.', '', src)
    return src


def rename_excel_ops(src: str) -> str:
    """Inline-adapt excel_ops.py source: rename `attach`/`save` at their
    definitions, then strip any residual `excel_ops.` self-refs (rare;
    usually only inside diagnostic strings)."""
    src = re.sub(r'^def attach\(', 'def excel_attach(', src, flags=re.MULTILINE)
    src = re.sub(r'^def save\(', 'def excel_save(', src, flags=re.MULTILINE)
    return _strip_module_refs(src)


def rename_sap_ops(src: str) -> str:
    """Inline-adapt sap_ops.py source: rename `attach` at its definition,
    then strip any residual `sap_ops.` self-refs (e.g. the fallback error
    message that names `sap_ops.export_alv_to_file`)."""
    src = re.sub(r'^def attach\(', 'def sap_attach(', src, flags=re.MULTILINE)
    return _strip_module_refs(src)


def rewrite_pipeline_refs(src: str) -> str:
    """Adapt pipeline.py: replace every excel_ops.X / sap_ops.X reference
    with the bare identifier the inlined notebook code binds."""
    return _strip_module_refs(src)


# ---------------------------------------------------------------------------
# Cell content (notebook-only pieces + consolidated imports)
# ---------------------------------------------------------------------------

IMPORTS = """\
from __future__ import annotations

import contextlib
import difflib
import io
import os
import re
import sys
import tempfile
import threading
import time
import traceback
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from typing import Callable

import pandas as pd
import pythoncom
import win32com.client
from win32com.client import constants  # noqa: F401 (loaded lazily by pywin32)
"""

ENGINE_TAIL = """\
# ---------------------------------------------------------------------------
# Notebook front end: message boxes + run_process(), the ONLY function the
# process cells call. Everything above this line is shared machinery.
# ---------------------------------------------------------------------------

_MB_OKCANCEL, _MB_ICONERROR, _MB_ICONINFO, _MB_TOPMOST = 0x1, 0x10, 0x40, 0x40000


def _msgbox(title: str, message: str, flags: int) -> int:
    \"\"\"The original notebook's ctypes message box. Returns the button id
    (1 = OK, 2 = Cancel).\"\"\"
    import ctypes
    return ctypes.windll.user32.MessageBoxW(0, message, title, flags)


def make_event_handler(popups: bool):
    \"\"\"Log lines print into the cell output; popup events become real
    Windows message boxes when popups=True, printed banners when False.
    ERROR popups always box regardless -- they only appear when something is
    wrong, which is exactly when the box is wanted.\"\"\"
    boxes_possible = sys.platform == "win32"

    def on_event(kind, payload):
        if kind == "log":
            msg, level = payload
            prefix = {"ok": "[OK]  ", "warn": "[WARN]", "error": "[ERR] ",
                      "info": "      "}.get(level, "      ")
            print(f"{prefix} {msg}")
        elif kind == "status":
            print(f"...    {payload}")
        elif kind == "popup":
            p = payload
            banner = "!" if p.get("kind") == "error" else "-"
            print(banner * 60)
            print(f"[{p.get('kind', 'info').upper()}] {p['title']}")
            print(p["message"])
            print(banner * 60)
            if p.get("ack") is not None:                     # step popup
                proceed = True
                if popups and boxes_possible:
                    ret = _msgbox(p["title"], p["message"],
                                  _MB_OKCANCEL | _MB_ICONINFO | _MB_TOPMOST)
                    proceed = (ret == 1)
                if p.get("result") is not None:
                    p["result"]["proceed"] = proceed
                p["ack"].set()
            elif p.get("kind") == "error":
                if boxes_possible:
                    _msgbox(p["title"], p["message"],
                            _MB_ICONERROR | _MB_TOPMOST)
            elif popups and boxes_possible:
                _msgbox(p["title"], p["message"],
                        _MB_ICONINFO | _MB_TOPMOST)
        elif kind == "done":
            ok = payload
            print(f"===== process {'succeeded' if ok else 'FAILED'} =====")

    return on_event


# Remembered across cells, so Run All asks for the workbook ONCE and the
# Notification cell re-uses the file the TO cell picked.
_last_browsed = ""


def run_process(workflow: str, excel_path: str = "", popups: bool = True,
                dry_run: bool = False, diagnose: bool = False) -> bool:
    \"\"\"Run one process end to end.

    workflow    "TO" or "NOTIF"
    excel_path  full path to the workbook; "" opens a Browse dialog the
                first time and re-uses that file for later cells in the
                same kernel session (restart the kernel, or set EXCEL_PATH,
                to pick a different file)
    popups      message box after each step (OK = continue, Cancel = stop).
                One-line off switch in the process cell. Error boxes always
                show regardless.
    dry_run     look everything up, report counts, write nothing
    diagnose    dump the ZTBV selection screens instead of processing
                (no Excel file needed)

    On failure or Cancel this raises SystemExit, so a Run All stops HERE
    instead of blindly executing the next process cell.
    \"\"\"
    global _last_browsed
    if workflow not in WORKFLOWS:
        raise SystemExit(
            f"workflow must be one of {list(WORKFLOWS)}, got {workflow!r}")
    if not diagnose and not excel_path:
        if _last_browsed:
            excel_path = _last_browsed
            print(f"Re-using the workbook selected earlier this session:")
            print(f"  {excel_path}")
            print("(set EXCEL_PATH in the cell, or restart the kernel, to "
                  "pick a different file)")
        else:
            import tkinter as tk
            from tkinter import filedialog
            _root = tk.Tk()
            _root.withdraw()
            _root.attributes("-topmost", True)
            try:
                excel_path = filedialog.askopenfilename(
                    title="Select the Excel workbook to process",
                    filetypes=[("Excel workbooks", "*.xlsx *.xlsm *.xlsb"),
                               ("All files", "*.*")],
                )
            finally:
                _root.destroy()
            if not excel_path:
                raise SystemExit(
                    "No file selected -- nothing was run, and the cells "
                    "after this one were stopped.")
            _last_browsed = excel_path
    cfg = RunConfig(
        excel_path=excel_path,
        workflow=workflow,
        stop_event=threading.Event(),
        dry_run=dry_run,
        diagnose=diagnose,
        step_popups=popups,
    )
    ok = run(cfg, on_event=make_event_handler(popups))
    if not ok:
        # Stop a Run All at the point of failure. The log above (and the
        # log file) say what happened; the next process must not run on top
        # of a failed or cancelled one.
        raise SystemExit(
            "This process did not complete (failed or cancelled) -- the "
            "cells after this one were stopped. Fix the issue and re-run "
            "this cell.")
    return ok


print("Engine loaded. Now run your process cell below")
print("(TO Number Process or Notification Number Process).")
"""

TO_CELL = """\
# =========================================================
# TO Number Process        (run the Engine cell first, once)
# =========================================================
POPUPS = True     # message box after each step (OK = continue, Cancel =
                  # stop). ONE-LINE OFF SWITCH: set to False once the first
                  # days have gone smoothly. Error boxes always show.
DRY_RUN = False   # True = report match counts only, write nothing
EXCEL_PATH = r""  # paste the workbook's full path here, or leave "" to browse

run_process("TO", excel_path=EXCEL_PATH, popups=POPUPS, dry_run=DRY_RUN)
"""

NOTIF_CELL = """\
# =========================================================
# Notification Number Process   (run the Engine cell first)
# =========================================================
POPUPS = True     # message box after each step -- same one-line switch
DRY_RUN = False   # True = report match counts only, write nothing
EXCEL_PATH = r""  # paste the workbook's full path here, or leave "" to browse

run_process("NOTIF", excel_path=EXCEL_PATH, popups=POPUPS, dry_run=DRY_RUN)
"""

DIAGNOSE_CELL = """\
# Only needed when a run stops with "Could not resolve the ... filter".
# Reads the ZTBV selection screens for the chosen process and prints which
# S<n> filter slot is which. No Excel file, no query, nothing written.
# Send the log file it names.
#
# Deliberately COMMENTED OUT so a Run All never triggers it -- uncomment
# the next line only when you need it, then run this cell:
# run_process("TO", diagnose=True)          # or "NOTIF"

# If Diagnose names the right slot, pin it here and re-run your process --
# example (uncomment and adjust):
# PUSH_BUTTONS[("Z50CFG_ENG_CRNT", "RSNUM")] = push_button_id("S15")
print("Diagnose cell: nothing to do (see the comments in this cell).")
"""

# ---------------------------------------------------------------------------
# Cell constructors
# ---------------------------------------------------------------------------

def md(source: str) -> dict:
    return {
        "cell_type": "markdown",
        "metadata": {},
        "source": source.splitlines(keepends=True),
    }


def code(source: str) -> dict:
    # Every code cell ends without a trailing blank line -- Jupyter is happier.
    src = source.rstrip("\n") + "\n"
    return {
        "cell_type": "code",
        "execution_count": None,
        "metadata": {},
        "outputs": [],
        "source": src.splitlines(keepends=True),
    }


# ---------------------------------------------------------------------------

CELLS = [
    md("""\
# esa-lookup -- TO Number & Notification Number processes

Fills the GTF SS Database workbook from SAP (transaction **ZTBV**, plant
**ESA1**). Same steps, same message boxes, and the same results as the
original per-step notebook -- but each step runs as **one SAP query + one
bulk Excel write**, so a full sheet takes seconds instead of minutes, the
multi-select filter is cleared of leftovers before every paste, and the
workbook is written **once, at the end** (never left half-filled: if a step
fails, the completed steps are written and the rest of the columns are left
exactly as they were).

## How to use -- 3 actions

1. Log into SAP GUI and open the workbook in Excel (or know its path).
2. Run the **Engine** cell below **once** -- then collapse that section and
   forget it.
3. Run YOUR process cell: **TO Number Process** or **Notification Number
   Process**. A message box reports each step: **OK** = continue,
   **Cancel** = stop with the workbook untouched.

On a combined sheet run the **TO process FIRST**, then the Notification
process -- the second pass fills the notification-only rows (usually most
of the sheet).

**Cell -> Run All works too**, and runs exactly that order: Engine, TO
process, Notification process. You browse to the workbook once; the second
process re-uses it automatically. If a process fails or you press Cancel,
the cells after it are stopped. The Diagnose cell at the bottom never runs
unless you uncomment it.

Every run also writes a log file to `%LOCALAPPDATA%\\esa-lookup\\logs\\`
(last 20 kept). Attach it whenever reporting a problem.
"""),

    md("""\
## Engine -- run once, then collapse

Everything both processes share: Excel bulk read/write, SAP GUI scripting,
and the step runner. **Nothing in here needs reading or editing.** It is
auto-generated from the app's source files by `build_notebook.py`; to change
behavior, change those files and regenerate. Do not hand-edit this notebook.
"""),
    code(IMPORTS
         + "\n\n" + rename_excel_ops(strip_header(read_py("excel_ops.py")))
         + "\n\n" + rename_sap_ops(strip_header(read_py("sap_ops.py")))
         + "\n\n" + rewrite_pipeline_refs(strip_header(read_py("pipeline.py")))
         + "\n\n" + ENGINE_TAIL),

    md("""\
# TO Number Process

| Step | Reads | SAP table | Writes |
|------|-------|-----------|--------|
| 1 | Col **K** (TO Number) | `LTAP` | Unloading Point -> **M** |
| 2 | Col **N + O** (Reservation + Item) | `Z50CFG_ENG_CRNT` | Notification -> **A** (matched rows only), Object / Material / Qty -> **C, D, E**, match key -> **P** |
| 3 | Col **C** (Object Number) | `Z50CFG_ENG_VALD` | Section / Module / Description / Sales Doc. -> **F, G, H, I** |

Rows without a TO number or reservation pair are skipped -- and their
column A is left untouched, so the Notification process can fill them next.
"""),
    code(TO_CELL),

    md("""\
# Notification Number Process

| Step | Reads | SAP table | Writes |
|------|-------|-----------|--------|
| 1 | Col **A** (Notification) | `Z50CFG_ENG_CRNT` | Object / Material / Qty -> **C, D, E** |
| 2 | Col **C** (Object Number) | `Z50CFG_ENG_VALD` | Section / Module / Description / Sales Doc. / LID -> **F..J** |

On a combined sheet, run this **after** the TO process -- this is the pass
that fills the notification-only rows.
"""),
    code(NOTIF_CELL),

    md("""\
## Troubleshooting

- **A step looks wrong**: press **Cancel** on its message box -- the run
  stops and the workbook is untouched. Then re-run with `DRY_RUN = True`
  to inspect match counts without writing.
- **"Could not resolve the ... filter"**: run the Diagnose cell below and
  send the log file it names.
- Every popup's text is also in the log file, so nothing is lost when a
  box is dismissed. Logs: `%LOCALAPPDATA%\\esa-lookup\\logs\\`.
"""),
    code(DIAGNOSE_CELL),
]


NOTEBOOK = {
    "cells": CELLS,
    "metadata": {
        "kernelspec": {
            "display_name": "Python 3",
            "language": "python",
            "name": "python3",
        },
        "language_info": {
            "codemirror_mode": {"name": "ipython", "version": 3},
            "file_extension": ".py",
            "mimetype": "text/x-python",
            "name": "python",
            "nbconvert_exporter": "python",
            "pygments_lexer": "ipython3",
            "version": "3.10",
        },
    },
    "nbformat": 4,
    "nbformat_minor": 5,
}


def main() -> None:
    out_path = ROOT / "esa_lookup.ipynb"
    with open(out_path, "w", encoding="utf-8", newline="\n") as f:
        json.dump(NOTEBOOK, f, indent=1, ensure_ascii=False)
        f.write("\n")
    print(f"wrote {out_path}")
    print(f"cells: {len(CELLS)}")


if __name__ == "__main__":
    main()
