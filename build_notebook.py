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
- Adds a notebook-only `print_event` handler + tkinter file-picker
  config cell + a final `run()` execute cell

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

PRINT_EVENT = """\
def print_event(kind: str, payload) -> None:
    \"\"\"Simple stdout event handler for notebook use. Same (kind, payload)
    contract the .py app's tkinter callback receives.\"\"\"
    if kind == "log":
        msg, level = payload
        prefix = {"ok": "[OK]  ", "warn": "[WARN]", "error": "[ERR] ",
                  "info": "      "}.get(level, "      ")
        print(f"{prefix} {msg}")
    elif kind == "status":
        print(f"...    {payload}")
    elif kind == "progress":
        pass  # too noisy for stdout
    elif kind == "done":
        ok = payload
        print(f"===== workflow {'succeeded' if ok else 'FAILED'} =====")
"""

CONFIG_CELL = """\
# ---- EDIT THIS -----------------------------------------------------------
WORKFLOW = "TO"   # "TO" or "NOTIF"
DRY_RUN = False   # True = run all SAP lookups + report match counts, but
                  # leave the workbook completely untouched
# --------------------------------------------------------------------------

# Native file-picker (same widget the .py GUI uses, via tkinter). Skip
# manual path-editing -- browse to the workbook you want to process.
import tkinter as tk
from tkinter import filedialog

_picker_root = tk.Tk()
_picker_root.withdraw()
_picker_root.attributes("-topmost", True)
try:
    EXCEL_PATH = filedialog.askopenfilename(
        title="Select the Excel workbook to process",
        filetypes=[("Excel workbooks", "*.xlsx *.xlsm *.xlsb"),
                   ("All files", "*.*")],
    )
finally:
    _picker_root.destroy()

if not EXCEL_PATH:
    raise SystemExit("No file selected -- aborting. Re-run this cell to retry.")

if WORKFLOW not in ("TO", "NOTIF"):
    raise SystemExit(f"WORKFLOW must be 'TO' or 'NOTIF', got {WORKFLOW!r}")

print(f"Workflow:    {WORKFLOW}")
print(f"Dry run:     {DRY_RUN}")
print(f"Excel file:  {EXCEL_PATH}")
"""

EXECUTE_CELL = """\
cfg = RunConfig(
    excel_path=EXCEL_PATH,
    workflow=WORKFLOW,
    stop_event=threading.Event(),
    dry_run=DRY_RUN,
)
run(cfg, on_event=print_event)
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
# esa-lookup (self-contained notebook)

**Auto-generated** from `pipeline.py` + `excel_ops.py` + `sap_ops.py` by
`build_notebook.py`. To update after a code change: edit the .py file(s),
then run `py build_notebook.py`. Do not hand-edit this notebook.

Mass-lookup SAP data into an Excel workbook via transaction **ZTBV**,
plant **ESA1**. Two workflows:

- **TO** (3 steps): K → LTAP → M ; N+O → Z50CFG_ENG_CRNT → C/D/E + A + P ;
  C → Z50CFG_ENG_VALD → F..I
- **NOTIF** (2 steps): A → Z50CFG_ENG_CRNT → C/D/E ; C → Z50CFG_ENG_VALD
  → F..I + LID to J

## Prerequisites

- Windows + **SAP GUI for Windows** + **Microsoft Excel** (desktop)
- SAP GUI Scripting enabled (Options → Accessibility & Scripting)
- Python 3.10+ with `pip install pandas openpyxl pywin32`

## How to run

1. Log into SAP GUI (any session; the script attaches to the first one).
2. Cell → Run All. The config cell shows a **Browse** dialog to pick
   the workbook; Excel will auto-open it if it is not already open.

Per-run logs land in `%LOCALAPPDATA%\\esa-lookup\\logs\\` (last 20 runs
retained).
"""),

    md("## Imports"),
    code(IMPORTS),

    md("## Excel COM helpers *(inlined from `excel_ops.py`)*"),
    code(rename_excel_ops(strip_header(read_py("excel_ops.py")))),

    md("## SAP GUI scripting helpers *(inlined from `sap_ops.py`)*"),
    code(rename_sap_ops(strip_header(read_py("sap_ops.py")))),

    md("""\
## Pipeline core *(inlined from `pipeline.py`)*

Key normalization, workflow definitions, orchestrator, and the per-run
file-log subsystem. All 10 notebook-parity fixes present.

**Gen 4:** steps chain in memory (`VirtualSheet`) and the workbook is
written once, in a single final pass -- a SAP failure or Stop leaves the
file untouched.
"""),
    code(rewrite_pipeline_refs(strip_header(read_py("pipeline.py")))),

    md("""\
## Notebook event handler

Replaces the .py app's tkinter callback with a plain `print()` so log
lines appear in cell output. Same `(kind, payload)` protocol the pipeline
emits.
"""),
    code(PRINT_EVENT),

    md("""\
---

## Configure your run

Set `WORKFLOW` below (`\"TO\"` or `\"NOTIF\"`), then run the cell. A native
Windows file-picker dialog pops up so you can browse to the Excel workbook.
Cancel the dialog to abort.

> If the dialog does not appear on top, look for it in the taskbar — some
> Jupyter front-ends push new native windows behind the browser.
"""),
    code(CONFIG_CELL),

    md("""\
## Execute

Runs the full workflow against the file you just picked. Log lines appear
below as each step progresses; the final `===== workflow succeeded =====`
banner means the workbook has been saved.
"""),
    code(EXECUTE_CELL),
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
