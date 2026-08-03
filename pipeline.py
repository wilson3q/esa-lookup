"""Orchestrator: chain SAP lookups + Excel bulk I/O for one workflow.

Gen 4: all SAP lookups run first and chain IN MEMORY (a step whose key
column was produced by an earlier step reads that step's fetched values
directly, not the workbook); the workbook is then written once, in a single
final pass. A failure or Stop during the SAP phase leaves the workbook
completely untouched.

Emits progress/log/status events through a callback so the tkinter GUI can
render them from its main thread.
"""
from __future__ import annotations

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
import pythoncom  # Fix H: worker-thread COM apartment init

import excel_ops
import sap_ops


class Cancelled(Exception):
    """Raised inside a worker step when the user has pressed Stop."""


# ---------------------------------------------------------------------------
# Key normalization (ported from the notebook -- same rules, single copy)
# ---------------------------------------------------------------------------

# Fix E: only expand scientific notation when the whole string matches it,
# not any string that happens to contain the letter 'E'. Otherwise object
# numbers like "1E2000" get mangled to "100" + trailing garbage.
_SCI_NOTATION_RE = re.compile(r"^-?\d+(\.\d+)?[eE][+-]?\d+$")

# Rows whose primary key cell equals one of these (case-insensitive) are
# not sent to SAP as filter values, mirroring the notebook's paste-side
# guard: `if v != "" and v.upper() not in {"NOT FOUND","NOTFOUND"}`.
# Note: the notebook's WRITE-BACK loop still clears their output cells (via
# ClearContents + the non-match else branch), so this refactor does too --
# skip rows are treated as non-matches for the write path, not preserved.
_SKIP_KEY_MARKERS = frozenset({"NOT FOUND", "NOTFOUND"})

# Item 2: maximum filter values pasted into one SAP query. Larger key lists
# are split into multiple navigate->paste->execute->export rounds and the
# exports concatenated before matching. 2000 keeps the multi-select dialog
# and the ALV export comfortably inside SAP's practical limits.
SAP_FILTER_CHUNK_SIZE = 2000


def _clean_cell(value) -> str:
    """Turn a raw Excel/pandas value into a stripped string, treating None,
    NaN, and pywintypes cell-error ints as empty.
    - Fix B: str(float('nan')) is 'nan' -- filter NaN before it becomes a key.
    - Fix G: #N/A / #REF! / #VALUE! come back from Excel COM as large-negative
      int codes (e.g. -2146826281). Treat those as blanks too.
    """
    if value is None:
        return ""
    if isinstance(value, float):
        # NaN check without importing math.isnan (works for float NaN)
        if value != value:
            return ""
    if isinstance(value, int) and not isinstance(value, bool) and value < -2_000_000_000:
        return ""
    return str(value).strip()


def _canonicalize_number_shape(txt: str) -> str:
    """Shared body of normalize_key / clean_numeric_for_sap."""
    for ch in (" ", "\t", chr(160), "'", ","):
        txt = txt.replace(ch, "")
    if _SCI_NOTATION_RE.match(txt):
        try:
            txt = format(Decimal(txt), "f")
        except InvalidOperation:
            pass
    if "." in txt:
        left, right = txt.split(".", 1)
        if right == "" or set(right) <= {"0"}:
            txt = left
    return txt


def normalize_key(value) -> str:
    """Aggressively normalize an Excel value so it matches SAP's key form.

    - drop None, NaN, and Excel cell-error codes
    - strip whitespace / NBSP / apostrophes / commas
    - collapse scientific notation ONLY when the whole string is sci-notation
    - drop trailing '.0' / '.00' / '.'
    - drop leading zeros (keep at least one char)

    The notebook applied different (weaker) rules per step, but every rule
    used here is applied identically to both the Excel side and the SAP side
    inside `_build_lookup`, so this can only add matches, never subtract.
    """
    txt = _clean_cell(value)
    if not txt:
        return ""
    txt = _canonicalize_number_shape(txt)
    while len(txt) > 1 and txt.startswith("0"):
        txt = txt[1:]
    return txt


def clean_numeric_for_sap(value) -> str:
    """Like normalize_key but preserves leading zeros -- used when pasting
    numeric identifiers into SAP where SAP itself will canonicalize them.
    """
    txt = _clean_cell(value)
    if not txt:
        return ""
    return _canonicalize_number_shape(txt)


def _is_skip_key_cell(value) -> bool:
    """True when this Excel key cell should NOT be pasted into SAP's filter
    dialog (blank, or one of the 'already resolved' markers). Matches the
    notebook's paste-side guard. Skipped rows are still processed by the
    write-back loop as non-matches.
    """
    txt = _clean_cell(value)
    if not txt:
        return True
    return txt.upper() in _SKIP_KEY_MARKERS


# ---------------------------------------------------------------------------
# Workflow definitions
# ---------------------------------------------------------------------------

@dataclass
class ExtraOutput:
    """A per-step output column that lives outside the main output block.

    `sap_col=None` means "no SAP source" -- the column is only cleared on a
    non-match (used to mirror the notebook's TO Step 3 behavior on column J,
    which is cleared when a row fails to match but preserved when it does).

    `preserve_on_nonmatch=True` means "on non-match, do not touch this cell"
    (mirrors TO Step 2's column A: notebook comment "Do not touch Column A
    if there is no match").

    Skip rows (blank / "NOT FOUND" key) are treated identically to non-match
    rows for the extras, matching the notebook where the write-back loop
    iterates every row and only the key check inside the loop distinguishes
    match vs. else (skip rows fall into the else branch).
    """
    excel_col: int
    sap_col: str | None = None
    preserve_on_nonmatch: bool = False


@dataclass
class LookupStep:
    name: str                     # human label for logs
    sap_table: str                # e.g. "LTAP"
    push_button_field: str        # e.g. "TO_NUMBER" -- indexes sap_ops.PUSH_BUTTONS
    key_columns: list[int]        # 1-based Excel column indices used as key
    key_joiner: str = "|"         # how multi-column keys are joined
    sap_key_columns: list[str] = field(default_factory=list)     # SAP df cols composing the match key
    sap_output_columns: list[str] = field(default_factory=list)  # SAP df cols we copy to Excel
    excel_output_columns: list[int] = field(default_factory=list)  # 1-based Excel columns to write
    extras: list[ExtraOutput] = field(default_factory=list)      # per-cell exceptions to the main block
    # If set, write header + composite match key (normalize_key of every part
    # joined by key_joiner) to this 1-based column, rows 1..last_row.
    # Notebook TO-2 does this for column P as an audit trail.
    match_key_column: int | None = None
    match_key_header: str = "Excel Match Key Used"


WORKFLOWS = {
    "TO": [
        LookupStep(
            name="LTAP -> Unloading Point",
            sap_table="LTAP",
            push_button_field="TO_NUMBER",
            key_columns=[11],                     # K
            sap_key_columns=["TANUM"],            # LTAP TO number field
            sap_output_columns=["ABLAD"],
            excel_output_columns=[13],            # M
        ),
        LookupStep(
            name="Z50CFG_ENG_CRNT (Reservation) -> QMNUM/OBJNR/DISP",
            sap_table="Z50CFG_ENG_CRNT",
            push_button_field="RSNUM",
            key_columns=[14, 15],                 # N | O
            sap_key_columns=["RSNUM", "RSPOS"],
            sap_output_columns=["OBJNR", "DISP_MATNR", "DISP_QTY"],
            excel_output_columns=[3, 4, 5],       # C, D, E
            # Notebook: A gets written on match with QMNUM but MUST NOT be
            # touched on non-match ("Do not touch Column A if there is no
            # match" -- line 874 in the original notebook). ClearContents
            # covers C..E only, so A is preserved on skip as well.
            extras=[ExtraOutput(excel_col=1, sap_col="QMNUM", preserve_on_nonmatch=True)],
            match_key_column=16,                  # P: audit-trail composite key (notebook line 851, 863)
        ),
        LookupStep(
            name="Z50CFG_ENG_VALD -> Section/Module/Description/SalesDoc",
            sap_table="Z50CFG_ENG_VALD",
            push_button_field="OBJNR",
            key_columns=[3],                      # C
            sap_key_columns=["OBJNR"],
            # Notebook TO-3 sap_data_map only carries these 4 (LID is NOT
            # populated on match). Column J is cleared in the non-match
            # branch and left alone on match -- modeled by the extras entry.
            sap_output_columns=["Z_SECTION", "Z_MODULE", "DESCRIPT", "SALES_ORDER"],
            excel_output_columns=[6, 7, 8, 9],    # F..I
            extras=[ExtraOutput(excel_col=10)],   # J: preserve on match, clear on non-match/skip
        ),
    ],
    "NOTIF": [
        LookupStep(
            name="Z50CFG_ENG_CRNT (Notification) -> OBJNR/DISP_MATNR/DISP_QTY",
            sap_table="Z50CFG_ENG_CRNT",
            push_button_field="QMNUM",
            key_columns=[1],                      # A
            sap_key_columns=["QMNUM"],
            sap_output_columns=["OBJNR", "DISP_MATNR", "DISP_QTY"],
            excel_output_columns=[3, 4, 5],       # C, D, E
        ),
        LookupStep(
            name="Z50CFG_ENG_VALD -> Section/Module/Description/SalesDoc/LID",
            sap_table="Z50CFG_ENG_VALD",
            push_button_field="OBJNR",
            key_columns=[3],                      # C
            sap_key_columns=["OBJNR"],
            # Notebook NOTIF-2 fetches LID from SAP but never writes it (bug
            # -- match branch writes only F..I). The completion message and
            # this project's README both say LID -> J, so we treat that as
            # the user's true intent and route LID through an ExtraOutput.
            # ClearContents range therefore matches notebook (F..I only).
            sap_output_columns=["Z_SECTION", "Z_MODULE", "DESCRIPT", "SALES_ORDER"],
            excel_output_columns=[6, 7, 8, 9],    # F..I
            extras=[ExtraOutput(excel_col=10, sap_col="LID", preserve_on_nonmatch=False)],
        ),
    ],
}


# ---------------------------------------------------------------------------
# Event callback protocol
# ---------------------------------------------------------------------------

# on_event(kind, payload):
#   ("log", (message: str, level: "info"|"ok"|"warn"|"error"))
#   ("status", message: str)
#   ("progress", fraction: float 0..1)
#   ("done", ok: bool)

EventFn = Callable[[str, object], None]


# ---------------------------------------------------------------------------
# Per-run file log -- persistent record so a mid-run crash can be diagnosed
# after the fact even if the GUI closed. Path is echoed into the GUI log at
# run start.
# ---------------------------------------------------------------------------
_LOG_FH = None       # file handle
_LOG_FILE = ""       # current log path


def _open_run_log() -> str:
    """Open a timestamped log file for this run. Never raises; returns the
    path (empty on failure). Also prunes older runs to keep at most 20 files.
    """
    global _LOG_FH, _LOG_FILE
    _LOG_FH = None
    _LOG_FILE = ""
    try:
        base = os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()
        log_dir = os.path.join(base, "esa-lookup", "logs")
        os.makedirs(log_dir, exist_ok=True)
        # Prune to the 20 most recent .log files.
        try:
            existing = sorted(
                (os.path.join(log_dir, n) for n in os.listdir(log_dir) if n.endswith(".log")),
                key=os.path.getmtime,
            )
            for old in existing[:-19]:
                try:
                    os.remove(old)
                except OSError:
                    pass
        except OSError:
            pass
        stamp = time.strftime("%Y%m%d-%H%M%S")
        path = os.path.join(log_dir, f"esa-lookup-{stamp}.log")
        _LOG_FH = open(path, "w", encoding="utf-8", buffering=1)
        _LOG_FILE = path
        return path
    except Exception:
        _LOG_FH = None
        _LOG_FILE = ""
        return ""


def _close_run_log() -> None:
    global _LOG_FH
    if _LOG_FH is not None:
        try:
            _LOG_FH.close()
        except Exception:
            pass
    _LOG_FH = None


def _file_log(level: str, msg: str) -> None:
    """Best-effort write to the current run's log file. Never raises."""
    if _LOG_FH is None:
        return
    try:
        _LOG_FH.write(
            f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {level.upper():5s} {msg}\n"
        )
    except Exception:
        pass


def _file_log_traceback() -> None:
    """Persist a full traceback of the current exception to the log file
    only -- the GUI shows the one-line summary, the file gets the full chain
    so a post-mortem can see what really happened."""
    if _LOG_FH is None:
        return
    try:
        _LOG_FH.write(traceback.format_exc())
        _LOG_FH.write("\n")
    except Exception:
        pass


def _log(on_event: EventFn, msg: str, level: str = "info") -> None:
    _file_log(level, msg)
    on_event("log", (msg, level))


def _status(on_event: EventFn, msg: str) -> None:
    _file_log("STAT", f"[status] {msg}")
    on_event("status", msg)


def _progress(on_event: EventFn, frac: float) -> None:
    # Progress ticks are too noisy for the file log.
    on_event("progress", max(0.0, min(1.0, frac)))


# ---------------------------------------------------------------------------
# Main pipeline
# ---------------------------------------------------------------------------

def _col_letter(idx: int) -> str:
    """1-based column index -> Excel letter (only up to ZZ, plenty here)."""
    result = ""
    n = idx
    while n > 0:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result


def _range(col_start: int, col_end: int, row_start: int, row_end: int) -> str:
    return (
        f"{_col_letter(col_start)}{row_start}:{_col_letter(col_end)}{row_end}"
    )


# Fix 3: SAP ALV exports typically use the ALV column TITLE, which is often a
# short description rather than the technical field name. Try the technical
# name first, then any known aliases. Extend per-site if needed.
SAP_COLUMN_ALIASES: dict[str, list[str]] = {
    "TANUM":       ["TANUM", "TO Number", "Transfer Order", "TrfOrd", "TrfOrdNo"],
    "ABLAD":       ["ABLAD", "Unloading Point", "UnloadPt"],
    "QMNUM":       ["QMNUM", "Notification", "Notification No", "Notification Number"],
    "RSNUM":       ["RSNUM", "Reservation", "Reservation No", "Reservation Number", "Res.Number"],
    "RSPOS":       ["RSPOS", "Item", "Item No", "Item Number", "Res.Item"],
    "OBJNR":       ["OBJNR", "Object Number", "Obj.Number", "Object No"],
    "DISP_MATNR":  ["DISP_MATNR", "Disp Material", "Disposition Material", "Disp.Material"],
    "DISP_QTY":    ["DISP_QTY", "Disp Qty", "Disposition Qty", "Disposition Quantity", "Disp.Qty"],
    "Z_SECTION":   ["Z_SECTION", "Section"],
    "Z_MODULE":    ["Z_MODULE", "Module"],
    "DESCRIPT":    ["DESCRIPT", "Description", "Descr."],
    "SALES_ORDER": ["SALES_ORDER", "Sales Order", "Sales Doc.", "Sales Doc", "Sales Doc. No."],
    "LID":         ["LID"],
}


def _resolve_column(df: pd.DataFrame, canonical: str) -> str | None:
    """Return whichever of `canonical`'s alias names is present in df, or None."""
    for name in SAP_COLUMN_ALIASES.get(canonical, [canonical]):
        if name in df.columns:
            return name
    return None


def _build_lookup(
    df: pd.DataFrame,
    key_cols: list[str],
    value_cols: list[str],
) -> tuple[dict, int, int]:
    """Return (lookup, dup_count, blank_key_count) built from df.

    Resolves each canonical SAP field name (e.g. "TANUM") to whichever alias
    (e.g. "TO Number") the ALV export actually used.

    Fix C: duplicate composite keys keep the FIRST occurrence and increment
    dup_count so the caller can warn -- overwriting silently loses data when
    a lookup table legitimately has multiple rows per key.

    Fix D: rows where ANY key part is blank are skipped, not just rows where
    ALL parts are blank. A composite of "12345|" would otherwise falsely
    match every SAP row with the same first part and a blank second.
    """
    needed = list(dict.fromkeys(key_cols + value_cols))
    resolved: dict[str, str] = {}
    missing: list[str] = []
    for c in needed:
        r = _resolve_column(df, c)
        if r is None:
            missing.append(c)
        else:
            resolved[c] = r
    if missing:
        raise RuntimeError(
            "SAP export is missing these expected columns: "
            f"{missing}\n"
            f"Columns present in the export: {list(df.columns)}\n"
            "Fix: edit the ALV layout in ZTBV so each missing field is shown "
            "(prefer 'Technical Name' as the column title) and re-save the "
            "default variant, OR add another alias to SAP_COLUMN_ALIASES in "
            "pipeline.py."
        )
    out: dict[str, dict] = {}
    dup_count = 0
    blank_key_count = 0
    for _, row in df.iterrows():
        parts = [normalize_key(row[resolved[c]]) for c in key_cols]
        if any(p == "" for p in parts):
            blank_key_count += 1
            continue
        composite = "|".join(parts)
        if composite in out:
            dup_count += 1
            continue  # keep first occurrence -- do not silently overwrite
        out[composite] = {
            c: ("" if pd.isna(row[resolved[c]]) else row[resolved[c]])
            for c in value_cols
        }
    return out, dup_count, blank_key_count


class VirtualSheet:
    """Gen 4 key-column resolver.

    Columns PRODUCED by an earlier step in this run are served from memory
    (`publish`); everything else is read from the workbook. This is what
    lets TO step 3 key off the OBJNR values step 2 just fetched, without an
    intermediate Excel write/read round-trip. A column no step has produced
    falls back to the workbook -- which also preserves the Gen 2/3 behavior
    for a skipped/no-op producer step (the consumer then sees whatever was
    already in the sheet, exactly as before).

    `last_row` mirrors Excel's End(xlUp) for in-memory columns by taking the
    last row whose value is non-empty (the Excel path wrote "" into
    non-matched cells, which End(xlUp) skips over as blanks).
    """

    def __init__(self, sheet):
        self.sheet = sheet
        self._mem: dict[int, list] = {}

    def publish(self, col: int, values: list) -> None:
        """Register `values` as the would-be content of rows 2..2+len-1."""
        self._mem[col] = values

    def is_virtual(self, col: int) -> bool:
        return col in self._mem

    def last_row(self, col: int) -> int:
        if col in self._mem:
            last = 1
            for i, v in enumerate(self._mem[col]):
                if _clean_cell(v) != "":
                    last = i + 2
            return last
        return excel_ops.last_row_in_column(self.sheet, col)

    def get_col(self, col: int, last_row: int) -> list:
        """Values for rows 2..last_row, padded with '' beyond available data
        (Excel semantics: cells below a produced/short column read as blank
        because the producing step's clear ran down to Rows.Count)."""
        n = max(0, last_row - 1)
        if col in self._mem:
            vals = list(self._mem[col][:n])
        else:
            rng = f"{_col_letter(col)}2:{_col_letter(col)}{last_row}"
            rows = excel_ops.read_range_2d(self.sheet, rng)
            vals = [(r[0] if r else "") for r in rows]
        return vals + [""] * (n - len(vals))


@dataclass
class StepResult:
    """Everything a fetched step carries into the deferred write-back."""
    step: LookupStep
    step_index: int
    last_row: int
    excel_keys: list[str]
    entries_by_row: list[dict | None]
    n_skipped: int
    matched: int
    unmatched: int


def _fetch_step(
    step: LookupStep,
    vsheet: VirtualSheet,
    xl: excel_ops.ExcelCtx,
    sap: sap_ops.SapSession,
    tmp_dir: str,
    on_event: EventFn,
    step_index: int,
    total_steps: int,
    stop: threading.Event,
) -> StepResult | None:
    """Run the SAP side of one lookup step and match it against the step's
    key column -- WITHOUT touching the workbook. Returns None when the step
    is a no-op (empty key column / no usable values).

    Gen 4: key columns written by an earlier step in this run resolve from
    memory via `vsheet`, so steps chain without an Excel round-trip.
    """

    # Progress: total_steps fetch segments + one final write segment.
    n_segments = total_steps + 1

    def sub_progress(sub_i: int, sub_n: int = 6):
        _progress(on_event, (step_index + sub_i / sub_n) / n_segments)

    step_started = time.time()
    if stop.is_set():
        raise Cancelled()

    # --- 1. Resolve key columns (memory first, Excel otherwise) ----------
    _status(on_event, f"[{step_index + 1}/{total_steps}] {step.name}: reading keys")
    _log(on_event, f"--- Step {step_index + 1}/{total_steps}: {step.name}  "
                    f"(table={step.sap_table}, key_cols={step.key_columns})")

    # Fix F: each step derives its own last_row from its own primary key
    # column (matches notebook, which re-computes last_row per step). A short
    # column later in the workflow doesn't process phantom rows from a
    # longer column earlier in the workflow.
    primary_col = step.key_columns[0]
    if vsheet.is_virtual(primary_col):
        _log(on_event,
             f"key column {_col_letter(primary_col)} resolved from an earlier "
             f"step's in-memory result (no Excel round-trip)")
    last_row = vsheet.last_row(primary_col)
    if last_row < 2:
        _log(on_event,
             f"No data below the header in column {_col_letter(primary_col)} "
             f"(#{primary_col}); step skipped", "warn")
        _progress(on_event, (step_index + 1) / n_segments)
        return None

    key_vals = {c: vsheet.get_col(c, last_row) for c in step.key_columns}
    n_rows = last_row - 1
    sub_progress(1)

    # excel_keys is the normalized composite for every row -- used for both
    # lookup matching and (when match_key_column is set) the P-column audit
    # write. Notebook builds the identical string via `normalize_key(...)|
    # normalize_key(...)` inside its per-row loop. Parts are joined in
    # ascending column order (matches the Gen 2/3 offset-sorted read).
    ordered_cols = sorted(step.key_columns)
    excel_keys = [
        step.key_joiner.join(normalize_key(key_vals[c][i]) for c in ordered_cols)
        for i in range(n_rows)
    ]
    # Skip flags: only used to (a) exclude from SAP paste, (b) count for the
    # log summary. The write-back treats skip rows as non-matches, which
    # is what the notebook does implicitly (its else branch fires for any
    # composite key that's not in the SAP result set, including "NOTFOUND|").
    skip_flags = [_is_skip_key_cell(key_vals[primary_col][i]) for i in range(n_rows)]

    unique_paste_values: list[str] = []
    seen: set[str] = set()
    for i in range(n_rows):
        if skip_flags[i]:
            continue
        # For paste, use the FIRST key column value (SAP filter dialog is
        # single-column). Multi-column keys still work because SAP returns
        # the full result set and we filter by composite key on our side.
        v = clean_numeric_for_sap(key_vals[primary_col][i])
        if v and v not in seen:
            seen.add(v)
            unique_paste_values.append(v)

    n_skipped = sum(skip_flags)
    if n_skipped:
        _log(on_event,
             f"note: {n_skipped} row(s) in column {_col_letter(primary_col)} "
             f"are blank or 'NOT FOUND' -- not sent to SAP; their output "
             f"cells will be cleared just like any other non-match")

    if not unique_paste_values:
        _log(on_event,
             f"No usable values in column(s) {step.key_columns} "
             f"(after skipping blanks / 'NOT FOUND'); step is a no-op",
             "warn")
        _progress(on_event, (step_index + 1) / n_segments)
        return None

    _log(on_event, f"{len(unique_paste_values)} unique key(s) will be sent to SAP")
    sub_progress(2)

    # --- 2. Query SAP in chunks: navigate, paste, execute, verify, export
    # Item 2: large key lists are split so a single oversized multi-select
    # paste can't overload the selection screen or produce an unmanageable
    # export. Each chunk is a full navigate->paste->execute->export round;
    # chunk results are concatenated before matching.
    chunks = [
        unique_paste_values[i:i + SAP_FILTER_CHUNK_SIZE]
        for i in range(0, len(unique_paste_values), SAP_FILTER_CHUNK_SIZE)
    ]
    if len(chunks) > 1:
        _log(on_event,
             f"splitting {len(unique_paste_values)} key(s) into {len(chunks)} "
             f"chunk(s) of <= {SAP_FILTER_CHUNK_SIZE}")

    push_id = sap_ops.PUSH_BUTTONS[(step.sap_table, step.push_button_field)]
    frames: list[pd.DataFrame] = []
    ts = int(time.time())
    for ci, chunk in enumerate(chunks):
        if stop.is_set():
            raise Cancelled()
        tag = f" (chunk {ci + 1}/{len(chunks)})" if len(chunks) > 1 else ""

        _status(on_event, f"[{step_index + 1}/{total_steps}] SAP: loading {step.sap_table}{tag}")
        sap_ops.open_ztbv_table(sap, step.sap_table, log=lambda m: _log(on_event, m))
        _status(on_event, f"[{step_index + 1}/{total_steps}] SAP: sending {len(chunk)} filter values{tag}")
        try:
            # Item 3: primary transport is a temp text file ('Import from
            # Text File' in the multi-select dialog) -- deterministic even
            # if the user copies something to the clipboard mid-run.
            sap_ops.fill_multi_value_filter_from_file(
                sap, push_id, chunk, tmp_dir, log=lambda m: _log(on_event, m))
        except sap_ops.SapError as e:
            # Fallback: the Gen 2/3 clipboard path (values staged via an
            # Excel scratch workbook). Kept for SAP GUI versions where the
            # import dialog is not scriptable.
            _log(on_event,
                 f"file import unavailable ({e}); falling back to "
                 f"clipboard paste", "warn")
            scratch = excel_ops.stage_values_on_clipboard(xl.app, chunk)
            try:
                sap_ops.paste_multi_value_filter(
                    sap, push_id, chunk, log=lambda m: _log(on_event, m))
            finally:
                excel_ops.close_scratch(scratch)
        _status(on_event, f"[{step_index + 1}/{total_steps}] SAP: executing query{tag}")
        sap_ops.execute_query(sap, log=lambda m: _log(on_event, m))

        # Item 2: read the status bar + confirm a result grid exists. An SAP
        # error message or missing grid fails HERE with SAP's own words
        # instead of a confusing export failure two calls later.
        grid_rows = sap_ops.query_result_check(sap, log=lambda m: _log(on_event, m))
        if grid_rows == 0:
            _log(on_event, f"SAP returned 0 rows{tag}; nothing to export", "warn")
            sub_progress(2 + 4 * (ci + 1) / len(chunks))
            continue

        _status(on_event,
                f"[{step_index + 1}/{total_steps}] SAP: exporting ALV grid "
                f"({grid_rows} rows){tag}")
        export_name = f"{step.sap_table}_{step.push_button_field}_{ts}_{ci}.xlsx"
        export_start = time.time()
        export_path = sap_ops.export_alv_to_file(
            sap, tmp_dir, export_name, log=lambda m: _log(on_event, m)
        )
        try:
            size_kb = os.path.getsize(export_path) / 1024.0
        except OSError:
            size_kb = 0.0
        _log(on_event, f"export finished in {time.time() - export_start:.1f}s "
                        f"({size_kb:.0f} KB)")

        df_chunk = pd.read_excel(export_path)
        # Item 2: verify the file matches what the grid showed. Fewer rows
        # than the grid = truncated export = silent data loss downstream.
        if len(df_chunk) < grid_rows:
            raise RuntimeError(
                f"Export row-count mismatch{tag}: the SAP grid shows "
                f"{grid_rows} row(s) but the export file contains "
                f"{len(df_chunk)} -- the export is incomplete. Re-run; if it "
                f"persists, export once manually from ZTBV and compare with "
                f"{export_path}.")
        if len(df_chunk) > grid_rows:
            _log(on_event,
                 f"note: export has {len(df_chunk)} row(s) vs {grid_rows} in "
                 f"the grid (totals/subtotal lines?); harmless unless the "
                 f"extra rows carry key values", "warn")
        frames.append(df_chunk)
        sub_progress(2 + 4 * (ci + 1) / len(chunks))

    # --- 3. Combine chunk results, build lookup dict ---------------------
    _status(on_event, f"[{step_index + 1}/{total_steps}] loading SAP export")
    extra_sap_cols = [e.sap_col for e in step.extras if e.sap_col]
    if not frames:
        # Every chunk came back empty: a legitimate "no matches" outcome.
        # (Previously this crashed in _build_lookup with a missing-columns
        # error, because the export of an empty grid carries no data.)
        _log(on_event,
             "SAP returned no rows for any of the requested keys -- every "
             "row will be treated as unmatched", "warn")
        df = None
        lookup, dup_count, blank_key_count = {}, 0, 0
    else:
        df = frames[0] if len(frames) == 1 else pd.concat(frames, ignore_index=True)
        _log(on_event, f"SAP returned {len(df)} row(s) with columns {list(df.columns)[:8]}{'...' if len(df.columns) > 8 else ''}")
        lookup, dup_count, blank_key_count = _build_lookup(
            df,
            step.sap_key_columns,
            step.sap_output_columns + extra_sap_cols,
        )
    if dup_count:
        _log(on_event,
             f"WARNING: SAP returned {dup_count} duplicate key(s); only the "
             f"FIRST row per key is used. Check the run log file (path echoed "
             f"above) for the raw SAP columns, or tighten your ALV filter.",
             "warn")
    if blank_key_count:
        _log(on_event,
             f"note: skipped {blank_key_count} SAP row(s) with blank/partial "
             f"key columns", "info")

    # --- 5. Resolve each row's fate: matched vs. else --------------------
    # Skip rows use the same else branch as SAP non-matches -- the notebook
    # write loop makes no distinction (both fall through to the same clear
    # branch). See _SKIP_KEY_MARKERS docstring.
    _status(on_event, f"[{step_index + 1}/{total_steps}] matching keys in memory")
    matched = 0
    entries_by_row: list[dict | None] = []
    for k, skipped in zip(excel_keys, skip_flags):
        if skipped:
            entries_by_row.append(None)
            continue
        entry = lookup.get(k)
        entries_by_row.append(entry)
        if entry is not None:
            matched += 1

    # --- 6. Publish outputs for downstream steps (Gen 4) -----------------
    # Later steps key off these values from memory instead of re-reading an
    # Excel write-back. Only main-block outputs are published -- no current
    # workflow keys off an extras column (extras also depend on EXISTING
    # cell content, which only the write-back phase reads).
    for j, sap_c in enumerate(step.sap_output_columns):
        col_values = []
        for entry in entries_by_row:
            v = entry[sap_c] if entry is not None else ""
            col_values.append("" if v is None else v)
        vsheet.publish(step.excel_output_columns[j], col_values)

    non_skip = n_rows - n_skipped
    unmatched = non_skip - matched
    # Sanity: if we sent >20 unique keys and SAP came back with <10% as many
    # rows as we asked about, something is off -- surface a warning so the
    # user does not silently accept mostly-empty output.
    if len(unique_paste_values) > 20 and len(lookup) < max(1, len(unique_paste_values) // 10):
        _log(on_event,
             f"WARNING: sent {len(unique_paste_values)} unique key(s) to SAP but "
             f"only {len(lookup)} matched. Common causes: (a) query returned an "
             f"error screen, (b) the ALV filter rejected the values, (c) plant "
             f"or table wrong for this environment.", "warn")
    _log(on_event,
         f"Step {step_index + 1} fetched in {time.time() - step_started:.1f}s: "
         f"{matched}/{non_skip} rows matched, {unmatched} unmatched"
         + (f", {n_skipped} skipped" if n_skipped else "")
         + " (write deferred to final pass)",
         "ok")
    _progress(on_event, (step_index + 1) / n_segments)
    return StepResult(
        step=step,
        step_index=step_index,
        last_row=last_row,
        excel_keys=excel_keys,
        entries_by_row=entries_by_row,
        n_skipped=n_skipped,
        matched=matched,
        unmatched=unmatched,
    )


def _write_back(
    results: list[StepResult],
    xl: excel_ops.ExcelCtx,
    on_event: EventFn,
    total_steps: int,
) -> None:
    """Apply every fetched step's writes to the workbook in ONE pass.

    Gen 4: nothing reaches the workbook unless every step fetched
    successfully -- a SAP failure or user Stop during the fetch phase leaves
    the file completely untouched. Per-step write semantics (clear-to-bottom,
    text formats, extras preserve/clear rules, audit match-key column) are
    unchanged from Gen 2/3; only WHEN they run has moved. Step column sets
    are disjoint in both workflows, so applying them sequentially here is
    identical to the old interleaved order.
    """
    n_segments = total_steps + 1
    _status(on_event, "writing all step results to Excel (single pass)")
    sheet = xl.sheet
    rows_count = int(sheet.Rows.Count)

    with excel_ops.bulk_write(xl.app):
        for done, res in enumerate(results):
            step = res.step
            last_row = res.last_row
            excel_keys = res.excel_keys
            entries_by_row = res.entries_by_row
            n_rows = len(excel_keys)

            # -------- Main output block ---------------------------------
            oc_min = min(step.excel_output_columns)
            oc_max = max(step.excel_output_columns)
            oc_span = oc_max - oc_min + 1
            target_write = _range(oc_min, oc_max, 2, last_row)
            # Match notebook: clear the ENTIRE column range down to
            # Rows.Count so stale rows below last_row (from a prior, longer
            # run) are wiped.
            target_clear = _range(oc_min, oc_max, 2, rows_count)

            main_out: list[list] = []
            for i in range(n_rows):
                entry = entries_by_row[i]
                row_vals = ["" for _ in range(oc_span)]
                if entry is not None:
                    for j, sap_c in enumerate(step.sap_output_columns):
                        excel_c = step.excel_output_columns[j]
                        offset = excel_c - oc_min
                        val = entry[sap_c]
                        row_vals[offset] = "" if val is None else val
                main_out.append(row_vals)

            excel_ops.clear_range(sheet, target_clear)
            excel_ops.set_column_format_text(
                sheet, f"{_col_letter(oc_min)}:{_col_letter(oc_max)}"
            )
            excel_ops.write_range_2d(sheet, target_write, main_out)

            # -------- Extras (per-column, with per-column semantics) ----
            for extra in step.extras:
                letter = _col_letter(extra.excel_col)
                xrange = f"{letter}2:{letter}{last_row}"

                # Read existing only when at least one row will preserve it.
                # Safe to read here (not in the fetch phase): no step writes
                # another step's extras column, so the pre-run content is
                # still intact at this point.
                need_existing = extra.preserve_on_nonmatch or extra.sap_col is None
                existing_extra = (
                    excel_ops.read_range_2d(sheet, xrange) if need_existing else None
                )

                col_data: list[list] = []
                for i in range(n_rows):
                    existing_val = (
                        existing_extra[i][0]
                        if (existing_extra and i < len(existing_extra)
                            and existing_extra[i])
                        else ""
                    )
                    entry = entries_by_row[i]
                    if entry is not None:
                        # Matched row.
                        if extra.sap_col is None:
                            # No SAP source -> preserve on match
                            # (mirrors notebook TO-3 col J behavior).
                            col_data.append([existing_val])
                        else:
                            v = entry[extra.sap_col]
                            col_data.append(["" if v is None else v])
                    else:
                        # Non-match OR skip (treated the same, per notebook).
                        if extra.preserve_on_nonmatch:
                            col_data.append([existing_val])
                        else:
                            col_data.append([""])

                # Only set text format on columns where we're writing SAP
                # data; a preserve-only column (sap_col=None) should keep
                # whatever NumberFormat the user had -- notebook doesn't
                # touch it.
                if extra.sap_col is not None:
                    excel_ops.set_column_format_text(sheet, f"{letter}:{letter}")
                excel_ops.write_range_2d(sheet, xrange, col_data)

            # -------- Match key column (P for TO-2) ----------------------
            if step.match_key_column is not None:
                mk_letter = _col_letter(step.match_key_column)
                # Header at row 1 + text format for the whole column,
                # matching notebook lines 851-852.
                sheet.Cells(1, step.match_key_column).Value = step.match_key_header
                excel_ops.set_column_format_text(sheet, f"{mk_letter}:{mk_letter}")
                # Notebook line 863 writes the composite key to P for every
                # row in the loop (including skip rows, which get
                # "NOTFOUND|..." or "|" written). Match that by writing
                # excel_keys[i] for every row 2..last_row.
                mk_range = f"{mk_letter}2:{mk_letter}{last_row}"
                mk_data = [[k] for k in excel_keys]
                excel_ops.write_range_2d(sheet, mk_range, mk_data)

            _progress(
                on_event,
                (total_steps + (done + 1) / max(1, len(results))) / n_segments,
            )

    _log(on_event,
         f"write-back complete: {len(results)} step block(s) written in one pass",
         "ok")


def _log_write_state(on_event: EventFn, write_state: str) -> None:
    """After a cancel/failure, tell the user exactly what state the file is
    in. Gen 4 defers all writes, so the common case is 'untouched'."""
    if write_state == "none":
        _log(on_event,
             "the workbook was NOT modified (all writes are deferred until "
             "every SAP step succeeds)", "info")
    elif write_state == "partial":
        _log(on_event,
             "WARNING: failure happened DURING the final write-back -- the "
             "workbook may be partially updated and has not been saved. "
             "Review it before saving manually.", "warn")


def _dry_run_report(results: list[StepResult], on_event: EventFn) -> None:
    """Item 4: summarize what the write-back WOULD do, without doing it."""
    _log(on_event, "DRY RUN -- the workbook will not be modified. Planned writes:")
    for res in results:
        step = res.step
        cols = [_col_letter(c) for c in step.excel_output_columns]
        cols += [_col_letter(e.excel_col) for e in step.extras]
        if step.match_key_column is not None:
            cols.append(_col_letter(step.match_key_column))
        _log(on_event,
             f"  step {res.step_index + 1} ({step.sap_table}): "
             f"{res.matched} matched / {res.unmatched} unmatched"
             + (f" / {res.n_skipped} skipped" if res.n_skipped else "")
             + f" -> columns {'/'.join(cols)}, rows 2..{res.last_row}")
        shown = 0
        for i, entry in enumerate(res.entries_by_row):
            if entry is None:
                continue
            preview = ", ".join(f"{k}={entry[k]}" for k in entry)
            _log(on_event, f"    e.g. row {i + 2}: {preview}")
            shown += 1
            if shown >= 3:
                break


@dataclass
class RunConfig:
    excel_path: str
    workflow: str            # "TO" or "NOTIF"
    stop_event: threading.Event
    dry_run: bool = False    # Item 4: fetch + report matches, write nothing


def run(cfg: RunConfig, on_event: EventFn) -> bool:
    """Entry point invoked on a background thread. Returns True on success."""
    # Fix H: initialize this worker thread's COM apartment. pywin32 does an
    # implicit CoInitialize on Dispatch, but the second Run click (new
    # worker thread, same process) can hit CO_E_NOTINITIALIZED on
    # GetActiveObject / GetObject("SAPGUI") without an explicit init here.
    # Uninitialize in finally so the thread exits clean.
    pythoncom.CoInitialize()
    run_started = time.time()
    log_path = _open_run_log()
    # Gen 4 write state, used by the error/cancel handlers to tell the user
    # exactly what happened to the file: "none" -> nothing written (fetch
    # phase), "partial" -> failure mid write-back, "done" -> fully written.
    write_state = "none"
    try:
        _log(on_event, f"esa-lookup starting workflow '{cfg.workflow}'", "info")
        if log_path:
            _log(on_event, f"detailed log file: {log_path}", "info")
        # Diagnostic: freeze the environment into the file so a post-mortem
        # can tell which Python / pandas / openpyxl the run used.
        try:
            import openpyxl as _openpyxl
            openpyxl_v = _openpyxl.__version__
        except Exception:
            openpyxl_v = "?"
        _file_log("info",
                  f"env: Python {sys.version.split()[0]} on "
                  f"{sys.platform}, pandas {pd.__version__}, openpyxl {openpyxl_v}, "
                  f"cwd={os.getcwd()}, excel={cfg.excel_path}")

        _status(on_event, "opening Excel")
        xl = excel_ops.attach(cfg.excel_path)
        _log(on_event, f"attached to Excel: {os.path.basename(cfg.excel_path)}", "ok")

        _status(on_event, "attaching to SAP GUI")
        sap = sap_ops.attach()
        _log(on_event, "attached to SAP GUI session", "ok")

        steps = WORKFLOWS[cfg.workflow]
        total_steps = len(steps)
        tmp_dir = os.path.join(tempfile.gettempdir(), "esa_lookup")

        # Pre-flight: verify the workflow's primary input column has data
        # before we even spin up SAP navigation. Later steps derive their
        # own last_row (Fix F) and skip themselves if their column is empty.
        primary_col = steps[0].key_columns[0]
        primary_last = excel_ops.last_row_in_column(xl.sheet, primary_col)
        if primary_last < 2:
            _log(on_event,
                 f"No data below the header in column "
                 f"{_col_letter(primary_col)} (#{primary_col}) -- the "
                 f"workflow's primary input. Aborting.", "error")
            on_event("done", False)
            return False
        _log(on_event,
             f"step 1 will process rows 2..{primary_last} (column "
             f"{_col_letter(primary_col)}); each subsequent step derives "
             f"its own row range from its own key column")
        _log(on_event,
             "Gen 4: steps chain in memory; the workbook is written once, "
             "after every SAP step has succeeded")

        # ---- Phase 1: fetch every step from SAP (no workbook writes) ----
        vsheet = VirtualSheet(xl.sheet)
        results: list[StepResult] = []
        totals_matched = 0
        totals_seen = 0
        for i, step in enumerate(steps):
            if cfg.stop_event.is_set():
                raise Cancelled()
            res = _fetch_step(
                step, vsheet, xl, sap, tmp_dir, on_event, i, total_steps,
                cfg.stop_event,
            )
            if res is None:
                continue  # no-op step (empty key column)
            results.append(res)
            totals_matched += res.matched
            totals_seen += res.matched + res.unmatched

        if cfg.stop_event.is_set():
            raise Cancelled()

        # ---- Dry run: report what would be written, touch nothing -------
        if cfg.dry_run:
            _dry_run_report(results, on_event)
            _status(on_event, "dry run complete -- nothing written")
            _log(on_event,
                 f"DRY RUN complete: {totals_matched}/{totals_seen} "
                 f"row-matches across {total_steps} step(s); the workbook "
                 f"was not modified", "ok")
            _progress(on_event, 1.0)
            on_event("done", True)
            return True

        # ---- Phase 2: single write-back pass + save ---------------------
        write_state = "partial"
        _write_back(results, xl, on_event, total_steps)
        write_state = "done"

        # Fix L: save() now raises ExcelError on failure; warn the user
        # rather than silently pretending the write persisted.
        try:
            excel_ops.save(xl.book)
        except excel_ops.ExcelError as e:
            _log(on_event, f"WARNING: {e}", "warn")

        _status(on_event, "done")
        _log(on_event, f"all steps complete: {totals_matched}/{totals_seen} row-matches across {total_steps} step(s)", "ok")
        _progress(on_event, 1.0)
        on_event("done", True)
        return True

    except Cancelled:
        _log(on_event, "cancelled by user", "warn")
        _log_write_state(on_event, write_state)
        _status(on_event, "cancelled")
        on_event("done", False)
        return False
    except sap_ops.SapError as e:
        _log(on_event, f"SAP error: {e}", "error")
        _file_log_traceback()  # full chain into the log file for post-mortem
        _log_write_state(on_event, write_state)
        _status(on_event, "SAP error")
        on_event("done", False)
        return False
    except excel_ops.ExcelError as e:
        _log(on_event, f"Excel error: {e}", "error")
        _file_log_traceback()
        _log_write_state(on_event, write_state)
        _status(on_event, "Excel error")
        on_event("done", False)
        return False
    except Exception as e:
        _log(on_event, f"unexpected error: {e}", "error")
        _log(on_event, traceback.format_exc(), "error")
        _file_log_traceback()
        _log_write_state(on_event, write_state)
        _status(on_event, "failed")
        on_event("done", False)
        return False
    finally:
        _log(on_event, f"total elapsed: {time.time() - run_started:.1f}s", "info")
        _close_run_log()
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
