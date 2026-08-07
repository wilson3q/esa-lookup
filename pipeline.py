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
    """Turn a raw Excel value into a stripped string, treating None,
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
    # Optional per-value transform applied to the fetched SAP value before
    # writing (and before publishing for downstream steps). Used to derive
    # the reservation pair out of the unloading point -- the job the Excel
    # template used to do with a formula.
    transform: Callable[[str], str] | None = None


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


# The ESA unloading point (LTAP ABLAD) encodes the reservation pair.
# The Excel template split it with formulas in N and O (verbatim, from the
# ESA operator, 2026-08-07):
#
#     N: =IFERROR(VALUE(MID(M12,2,9)),"")     chars 2..10 -> number
#     O: =IFERROR(VALUE(MID(M4,11,4)),"")     chars 11..14 -> number
#
# i.e. a 14-char value '05175455500001' -> N 517545550, O 1. The app
# cannot rely on that formula: all writes are deferred to the end of the
# run, so step 2 would read N/O before M exists in the sheet -- and the
# formula is easily lost when a sheet is copied (which is how this
# surfaced). The split lives here instead.
#
# One deliberate divergence: the formula ALWAYS discards character 1;
# this code keeps it and strips leading zeros. Identical output for every
# value starting with '0' (all observed data), but SAP's RSNUM is 10
# digits -- if one ever uses all 10, the formula silently truncates it
# while this code still matches. Anything that does not look like the
# encoded form (blank, 'NOT FOUND', a plain dock name) yields "", which
# step 2 then skips as a row without a reservation.

# The scanned transfer order (column B) encodes the TO pair the same way.
# Template formulas (verbatim, from the ESA operator, 2026-08-07):
#
#     K (TO Number):      =LEFT(B14,10)        chars 1..10
#     L (TO Item Number): =MID(B14,11,4)       chars 11..14, keeps zeros
#
# '10098436690001' -> K '1009843669', L '0001'. Same fragility as the N/O
# formulas (lost when a sheet is copied), so the split lives here too --
# applied at run start, BEFORE step 1 keys on K. Unlike the formula, an
# existing K/L value is kept when B is blank, so hand-typed TO numbers
# still work on rows that were never scanned.

def _to_number_from_scan(v) -> str:
    s = clean_numeric_for_sap(v)
    return s[:10] if s else ""


def _to_item_from_scan(v) -> str:
    s = clean_numeric_for_sap(v)
    return s[10:14] if s else ""


@dataclass
class SheetDerivation:
    """A template formula moved into the pipeline: target columns computed
    from a source column at run start, published for the steps to key on,
    and written to the sheet in the final pass (so the audit view the
    operator knows is preserved). Per row, a target keeps its existing
    sheet value when the derivation yields "" -- the derivations only ever
    ADD information."""
    source_col: int                                   # 1-based Excel column
    targets: list[tuple[int, Callable[[str], str]]]   # (excel_col, fn)


WORKFLOW_DERIVATIONS: dict[str, list[SheetDerivation]] = {
    "TO": [SheetDerivation(
        source_col=2,                                 # B: scanned TO
        targets=[(11, _to_number_from_scan),          # K
                 (12, _to_item_from_scan)],           # L
    )],
    "NOTIF": [],
}


def _reservation_from_unloading_point(v) -> str:
    s = str(v or "").strip()
    if len(s) < 5 or not s.isdigit():
        return ""
    return s[:-4].lstrip("0") or "0"


def _item_from_unloading_point(v) -> str:
    s = str(v or "").strip()
    if len(s) < 5 or not s.isdigit():
        return ""
    return s[-4:].lstrip("0") or "0"


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
            extras=[
                # The template formula, moved into the pipeline: N and O
                # are split out of the unloading point so step 2 can key on
                # them (in memory -- no Excel round trip) and the sheet
                # still shows them for auditing. Non-match rows get ""
                # like every other output.
                ExtraOutput(excel_col=14, sap_col="ABLAD",
                            transform=_reservation_from_unloading_point),
                ExtraOutput(excel_col=15, sap_col="ABLAD",
                            transform=_item_from_unloading_point),
            ],
        ),
        LookupStep(
            # Keyed on the reservation pair, exactly like the ESA
            # developer's own notebook (his code is the spec, 2026-08). A
            # 2026-08 rework briefly keyed this on the TO number instead
            # (905147d); reverted -- the reservation path is the one proven
            # in the field, N/O are operator-provided inputs in the live
            # workbook, and preserving A on non-match is what lets the
            # NOTIF process later fill the notification-only rows (the
            # either/or model: most rows have a notification but no TO).
            name="Z50CFG_ENG_CRNT (Reservation) -> QMNUM/OBJNR/DISP",
            sap_table="Z50CFG_ENG_CRNT",
            push_button_field="RSNUM",
            key_columns=[14, 15],                 # N | O
            sap_key_columns=["RSNUM", "RSPOS"],
            sap_output_columns=["OBJNR", "DISP_MATNR", "DISP_QTY"],
            excel_output_columns=[3, 4, 5],       # C, D, E
            extras=[
                # Notebook rule: A gets QMNUM on match but MUST NOT be
                # touched on non-match ("Do not touch Column A if there is
                # no match") -- the untouched rows carry the pre-existing
                # notification numbers the NOTIF process keys on.
                ExtraOutput(excel_col=1, sap_col="QMNUM", preserve_on_nonmatch=True),
            ],
            match_key_column=16,                  # P: audit-trail composite key
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
#   ("popup", {"kind": "step"|"info"|"error", "title": str, "message": str,
#              "ack": threading.Event|None, "result": {"proceed": bool}|None})
#     The original notebook spoke to its operator through blocking message
#     boxes -- a counts popup after every step, an error popup on failure --
#     and the ESA operator troubleshoots by screenshotting those together
#     with the SAP screen. These events reproduce that behavior. "step"
#     popups carry an ack Event: the handler MUST show OK/Cancel, write the
#     choice into result["proceed"], and set ack. "info"/"error" popups are
#     fire-and-forget.
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


def _popup(on_event: EventFn, kind: str, title: str, message: str,
           ack=None, result=None) -> None:
    _file_log("POPUP", f"[{kind}] {title}: " + message.replace("\n", " | "))
    on_event("popup", {"kind": kind, "title": title, "message": message,
                       "ack": ack, "result": result})


def _step_popup(on_event: EventFn, stop: threading.Event,
                title: str, message: str) -> None:
    """Blocking per-step popup, the original notebook's rhythm: show the
    counts, wait for the operator's OK before touching SAP again -- so they
    can inspect the result grid still on screen -- or Cancel to stop the
    run (the workbook is untouched during the fetch phase).

    The wait polls so a Stop pressed through other means still interrupts.
    If the event handler never acks (headless caller that ignores popups),
    the run would wait forever -- which is why step popups are opt-in via
    RunConfig.step_popups and every shipped handler acks.
    """
    ack = threading.Event()
    result = {"proceed": True}
    _popup(on_event, "step", title, message, ack=ack, result=result)
    while not ack.wait(0.1):
        if stop.is_set():
            raise Cancelled()
    if not result["proceed"]:
        stop.set()
        raise Cancelled()


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


def _build_lookup(
    records: list[dict],
    key_cols: list[str],
    value_cols: list[str],
) -> tuple[dict, int, int]:
    """Turn the grid rows into {composite_key: {value_col: value}}.

    Rows come straight off the ALV grid by TECHNICAL field name, so the
    keys of each record dict are exactly the names in key_cols/value_cols.

    Returns (lookup, dup_count, blank_key_count):
    - duplicate composite keys keep the FIRST occurrence (dup_count tells
      the caller to warn -- silently overwriting would lose data);
    - rows where ANY key part is blank are skipped (a composite "12345|"
      would falsely match every row sharing the first part).
    """
    out: dict[str, dict] = {}
    dup_count = 0
    blank_key_count = 0
    for row in records:
        parts = [normalize_key(row.get(c)) for c in key_cols]
        if any(p == "" for p in parts):
            blank_key_count += 1
            continue
        composite = "|".join(parts)
        if composite in out:
            dup_count += 1
            continue  # keep first occurrence -- do not silently overwrite
        out[composite] = {c: ("" if row.get(c) is None else row.get(c))
                          for c in value_cols}
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
    sap_rows: int = 0     # total rows SAP's grid reported across all chunks


def _fetch_step(
    step: LookupStep,
    vsheet: VirtualSheet,
    xl: excel_ops.ExcelCtx,
    sap: sap_ops.SapSession,
    on_event: EventFn,
    step_index: int,
    total_steps: int,
    stop: threading.Event,
    step_popups: bool = False,
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
        if step_popups:
            _step_popup(
                on_event, stop,
                f"Step {step_index + 1} of {total_steps} skipped",
                f"{step.name}\n\n"
                f"No data below the header in column "
                f"{_col_letter(primary_col)} -- nothing to look up.\n\n"
                f"OK = continue with the next step, Cancel = stop the run.")
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
    # ESA rule: a row is processed only when EVERY part of its key has a
    # value. A sheet legitimately mixes rows that carry a reservation with
    # rows that do not -- the complete ones must still be processed, and the
    # incomplete ones are skipped rather than failing anything. Previously
    # only the PRIMARY cell was tested, so "N filled, O blank" was sent to
    # SAP and then reported as unmatched.
    skip_flags = [
        any(_is_skip_key_cell(key_vals[c][i]) for c in ordered_cols)
        for i in range(n_rows)
    ]

    # Say how many rows were dropped for a PARTIAL key specifically -- those
    # look like usable rows to the operator, so silently folding them into
    # the skip count would hide a half-filled sheet.
    if len(ordered_cols) > 1:
        partial = [
            i for i in range(n_rows)
            if skip_flags[i] and not _is_skip_key_cell(key_vals[primary_col][i])
        ]
        if partial:
            blank_cols = sorted({
                c for c in ordered_cols for i in partial
                if _is_skip_key_cell(key_vals[c][i])
            })
            _log(on_event,
                 f"note: {len(partial)} row(s) have column "
                 f"{_col_letter(primary_col)} filled but column(s) "
                 f"{', '.join(_col_letter(c) for c in blank_cols)} blank -- "
                 f"skipped, since a complete key needs every one of "
                 f"{[_col_letter(c) for c in ordered_cols]}. Rows with a "
                 f"complete key are unaffected.")

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
        if step_popups:
            _step_popup(
                on_event, stop,
                f"Step {step_index + 1} of {total_steps} skipped",
                f"{step.name}\n\n"
                f"Column(s) "
                f"{'/'.join(_col_letter(c) for c in step.key_columns)} hold "
                f"no usable keys (all blank or 'NOT FOUND') -- nothing was "
                f"sent to SAP.\n\n"
                f"OK = continue with the next step, Cancel = stop the run.")
        return None

    _log(on_event, f"{len(unique_paste_values)} unique key(s) will be sent to SAP")
    sub_progress(2)

    # --- 2. Query SAP in chunks: navigate, paste, execute, read grid ----
    # Large key lists are split so one oversized multi-select paste cannot
    # overload the selection screen. Each chunk is a full navigate ->
    # paste -> execute -> read round; results are combined before matching.
    chunks = [
        unique_paste_values[i:i + SAP_FILTER_CHUNK_SIZE]
        for i in range(0, len(unique_paste_values), SAP_FILTER_CHUNK_SIZE)
    ]
    if len(chunks) > 1:
        _log(on_event,
             f"splitting {len(unique_paste_values)} key(s) into {len(chunks)} "
             f"chunk(s) of <= {SAP_FILTER_CHUNK_SIZE}")

    push_id = sap_ops.resolve_push_button(
        sap, step.sap_table, step.push_button_field)
    # Every SAP field this step needs, by technical name -- what the grid
    # reader asks for, exactly as the original notebook read them.
    wanted_cols = list(dict.fromkeys(
        step.sap_key_columns
        + step.sap_output_columns
        + [e.sap_col for e in step.extras if e.sap_col]
    ))
    records: list[dict] = []
    sap_rows_total = 0
    for ci, chunk in enumerate(chunks):
        if stop.is_set():
            raise Cancelled()
        tag = f" (chunk {ci + 1}/{len(chunks)})" if len(chunks) > 1 else ""

        _status(on_event, f"[{step_index + 1}/{total_steps}] SAP: loading {step.sap_table}{tag}")
        sap_ops.open_ztbv_table(sap, step.sap_table, log=lambda m: _log(on_event, m))
        _status(on_event, f"[{step_index + 1}/{total_steps}] SAP: sending {len(chunk)} filter values{tag}")
        # The original notebook's transport: stage the keys in a scratch
        # workbook, Copy(), upload from clipboard in the multi-select
        # dialog (after clearing leftovers from a previous run).
        scratch = excel_ops.stage_values_on_clipboard(xl.app, chunk)
        try:
            sap_ops.paste_multi_value_filter(
                sap, push_id, chunk, log=lambda m: _log(on_event, m))
        finally:
            excel_ops.close_scratch(scratch)
        _status(on_event, f"[{step_index + 1}/{total_steps}] SAP: executing query{tag}")
        sap_ops.execute_query(sap, log=lambda m: _log(on_event, m))

        # Read the status bar + confirm a result grid exists: a bad query
        # fails HERE with SAP's own words.
        grid_rows = sap_ops.query_result_check(sap, log=lambda m: _log(on_event, m))
        sap_rows_total += grid_rows
        if grid_rows == 0:
            _log(on_event, f"SAP returned 0 rows{tag}; nothing to read", "warn")
            sub_progress(2 + 4 * (ci + 1) / len(chunks))
            continue

        # Read the grid by technical field name -- the original notebook's
        # method. Technical names are unique per field, so there is no
        # column-title ambiguity, and no export file to parse.
        _status(on_event,
                f"[{step_index + 1}/{total_steps}] SAP: reading ALV grid "
                f"({grid_rows} rows){tag}")
        read_start = time.time()
        chunk_records = sap_ops.read_alv_grid(
            sap, wanted_cols, log=lambda m: _log(on_event, m), stop=stop)
        if stop.is_set():
            raise Cancelled()
        _log(on_event,
             f"grid read finished in {time.time() - read_start:.1f}s "
             f"({len(chunk_records)} rows)")
        if len(chunk_records) < grid_rows:
            raise RuntimeError(
                f"Grid read row-count mismatch{tag}: the status bar reports "
                f"{grid_rows} row(s) but only {len(chunk_records)} were "
                f"read.")
        records.extend(chunk_records)
        sub_progress(2 + 4 * (ci + 1) / len(chunks))

    # --- 3. Build the lookup dict from the grid rows ---------------------
    _status(on_event, f"[{step_index + 1}/{total_steps}] matching keys in memory")
    extra_sap_cols = [e.sap_col for e in step.extras if e.sap_col]
    if not records:
        _log(on_event,
             "SAP returned no rows for any of the requested keys -- every "
             "row will be treated as unmatched", "warn")
        lookup, dup_count, blank_key_count = {}, 0, 0
    else:
        _log(on_event, f"SAP returned {len(records)} row(s) with fields {wanted_cols}")
        lookup, dup_count, blank_key_count = _build_lookup(
            records,
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
    # Excel write-back.
    for j, sap_c in enumerate(step.sap_output_columns):
        col_values = []
        for entry in entries_by_row:
            v = entry[sap_c] if entry is not None else ""
            col_values.append("" if v is None else v)
        vsheet.publish(step.excel_output_columns[j], col_values)
    # Extras are published too when they are a pure function of the fetch
    # (a SAP source and no preserve-on-nonmatch dependence on existing cell
    # content): TO step 2 keys on N/O, which step 1 derives from ABLAD.
    for extra in step.extras:
        if extra.sap_col is None or extra.preserve_on_nonmatch:
            continue
        col_values = []
        for entry in entries_by_row:
            if entry is None:
                col_values.append("")
                continue
            v = entry[extra.sap_col]
            v = "" if v is None else v
            if extra.transform is not None:
                v = extra.transform(v)
            col_values.append(v)
        vsheet.publish(extra.excel_col, col_values)

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
    result = StepResult(
        step=step,
        step_index=step_index,
        last_row=last_row,
        excel_keys=excel_keys,
        entries_by_row=entries_by_row,
        n_skipped=n_skipped,
        matched=matched,
        unmatched=unmatched,
        sap_rows=sap_rows_total,
    )
    return result


def _step_summary(res: StepResult, total_steps: int, tail: str) -> str:
    """The original notebook's per-step message box body, counts and all."""
    step = res.step
    key_cols = "/".join(_col_letter(c) for c in sorted(step.key_columns))
    out_cols = ", ".join(_step_columns(step))
    return (f"{step.name}\n\n"
            f"SAP {step.sap_table} rows detected: {res.sap_rows}\n"
            f"Matched: {res.matched}\n"
            f"Not matched: {res.unmatched}\n"
            f"Skipped (blank / NOT FOUND): {res.n_skipped}\n\n"
            f"Keys were read from column(s) {key_cols}; results go to "
            f"column(s) {out_cols}.\n" + tail)


def _write_back(
    results: list[StepResult],
    xl: excel_ops.ExcelCtx,
    on_event: EventFn,
    total_steps: int,
    derived_cols: list[tuple[int, list]] | None = None,
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
        # Template-formula derivations first (K/L from the scan in B): the
        # values the steps keyed on, written where the formulas used to
        # put them. Rows 2..len only -- these are input-side columns, so
        # no clear-to-bottom.
        for col, values in (derived_cols or []):
            letter = _col_letter(col)
            excel_ops.set_column_format_text(sheet, f"{letter}:{letter}")
            excel_ops.write_range_2d(
                sheet, f"{letter}2:{letter}{len(values) + 1}",
                [[v] for v in values])

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
                            v = "" if v is None else v
                            if extra.transform is not None:
                                v = extra.transform(v)
                            col_data.append([v])
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


def _step_columns(step: LookupStep) -> list[str]:
    """Every Excel column a step writes, as letters, in write order."""
    cols = [_col_letter(c) for c in step.excel_output_columns]
    cols += [_col_letter(e.excel_col) for e in step.extras]
    if step.match_key_column is not None:
        cols.append(_col_letter(step.match_key_column))
    return cols


def _salvage_completed_steps(
    results: list[StepResult],
    steps: list[LookupStep],
    xl,
    on_event: EventFn,
    write_state: str,
    dry_run: bool,
    derived_cols: list[tuple[int, list]] | None = None,
) -> str:
    """A step failed. Write the steps that DID complete, and say plainly
    which columns are filled and which are not.

    Gen 4 originally discarded everything on any failure. That is the safest
    rule but it also threw away work that was already correct -- a run whose
    step 1 matched every row still left the workbook untouched because step 2
    could not find its SAP filter. Steps write disjoint column blocks and
    `results` only ever holds a PREFIX of the workflow (a step that fails
    stops the ones after it, which is also what feeds them their keys), so
    applying that prefix is exactly what a successful run would have written
    for those steps.

    Returns (new_write_state, salvage_error_text). Never raises: a failure
    to salvage must not replace the original error -- but it must not be
    silent either, so its text travels back for the error popup (two
    failures in one run usually share one cause).
    """
    if dry_run or not results or xl is None:
        return write_state, ""
    if write_state == "stepwise":
        # Step-by-step mode already wrote (and saved) every completed step
        # the moment it finished -- nothing to salvage.
        return write_state, ""
    if write_state != "none":
        # The failure was already inside the final write-back; re-running it
        # would double-apply. Leave it to the 'partial' warning.
        return write_state, ""

    done = {res.step_index for res in results}
    missing = [(i, s) for i, s in enumerate(steps) if i not in done]
    try:
        _log(on_event,
             f"salvage: {len(results)} of {len(steps)} step(s) completed "
             f"before the failure -- writing those, leaving the rest alone",
             "warn")
        _write_back(results, xl, on_event, len(steps), derived_cols)
        try:
            excel_ops.save(xl.book)
        except excel_ops.ExcelError as e:
            _log(on_event, f"WARNING: {e}", "warn")
        for res in results:
            _log(on_event,
                 f"  WRITTEN  step {res.step_index + 1} "
                 f"({res.step.sap_table}): columns "
                 f"{', '.join(_step_columns(res.step))} -- "
                 f"{res.matched} matched / {res.unmatched} unmatched", "ok")
        for i, s in missing:
            _log(on_event,
                 f"  NOT RUN  step {i + 1} ({s.sap_table}): columns "
                 f"{', '.join(_step_columns(s))} left as they were", "warn")
        return "salvaged", ""
    except Exception as e:
        _log(on_event,
             f"WARNING: could not write the completed steps either ({e}). "
             f"The workbook may be partially updated -- review it before "
             f"saving.", "warn")
        _file_log_traceback()
        return "partial", f"{type(e).__name__}: {e}"


def _write_state_text(write_state: str) -> str:
    """One plain sentence about the workbook, shared by the log lines and
    the popups so the two can never disagree."""
    return {
        "none": "The workbook was NOT modified.",
        "salvaged": ("The workbook holds the completed steps ONLY (see "
                     "WRITTEN / NOT RUN in the log) and has been saved. "
                     "Re-running after the fix is safe -- every step "
                     "rewrites its own columns from scratch."),
        "stepwise": ("Steps completed so far are ALREADY written and "
                     "saved in the workbook (step-by-step mode); later "
                     "steps were not run. Re-running is safe -- every step "
                     "rewrites its own columns from scratch."),
        "partial": ("Failure happened DURING the final write-back -- the "
                    "workbook may be partially updated and has NOT been "
                    "saved. Review it before saving manually."),
        "done": ("All output columns were written in one pass and the "
                 "workbook was SAVED."),
    }[write_state]


def _run_summary_lines(results: list[StepResult]) -> str:
    """Per-step counts for the completion popup, one line per step ran."""
    if not results:
        return "No step had anything to look up."
    return "\n".join(
        f"Step {r.step_index + 1} ({r.step.sap_table}): "
        f"{r.matched} matched / {r.unmatched} not matched"
        + (f" / {r.n_skipped} skipped" if r.n_skipped else "")
        + f" -> columns {', '.join(_step_columns(r.step))}"
        for r in results)


def _log_write_state(on_event: EventFn, write_state: str) -> None:
    """After a cancel/failure, tell the user exactly what state the file is
    in."""
    if write_state == "none":
        _log(on_event,
             _write_state_text("none") + " (no step completed before the "
             "failure)", "info")
    elif write_state == "salvaged":
        _log(on_event, _write_state_text("salvaged"), "warn")
    elif write_state == "stepwise":
        _log(on_event, _write_state_text("stepwise"), "warn")
    elif write_state == "partial":
        _log(on_event, "WARNING: " + _write_state_text("partial"), "warn")


def _dry_run_report(results: list[StepResult], on_event: EventFn) -> None:
    """Item 4: summarize what the write-back WOULD do, without doing it."""
    _log(on_event, "DRY RUN -- the workbook will not be modified. Planned writes:")
    for res in results:
        step = res.step
        cols = _step_columns(step)
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


def _diagnose(workflow: str, sap: sap_ops.SapSession, on_event: EventFn) -> None:
    """Dump every ZTBV selection screen this workflow touches.

    ZTBV names its select-options generically (S3, S15, S29), so a filter
    can only be identified by the texts printed on its row. This lists
    every slot with those texts, and says whether each field the workflow
    needs has a recorded slot in PUSH_BUTTONS. Read-only: no filter is
    pasted, no query executed, Excel never opened.
    """
    tables: list[str] = []
    for step in WORKFLOWS[workflow]:
        if step.sap_table not in tables:
            tables.append(step.sap_table)

    _log(on_event,
         f"DIAGNOSE: listing filter slots for {len(tables)} table(s): "
         f"{', '.join(tables)}. Nothing is pasted, executed, or written.")

    for table in tables:
        _log(on_event, "")
        _log(on_event, f"===== {table} =====")
        sap_ops.open_ztbv_table(sap, table, log=lambda m: _log(on_event, m))
        try:
            sap_ops.list_filter_slots(sap, log=lambda m: _log(on_event, m))
        except sap_ops.SapError as e:
            _log(on_event, f"{table}: {e}", "error")
            continue
        for step in WORKFLOWS[workflow]:
            if step.sap_table != table:
                continue
            key = (table, step.push_button_field)
            slot = sap_ops.PUSH_BUTTONS.get(key)
            if slot:
                _log(on_event,
                     f"  -> {step.push_button_field} uses recorded slot "
                     f"{slot}", "ok")
            else:
                _log(on_event,
                     f"  -> {step.push_button_field} has NO recorded slot "
                     f"-- pick it from the list above and add it to "
                     f"PUSH_BUTTONS (SAP utilities)", "error")

    _log(on_event, "")
    _log(on_event,
         "DIAGNOSE complete -- compare the slots above with PUSH_BUTTONS "
         "and send the log file when reporting.", "ok")


@dataclass
class RunConfig:
    excel_path: str
    workflow: str            # "TO" or "NOTIF"
    stop_event: threading.Event
    dry_run: bool = False    # Item 4: fetch + report matches, write nothing
    diagnose: bool = False   # dump SAP selection screens; touch nothing else
    # Original-notebook rhythm: a blocking counts popup after every step
    # (OK = continue, Cancel = stop). Error and completion popups are always
    # emitted regardless of this flag; this only controls the per-step ones.
    step_popups: bool = False


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
    # Declared up here so the failure handlers can salvage whatever the fetch
    # phase managed to complete before it died, and so the error popup can
    # name the step and snapshot the SAP screen.
    xl = None
    sap = None
    steps: list[LookupStep] = []
    results: list[StepResult] = []
    derived_cols: list[tuple[int, list]] = []
    step_label = "startup (before any SAP step)"
    salvage_error = ""

    # Remember the last status line: when a run dies, "what was it DOING"
    # is the single most valuable fact on the error popup -- a com_error
    # during 'reading keys' points at Excel, during 'sending filter values'
    # at SAP, without waiting for the log file to travel.
    last_action = {"text": ""}
    _caller_on_event = on_event

    def on_event(kind, payload):  # noqa: shadows the parameter on purpose
        if kind == "status":
            last_action["text"] = payload
        _caller_on_event(kind, payload)

    def _error_popup(err_text: str) -> None:
        """The original notebook's 'Error occurred' message box, upgraded
        with what a remote debugger needs: the step, the exact action in
        flight, the SAP screen at the moment of failure, the workbook
        state, and the log path."""
        snapshot = ""
        if sap is not None:
            try:
                snapshot = sap_ops.screen_snapshot(sap)
            except Exception:
                snapshot = ""
        _popup(
            on_event, "error", "esa-lookup -- error",
            f"Error occurred in {step_label}:\n\n{err_text}\n\n"
            + (f"While doing: {last_action['text']}\n\n"
               if last_action["text"] else "")
            + (f"SAP screen at the moment of failure:\n{snapshot}\n\n"
               if snapshot else "")
            + "The SAP window has been left exactly where it stopped.\n"
              "Please screenshot BOTH the SAP window and this message.\n\n"
            + _write_state_text(write_state)
            + (f"\n\nALSO: writing the completed steps to Excel failed "
               f"too:\n{salvage_error}\nTwo failures in one run usually "
               f"share one cause -- most often Excel or the workbook was "
               f"closed, or was being clicked/edited, while the run was in "
               f"flight. Close the workbook WITHOUT saving and re-run."
               if salvage_error else "")
            + (f"\n\nLog file (attach it when reporting):\n{log_path}"
               if log_path else ""))

    try:
        _log(on_event, f"esa-lookup starting workflow '{cfg.workflow}'", "info")
        if log_path:
            _log(on_event, f"detailed log file: {log_path}", "info")
        # Diagnostic: freeze the environment into the file for post-mortems.
        _file_log("info",
                  f"env: Python {sys.version.split()[0]} on {sys.platform}, "
                  f"cwd={os.getcwd()}, excel={cfg.excel_path}")

        # Diagnose is SAP-only and read-only: no workbook is opened, so it
        # runs even when the operator has no file picked.
        if cfg.diagnose:
            _status(on_event, "attaching to SAP GUI")
            sap = sap_ops.attach()
            _log(on_event, "attached to SAP GUI session", "ok")
            _diagnose(cfg.workflow, sap, on_event)
            _status(on_event, "diagnose complete -- nothing written")
            _popup(on_event, "info", "Diagnose complete",
                   "The ZTBV selection screens were read and dumped to the "
                   "log. Nothing was pasted, executed, or written.\n\n"
                   + (f"Send this log file:\n{log_path}" if log_path else
                      "See the log pane for the results."))
            _progress(on_event, 1.0)
            on_event("done", True)
            return True

        _status(on_event, "opening Excel")
        xl = excel_ops.attach(cfg.excel_path)
        _log(on_event, f"attached to Excel: {os.path.basename(cfg.excel_path)}", "ok")

        _status(on_event, "attaching to SAP GUI")
        sap = sap_ops.attach()
        _log(on_event, "attached to SAP GUI session", "ok")

        steps = WORKFLOWS[cfg.workflow]
        total_steps = len(steps)
        vsheet = VirtualSheet(xl.sheet)

        # ---- Phase 0: template-formula derivations (sheet -> sheet) -----
        # The Excel template computed some columns from others with
        # formulas (K/L out of the scan in B). Those formulas vanish when
        # sheets are copied, so the app derives the values itself: merged
        # with any existing cell content, published for the steps to key
        # on, and written in the final pass. Must run BEFORE the pre-flight
        # check -- a sheet with scans in B but a lost K formula is valid
        # input.
        for d in WORKFLOW_DERIVATIONS.get(cfg.workflow, []):
            last = max([vsheet.last_row(d.source_col)]
                       + [vsheet.last_row(c) for c, _ in d.targets])
            if last < 2:
                continue
            src_vals = vsheet.get_col(d.source_col, last)
            for t_col, fn in d.targets:
                existing = vsheet.get_col(t_col, last)
                merged, n_derived = [], 0
                for s_v, e_v in zip(src_vals, existing):
                    v = fn(s_v)
                    if v:
                        n_derived += 1
                    else:
                        v = "" if e_v is None else e_v
                    merged.append(v)
                if not n_derived:
                    continue   # nothing to add -- leave column untouched
                vsheet.publish(t_col, merged)
                derived_cols.append((t_col, merged))
                _log(on_event,
                     f"derived column {_col_letter(t_col)} from column "
                     f"{_col_letter(d.source_col)} for {n_derived} row(s) "
                     f"(the template's formula, applied by the app)")

        # Pre-flight: verify the workflow's primary input column has data
        # before we even spin up SAP navigation. Later steps derive their
        # own last_row (Fix F) and skip themselves if their column is empty.
        primary_col = steps[0].key_columns[0]
        primary_last = vsheet.last_row(primary_col)
        if primary_last < 2:
            _log(on_event,
                 f"No data below the header in column "
                 f"{_col_letter(primary_col)} (#{primary_col}) -- the "
                 f"workflow's primary input. Aborting.", "error")
            # Notebook parity: "No TO Number found in Column K." was a
            # message box, not a log line.
            _popup(on_event, "error", "esa-lookup -- nothing to do",
                   f"No data found below the header in column "
                   f"{_col_letter(primary_col)} -- the {cfg.workflow} "
                   f"workflow's primary input.\n\nNothing was written.")
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
        totals_matched = 0
        totals_seen = 0
        for i, step in enumerate(steps):
            if cfg.stop_event.is_set():
                raise Cancelled()
            step_label = f"Step {i + 1} of {total_steps} ({step.sap_table})"
            res = _fetch_step(
                step, vsheet, xl, sap, on_event, i, total_steps,
                cfg.stop_event, step_popups=cfg.step_popups,
            )
            if res is None:
                continue  # no-op step (empty key column)
            results.append(res)
            totals_matched += res.matched
            totals_seen += res.matched + res.unmatched

            if cfg.step_popups and not cfg.dry_run:
                # Step-by-step mode = the original notebook, faithfully:
                # this step's results are written AND SAVED before its
                # message box, so the operator can switch to Excel and
                # inspect real cells while the box waits. The price is
                # that Cancel keeps the steps already written -- said in
                # the box itself.
                step_label = (f"writing step {i + 1} results to Excel "
                              f"(step-by-step mode)")
                _status(on_event, f"[{i + 1}/{total_steps}] writing this "
                                  f"step's results to Excel")
                _write_back([res], xl, on_event, total_steps,
                            derived_cols if write_state == "none" else None)
                try:
                    excel_ops.save(xl.book)
                except excel_ops.ExcelError as e:
                    _log(on_event, f"WARNING: {e}", "warn")
                write_state = "stepwise"
                step_label = f"Step {i + 1} of {total_steps} ({step.sap_table})"
                _step_popup(
                    on_event, cfg.stop_event,
                    f"Step {i + 1} of {total_steps} completed",
                    _step_summary(
                        res, total_steps,
                        "\nThis step is ALREADY written and saved -- switch "
                        "to Excel to inspect it now.\n\n"
                        "OK = continue, Cancel = stop the run (steps written "
                        "so far stay in the workbook)."))
            elif cfg.step_popups:
                _step_popup(
                    on_event, cfg.stop_event,
                    f"Step {i + 1} of {total_steps} completed",
                    _step_summary(
                        res, total_steps,
                        "\nDRY RUN -- nothing is written.\n\n"
                        "OK = continue, Cancel = stop the run."))

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
            _popup(on_event, "info", "Dry run complete",
                   _run_summary_lines(results)
                   + "\n\nDRY RUN -- the workbook was NOT modified.")
            _progress(on_event, 1.0)
            on_event("done", True)
            return True

        # ---- Phase 2: write + save ---------------------------------------
        if write_state == "stepwise":
            # Step-by-step mode wrote and saved each step as it finished.
            _log(on_event,
                 "all steps were already written and saved one at a time "
                 "(step-by-step mode); nothing left to write", "ok")
            write_state = "done"
        else:
            step_label = "the final write-back to Excel"
            write_state = "partial"
            _write_back(results, xl, on_event, total_steps, derived_cols)
            write_state = "done"

        # Fix L: save() now raises ExcelError on failure; warn the user
        # rather than silently pretending the write persisted.
        try:
            excel_ops.save(xl.book)
        except excel_ops.ExcelError as e:
            _log(on_event, f"WARNING: {e}", "warn")

        _status(on_event, "done")
        _log(on_event, f"all steps complete: {totals_matched}/{totals_seen} row-matches across {total_steps} step(s)", "ok")
        # Notebook parity: the "Completed." message box with the counts.
        _popup(on_event, "info", "Completed",
               _run_summary_lines(results) + "\n\n"
               + ("All steps were written and SAVED one at a time "
                  "(step-by-step mode)."
                  if cfg.step_popups and results else
                  _write_state_text("done")))
        _progress(on_event, 1.0)
        on_event("done", True)
        return True

    except Cancelled:
        # Stop stays strictly all-or-nothing: the GUI promises the file will
        # not be modified, and an explicit abort is not a partial result the
        # operator asked to keep.
        _log(on_event, "cancelled by user", "warn")
        if write_state == "none":
            _log(on_event,
                 "the workbook was NOT modified (Stop leaves the file "
                 "untouched)", "info")
        else:
            _log_write_state(on_event, write_state)
        _popup(on_event, "info", "Cancelled",
               "The run was stopped.\n\n" + _write_state_text(write_state))
        _status(on_event, "cancelled")
        on_event("done", False)
        return False
    except sap_ops.SapError as e:
        _log(on_event, f"SAP error: {e}", "error")
        _file_log_traceback()  # full chain into the log file for post-mortem
        write_state, salvage_error = _salvage_completed_steps(
            results, steps, xl, on_event, write_state, cfg.dry_run,
            derived_cols)
        _log_write_state(on_event, write_state)
        _error_popup(str(e))
        _status(on_event, "SAP error")
        on_event("done", False)
        return False
    except excel_ops.ExcelError as e:
        _log(on_event, f"Excel error: {e}", "error")
        _file_log_traceback()
        write_state, salvage_error = _salvage_completed_steps(
            results, steps, xl, on_event, write_state, cfg.dry_run,
            derived_cols)
        _log_write_state(on_event, write_state)
        _error_popup(str(e))
        _status(on_event, "Excel error")
        on_event("done", False)
        return False
    except Exception as e:
        _log(on_event, f"unexpected error: {e}", "error")
        _log(on_event, traceback.format_exc(), "error")
        _file_log_traceback()
        write_state, salvage_error = _salvage_completed_steps(
            results, steps, xl, on_event, write_state, cfg.dry_run,
            derived_cols)
        _log_write_state(on_event, write_state)
        _error_popup(f"{type(e).__name__}: {e}")
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
