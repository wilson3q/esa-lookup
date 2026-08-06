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
            # ESA rework (2026-08): the original keyed this step on N|O
            # (reservation number + item), which the operator had to fill by
            # hand between steps. Per the ESA developers' demo, the same
            # table answers a TO-number query directly -- Z50CFG_ENG_CRNT
            # carries the TO pair and the reservation pair on one row -- so
            # the step now keys on column K (already filled for step 1) and
            # DERIVES the rest: QMNUM -> A, RSNUM -> N, RSPOS -> O. Rows
            # whose TO number found no CRNT row are simply unmatched, same
            # as the demo's pasted blanks being ignored by SAP.
            name="Z50CFG_ENG_CRNT (TO Number) -> QMNUM/OBJNR/DISP",
            sap_table="Z50CFG_ENG_CRNT",
            push_button_field="TO_NUMBER",
            key_columns=[11],                     # K (same input as step 1)
            sap_key_columns=["TANUM"],
            sap_output_columns=["OBJNR", "DISP_MATNR", "DISP_QTY"],
            excel_output_columns=[3, 4, 5],       # C, D, E
            extras=[
                # Notebook rule kept: A gets QMNUM on match but MUST NOT be
                # touched on non-match ("Do not touch Column A if there is
                # no match"). N/O are now outputs (the derived reservation
                # pair) -- also preserved on non-match so a hand-entered
                # value is never clobbered by a miss.
                ExtraOutput(excel_col=1, sap_col="QMNUM", preserve_on_nonmatch=True),
                ExtraOutput(excel_col=14, sap_col="RSNUM", preserve_on_nonmatch=True),
                ExtraOutput(excel_col=15, sap_col="RSPOS", preserve_on_nonmatch=True),
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


# Fix 3: SAP ALV exports typically use the ALV column TITLE, which is often a
# short description rather than the technical field name. Try the technical
# name first, then any known aliases. Extend per-site if needed.
SAP_COLUMN_ALIASES: dict[str, list[str]] = {
    "TANUM":       ["TANUM", "TO Number", "TO No.", "Transfer Order", "TrfOrd", "TrfOrdNo"],
    # "Unl. Point" is what the ESA ZTBV layout actually prints for ABLAD --
    # the ALV shows English short labels, not technical names.
    "ABLAD":       ["ABLAD", "Unl. Point", "Unloading Point", "UnloadPt"],
    "QMNUM":       ["QMNUM", "Notifctn", "Notification", "Notification No", "Notification Number"],
    "RSNUM":       ["RSNUM", "Reserv.No.", "Reservation", "Reservation No", "Reservation Number", "Res.Number"],
    # ORDER MATTERS. The Z50CFG_ENG_CRNT layout carries BOTH pairs side by
    # side: 'Reserv.No.' + 'Itm' (the reservation) and 'TO Number' + 'Item'
    # (the transfer order). RSPOS is the RESERVATION item, so 'Itm' has to
    # be tried before the generic 'Item' -- otherwise RSPOS silently binds
    # to the TO item, resolution "succeeds", and step 2 keys on a mismatched
    # pair that quietly matches nothing.
    "RSPOS":       ["RSPOS", "Itm", "Res.Item", "Item No", "Item Number", "Item"],
    "OBJNR":       ["OBJNR", "Object number", "Object Number", "Obj.Number", "Object No"],
    "DISP_MATNR":  ["DISP_MATNR", "Disp Matl", "Disp Material", "Disposition Material", "Disp.Material"],
    "DISP_QTY":    ["DISP_QTY", "Disp Qty", "Disposition Qty", "Disposition Quantity", "Disp.Qty"],
    "Z_SECTION":   ["Z_SECTION", "Section"],
    "Z_MODULE":    ["Z_MODULE", "Module"],
    "DESCRIPT":    ["DESCRIPT", "Description", "Descr."],
    "SALES_ORDER": ["SALES_ORDER", "Sales Order", "Sales Doc.", "Sales Doc", "Sales Doc. No."],
    "LID":         ["LID"],
}


def _norm_col(name) -> str:
    """Fold an ALV column title for tolerant comparison: case, spaces, and
    the punctuation SAP sprinkles through its abbreviations all ignored, so
    "Unl. Point" / "Unl.Point" / "UNL POINT" are one name.
    """
    return re.sub(r"[\s._\-/]+", "", str(name)).lower()


def _resolve_column(df: pd.DataFrame, canonical: str) -> str | None:
    """Return whichever of `canonical`'s alias names is present in df, or None.

    Exact match first, then the folded comparison -- an ALV variant that
    prints "Unl.Point" should still resolve an alias spelled "Unl. Point"
    without needing a separate entry for every punctuation variant.
    """
    aliases = SAP_COLUMN_ALIASES.get(canonical, [canonical])
    for name in aliases:
        if name in df.columns:
            return name
    folded: dict[str, list[str]] = {}
    for actual in df.columns:
        folded.setdefault(_norm_col(actual), []).append(actual)
    for name in aliases:
        hits = folded.get(_norm_col(name))
        if hits:
            return hits[0]
    return None


# ---------------------------------------------------------------------------
# Reading whatever SAP actually exported
# ---------------------------------------------------------------------------

# SAP names every ALV export ".xlsx", but only the `&XXL` path really writes
# one. When `&XXL` is unavailable (older SAP GUI build, or no Excel/XXL
# frontend integration -- it fails with "The control could not be found"),
# `sap_ops.export_alv_to_file` falls back to `&PC` "Save as Local File",
# which writes a DELIMITED TEXT file under that same .xlsx name. Feeding
# that to pd.read_excel dies with "Excel file format cannot be determined,
# you must specify an engine manually".
#
# Independently, on sites whose Office integration is enabled, the `&XXL`
# path itself can write an MHTML-wrapped <table> under the .xlsx name --
# openpyxl rejects that one with "not a zip file".
#
# So the extension tells us nothing: sniff the real format from the leading
# bytes and dispatch on that.

_ZIP_MAGIC = b"PK\x03\x04"          # .xlsx / OOXML (a zip container)
_OLE2_MAGIC = b"\xd0\xcf\x11\xe0"   # legacy .xls (BIFF8 inside OLE2)


def _decode_sap_text(raw: bytes) -> str:
    """Decode an SAP text export: honour the BOM SAP writes on Unicode
    systems, else fall back to the Windows frontend codepage.
    """
    for bom, enc in ((b"\xff\xfe\x00\x00", "utf-32"),
                     (b"\x00\x00\xfe\xff", "utf-32"),
                     (b"\xff\xfe", "utf-16"),
                     (b"\xfe\xff", "utf-16"),
                     (b"\xef\xbb\xbf", "utf-8-sig")):
        if raw.startswith(bom):
            return raw.decode(enc)
    for enc in ("utf-8", "cp1252"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("latin-1", "replace")


def _sap_text_to_frame(text: str) -> pd.DataFrame:
    """Parse an SAP `&PC` export -- either the tab-delimited "Spreadsheet"
    format or the pipe-ruled "unconverted" list format -- into a DataFrame.

    Every cell stays a string. SAP already formatted the values the way the
    ALV displayed them, and keeping them verbatim preserves leading zeros on
    material / TO numbers that pandas' numeric inference would eat.
    `normalize_key` canonicalizes both sides of the join anyway, so string
    cells match exactly the rows numeric cells would.
    """
    lines = []
    for ln in text.splitlines():
        s = ln.strip()
        # skip blanks and ALV list rulers: "--------" / "|----+----|"
        if not s or set(s) <= set("-+| "):
            continue
        lines.append(ln)
    if not lines:
        return pd.DataFrame()

    if sum(ln.count("|") for ln in lines) > sum(ln.count("\t") for ln in lines):
        rows = [[c.strip() for c in ln.strip().strip("|").split("|")]
                for ln in lines]
    else:
        rows = [[c.strip() for c in ln.split("\t")] for ln in lines]

    # SAP prefixes some list exports with report title / run-date lines that
    # carry fewer fields than the table. The real header row is the first one
    # with the table's own field count (the most common width).
    widths = [len(r) for r in rows]
    modal = max(set(widths), key=widths.count)
    first = widths.index(modal)
    # Rows of a different width are footers ("3 record(s) selected") or wrapped
    # lines; dropping them can only UNDER-count, which the caller's
    # grid_rows comparison turns into a loud error rather than silent loss.
    rows = [r for r in rows[first:] if len(r) == modal]

    header = [h or f"Unnamed: {i}" for i, h in enumerate(rows[0])]
    return pd.DataFrame(rows[1:], columns=header)


def _read_sap_export(path: str, log=None) -> pd.DataFrame:
    """Read the file `sap_ops.export_alv_to_file` produced, whatever format
    SAP actually chose for it.
    """
    with open(path, "rb") as fh:
        raw = fh.read()
    head = raw[:2048]

    if head.startswith(_ZIP_MAGIC):
        return pd.read_excel(path, engine="openpyxl")

    if head.startswith(_OLE2_MAGIC):
        try:
            return pd.read_excel(path, engine="xlrd")
        except ImportError as e:
            raise RuntimeError(
                f"SAP exported a legacy .xls workbook ({path}); reading it "
                f"needs the 'xlrd' package -- run: pip install xlrd"
            ) from e

    # MHTML / HTML: some sites' Office integration makes "Export as
    # Spreadsheet" write an MHTML-wrapped <table> under the .xlsx name.
    # The MIME preamble has to be sliced off before read_html sees it.
    lowered = head.lower()
    payload = None
    if b"mime-version:" in lowered or b"content-location:" in lowered:
        i = raw.lower().find(b"<html")
        if i < 0:
            raise RuntimeError(
                f"SAP export {path} looks like MHTML but no <html> body was "
                f"found.")
        payload = raw[i:]
    elif b"<html" in lowered or b"<table" in lowered:
        payload = raw

    if payload is not None:
        # Must be a file-like object: pandas >=3 treats a bytes/str argument
        # as a PATH, so passing the markup itself raises FileNotFoundError.
        try:
            tables = pd.read_html(io.BytesIO(payload))
        except ImportError as e:
            raise RuntimeError(
                f"SAP exported an HTML/MHTML table ({path}); reading it needs "
                f"the 'lxml' package -- run: pip install lxml"
            ) from e
        if not tables:
            raise RuntimeError(
                f"SAP export {path}: no <table> found in the HTML body.")
        return tables[0]

    text = _decode_sap_text(raw)
    if log:
        log("SAP: export is a delimited text file (&PC fallback), not xlsx -- "
            "parsing as text")
    df = _sap_text_to_frame(text)
    # We only get here for a grid that reported rows (the caller skips empty
    # grids), so a frame with no data rows means the file is not a delimited
    # export at all -- say so instead of failing later on missing columns.
    if df.empty:
        raise RuntimeError(
            f"Could not parse the SAP export at {path}: it is neither an "
            f"xlsx/xls workbook nor a recognizable delimited text export. "
            f"Open it manually to see what ZTBV actually wrote."
        )
    return df


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
        # Name the closest present titles: the ALV prints English short
        # labels ("Unl. Point" for ABLAD), so the field is usually right
        # there under a name no alias lists yet. Guessing it here saves a
        # round trip to whoever is standing at the customer's machine.
        hints = []
        for c in missing:
            near = difflib.get_close_matches(
                c, [str(x) for x in df.columns], n=3, cutoff=0.4)
            for alias in SAP_COLUMN_ALIASES.get(c, []):
                near += difflib.get_close_matches(
                    alias, [str(x) for x in df.columns], n=3, cutoff=0.6)
            near = list(dict.fromkeys(near))[:3]
            if near:
                hints.append(f"  {c}: closest titles present are {near}")
        raise RuntimeError(
            "SAP export is missing these expected columns: "
            f"{missing}\n"
            f"Columns present in the export: {list(df.columns)}\n"
            + ("Did you mean:\n" + "\n".join(hints) + "\n" if hints else "")
            + "Fix: edit the ALV layout in ZTBV so each missing field is shown "
            "(prefer 'Technical Name' as the column title) and re-save the "
            "default variant, OR add another alias to SAP_COLUMN_ALIASES in "
            "pipeline.py."
        )

    # A title repeated in the layout makes df[title] a DataFrame slice, so
    # row[title] is a Series -- normalize_key would stringify the whole
    # Series and pd.isna would raise "truth value is ambiguous". The ESA
    # LTAP layout repeats 'Item', 'Typ', 'Sec' and 'B.pos' several times,
    # so fail loudly here rather than write nonsense into the workbook.
    titles = [str(x) for x in df.columns]
    ambiguous = {c: r for c, r in resolved.items() if titles.count(str(r)) > 1}
    if ambiguous:
        raise RuntimeError(
            "These SAP fields resolved to a column title that appears more "
            "than once in the export, so the right one cannot be told apart:\n"
            + "\n".join(f"  {c} -> {r!r} (appears {titles.count(str(r))} times)"
                        for c, r in ambiguous.items())
            + "\nFix: edit the ALV layout in ZTBV to show 'Technical Name' as "
            "the column title (which is unique per field), or remove the "
            "duplicate columns from the layout and re-save the variant."
        )

    # Two SAP fields landing on one column means an alias list is too greedy
    # (the way RSPOS's generic "Item" once outranked the reservation's own
    # "Itm"). Resolution would "succeed" and the step would then key on a
    # field it was never meant to read, so refuse rather than guess.
    collisions: dict[str, list[str]] = {}
    for c, r in resolved.items():
        collisions.setdefault(str(r), []).append(c)
    clashing = {r: cs for r, cs in collisions.items() if len(cs) > 1}
    if clashing:
        raise RuntimeError(
            "These SAP fields all resolved to the SAME export column, so at "
            "least one of them is bound to the wrong field:\n"
            + "\n".join(f"  {cs} -> {r!r}" for r, cs in clashing.items())
            + "\nFix: tighten the alias lists in SAP_COLUMN_ALIASES so each "
            "field matches its own column title first."
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
    sap_rows: int = 0     # total rows SAP's grid reported across all chunks


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

    # Resolved on the first chunk, after open_ztbv_table has the selection
    # screen up -- a (table, field) pair missing from PUSH_BUTTONS is then
    # found by its on-screen label instead of failing.
    push_id: str | None = None
    frames: list[pd.DataFrame] = []
    ts = int(time.time())
    # Every SAP field this step needs, by technical name -- what the grid
    # reader asks for, and exactly what the original notebook read.
    wanted_cols = list(dict.fromkeys(
        step.sap_key_columns
        + step.sap_output_columns
        + [e.sap_col for e in step.extras if e.sap_col]
    ))
    # One export fallback decision per step, not per chunk: if the grid read
    # is unavailable here it will be unavailable for every other chunk too,
    # and flip-flopping between paths mid-step would mix column namings.
    use_grid = True
    sap_rows_total = 0
    for ci, chunk in enumerate(chunks):
        if stop.is_set():
            raise Cancelled()
        tag = f" (chunk {ci + 1}/{len(chunks)})" if len(chunks) > 1 else ""

        _status(on_event, f"[{step_index + 1}/{total_steps}] SAP: loading {step.sap_table}{tag}")
        sap_ops.open_ztbv_table(sap, step.sap_table, log=lambda m: _log(on_event, m))
        if push_id is None:
            push_id = sap_ops.resolve_push_button(
                sap, step.sap_table, step.push_button_field,
                log=lambda m: _log(on_event, m))
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
        sap_rows_total += grid_rows
        if grid_rows == 0:
            _log(on_event, f"SAP returned 0 rows{tag}; nothing to export", "warn")
            sub_progress(2 + 4 * (ci + 1) / len(chunks))
            continue

        # --- Primary read: the grid itself, by technical field name ------
        # Ported from the original notebook. It sidesteps the export file
        # format entirely and, because technical names are unique per field,
        # every alias / duplicate-title / short-label problem with it.
        if use_grid:
            try:
                _status(on_event,
                        f"[{step_index + 1}/{total_steps}] SAP: reading ALV grid "
                        f"({grid_rows} rows){tag}")
                read_start = time.time()
                records = sap_ops.read_alv_grid(
                    sap, wanted_cols, log=lambda m: _log(on_event, m), stop=stop)
                if stop.is_set():
                    raise Cancelled()
                df_chunk = pd.DataFrame(records, columns=wanted_cols)
                _log(on_event,
                     f"grid read finished in {time.time() - read_start:.1f}s "
                     f"({len(df_chunk)} rows)")
                if len(df_chunk) < grid_rows:
                    raise RuntimeError(
                        f"Grid read row-count mismatch{tag}: the status bar "
                        f"reports {grid_rows} row(s) but only "
                        f"{len(df_chunk)} were read.")
                frames.append(df_chunk)
                sub_progress(2 + 4 * (ci + 1) / len(chunks))
                continue
            except sap_ops.SapError as e:
                # A missing technical name is the expected reason to land
                # here; fall back to the export for the rest of the step.
                _log(on_event,
                     f"grid read unavailable ({e}); falling back to the file "
                     f"export for this step", "warn")
                use_grid = False

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

        df_chunk = _read_sap_export(
            export_path, log=lambda m: _log(on_event, m)
        )
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
    if step_popups:
        # The original notebook's per-step message box, counts and all. The
        # operator can look at the SAP result grid (still on screen) before
        # clicking OK; Cancel stops the run with the workbook untouched.
        key_cols = "/".join(_col_letter(c) for c in sorted(step.key_columns))
        out_cols = ", ".join(_step_columns(step))
        _step_popup(
            on_event, stop,
            f"Step {step_index + 1} of {total_steps} completed",
            f"{step.name}\n\n"
            f"SAP {step.sap_table} rows detected: {sap_rows_total}\n"
            f"Matched: {matched}\n"
            f"Not matched: {unmatched}\n"
            f"Skipped (blank / NOT FOUND): {n_skipped}\n\n"
            f"Keys were read from column(s) {key_cols}.\n"
            f"Results go to column(s) {out_cols}, written together at the "
            f"end of the run.\n\n"
            f"OK = continue, Cancel = stop the run (workbook untouched).")
    return result


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

    Returns the new write_state. Never raises: a failure to salvage must not
    replace the original error, which is the one worth reporting.
    """
    if dry_run or not results or xl is None:
        return write_state
    if write_state != "none":
        # The failure was already inside the final write-back; re-running it
        # would double-apply. Leave it to the 'partial' warning.
        return write_state

    done = {res.step_index for res in results}
    missing = [(i, s) for i, s in enumerate(steps) if i not in done]
    try:
        _log(on_event,
             f"salvage: {len(results)} of {len(steps)} step(s) completed "
             f"before the failure -- writing those, leaving the rest alone",
             "warn")
        _write_back(results, xl, on_event, len(steps))
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
        return "salvaged"
    except Exception as e:
        _log(on_event,
             f"WARNING: could not write the completed steps either ({e}). "
             f"The workbook may be partially updated -- review it before "
             f"saving.", "warn")
        _file_log_traceback()
        return "partial"


def _write_state_text(write_state: str) -> str:
    """One plain sentence about the workbook, shared by the log lines and
    the popups so the two can never disagree."""
    return {
        "none": "The workbook was NOT modified.",
        "salvaged": ("The workbook holds the completed steps ONLY (see "
                     "WRITTEN / NOT RUN in the log) and has been saved. "
                     "Re-running after the fix is safe -- every step "
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

    ZTBV's select-options are named generically (S3, S15, S29), so a filter
    can only be identified by the label printed beside it. When that pairing
    fails -- a screen built without GuiLabels, an unusual SAP GUI build --
    the run dies with nothing to act on. This walks each table's screen and
    logs BOTH the resolved filter list and the raw control inventory, so the
    S<n> -> field mapping can be read by eye from the log file and pinned
    with ESA_LOOKUP_PUSH_<TABLE>_<FIELD>.

    Read-only: it navigates and reads. No filter is pasted, no query is
    executed, and Excel is never opened.
    """
    tables: list[str] = []
    for step in WORKFLOWS[workflow]:
        if step.sap_table not in tables:
            tables.append(step.sap_table)

    _log(on_event,
         f"DIAGNOSE: dumping the ZTBV selection screen for {len(tables)} "
         f"table(s): {', '.join(tables)}. Nothing is pasted, executed, or "
         f"written -- this only reads the screens.")

    wanted_by_table: dict[str, list[str]] = {}
    for step in WORKFLOWS[workflow]:
        wanted_by_table.setdefault(step.sap_table, [])
        if step.push_button_field not in wanted_by_table[step.sap_table]:
            wanted_by_table[step.sap_table].append(step.push_button_field)

    for table in tables:
        _log(on_event, "")
        _log(on_event, f"===== {table} =====")
        sap_ops.open_ztbv_table(sap, table, log=lambda m: _log(on_event, m))
        try:
            filters = sap_ops.describe_selection_screen(sap)
        except sap_ops.SapError as e:
            _log(on_event, f"{table}: cannot read the selection screen: {e}",
                 "error")
            continue

        _log(on_event, f"{table}: {len(filters)} multi-value filter(s)")
        for f in filters:
            _log(on_event,
                 f"  {f['param']:<6} row={f['row']:<4} label={f['label']!r}"
                 + (f"  tooltip={f['tooltip']!r}" if f["tooltip"] else ""))

        # What this workflow actually needs off this screen, and whether it
        # would resolve right now.
        for field in wanted_by_table[table]:
            try:
                pid = sap_ops.resolve_push_button(sap, table, field)
                _log(on_event, f"  -> {field} resolves to {pid}", "ok")
            except sap_ops.SapError as e:
                _log(on_event, f"  -> {field} DOES NOT RESOLVE: {e}", "error")
                _log(on_event,
                     f"     pin it with: set "
                     f"{sap_ops.env_override_var(table, field)}=S<n>", "warn")

        # Raw inventory -- the fallback when no label paired. Only text-bearing
        # controls; the rest is noise for this purpose.
        _file_log("info", f"--- {table}: raw control inventory ---")
        try:
            for it in sap_ops.screen_inventory(sap):
                if not (it["text"] or it["tooltip"]):
                    continue
                _file_log("info",
                          f"  {it['type']:<16} row={it['row']:<4} "
                          f"col={it['col']:<4} name={it['name']!r} "
                          f"text={it['text']!r} tooltip={it['tooltip']!r}")
        except sap_ops.SapError as e:
            _file_log("info", f"  inventory unavailable: {e}")

    _log(on_event, "")
    _log(on_event,
         "DIAGNOSE complete. The full per-control dump is in the log FILE "
         "(path echoed at the top) -- send that file on.", "ok")


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
    step_label = "startup (before any SAP step)"

    def _error_popup(err_text: str) -> None:
        """The original notebook's 'Error occurred' message box, upgraded
        with what a remote debugger needs: the step, the SAP screen at the
        moment of failure, the workbook state, and the log path."""
        snapshot = ""
        if sap is not None:
            try:
                snapshot = sap_ops.screen_snapshot(sap)
            except Exception:
                snapshot = ""
        _popup(
            on_event, "error", "esa-lookup -- error",
            f"Error occurred in {step_label}:\n\n{err_text}\n\n"
            + (f"SAP screen at the moment of failure:\n{snapshot}\n\n"
               if snapshot else "")
            + "The SAP window has been left exactly where it stopped.\n"
              "Please screenshot BOTH the SAP window and this message.\n\n"
            + _write_state_text(write_state)
            + (f"\n\nLog file (attach it when reporting):\n{log_path}"
               if log_path else ""))

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
        vsheet = VirtualSheet(xl.sheet)
        totals_matched = 0
        totals_seen = 0
        for i, step in enumerate(steps):
            if cfg.stop_event.is_set():
                raise Cancelled()
            step_label = f"Step {i + 1} of {total_steps} ({step.sap_table})"
            res = _fetch_step(
                step, vsheet, xl, sap, tmp_dir, on_event, i, total_steps,
                cfg.stop_event, step_popups=cfg.step_popups,
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
            _popup(on_event, "info", "Dry run complete",
                   _run_summary_lines(results)
                   + "\n\nDRY RUN -- the workbook was NOT modified.")
            _progress(on_event, 1.0)
            on_event("done", True)
            return True

        # ---- Phase 2: single write-back pass + save ---------------------
        step_label = "the final write-back to Excel"
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
        # Notebook parity: the "Completed." message box with the counts.
        _popup(on_event, "info", "Completed",
               _run_summary_lines(results) + "\n\n"
               + _write_state_text("done"))
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
        write_state = _salvage_completed_steps(
            results, steps, xl, on_event, write_state, cfg.dry_run)
        _log_write_state(on_event, write_state)
        _error_popup(str(e))
        _status(on_event, "SAP error")
        on_event("done", False)
        return False
    except excel_ops.ExcelError as e:
        _log(on_event, f"Excel error: {e}", "error")
        _file_log_traceback()
        write_state = _salvage_completed_steps(
            results, steps, xl, on_event, write_state, cfg.dry_run)
        _log_write_state(on_event, write_state)
        _error_popup(str(e))
        _status(on_event, "Excel error")
        on_event("done", False)
        return False
    except Exception as e:
        _log(on_event, f"unexpected error: {e}", "error")
        _log(on_event, traceback.format_exc(), "error")
        _file_log_traceback()
        write_state = _salvage_completed_steps(
            results, steps, xl, on_event, write_state, cfg.dry_run)
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
