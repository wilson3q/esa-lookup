"""SAP GUI Scripting helpers for ZTBV table lookup.

All calls attach to an already-running SAP GUI session -- the user must be
logged in before the app is launched.
"""
from __future__ import annotations

import os
import re
import time
from dataclasses import dataclass

import win32com.client


PLANT = "ESA1"
TRANSACTION = "ZTBV"


class SapError(RuntimeError):
    pass


@dataclass
class SapSession:
    session: object

    def find(self, oid):
        return self.session.findById(oid)


def attach() -> SapSession:
    """Attach to the first running SAP GUI session. Raise SapError if none."""
    try:
        sap_gui = win32com.client.GetObject("SAPGUI")
    except Exception as e:
        raise SapError(
            "Cannot reach SAP GUI. Log into SAP GUI first and confirm "
            "'Enable scripting' is on (Options -> Accessibility & Scripting)."
        ) from e
    try:
        app = sap_gui.GetScriptingEngine
        conn = app.Children(0)
        sess = conn.Children(0)
    except Exception as e:
        raise SapError(
            "SAP GUI is running but no active session was found. Open a "
            "connection and log in, then retry."
        ) from e
    return SapSession(sess)


def close_lingering_modals(s: SapSession, log=None) -> int:
    """Fix K: close any wnd[1..N] popups left over from a previous run.
    Typing an okcd on wnd[0] while a modal is open is silently ignored,
    which then causes the next few button-presses to hit the WRONG screen.
    Returns the number of modals actually closed (for diagnostic logging).
    """
    closed = 0
    for i in range(9, 0, -1):
        try:
            title = ""
            try:
                title = str(s.find(f"wnd[{i}]").Text or "")
            except Exception:
                pass
            s.find(f"wnd[{i}]").close()
            closed += 1
            if log:
                log(f"SAP: closed leftover modal wnd[{i}]" +
                    (f" (title: {title!r})" if title else ""))
        except Exception:
            pass
    return closed


def open_ztbv_table(s: SapSession, table: str, log=None) -> None:
    """Navigate to /nZTBV -> plant + table -> F8 (enter table view)."""
    if log:
        log(f"SAP: /n{TRANSACTION} -> {table} @ plant {PLANT}")
    n_closed = close_lingering_modals(s, log=log)
    if n_closed and log:
        log(f"SAP: {n_closed} lingering modal(s) closed before navigation")
    s.find("wnd[0]").maximize()
    s.find("wnd[0]/tbar[0]/okcd").Text = f"/n{TRANSACTION}"
    s.find("wnd[0]").sendVKey(0)
    time.sleep(0.3)
    s.find("wnd[0]/usr/txtD_WERKS").Text = PLANT
    s.find("wnd[0]/usr/ctxtD_TAB").Text = table
    s.find("wnd[0]/usr/ctxtD_TAB").SetFocus()
    s.find("wnd[0]/usr/ctxtD_TAB").caretPosition = len(table)
    s.find("wnd[0]/tbar[1]/btn[8]").press()  # F8 -> selection screen
    time.sleep(0.3)


_VALU_PUSH_RE = re.compile(r"btn%_(.+?)_%_APP_%-VALU_PUSH")
_SELOPT_INPUT_RE = re.compile(r"(S\d+)-(LOW|HIGH)$", re.I)


def _attr(c, name: str, default: str = "") -> str:
    try:
        v = getattr(c, name)
        return "" if v is None else str(v)
    except Exception:
        return default


def _int_attr(c, name: str, default: int = -1) -> int:
    try:
        return int(getattr(c, name))
    except Exception:
        return default


def _walk_controls(root, max_depth: int = 6) -> list:
    """Every control under `root`, depth-first. Individual failures are
    skipped -- a selection screen with one unreadable control should still
    yield the other 200."""
    out = []

    def walk(node, depth):
        if depth > max_depth:
            return
        try:
            children = node.Children
            count = int(children.Count)
        except Exception:
            return
        for i in range(count):
            try:
                c = children(i)
            except Exception:
                continue
            out.append(c)
            walk(c, depth + 1)

    walk(root, 0)
    return out


def screen_inventory(s, log=None) -> list[dict]:
    """Flat dump of every control on wnd[0]/usr, with both coordinate systems.

    This is the raw material for `describe_selection_screen` and for the
    Diagnose run -- when label pairing fails, the inventory is what lets a
    human map S<n> -> field by eye from the log file.
    """
    try:
        usr = s.find("wnd[0]/usr")
    except Exception as e:
        raise SapError(f"No selection screen on wnd[0]/usr: {e}") from e

    items = []
    for c in _walk_controls(usr):
        cid = _attr(c, "Id")
        items.append({
            "id": cid[cid.find("wnd[0]"):] if "wnd[0]" in cid else cid,
            "type": _attr(c, "Type"),
            "name": _attr(c, "Name"),
            "text": _attr(c, "Text").strip(),
            "tooltip": _attr(c, "Tooltip").strip(),
            # Character metric: exact per dynpro row/column. Only meaningful
            # for controls inside the user area -- which is all of these.
            "row": _int_attr(c, "CharTop"),
            "col": _int_attr(c, "CharLeft"),
            # Pixel metric: the fallback when CharTop/CharLeft read as 0 on
            # this SAP GUI build.
            "ptop": _int_attr(c, "Top"),
            "pleft": _int_attr(c, "Left"),
        })
    if log:
        log(f"SAP: selection screen carries {len(items)} control(s)")
    return items


def describe_selection_screen(s, log=None) -> list[dict]:
    """List every multi-value ('=>' arrow) filter on the current ZTBV
    selection screen, with the on-screen label next to each one.

    ZTBV names its select-options generically -- S3, S15, S29 -- so which
    slot is which field cannot be inferred from the id alone, and guessing
    would paste values into the wrong filter and return confidently wrong
    data. Run this once against a table to read the mapping off the screen
    instead of recording it by hand.

    Pairing runs in CHARACTER metric (CharTop/CharLeft), not pixels. A
    selection screen row reads [label] [LOW] [HIGH] [=> button], so the
    field name is the nearest label to the LEFT on the SAME dynpro row.
    Pixel `Top` was what the previous version used, and it does not agree
    between a label and a push button drawn on one row -- every filter came
    back unlabelled.

    Only GuiLabel controls are treated as labels. The select-option input
    fields on the same row are GuiTextField/GuiCTextField and hold retained
    FILTER VALUES, so accepting them would let a leftover TO number pose as
    a field name.

    Each filter's `tooltip` (read off its own -LOW field) is carried as an
    independent second source of the field's identity, used by
    `resolve_push_button` only when no label matched.

    Call it AFTER open_ztbv_table(), while the selection screen is up.
    Returns [{param, label, tooltip, push_id, low_field, row, col}].
    """
    items = screen_inventory(s, log=None)

    # CharTop/CharLeft are 0 on some SAP GUI builds for some control types.
    # Fall back to pixels wholesale rather than mixing the two metrics.
    use_char = any(it["row"] > 0 for it in items)
    row_of = (lambda it: it["row"]) if use_char else (lambda it: it["ptop"])
    col_of = (lambda it: it["col"]) if use_char else (lambda it: it["pleft"])
    # Row tolerance for the "near miss" pass: 1 dynpro row, or roughly one
    # row of pixels when we are stuck in pixel metric.
    row_slack = 1 if use_char else 12

    labels = [it for it in items if it["type"] == "GuiLabel" and it["text"]]
    if not labels:
        # No GuiLabel at all: widen to text controls that are NOT select-option
        # inputs, so a screen built entirely from output fields still resolves.
        # Noted in the log because these are a weaker signal than a real label.
        labels = [
            it for it in items
            if it["text"]
            and it["type"] in ("GuiTextField", "GuiCTextField")
            and not _SELOPT_INPUT_RE.search(it["name"] or "")
        ]
        if labels and log:
            log(f"SAP: no GuiLabel on this screen; falling back to "
                f"{len(labels)} non-input text control(s) as label sources")

    # Tooltip of each select-option's own -LOW field, keyed by S<n>.
    tooltip_by_param: dict[str, str] = {}
    low_field_by_param: dict[str, str] = {}
    for it in items:
        m = _SELOPT_INPUT_RE.search(it["name"] or "")
        if not m or m.group(2).upper() != "LOW":
            continue
        param = m.group(1).upper()
        low_field_by_param[param] = it["id"]
        if it["tooltip"]:
            tooltip_by_param[param] = it["tooltip"]

    def label_for(btn) -> str:
        brow, bcol = row_of(btn), col_of(btn)
        same_row = [l for l in labels if abs(row_of(l) - brow) == 0]
        left = [l for l in same_row if col_of(l) < bcol]
        if left:
            return max(left, key=col_of)["text"]
        if same_row:
            return min(same_row, key=col_of)["text"]
        near = [l for l in labels if abs(row_of(l) - brow) <= row_slack]
        if near:
            return min(near, key=lambda l: (abs(row_of(l) - brow), -col_of(l)))["text"]
        return ""

    out: list[dict] = []
    for it in items:
        m = _VALU_PUSH_RE.search(it["id"])
        if not m:
            continue
        param = m.group(1).upper()
        label = label_for(it)
        tooltip = tooltip_by_param.get(param, "")
        out.append({
            "param": param,
            "label": label or "(no label found)",
            "tooltip": tooltip,
            "push_id": it["id"],
            "low_field": low_field_by_param.get(
                param, f"wnd[0]/usr/ctxt{param}-LOW"),
            "row": row_of(it),
            "col": col_of(it),
        })
        if log:
            log(f"SAP: filter {param:<6} label={label or '(none)'!r}"
                + (f" tooltip={tooltip!r}" if tooltip else ""))
    if not out:
        raise SapError(
            "No multi-value filter buttons found. Is the ZTBV selection "
            "screen actually displayed (call open_ztbv_table first)?")
    if log:
        named = sum(1 for f in out if f["label"] != "(no label found)")
        log(f"SAP: {len(out)} filter(s) found, {named} with a label, "
            f"{sum(1 for f in out if f['tooltip'])} with a tooltip "
            f"({'character' if use_char else 'pixel'} metric)")
    return out


def paste_multi_value_filter(
    s: SapSession, push_button_id: str, values: list[str], log=None
) -> None:
    """Click multi-select push button, upload clipboard, OK.

    Caller is responsible for having already staged `values` on the OS
    clipboard (typically via Excel: put values in Column A of a temp
    workbook then .Copy()). This mirrors the notebook's approach.
    """
    if log:
        log(f"SAP: pasting {len(values)} filter values via clipboard")
    s.find(push_button_id).press()
    time.sleep(0.3)
    # Item 2: clear values retained from a previous chunk/run first -- SAP
    # keeps the selection dialog's contents within a session, and the
    # clipboard upload can append rather than replace. Deleting everything
    # up front makes the paste deterministic. btn[16] = "Delete Entire
    # Selection" on the standard multi-select dialog toolbar.
    try:
        s.find("wnd[1]/tbar[0]/btn[16]").press()
        time.sleep(0.2)
    except Exception:
        pass  # button absent on this SAP GUI version -- old behavior
    # In the multi-value dialog: btn[24] = Upload from Clipboard, btn[8] = OK
    s.find("wnd[1]/tbar[0]/btn[24]").press()
    time.sleep(0.3)
    s.find("wnd[1]/tbar[0]/btn[8]").press()
    time.sleep(0.2)


def fill_multi_value_filter_from_file(
    s: SapSession, push_button_id: str, values: list[str], work_dir: str,
    log=None,
) -> None:
    """Item 3: feed the multi-select filter from a temp TEXT FILE instead of
    the OS clipboard.

    The clipboard is global state -- a user copying anything while a run is
    in flight silently replaces the filter values (wrong results, not a
    crash). The multi-select dialog's 'Import from Text File' button
    (btn[23]) reads from a file only we control.

    Requires 'Show native Microsoft Windows dialogs' = Off (README end-user
    setup) so the file-open dialog is the scriptable SAP one (same
    DY_PATH / DY_FILENAME layout as the ALV export save dialog). Raises
    SapError on any failure -- after closing the dialogs it opened -- so
    the caller can fall back to the clipboard path.
    """
    os.makedirs(work_dir, exist_ok=True)
    filename = "esa_lookup_filter.txt"
    path = os.path.join(work_dir, filename)
    # One value per line, CRLF. SAP keys are plain ASCII; anything else
    # fails fast here and the caller drops to the clipboard path.
    try:
        with open(path, "w", encoding="ascii", newline="") as f:
            f.write("\r\n".join(str(v) for v in values))
            f.write("\r\n")
    except (OSError, UnicodeEncodeError) as e:
        raise SapError(f"cannot write filter file {path}: {e}") from e

    if log:
        log(f"SAP: importing {len(values)} filter values from file")
    s.find(push_button_id).press()
    time.sleep(0.3)
    try:
        # Clear leftover values first (same rationale as the clipboard
        # path: the dialog retains contents within a session).
        try:
            s.find("wnd[1]/tbar[0]/btn[16]").press()  # Delete Entire Selection
            time.sleep(0.2)
        except Exception:
            pass
        s.find("wnd[1]/tbar[0]/btn[23]").press()      # Import from Text File
        time.sleep(0.4)
        s.find("wnd[2]/usr/ctxtDY_PATH").Text = work_dir
        s.find("wnd[2]/usr/ctxtDY_FILENAME").Text = filename
        s.find("wnd[2]/tbar[0]/btn[0]").press()       # Open / OK
        time.sleep(0.3)
        s.find("wnd[1]/tbar[0]/btn[8]").press()       # Copy -> selection screen
        time.sleep(0.2)
    except Exception as e:
        # Leave the screen usable for the clipboard fallback: close the
        # file dialog and the multi-select dialog if still open.
        for wid in ("wnd[2]", "wnd[1]"):
            try:
                s.find(wid).close()
            except Exception:
                pass
        raise SapError(
            f"file import into the multi-select dialog failed: "
            f"{type(e).__name__}: {e}"
        ) from e


def execute_query(s: SapSession, log=None) -> None:
    """Press F8 on the selection screen to run the query."""
    if log:
        log("SAP: executing query (F8)")
    s.find("wnd[0]/tbar[1]/btn[8]").press()
    time.sleep(0.5)


def read_statusbar(s: SapSession) -> tuple[str, str]:
    """Return (message_type, text) from the main window's status bar.

    message_type is one of '' / 'S' (success) / 'W' (warning) / 'E' (error)
    / 'A' (abort) / 'I' (info). Returns ('', '') when the bar is empty or
    unreadable -- callers must treat that as "no message", not success.
    """
    try:
        sbar = s.find("wnd[0]/sbar")
        return (str(sbar.MessageType or "").strip().upper(),
                str(sbar.Text or "").strip())
    except Exception:
        return "", ""


def query_result_check(s: SapSession, log=None) -> int:
    """Item 2: after F8, confirm a result grid exists and return its row
    count (0 = query ran but matched nothing).

    Reads the status bar first: an 'E' (error) or 'A' (abort) message, or a
    missing result grid, raises SapError carrying SAP's own message -- so a
    bad query fails HERE with a precise reason instead of two calls later
    as a confusing export failure. Non-fatal messages are just logged.
    """
    msg_type, msg_text = read_statusbar(s)
    if msg_text and log:
        log(f"SAP status bar: [{msg_type or ' '}] {msg_text}")
    if msg_type in ("E", "A"):
        raise SapError(f"SAP reported an error after executing the query: {msg_text}")
    try:
        grid = s.find("wnd[0]/shellcont/shell")
        return int(grid.RowCount)
    except Exception as e:
        raise SapError(
            "No result grid appeared after executing the query"
            + (f" -- SAP said: {msg_text!r}" if msg_text else "")
            + ". Check the filter values, table name, and plant."
        ) from e


def export_alv_to_file(
    s: SapSession, target_dir: str, filename: str, log=None, timeout_s: int = 30
) -> str:
    """Export the current ALV grid to a spreadsheet file in `target_dir`.

    Tries the modern `&XXL` toolbar path first, then falls back to `&PC`
    (Save as Local File) with the spreadsheet format radio.

    Returns the full path to the exported file. Raises SapError on failure
    or timeout.
    """
    os.makedirs(target_dir, exist_ok=True)
    target_path = os.path.join(target_dir, filename)
    # Fix A: os.remove can silently fail (AV / lingering handle). Even if the
    # stale file survives, we only accept files whose mtime is newer than the
    # moment we triggered the export, so we can never return a prior step's
    # data as the current step's result.
    if os.path.exists(target_path):
        try:
            os.remove(target_path)
        except OSError:
            pass
    export_started_at = time.time()

    grid = s.find("wnd[0]/shellcont/shell")

    last_err: Exception | None = None
    for approach in ("XXL", "PC"):
        try:
            if approach == "XXL":
                if log:
                    log("SAP: exporting via &MB_EXPORT / &XXL")
                grid.pressToolbarContextButton("&MB_EXPORT")
                time.sleep(0.2)
                grid.selectContextMenuItem("&XXL")
            else:
                if log:
                    log("SAP: retrying export via &PC")
                grid.pressToolbarContextButton("&MB_EXPORT")
                time.sleep(0.2)
                grid.selectContextMenuItem("&PC")
                time.sleep(0.4)
                # Format picker -- select spreadsheet radio if present
                try:
                    s.find(
                        "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/"
                        "radSPOPLI-SELFLAG[1,0]"
                    ).Select()
                except Exception:
                    pass
                try:
                    s.find("wnd[1]/tbar[0]/btn[0]").press()
                except Exception:
                    pass
            time.sleep(0.5)
            # Save-As dialog: DY_PATH + DY_FILENAME + Save
            s.find("wnd[1]/usr/ctxtDY_PATH").Text = target_dir
            s.find("wnd[1]/usr/ctxtDY_FILENAME").Text = filename
            # btn[11] = "Replace" / Save on standard SAP save-as dialog
            s.find("wnd[1]/tbar[0]/btn[11]").press()
            # Wait for file, refusing any pre-existing stale copy (Fix A).
            deadline = time.time() + timeout_s
            while time.time() < deadline:
                if os.path.exists(target_path) and os.path.getsize(target_path) > 0:
                    try:
                        fresh = os.path.getmtime(target_path) >= export_started_at
                    except OSError:
                        fresh = False
                    if fresh:
                        if log:
                            log(f"SAP: export saved -> {target_path}")
                        return target_path
                time.sleep(0.25)
            raise SapError(f"Export timed out (>{timeout_s}s) waiting for {target_path}")
        except Exception as e:
            last_err = e
            if log:
                log(f"SAP: {approach} export attempt failed: "
                    f"{type(e).__name__}: {e}")
            # Best-effort: close any modal that might be hanging around so
            # the fallback attempt (or a subsequent step) can navigate.
            try:
                s.find("wnd[1]/tbar[0]/btn[12]").press()  # Cancel
                if log:
                    log("SAP: cancelled leftover modal to prepare fallback")
            except Exception:
                pass
            time.sleep(0.3)
            continue

    raise SapError(
        "ALV export failed via both &XXL and &PC. The exact save-dialog "
        "field IDs may differ on this SAP GUI version -- record one export "
        "via the SAP GUI Script Recorder and adjust `sap_ops.export_alv_to_file`."
    ) from last_err


def _visible_row_count(grid) -> int:
    """Rows the ALV control currently holds in the frontend buffer."""
    try:
        n = int(grid.VisibleRowCount)
        if n > 0:
            return n
    except Exception:
        pass
    return 20


def _scroll_grid(grid, row_number: int) -> bool:
    try:
        grid.firstVisibleRow = row_number
        return True
    except Exception:
        pass
    try:
        grid.VerticalScrollbar.Position = row_number
        return True
    except Exception:
        pass
    return False


def _safe_cell(grid, row: int, column: str) -> str:
    """GetCellValue with the original's scroll-and-retry: a row can fall out
    of the frontend buffer between the scroll and the read.
    """
    for _ in range(3):
        try:
            return grid.GetCellValue(row, column)
        except Exception:
            _scroll_grid(grid, row)
            time.sleep(0.2)
    return ""


def read_alv_grid(
    s: SapSession,
    columns: list[str],
    log=None,
    stop=None,
) -> list[dict]:
    """Read `columns` straight off the ALV grid, by TECHNICAL field name.

    This is the original notebook's read path (GetCellValue + scroll), and
    it is why that version never had to care what the layout titles its
    columns: 'TANUM' is 'TANUM' whether the ALV displays it as 'TO Number'
    or 'Transfer Order'. No alias table, no ambiguity when a layout repeats
    a title, no dependence on the export file format.

    Only the requested columns are fetched, so the cost is
    rows x len(columns) -- not rows x every field in the layout.

    Raises SapError if a name is not in the grid (wrong technical name, or
    the field is absent from the displayed variant) so the caller can fall
    back to the file export.
    """
    grid = s.find("wnd[0]/shellcont/shell")
    try:
        row_count = int(grid.RowCount)
    except Exception as e:
        raise SapError(f"ALV grid exposes no RowCount: {e}") from e
    if row_count == 0:
        return []

    # Fail fast on a bad column name: probe row 0 once per column before
    # committing to a full scroll-and-read pass.
    for c in columns:
        try:
            grid.GetCellValue(0, c)
        except Exception as e:
            raise SapError(
                f"ALV grid has no column {c!r} (technical name). Either the "
                f"field is missing from the displayed layout variant, or the "
                f"name differs on this system. Original error: {e}"
            ) from e

    visible = _visible_row_count(grid)
    if log:
        log(f"SAP: reading {row_count} row(s) x {len(columns)} column(s) "
            f"from the ALV grid by technical name")

    out: list[dict] = []
    for start in range(0, row_count, visible):
        if stop is not None and stop.is_set():
            break
        _scroll_grid(grid, start)
        time.sleep(0.05)
        for r in range(start, min(start + visible, row_count)):
            out.append({c: _safe_cell(grid, r, c) for c in columns})
    return out


# Multi-value push button IDs on the ZTBV selection screen per (table, field).
# These match the notebook. If a site uses a customized ZTBV layout the S<n>
# indices may shift -- record once with the Script Recorder and update here.
PUSH_BUTTONS = {
    ("LTAP", "TO_NUMBER"): "wnd[0]/usr/btn%_S3_%_APP_%-VALU_PUSH",
    ("Z50CFG_ENG_CRNT", "RSNUM"): "wnd[0]/usr/btn%_S15_%_APP_%-VALU_PUSH",
    ("Z50CFG_ENG_CRNT", "QMNUM"): "wnd[0]/usr/btn%_S29_%_APP_%-VALU_PUSH",
    ("Z50CFG_ENG_VALD", "OBJNR"): "wnd[0]/usr/btn%_S2_%_APP_%-VALU_PUSH",
}

# On-screen label fragments per logical field, all lowercase. Used by
# resolve_push_button to find a filter slot that PUSH_BUTTONS does not list
# yet -- ZTBV's S<n> parameter names are generic, but the label printed next
# to each filter names the field, so match on that.
FIELD_LABEL_SYNONYMS = {
    "TO_NUMBER": ["transfer order", "to number", "to no"],
    "RSNUM": ["reservation"],
    "QMNUM": ["notification", "notifctn"],
    "OBJNR": ["object"],
}


def push_button_id(param: str) -> str:
    """S7 -> the full multi-value push button id for that select-option."""
    return f"wnd[0]/usr/btn%_{param.upper()}_%_APP_%-VALU_PUSH"


def env_override_var(table: str, field: str) -> str:
    return "ESA_LOOKUP_PUSH_" + re.sub(
        r"[^A-Z0-9]+", "_", f"{table}_{field}".upper())


def _env_override(table: str, field: str) -> str | None:
    """Read a (table, field) -> filter mapping out of the environment.

    Lets whoever is standing at the customer's machine correct a mapping
    without a rebuild-and-redistribute cycle:

        set ESA_LOOKUP_PUSH_Z50CFG_ENG_CRNT_TO_NUMBER=S7

    Accepts either the bare select-option name ("S7") or a full control id.
    """
    raw = (os.environ.get(env_override_var(table, field)) or "").strip()
    if not raw:
        return None
    if re.fullmatch(r"S\d+", raw, re.I):
        return push_button_id(raw)
    return raw


def resolve_push_button(s, table: str, field: str, log=None) -> str:
    """Return the multi-value push button id for (table, field).

    Resolution order:
      1. An ESA_LOOKUP_PUSH_<TABLE>_<FIELD> environment override.
      2. PUSH_BUTTONS, the recorded mappings.
      3. The CURRENTLY DISPLAYED selection screen: the filter whose on-screen
         label matches the field's synonyms, else -- only if no label matched
         at all -- the filter whose tooltip does.

    Exactly one candidate must match. Zero or several raise SapError listing
    every filter on the screen, because pasting keys into the wrong filter
    would return plausible-looking wrong data -- the one failure mode worse
    than stopping.

    A resolved id is cached into PUSH_BUTTONS for the rest of the process.
    Must be called AFTER open_ztbv_table() so the screen is displayed.
    """
    key = (table, field)
    override = _env_override(table, field)
    if override:
        PUSH_BUTTONS[key] = override
        if log:
            log(f"SAP: {field} on {table} -> {override} "
                f"(from {env_override_var(table, field)})")
        return override
    if key in PUSH_BUTTONS:
        return PUSH_BUTTONS[key]

    synonyms = FIELD_LABEL_SYNONYMS.get(field)
    if not synonyms:
        raise SapError(
            f"No push button known for {key} and no label synonyms defined "
            f"for {field!r} -- add it to FIELD_LABEL_SYNONYMS or "
            f"PUSH_BUTTONS in sap_ops.")

    filters = describe_selection_screen(s, log=log)
    hits = [f for f in filters
            if any(syn in f["label"].lower() for syn in synonyms)]
    matched_on = "label"
    if not hits:
        # Tooltips are only consulted when no label matched -- consulting both
        # at once would turn a clean single label hit into an ambiguity error.
        hits = [f for f in filters
                if any(syn in f["tooltip"].lower() for syn in synonyms)]
        matched_on = "tooltip"
    if len(hits) != 1:
        listing = "\n".join(
            f"  {f['param']:<6} label={f['label']!r}"
            + (f" tooltip={f['tooltip']!r}" if f["tooltip"] else "")
            for f in filters)
        raise SapError(
            f"Could not resolve the {field} filter on {table}: "
            f"{len(hits)} candidate(s) matched {synonyms}. Filters on this "
            f"screen:\n{listing}\n"
            f"Fix (either one):\n"
            f"  - set {env_override_var(table, field)}=S<n> in the "
            f"environment and re-run -- no rebuild needed; or\n"
            f"  - add PUSH_BUTTONS[({table!r}, {field!r})] = "
            f"push_button_id('S<n>') in sap_ops.\n"
            f"If every label above is '(no label found)', run the app's "
            f"Diagnose button and send the log -- the raw screen dump names "
            f"which S<n> is which.")
    resolved = hits[0]["push_id"]
    PUSH_BUTTONS[key] = resolved
    if log:
        log(f"SAP: resolved {field} on {table} -> {hits[0]['param']} "
            f"by on-screen {matched_on} "
            f"({hits[0][matched_on]!r})")
    return resolved
