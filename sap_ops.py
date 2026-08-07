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


def screen_snapshot(s: SapSession) -> str:
    """One plain-text snapshot of where SAP currently stands: main window
    title, any open popup titles, and the status bar message.

    Used by the error popup. The operator standing at the machine sees the
    frozen SAP screen; the person debugging remotely sees only what we
    capture here -- so grab everything cheap and never raise.
    """
    parts = []
    try:
        parts.append(f"window:  {str(s.find('wnd[0]').Text or '').strip()}")
    except Exception:
        parts.append("window:  (unreadable -- SAP GUI may be gone)")
    for i in (1, 2):
        try:
            title = str(s.find(f"wnd[{i}]").Text or "").strip()
            parts.append(f"popup wnd[{i}]: {title or '(untitled)'}")
        except Exception:
            pass  # no popup at this level -- the common case
    msg_type, msg_text = read_statusbar(s)
    if msg_text:
        parts.append(f"status bar: [{msg_type or ' '}] {msg_text}")
    return "\n".join(parts)


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
def push_button_id(param: str) -> str:
    """S15 -> the full id of that select-option's multi-value push button."""
    return f"wnd[0]/usr/btn%_{param.upper()}_%_APP_%-VALU_PUSH"


# The multi-value filter slot per (table, field) on the ZTBV selection
# screen. ZTBV names its select-options generically (S3, S15, ...), so
# these were read off the live screens once and recorded. If SAP's screen
# ever changes, run the Diagnose cell to list the current slots and edit
# the entry here -- this dict is the ONLY place a slot is defined.
#
# Known but unused, for reference: Z50CFG_ENG_CRNT also has a
# 'Transfer Order Number' filter at S19.
PUSH_BUTTONS = {
    ("LTAP", "TO_NUMBER"): push_button_id("S3"),
    ("Z50CFG_ENG_CRNT", "RSNUM"): push_button_id("S15"),
    ("Z50CFG_ENG_CRNT", "QMNUM"): push_button_id("S29"),
    ("Z50CFG_ENG_VALD", "OBJNR"): push_button_id("S2"),
}


def resolve_push_button(s, table: str, field: str, log=None) -> str:
    """The recorded filter slot for (table, field). Refuses -- with the fix
    spelled out -- rather than guessing: pasting keys into the wrong filter
    returns confidently wrong data, which is worse than stopping."""
    try:
        return PUSH_BUTTONS[(table, field)]
    except KeyError:
        raise SapError(
            f"No filter slot recorded for {field} on {table}. Run the "
            f"Diagnose cell to list the slots on that screen, then add the "
            f"entry to PUSH_BUTTONS (in the SAP utilities).") from None


def list_filter_slots(s: SapSession, log=None) -> list[str]:
    """Diagnose helper: every multi-value filter slot (S<n>) currently on
    screen, in screen order, with the texts on its row. Enough to map a
    slot to a field by eye; the operator screenshots the screen anyway."""
    out = []
    try:
        usr = s.find("wnd[0]/usr")
        children = usr.Children
        rows: dict[int, dict] = {}
        for i in range(int(children.Count)):
            c = children(i)
            try:
                cid, top = str(c.Id), int(c.Top)
            except Exception:
                continue
            row = rows.setdefault(top, {"param": "", "texts": []})
            if "VALU_PUSH" in cid:
                m = re.search(r"btn%_(.+?)_%_APP_%-VALU_PUSH", cid)
                row["param"] = m.group(1) if m else "?"
            else:
                try:
                    t = str(c.Text or "").strip()
                except Exception:
                    t = ""
                if t:
                    row["texts"].append(t)
        for top in sorted(rows):
            r = rows[top]
            if r["param"]:
                line = f"{r['param']:<6} {' | '.join(r['texts'])}"
                out.append(line)
                if log:
                    log(f"SAP: filter {line}")
    except Exception as e:
        raise SapError(f"Could not read the selection screen: {e}") from e
    if not out:
        raise SapError("No multi-value filter slots found on this screen.")
    return out
