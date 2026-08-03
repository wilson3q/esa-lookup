"""SAP GUI Scripting helpers for ZTBV table lookup.

All calls attach to an already-running SAP GUI session -- the user must be
logged in before the app is launched.
"""
from __future__ import annotations

import os
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


# Multi-value push button IDs on the ZTBV selection screen per (table, field).
# These match the notebook. If a site uses a customized ZTBV layout the S<n>
# indices may shift -- record once with the Script Recorder and update here.
PUSH_BUTTONS = {
    ("LTAP", "TO_NUMBER"): "wnd[0]/usr/btn%_S3_%_APP_%-VALU_PUSH",
    ("Z50CFG_ENG_CRNT", "RSNUM"): "wnd[0]/usr/btn%_S15_%_APP_%-VALU_PUSH",
    ("Z50CFG_ENG_CRNT", "QMNUM"): "wnd[0]/usr/btn%_S29_%_APP_%-VALU_PUSH",
    ("Z50CFG_ENG_VALD", "OBJNR"): "wnd[0]/usr/btn%_S2_%_APP_%-VALU_PUSH",
}
