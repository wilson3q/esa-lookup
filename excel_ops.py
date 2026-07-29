"""Excel COM helpers -- attach to running Excel or open the workbook,
bulk read/write ranges, and stage values on the OS clipboard for SAP paste.
"""
from __future__ import annotations

import contextlib
import os
from dataclasses import dataclass

import win32com.client
from win32com.client import constants  # noqa: F401  (loaded lazily by pywin32)


XL_UP = -4162  # Excel: xlUp
XL_CALC_MANUAL = -4135
XL_CALC_AUTOMATIC = -4105


class ExcelError(RuntimeError):
    pass


@dataclass
class ExcelCtx:
    app: object
    book: object
    sheet: object


def attach(path: str) -> ExcelCtx:
    """Return an ExcelCtx for `path`. If Excel already has the file open,
    reuse it; otherwise start Excel and open it read/write.
    """
    if not os.path.exists(path):
        raise ExcelError(f"Excel file not found:\n{path}")

    abs_path = os.path.abspath(path).lower()

    app = None
    try:
        app = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        app = win32com.client.Dispatch("Excel.Application")
        app.Visible = True

    # Fix I: OneDrive/SharePoint-hosted workbooks report FullName as an
    # https:// URL, not a local filesystem path. os.path.abspath on a URL
    # produces gibberish and never matches -- so we then Open() a second
    # (read-only) copy and abort. Match by URL-tail filename in that case.
    target_basename = os.path.basename(path).lower()
    book = None
    for i in range(1, app.Workbooks.Count + 1):
        b = app.Workbooks(i)
        try:
            fn = str(b.FullName)
        except Exception:
            continue
        fn_lower = fn.lower()
        if fn_lower.startswith(("http:", "https:")):
            tail = fn.rsplit("/", 1)[-1].lower()
            if tail == target_basename:
                book = b
                break
        else:
            try:
                if os.path.abspath(fn).lower() == abs_path:
                    book = b
                    break
            except Exception:
                continue

    # Fix 7: always pass an absolute path -- Excel resolves relative paths
    # against its own cwd, which is usually not what the caller wants.
    if book is None:
        book = app.Workbooks.Open(os.path.abspath(path), 0, False)

    if book.ReadOnly:
        raise ExcelError(
            "Workbook is open Read-Only. Close it in Excel or check "
            "OneDrive/SharePoint permissions, then retry.\n" + path
        )

    return ExcelCtx(app=app, book=book, sheet=book.Worksheets(1))


def last_row_in_column(sheet, col_index: int) -> int:
    return int(sheet.Cells(sheet.Rows.Count, col_index).End(XL_UP).Row)


def read_range_2d(sheet, range_str: str) -> list[list]:
    """Read a range in one COM call; normalize to a list-of-lists.

    Single-cell ranges come back as a scalar; single-row/single-column
    ranges come back as a 1-tuple of tuples in some cases -- normalize both.
    """
    raw = sheet.Range(range_str).Value
    if raw is None:
        return []
    if not isinstance(raw, tuple):
        return [[raw]]
    # Range with a single row: raw is a tuple of scalars
    if raw and not isinstance(raw[0], tuple):
        return [list(raw)]
    return [list(r) for r in raw]


def write_range_2d(sheet, range_str: str, values_2d: list[list]) -> None:
    """Write a 2D python list to a range in one COM call.

    Excel's COM interface accepts a tuple-of-tuples on the RHS of Range.Value.
    Fix J: pre-check for merged cells so a mid-write failure can never leave
    the sheet in a half-updated state.
    """
    if not values_2d:
        return
    rng = sheet.Range(range_str)
    try:
        merged = bool(rng.MergeCells)
    except Exception:
        # `MergeCells` returns None (not a bool) for a mixed-merge range;
        # that itself signals merged content is present.
        merged = True
    if merged:
        raise ExcelError(
            f"Cannot write to range {range_str}: it contains merged cells. "
            f"Unmerge those columns in Excel and re-run. (SAP lookup output "
            f"columns must be plain, unmerged cells.)"
        )
    payload = tuple(tuple(row) for row in values_2d)
    rng.Value = payload


def clear_range(sheet, range_str: str) -> None:
    sheet.Range(range_str).ClearContents()


def set_column_format_text(sheet, columns: str) -> None:
    """columns like 'A:E' or 'M:M'."""
    sheet.Range(f"{columns}").NumberFormat = "@"


def autofit(sheet, columns: str) -> None:
    sheet.Columns(columns).AutoFit()


@contextlib.contextmanager
def bulk_write(app):
    """Disable Excel screen updating and calc while writing large blocks."""
    prev_updating = app.ScreenUpdating
    prev_events = app.EnableEvents
    prev_calc = app.Calculation
    try:
        app.ScreenUpdating = False
        app.EnableEvents = False
        app.Calculation = XL_CALC_MANUAL
        yield
    finally:
        app.Calculation = prev_calc
        app.EnableEvents = prev_events
        app.ScreenUpdating = prev_updating


def stage_values_on_clipboard(app, values: list[str]) -> object:
    """Create a scratch workbook, dump `values` into Column A as text,
    Copy() them onto the OS clipboard, and return the scratch workbook so
    the caller can close it after SAP has consumed the paste.
    """
    scratch = app.Workbooks.Add()
    ws = scratch.Worksheets(1)
    ws.Columns("A").NumberFormat = "@"
    # Fix 1: one bulk COM assignment instead of N per-cell writes. On a
    # 5000-key input this is the difference between ~30 s and <1 s.
    if values:
        payload = tuple(("" if v is None else str(v),) for v in values)
        ws.Range(f"A1:A{len(values)}").Value = payload
        ws.Range(f"A1:A{len(values)}").Copy()
    return scratch


def close_scratch(scratch) -> None:
    try:
        scratch.Close(SaveChanges=False)
    except Exception:
        pass


def save(book) -> None:
    """Save the workbook. Fix L: raise ExcelError on failure instead of
    silently swallowing it, so the pipeline can log a warning and the user
    knows to save manually rather than assuming success.
    """
    try:
        book.Save()
    except Exception as e:
        raise ExcelError(
            f"Excel refused to save the workbook: {e}. Data is written but "
            f"unsaved -- press Ctrl+S in Excel to persist it."
        ) from e
