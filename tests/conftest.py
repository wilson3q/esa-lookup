"""Shared test fixtures.

Stubs the Windows-only COM modules (win32com, pythoncom) BEFORE the app
modules are imported, then provides a fake Excel sheet + fake SAP so the
full pipeline can run end-to-end on any OS -- including Linux CI.

pywin32 is intentionally never used even when installed: tests must be
deterministic and must not talk to a real Excel or SAP GUI.
"""
from __future__ import annotations

import contextlib
import os
import re
import sys
import threading
import types


def _stub_com_modules() -> None:
    if "win32com.client" in sys.modules and not isinstance(
            sys.modules["win32com.client"], types.ModuleType):
        return
    win32com = types.ModuleType("win32com")
    client = types.ModuleType("win32com.client")
    client.GetObject = lambda *a, **k: None
    client.Dispatch = lambda *a, **k: None
    client.GetActiveObject = lambda *a, **k: None
    client.constants = object()
    win32com.client = client
    sys.modules["win32com"] = win32com
    sys.modules["win32com.client"] = client
    pythoncom = types.ModuleType("pythoncom")
    pythoncom.CoInitialize = lambda: None
    pythoncom.CoUninitialize = lambda: None
    sys.modules["pythoncom"] = pythoncom


_stub_com_modules()
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pandas as pd  # noqa: E402
import pytest  # noqa: E402

import excel_ops  # noqa: E402
import sap_ops  # noqa: E402
import pipeline  # noqa: E402


ROWS_COUNT = 1000  # stand-in for sheet.Rows.Count in clear semantics


# ---------------------------------------------------------------------------
# Fake Excel sheet: a dict {(row, col): value} plus the minimal COM surface
# pipeline touches directly (Cells(...).Value and Rows.Count).
# ---------------------------------------------------------------------------

class FakeCell:
    def __init__(self, env, r, c):
        self.env, self.r, self.c = env, r, c

    @property
    def Value(self):
        return self.env.grid.get((self.r, self.c))

    @Value.setter
    def Value(self, v):
        self.env.events.append(("write", f"cell({self.r},{self.c})"))
        self.env.grid[(self.r, self.c)] = v


class FakeSheet:
    def __init__(self, env):
        self.env = env
        self.Rows = types.SimpleNamespace(Count=ROWS_COUNT)

    def Cells(self, r, c):
        return FakeCell(self.env, r, c)


def col_num(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def parse_range(rng: str):
    a, b = rng.split(":")
    m1 = re.match(r"([A-Z]+)(\d+)", a)
    m2 = re.match(r"([A-Z]+)(\d+)", b)
    return (int(m1.group(2)), col_num(m1.group(1)),
            int(m2.group(2)), col_num(m2.group(1)))


# ---------------------------------------------------------------------------
# Reference SAP data used by the fakes (fresh copy per FakeEnv)
# ---------------------------------------------------------------------------

def default_sap_tables() -> dict[str, pd.DataFrame]:
    return {
        "LTAP": pd.DataFrame({
            # ESA-encoded unloading points: 10-digit reservation + 4-digit
            # item. Splitting gives 100|1 and 200|2 -- the keys CRNT holds.
            "TANUM": ["T001", "T002"],
            "ABLAD": ["00000001000001", "00000002000002"]}),
        "Z50CFG_ENG_CRNT": pd.DataFrame({
            "TANUM": ["T001", "T002"],
            "QMNUM": ["QN1", "QN2"], "RSNUM": [100, 200], "RSPOS": [1, 2],
            "OBJNR": ["OBJ1", "OBJ2"], "DISP_MATNR": ["MAT1", "MAT2"],
            "DISP_QTY": [5, 7]}),
        "Z50CFG_ENG_VALD": pd.DataFrame({
            "OBJNR": ["OBJ1", "OBJ2"], "Z_SECTION": ["S1", "S2"],
            "Z_MODULE": ["M1", "M2"], "DESCRIPT": ["D1", "D2"],
            "SALES_ORDER": ["SO1", "SO2"], "LID": ["L1", "L2"]}),
    }


def make_to_grid() -> dict:
    """Workbook seed for the TO workflow. Column C holds sentinel garbage
    that must NEVER reach SAP (step 3 must key off memory, not the sheet)."""
    g = {}
    for c in range(1, 17):
        g[(1, c)] = f"H{c}"                      # headers
    g[(2, 11)], g[(3, 11)] = "T001", "T002"      # K: TO numbers
    g[(2, 14)], g[(3, 14)] = "STALE-N", "STALE-N"  # N: derived by step 1
    g[(2, 15)], g[(3, 15)] = "STALE-O", "STALE-O"  # O: derived by step 1
    g[(2, 3)], g[(3, 3)] = "STALE1", "STALE2"    # C: sentinel garbage
    g[(2, 1)], g[(3, 1)] = "keepA2", "keepA3"    # A: pre-existing
    g[(2, 10)], g[(3, 10)] = "keepJ2", "keepJ3"  # J: pre-existing
    return g


def make_notif_grid() -> dict:
    g = {}
    for c in range(1, 17):
        g[(1, c)] = f"H{c}"
    g[(2, 1)], g[(3, 1)] = "QN1", "QN2"          # A: notification numbers
    g[(2, 3)], g[(3, 3)] = "STALE1", "STALE2"    # C: sentinel garbage
    g[(2, 10)], g[(3, 10)] = "oldJ2", "oldJ3"    # J: overwritten by LID
    return g


EXPECTED_TO = {
    (2, 13): "00000001000001", (3, 13): "00000002000002",  # M: ABLAD
    (2, 1): "QN1", (3, 1): "QN2",                  # A: QMNUM derived on match
    (2, 14): "100", (3, 14): "200",                # N: split out of ABLAD
    (2, 15): "1", (3, 15): "2",                    # O: split out of ABLAD
    (2, 3): "OBJ1", (3, 3): "OBJ2",                # C from step 2
    (2, 4): "MAT1", (2, 5): 5,                     # D, E
    (2, 6): "S1", (2, 7): "M1", (2, 8): "D1", (2, 9): "SO1",  # F..I
    (2, 10): "keepJ2", (3, 10): "keepJ3",          # J preserved on match
    (2, 16): "100|1", (3, 16): "200|2",            # P audit key (RSNUM|RSPOS)
    (1, 16): "Excel Match Key Used",
}

EXPECTED_NOTIF = {
    (2, 3): "OBJ1", (3, 3): "OBJ2",                # C
    (2, 4): "MAT1", (2, 5): 5,                     # D, E
    (2, 6): "S1", (2, 7): "M1", (2, 8): "D1", (2, 9): "SO1",  # F..I
    (2, 10): "L1", (3, 10): "L2",                  # J = LID on match
    (2, 1): "QN1", (3, 1): "QN2",                  # A untouched (key column)
}


# ---------------------------------------------------------------------------
# FakeEnv: installs all COM-boundary fakes via pytest's monkeypatch (so
# every test gets a pristine module state) and drives pipeline.run().
# ---------------------------------------------------------------------------

class FakeEnv:
    def __init__(self, monkeypatch):
        self.grid: dict = {}
        self.events: list = []          # ("export"|"paste"|"clear"|"write", detail)
        self.sap_tables = default_sap_tables()
        self.current_table: str | None = None
        self.fail_table: str | None = None       # query_result_check raises
        self.zero_tables: set = set()            # query returns 0 rows
        self.truncate_tables: set = set()        # grid claims more than export
        self.no_grid_tables: set = set()         # force the export fallback
        self.file_import_fails = False           # simulate unscriptable dialog
        self.clipboard_stages = 0
        self.sheet = FakeSheet(self)
        self._install(monkeypatch)

    # -- fake COM boundary -------------------------------------------------
    def _install(self, mp):
        env = self

        def last_row(sheet, col):
            rows = [r for (r, c), v in env.grid.items()
                    if c == col and str(v or "").strip() != ""]
            return max(rows) if rows else 1

        def read(sheet, rng):
            r1, c1, r2, c2 = parse_range(rng)
            return [[env.grid.get((r, c)) for c in range(c1, c2 + 1)]
                    for r in range(r1, r2 + 1)]

        def write(sheet, rng, vals):
            if getattr(env, "write_fails", False):
                raise RuntimeError("simulated Excel write failure")
            env.events.append(("write", rng))
            r1, c1, r2, c2 = parse_range(rng)
            for i, r in enumerate(range(r1, r2 + 1)):
                for j, c in enumerate(range(c1, c2 + 1)):
                    env.grid[(r, c)] = vals[i][j]

        def clear(sheet, rng):
            env.events.append(("clear", rng))
            r1, c1, r2, c2 = parse_range(rng)
            for (r, c) in list(env.grid):
                if r1 <= r <= min(r2, ROWS_COUNT) and c1 <= c <= c2:
                    del env.grid[(r, c)]

        def stage(app, values):
            env.clipboard_stages += 1
            return object()

        def open_table(s, table, log=None):
            env.current_table = table

        def fill_from_file(s, pid, values, work_dir, log=None):
            if env.file_import_fails:
                raise sap_ops.SapError("simulated: import dialog not scriptable")
            env.events.append(("paste", list(values)))

        def paste_clipboard(s, pid, values, log=None):
            env.events.append(("paste", list(values)))

        def check(s, log=None):
            t = env.current_table
            if t == env.fail_table:
                raise sap_ops.SapError(f"simulated SAP error on {t}")
            if t in env.zero_tables:
                return 0
            n = len(env.sap_tables[t])
            return n + 5 if t in env.truncate_tables else n

        def read_grid(sap, columns, log=None, stop=None):
            """Primary path: the ALV grid, addressed by technical name."""
            table = env.current_table
            if table in env.no_grid_tables:
                raise sap_ops.SapError("simulated: grid read unavailable")
            df = env.sap_tables[table]
            missing = [c for c in columns if c not in df.columns]
            if missing:
                raise sap_ops.SapError(
                    f"ALV grid has no column {missing[0]!r} (technical name)")
            env.events.append(("fetch", table))
            env.events.append(("gridread", table))
            return df[list(columns)].to_dict("records")

        def export(sap, target_dir, filename, log=None, timeout_s=30):
            table = env.current_table
            env.events.append(("fetch", table))
            env.events.append(("export", table))
            os.makedirs(target_dir, exist_ok=True)
            path = os.path.join(target_dir, filename)
            env.sap_tables[table].to_excel(path, index=False)
            return path

        mp.setattr(excel_ops, "last_row_in_column", last_row)
        mp.setattr(excel_ops, "read_range_2d", read)
        mp.setattr(excel_ops, "write_range_2d", write)
        mp.setattr(excel_ops, "clear_range", clear)
        mp.setattr(excel_ops, "set_column_format_text", lambda s, c: None)
        mp.setattr(excel_ops, "bulk_write", lambda app: contextlib.nullcontext())
        mp.setattr(excel_ops, "stage_values_on_clipboard", stage)
        mp.setattr(excel_ops, "close_scratch", lambda s: None)
        mp.setattr(excel_ops, "save",
                   lambda b: env.events.append(("write", "SAVE")))
        mp.setattr(excel_ops, "attach",
                   lambda p: excel_ops.ExcelCtx(app=None, book=None, sheet=env.sheet))
        mp.setattr(sap_ops, "attach", lambda: sap_ops.SapSession(None))
        mp.setattr(sap_ops, "open_ztbv_table", open_table)
        mp.setattr(sap_ops, "fill_multi_value_filter_from_file", fill_from_file)
        mp.setattr(sap_ops, "paste_multi_value_filter", paste_clipboard)
        mp.setattr(sap_ops, "execute_query", lambda s, log=None: None)
        mp.setattr(sap_ops, "query_result_check", check)
        mp.setattr(sap_ops, "read_alv_grid", read_grid)
        mp.setattr(sap_ops, "resolve_push_button",
                   lambda s, table, field, log=None: f"btn_{table}_{field}")
        mp.setattr(sap_ops, "export_alv_to_file", export)

    # -- helpers -----------------------------------------------------------
    def run(self, workflow="TO", dry_run=False, stop_event=None,
            step_popups=False):
        """Run the pipeline; returns (ok, joined_log_text).

        Popup events are recorded in self.popups. A "step" popup (carrying
        an ack Event) is auto-acknowledged immediately -- with OK by
        default, or Cancel once self.popup_cancel_at matches its 1-based
        order -- because nobody is around to click a message box in a test.
        """
        self.events.clear()
        self.popups: list = []
        cfg = pipeline.RunConfig(
            excel_path=__file__, workflow=workflow,
            stop_event=stop_event or threading.Event(), dry_run=dry_run,
            step_popups=step_popups,
        )
        logs: list = []

        def on_event(kind, payload):
            logs.append((kind, payload))
            if kind == "popup":
                self.popups.append(payload)
                if payload.get("ack") is not None:
                    n_step = sum(1 for p in self.popups
                                 if p.get("ack") is not None)
                    if payload.get("result") is not None:
                        payload["result"]["proceed"] = (
                            n_step != getattr(self, "popup_cancel_at", 0))
                    payload["ack"].set()

        ok = pipeline.run(cfg, on_event)
        joined = " | ".join(p[0] for k, p in logs if k == "log")
        return ok, joined

    def pastes(self):
        return [p for kind, p in self.events if kind == "paste"]

    def assert_cells(self, expected: dict):
        for k, v in expected.items():
            assert self.grid.get(k) == v, (
                f"{k}: expected {v!r}, got {self.grid.get(k)!r}")


@pytest.fixture
def env(monkeypatch) -> FakeEnv:
    return FakeEnv(monkeypatch)
