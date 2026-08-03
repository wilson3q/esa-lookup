"""End-to-end workflow tests against the fake COM boundary (see conftest).

These pin the Gen 4 guarantees:
- in-memory step chaining (no Excel round-trip between steps)
- single deferred write-back (failures/Stop leave the workbook untouched)
- SAP self-checks (0-row results, truncated exports, chunking)
- temp-file filter transport with clipboard fallback
- dry-run mode
"""
import threading

import pipeline
from conftest import EXPECTED_NOTIF, EXPECTED_TO, make_notif_grid, make_to_grid


class TestToWorkflow:
    def test_happy_path_values(self, env):
        env.grid = make_to_grid()
        ok, logs = env.run("TO")
        assert ok, logs
        env.assert_cells(EXPECTED_TO)

    def test_step3_keys_come_from_memory(self, env):
        env.grid = make_to_grid()
        ok, _ = env.run("TO")
        assert ok
        pastes = env.pastes()
        assert pastes == [["T001", "T002"], ["100", "200"], ["OBJ1", "OBJ2"]]
        # the sentinel garbage in column C must never reach SAP
        assert not any("STALE" in v for p in pastes for v in p)

    def test_all_writes_after_last_export(self, env):
        env.grid = make_to_grid()
        ok, _ = env.run("TO")
        assert ok
        kinds = [k for k, _ in env.events]
        last_export = max(i for i, k in enumerate(kinds) if k == "export")
        first_mutation = min(
            i for i, k in enumerate(kinds) if k in ("clear", "write"))
        assert first_mutation > last_export
        assert env.events[-1] == ("write", "SAVE")

    def test_happy_path_never_touches_clipboard(self, env):
        env.grid = make_to_grid()
        ok, _ = env.run("TO")
        assert ok
        assert env.clipboard_stages == 0


class TestNotifWorkflow:
    def test_happy_path_values(self, env):
        env.grid = make_notif_grid()
        ok, logs = env.run("NOTIF")
        assert ok, logs
        env.assert_cells(EXPECTED_NOTIF)

    def test_step2_keys_come_from_memory(self, env):
        env.grid = make_notif_grid()
        ok, _ = env.run("NOTIF")
        assert ok
        assert env.pastes() == [["QN1", "QN2"], ["OBJ1", "OBJ2"]]


class TestFailureAtomicity:
    def test_sap_error_leaves_sheet_untouched(self, env):
        env.grid = make_to_grid()
        env.fail_table = "Z50CFG_ENG_VALD"      # last step fails
        ok, logs = env.run("TO")
        assert not ok
        assert env.grid == make_to_grid()
        assert "simulated SAP error" in logs
        assert "NOT modified" in logs

    def test_truncated_export_aborts_before_write(self, env):
        env.grid = make_to_grid()
        env.truncate_tables.add("Z50CFG_ENG_CRNT")
        ok, logs = env.run("TO")
        assert not ok
        assert env.grid == make_to_grid()
        assert "row-count mismatch" in logs

    def test_stop_before_run_leaves_sheet_untouched(self, env):
        env.grid = make_to_grid()
        stop = threading.Event()
        stop.set()
        ok, logs = env.run("TO", stop_event=stop)
        assert not ok
        assert env.grid == make_to_grid()
        assert "cancelled" in logs


class TestSapSelfChecks:
    def test_chunking_same_result_one_paste_per_key(self, env, monkeypatch):
        monkeypatch.setattr(pipeline, "SAP_FILTER_CHUNK_SIZE", 1)
        env.grid = make_to_grid()
        ok, logs = env.run("TO")
        assert ok, logs
        env.assert_cells(EXPECTED_TO)
        assert env.pastes() == [["T001"], ["T002"], ["100"], ["200"],
                                ["OBJ1"], ["OBJ2"]]
        assert "2 chunk(s) of <= 1" in logs

    def test_zero_row_result_is_clean_unmatched(self, env):
        env.grid = make_to_grid()
        env.zero_tables.add("LTAP")
        ok, logs = env.run("TO")
        assert ok, logs
        assert env.grid.get((2, 13)) == "" and env.grid.get((3, 13)) == ""
        assert env.grid.get((2, 3)) == "OBJ1"   # later steps still ran
        assert "returned 0 rows" in logs
        assert "no rows for any" in logs


class TestFilterTransport:
    def test_clipboard_fallback(self, env):
        env.grid = make_to_grid()
        env.file_import_fails = True
        ok, logs = env.run("TO")
        assert ok, logs
        env.assert_cells(EXPECTED_TO)
        assert env.clipboard_stages == 3        # one per step
        assert "falling back to clipboard paste" in logs


class TestDryRun:
    def test_dry_run_reports_but_writes_nothing(self, env):
        env.grid = make_to_grid()
        ok, logs = env.run("TO", dry_run=True)
        assert ok, logs
        assert env.grid == make_to_grid()       # byte-identical
        # SAP side ran fully (all three exports)
        assert [k for k, _ in env.events].count("export") == 3
        assert "DRY RUN" in logs
        assert "2 matched / 0 unmatched" in logs
        assert "was not modified" in logs

    def test_dry_run_chain_still_in_memory(self, env):
        env.grid = make_to_grid()
        ok, _ = env.run("TO", dry_run=True)
        assert ok
        assert env.pastes()[2] == ["OBJ1", "OBJ2"]
