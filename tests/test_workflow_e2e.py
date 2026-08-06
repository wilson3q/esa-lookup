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
        # steps 1 and 2 both key on K (TO numbers); step 3 keys on the OBJNR
        # values step 2 published in memory
        assert pastes == [["T001", "T002"], ["T001", "T002"], ["OBJ1", "OBJ2"]]
        # the sentinel garbage in column C must never reach SAP
        assert not any("STALE" in v for p in pastes for v in p)

    def test_all_writes_after_last_export(self, env):
        env.grid = make_to_grid()
        ok, _ = env.run("TO")
        assert ok
        kinds = [k for k, _ in env.events]
        last_export = max(i for i, k in enumerate(kinds) if k == "fetch")
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
    """A failed step must not cost you the steps that already succeeded.

    Gen 4 originally discarded every write when any step failed. That threw
    away correct work -- a run whose step 1 matched every row still left the
    workbook untouched because step 2 could not find its SAP filter. The
    contract now: completed steps are written and saved, steps that failed or
    never ran leave their columns exactly as they were, and the log says
    which is which.
    """

    def test_sap_error_writes_the_completed_steps_only(self, env):
        env.grid = make_to_grid()
        env.fail_table = "Z50CFG_ENG_VALD"      # last step fails
        ok, logs = env.run("TO")
        assert not ok
        assert "simulated SAP error" in logs

        # steps 1 and 2 completed -> their columns are written
        env.assert_cells({
            (2, 13): "DockA", (3, 13): "DockB",         # M   (step 1)
            (2, 3): "OBJ1", (3, 3): "OBJ2",             # C   (step 2)
            (2, 4): "MAT1", (2, 5): 5,                  # D,E (step 2)
            (2, 1): "QN1", (3, 1): "QN2",               # A   (step 2 extra)
            (2, 16): "T001", (3, 16): "T002",           # P   (step 2 audit)
        })
        # step 3 never ran -> F..I stay empty, J keeps its prior content
        for col in (6, 7, 8, 9):
            assert (2, col) not in env.grid and (3, col) not in env.grid
        assert env.grid[(2, 10)] == "keepJ2"
        assert env.grid[(3, 10)] == "keepJ3"
        assert "WRITTEN  step 1" in logs and "WRITTEN  step 2" in logs
        assert "NOT RUN  step 3" in logs

    def test_truncated_export_writes_only_the_completed_step(self, env):
        env.grid = make_to_grid()
        env.truncate_tables.add("Z50CFG_ENG_CRNT")   # step 2 fails
        seed = make_to_grid()
        ok, logs = env.run("TO")
        assert not ok
        assert "row-count mismatch" in logs
        # exactly step 1's column M changed; nothing else moved
        changed = {k: v for k, v in env.grid.items() if seed.get(k) != v}
        assert changed == {(2, 13): "DockA", (3, 13): "DockB"}
        # step 2's own key column keeps the sentinel garbage it started with
        assert env.grid[(2, 3)] == "STALE1"

    def test_failure_in_the_first_step_writes_nothing(self, env):
        env.grid = make_to_grid()
        env.fail_table = "LTAP"
        ok, logs = env.run("TO")
        assert not ok
        assert env.grid == make_to_grid()
        assert "NOT modified" in logs

    def test_dry_run_writes_nothing_even_when_a_later_step_fails(self, env):
        env.grid = make_to_grid()
        env.fail_table = "Z50CFG_ENG_VALD"
        ok, _ = env.run("TO", dry_run=True)
        assert not ok
        assert env.grid == make_to_grid()

    def test_salvaged_run_is_saved(self, env):
        env.grid = make_to_grid()
        env.fail_table = "Z50CFG_ENG_VALD"
        ok, _ = env.run("TO")
        assert not ok
        assert ("write", "SAVE") in env.events

    def test_rerunning_after_a_salvage_completes_the_sheet(self, env):
        """Salvage must leave a workbook a re-run can finish, not a mess."""
        env.grid = make_to_grid()
        env.fail_table = "Z50CFG_ENG_VALD"
        assert not env.run("TO")[0]
        env.fail_table = None
        ok, logs = env.run("TO")
        assert ok, logs
        env.assert_cells(EXPECTED_TO)

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
        assert env.pastes() == [["T001"], ["T002"], ["T001"], ["T002"],
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
        assert [k for k, _ in env.events].count("fetch") == 3
        assert "DRY RUN" in logs
        assert "2 matched / 0 unmatched" in logs
        assert "was not modified" in logs

    def test_dry_run_chain_still_in_memory(self, env):
        env.grid = make_to_grid()
        ok, _ = env.run("TO", dry_run=True)
        assert ok
        assert env.pastes()[2] == ["OBJ1", "OBJ2"]
