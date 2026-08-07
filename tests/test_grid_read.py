"""The ALV grid read, ported from the original notebook.

The original addressed SAP columns by TECHNICAL name via GetCellValue, which
is why it never needed an alias table: 'TANUM' is 'TANUM' whatever the
layout titles it. The export-based port had to guess titles, which produced
the 'Unl. Point' miss, the triplicate-'Item' ambiguity and the RSPOS
mis-binding. These tests pin the grid read as THE read path (the export
fallback was removed along with its pandas dependency).
"""
import pytest

from conftest import EXPECTED_TO, make_to_grid


class TestGridIsPrimary:
    def test_grid_read_used_and_no_export_happens(self, env):
        env.grid = make_to_grid()
        ok, _ = env.run("TO")
        assert ok
        kinds = [k for k, _ in env.events]
        assert kinds.count("gridread") == 3
        assert "export" not in kinds, "export should not run when the grid works"

    def test_values_match_the_export_path(self, env):
        """Same results as before the port -- this is a read swap, not a
        behavior change."""
        env.grid = make_to_grid()
        ok, _ = env.run("TO")
        assert ok
        env.assert_cells(EXPECTED_TO)

    def test_only_needed_columns_are_requested(self, env, monkeypatch):
        """Cost is rows x needed columns, not rows x the whole layout."""
        asked: list[list] = []
        import sap_ops
        orig = sap_ops.read_alv_grid

        def spy(sap, columns, log=None, stop=None):
            asked.append(list(columns))
            return orig(sap, columns, log=log, stop=stop)

        monkeypatch.setattr(sap_ops, "read_alv_grid", spy)
        env.grid = make_to_grid()
        ok, _ = env.run("TO")
        assert ok
        # Step 1 needs exactly TANUM (key) + ABLAD (output).
        assert asked[0] == ["TANUM", "ABLAD"]
        # Nothing asks for the whole table (step 2 is the widest: TANUM key
        # + 3 outputs + QMNUM/RSNUM/RSPOS extras = 7 of ~100 layout fields).
        assert all(len(cols) <= 7 for cols in asked), asked


class TestRowCountGuard:
    def test_short_grid_read_aborts_before_any_write(self, env):
        env.grid = make_to_grid()
        env.truncate_tables.add("Z50CFG_ENG_CRNT")   # status bar claims n+5
        seed = dict(env.grid)
        ok, logs = env.run("TO")
        assert not ok
        assert "row-count mismatch" in logs
        # The aborting step (2) must write nothing. Step 1 completed before
        # it, so its column M is salvaged -- see TestFailureAtomicity.
        changed = {k: v for k, v in env.grid.items() if seed.get(k) != v}
        assert changed == {(2, 13): "00000001000001", (3, 13): "00000002000002",
                           (2, 14): "100", (3, 14): "200",
                           (2, 15): "1", (3, 15): "2"}, (
            "only the completed step's columns may change on abort")


class TestGridIsTheOnlyPath:
    def test_grid_failure_is_a_clear_error_not_a_silent_switch(self, env):
        """The export-file fallback is gone (it existed for a case that
        never occurs on the ESA system). A grid problem must fail loudly
        with SAP's message, not silently change read methods."""
        env.grid = make_to_grid()
        env.sap_tables["LTAP"] = [{"TANUM": "T001"}]   # ABLAD missing
        ok, logs = env.run("TO")
        assert not ok
        assert "no column 'ABLAD'" in logs
