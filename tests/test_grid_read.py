"""The ALV grid read, ported from the original notebook.

The original addressed SAP columns by TECHNICAL name via GetCellValue, which
is why it never needed an alias table: 'TANUM' is 'TANUM' whatever the
layout titles it. The export-based port had to guess titles, which produced
the 'Unl. Point' miss, the triplicate-'Item' ambiguity and the RSPOS
mis-binding. These tests pin the grid read as the primary path and the
export as the fallback.
"""
import pandas as pd
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
        # Nothing asks for the whole table.
        assert all(len(cols) <= 6 for cols in asked), asked


class TestExportFallback:
    def test_falls_back_when_grid_unavailable(self, env):
        env.no_grid_tables.add("LTAP")
        env.grid = make_to_grid()
        ok, logs = env.run("TO")
        assert ok
        assert "falling back to the file export" in logs
        kinds = [k for k, _ in env.events]
        assert "export" in kinds          # LTAP took the file path
        assert "gridread" in kinds        # later steps still used the grid

    def test_fallback_produces_the_same_values(self, env):
        env.no_grid_tables.update(
            {"LTAP", "Z50CFG_ENG_CRNT", "Z50CFG_ENG_VALD"})
        env.grid = make_to_grid()
        ok, _ = env.run("TO")
        assert ok
        assert [k for k, _ in env.events].count("export") == 3
        env.assert_cells(EXPECTED_TO)

    def test_missing_technical_name_falls_back_rather_than_failing(self, env):
        """A field absent from the displayed variant must not kill the run."""
        env.sap_tables["LTAP"] = pd.DataFrame({
            "TANUM": ["T001", "T002"], "ABLAD": ["DockA", "DockB"]})
        env.no_grid_tables.add("LTAP")
        env.grid = make_to_grid()
        ok, _ = env.run("TO")
        assert ok


class TestRowCountGuard:
    def test_short_grid_read_aborts_before_any_write(self, env):
        env.grid = make_to_grid()
        env.truncate_tables.add("Z50CFG_ENG_CRNT")   # status bar claims n+5
        seed = dict(env.grid)
        ok, logs = env.run("TO")
        assert not ok
        assert "row-count mismatch" in logs
        assert env.grid == seed, "workbook must be untouched on abort"
