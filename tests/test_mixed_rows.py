"""Mixed sheets: some rows carry a key, some do not.

The ESA rule: rows missing their key are skipped, never fatal, and the rows
that DO have one must still be processed all the way through. Since the TO
rework, every TO step keys on column K, so "mixed" means K filled vs K
blank/'NOT FOUND' -- and the reservation pair N/O is now DERIVED into the
sheet for matched rows rather than required from the operator.
"""
import pipeline
from pipeline import ExtraOutput, LookupStep


def _mixed_grid():
    """4 rows: two with a TO number, one blank, one 'NOT FOUND'."""
    g = {}
    for c in range(1, 17):
        g[(1, c)] = f"H{c}"
    g[(2, 11)] = "T001"
    g[(3, 11)] = "T002"
    # row 4: K blank
    g[(5, 11)] = "NOT FOUND"
    return g


class TestMixedKeyRows:
    def test_run_succeeds_with_a_partial_sheet(self, env):
        env.grid = _mixed_grid()
        ok, logs = env.run("TO")
        assert ok, logs

    def test_keyed_rows_flow_through_all_three_steps(self, env):
        env.grid = _mixed_grid()
        ok, _ = env.run("TO")
        assert ok
        assert env.grid[(2, 13)] == "DockA"    # M: step 1
        assert env.grid[(2, 3)] == "OBJ1"      # C: step 2
        assert env.grid[(2, 6)] == "S1"        # F: step 3 keyed off C
        assert env.grid[(3, 13)] == "DockB"
        assert env.grid[(3, 3)] == "OBJ2"

    def test_reservation_pair_is_derived_not_required(self, env):
        """N/O start empty; step 2 fills them from SAP on match."""
        env.grid = _mixed_grid()
        ok, _ = env.run("TO")
        assert ok
        assert env.grid[(2, 14)] == 100 and env.grid[(2, 15)] == 1
        assert env.grid[(3, 14)] == 200 and env.grid[(3, 15)] == 2
        assert env.grid[(2, 1)] == "QN1"       # A: QMNUM derived on match

    def test_unkeyed_rows_are_ignored_not_fatal(self, env):
        env.grid = _mixed_grid()
        ok, _ = env.run("TO")
        assert ok
        for r in (4, 5):                       # blank and 'NOT FOUND'
            # "" = cleared, None = preserved-as-empty -- both mean no result
            assert env.grid.get((r, 3)) in ("", None)   # no OBJNR
            assert env.grid.get((r, 14)) in ("", None)  # no derived RSNUM
            assert env.grid.get((r, 6)) in ("", None)   # no Section

    def test_blanks_never_reach_the_sap_filter(self, env):
        """The demo behavior: blanks in the pasted list are ignored. We go
        one better and never send them at all."""
        env.grid = _mixed_grid()
        ok, _ = env.run("TO")
        assert ok
        for p in env.pastes():
            assert all(v not in ("", "NOT FOUND") for v in p), p


class TestCompositeKeyStillGuarded:
    """No shipping step keys on multiple columns since the TO rework, but
    the half-filled-composite skip logic must survive for any future one."""

    def _synthetic(self):
        return [LookupStep(
            name="synthetic N|O step",
            sap_table="Z50CFG_ENG_CRNT",
            push_button_field="RSNUM",
            key_columns=[14, 15],
            sap_key_columns=["RSNUM", "RSPOS"],
            sap_output_columns=["OBJNR"],
            excel_output_columns=[3],
        )]

    def test_half_filled_composite_is_skipped_and_named(self, env, monkeypatch):
        monkeypatch.setitem(pipeline.WORKFLOWS, "SYN", self._synthetic())
        g = {(1, c): f"H{c}" for c in range(1, 17)}
        g[(2, 14)], g[(2, 15)] = 100, 1        # complete
        g[(3, 14)] = 200                       # N filled, O blank
        env.grid = g
        ok, logs = env.run("SYN")
        assert ok, logs
        assert "skipped, since a complete key needs every one of" in logs
        assert env.pastes() == [["100"]]       # 200 never sent
        assert env.grid[(2, 3)] == "OBJ1"
        assert env.grid.get((3, 3), "") == ""
