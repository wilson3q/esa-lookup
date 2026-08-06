"""Mixed sheets: the either/or model, in miniature.

The live workbook mixes two kinds of rows: TO rows (K + reservation pair
N/O filled) and notification-only rows (just a notification in A). The TO
process must fill the TO rows all the way through, skip the rest without
failing, and -- critically -- leave column A of the notification-only rows
untouched, because the NOTIF process keys on it afterwards.
"""
import pipeline
from pipeline import LookupStep


def _mixed_grid():
    """4 data rows: two full TO rows, one TO row missing its reservation
    pair, one notification-only row."""
    g = {}
    for c in range(1, 17):
        g[(1, c)] = f"H{c}"
    # rows 2, 3: full TO rows (K + N + O)
    g[(2, 11)], g[(2, 14)], g[(2, 15)] = "T001", 100, 1
    g[(3, 11)], g[(3, 14)], g[(3, 15)] = "T002", 200, 2
    # row 4: TO number but no reservation pair -- step 1 serves it, step 2
    # must skip it (incomplete composite key)
    g[(4, 11)] = "T001"
    # row 5: notification-only row -- A filled, everything else empty
    g[(5, 1)] = "000429215427"
    return g


class TestMixedKeyRows:
    def test_run_succeeds_with_a_partial_sheet(self, env):
        env.grid = _mixed_grid()
        ok, logs = env.run("TO")
        assert ok, logs

    def test_full_to_rows_flow_through_all_three_steps(self, env):
        env.grid = _mixed_grid()
        ok, _ = env.run("TO")
        assert ok
        assert env.grid[(2, 13)] == "DockA"    # M: step 1
        assert env.grid[(2, 3)] == "OBJ1"      # C: step 2 (RSNUM|RSPOS)
        assert env.grid[(2, 6)] == "S1"        # F: step 3 keyed off C
        assert env.grid[(3, 13)] == "DockB"
        assert env.grid[(3, 3)] == "OBJ2"

    def test_notification_only_row_keeps_its_A(self, env):
        """The pivot of the either/or model: step 2 must NOT touch column A
        of a row it did not match -- the NOTIF process keys on it later."""
        env.grid = _mixed_grid()
        ok, _ = env.run("TO")
        assert ok
        assert env.grid[(5, 1)] == "000429215427"
        # ...and its output cells hold no TO-process data
        assert env.grid.get((5, 3)) in ("", None)
        assert env.grid.get((5, 13)) in ("", None)

    def test_partial_reservation_pair_is_skipped_not_fatal(self, env):
        env.grid = _mixed_grid()
        ok, logs = env.run("TO")
        assert ok
        # row 4: M filled by step 1 (K present)...
        assert env.grid[(4, 13)] == "DockA"
        # ...but step 2 skipped it (no reservation pair), so no OBJNR
        assert env.grid.get((4, 3)) in ("", None)

    def test_blanks_never_reach_the_sap_filter(self, env):
        env.grid = _mixed_grid()
        ok, _ = env.run("TO")
        assert ok
        for p in env.pastes():
            assert all(v not in ("", "NOT FOUND") for v in p), p


class TestCompositeKeyGuard:
    """The half-filled-composite skip logic, isolated."""

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
