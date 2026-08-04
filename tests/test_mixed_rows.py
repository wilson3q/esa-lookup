"""Mixed sheets: some rows carry a complete key, some do not.

The ESA rule: rows missing a key value are skipped, never fatal, and the
rows that DO have values must still be processed all the way through steps
2 and 3. A half-filled sheet is normal input, not an error.
"""
from conftest import make_to_grid


def _mixed_grid():
    """3 rows: row 2 complete, row 3 missing O, row 4 missing N and O.

    All three carry a TO number, so step 1 processes every one of them --
    only step 2 (keyed N|O) sees the difference.
    """
    g = {}
    for c in range(1, 17):
        g[(1, c)] = f"H{c}"
    for r, to in ((2, "T001"), (3, "T002"), (4, "T001")):
        g[(r, 11)] = to                       # K: TO numbers
    g[(2, 14)], g[(2, 15)] = 100, 1           # row 2: complete N|O
    g[(3, 14)] = 200                          # row 3: N only, O blank
    # row 4: neither
    return g


class TestMixedKeyCompleteness:
    def test_run_succeeds_with_a_partial_sheet(self, env):
        env.grid = _mixed_grid()
        ok, logs = env.run("TO")
        assert ok, logs

    def test_complete_row_flows_through_steps_2_and_3(self, env):
        env.grid = _mixed_grid()
        ok, _ = env.run("TO")
        assert ok
        # Row 2 has N=100, O=1 -> matches RSNUM 100 / RSPOS 1 in the fake SAP
        assert env.grid[(2, 3)] == "OBJ1"      # C: OBJNR from step 2
        assert env.grid[(2, 4)] == "MAT1"      # D
        assert env.grid[(2, 6)] == "S1"        # F: step 3 keyed off C
        assert env.grid[(2, 7)] == "M1"        # G

    def test_incomplete_rows_are_skipped_not_fatal(self, env):
        env.grid = _mixed_grid()
        ok, _ = env.run("TO")
        assert ok
        for r in (3, 4):
            # Either cleared or never touched -- both mean "no result", and
            # which one depends on how far each step's row range reached.
            assert env.grid.get((r, 3), "") == ""      # C: no OBJNR
            assert env.grid.get((r, 6), "") == ""      # F: no Section

    def test_partial_key_is_reported_as_skipped_not_unmatched(self, env):
        """Row 3 (N filled, O blank) must not be sent to SAP at all."""
        env.grid = _mixed_grid()
        ok, logs = env.run("TO")
        assert ok
        assert "skipped, since a complete key needs every one of" in logs
        # 200 is row 3's reservation number -- it must never reach the filter
        step2_pastes = [p for p in env.pastes() if "200" in [str(v) for v in p]]
        assert not step2_pastes, f"partial-key value was sent to SAP: {env.pastes()}"

    def test_step1_still_processes_every_row(self, env):
        """Step 1 keys on K alone, so the N/O gaps are irrelevant to it."""
        env.grid = _mixed_grid()
        ok, _ = env.run("TO")
        assert ok
        assert env.grid[(2, 13)] == "DockA"    # M
        assert env.grid[(3, 13)] == "DockB"
        assert env.grid[(4, 13)] == "DockA"
