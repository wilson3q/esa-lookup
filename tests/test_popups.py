"""Popup events: the original notebook's message-box behavior, reproduced.

The ESA operator troubleshoots by screenshotting a blocking popup together
with the SAP screen -- the original notebook ended every step and every
failure with one. The pipeline emits ("popup", {...}) events for the GUI to
render as modal boxes: "step" popups (ack Event, OK/Cancel) after each step
when RunConfig.step_popups is on, and fire-and-forget "info"/"error" popups
for completion / cancellation / failure regardless of that flag.
"""
import pipeline
from conftest import make_to_grid


def step_popups(env):
    return [p for p in env.popups if p.get("ack") is not None]


class TestStepPopups:
    def test_one_blocking_popup_per_step_plus_completion(self, env):
        env.grid = make_to_grid()
        ok, _ = env.run("TO", step_popups=True)
        assert ok
        assert len(step_popups(env)) == 3          # one per TO step
        final = env.popups[-1]
        assert final.get("ack") is None and final["kind"] == "info"
        assert final["title"] == "Completed"

    def test_step_popup_carries_the_notebook_counts(self, env):
        env.grid = make_to_grid()
        env.run("TO", step_popups=True)
        msg = step_popups(env)[0]["message"]
        assert "SAP LTAP rows detected: 2" in msg
        assert "Matched: 2" in msg
        assert "Not matched: 0" in msg
        assert "column(s) K" in msg                # where keys came from
        assert "M" in msg                          # where results go

    def test_step_results_are_in_the_workbook_at_popup_time(self, env):
        """Step-by-step mode IS the original notebook: the step is written
        and SAVED before its box, so the operator can switch to Excel and
        inspect real cells while the box waits."""
        env.grid = make_to_grid()
        seen_at_popup = []
        env.popup_probe = lambda p: seen_at_popup.append(
            (env.grid.get((2, 13)), ("write", "SAVE") in env.events))
        ok, _ = env.run("TO", step_popups=True)
        assert ok
        # at the FIRST popup, step 1's column M was already written+saved
        assert seen_at_popup[0] == ("00000001000001", True)

    def test_cancel_on_a_step_popup_keeps_the_written_steps(self, env):
        env.grid = make_to_grid()
        env.popup_cancel_at = 1                    # press Cancel on popup 1
        ok, logs = env.run("TO", step_popups=True)
        assert not ok
        assert "cancelled" in logs
        # step 1 was written (step-by-step mode) and STAYS...
        assert env.grid[(2, 13)] == "00000001000001"
        # ...steps 2 and 3 never ran
        assert env.grid[(2, 3)] == "STALE1"
        assert len(step_popups(env)) == 1
        # and the box said so
        assert "steps written so far stay" in step_popups(env)[0]["message"]

    def test_off_by_default_no_blocking_popups(self, env):
        env.grid = make_to_grid()
        ok, _ = env.run("TO")
        assert ok
        assert step_popups(env) == []
        # ...but the completion popup still shows
        assert any(p["title"] == "Completed" for p in env.popups)


class TestErrorPopup:
    def test_failure_pops_an_error_box_with_the_essentials(self, env):
        env.grid = make_to_grid()
        env.fail_table = "Z50CFG_ENG_VALD"
        ok, _ = env.run("TO")
        assert not ok
        errors = [p for p in env.popups if p["kind"] == "error"]
        assert len(errors) == 1
        msg = errors[0]["message"]
        assert "Step 3 of 3 (Z50CFG_ENG_VALD)" in msg   # which step died
        assert "simulated SAP error" in msg             # SAP's own words
        assert "screenshot" in msg.lower()              # what to do next
        # the workbook state must match what the log claims (salvage ran)
        assert "completed steps ONLY" in msg

    def test_error_popup_names_the_action_in_flight(self, env):
        """A com_error alone says nothing about WHERE; the popup must say
        what the app was doing -- 'reading keys' points at Excel, 'sending
        filter values' at SAP -- so a screenshot alone can be diagnosed."""
        env.grid = make_to_grid()
        env.fail_table = "Z50CFG_ENG_CRNT"
        env.run("TO")
        msg = [p for p in env.popups if p["kind"] == "error"][0]["message"]
        assert "While doing:" in msg

    def test_salvage_failure_is_named_in_the_error_popup(self, env):
        """Two failures in one run (the step AND the salvage write) usually
        share one cause -- the popup must show both, not silently fold the
        second into a 'partially updated' sentence."""
        env.grid = make_to_grid()
        env.fail_table = "Z50CFG_ENG_VALD"     # step 3 dies...
        env.write_fails = True                  # ...and so does the salvage
        ok, _ = env.run("TO")
        assert not ok
        msg = [p for p in env.popups if p["kind"] == "error"][0]["message"]
        assert "ALSO: writing the completed steps to Excel failed" in msg
        assert "simulated Excel write failure" in msg
        assert "Close the workbook WITHOUT saving" in msg

    def test_error_popup_fires_even_with_step_popups_off(self, env):
        env.grid = make_to_grid()
        env.fail_table = "LTAP"
        ok, _ = env.run("TO")
        assert not ok
        assert any(p["kind"] == "error" for p in env.popups)
        # first step failed -> nothing salvageable -> says NOT modified
        msg = [p for p in env.popups if p["kind"] == "error"][0]["message"]
        assert "NOT modified" in msg

    def test_empty_primary_column_pops_the_notebook_style_warning(self, env):
        env.grid = {(1, 1): "H1"}                  # headers only, no data
        ok, _ = env.run("TO")
        assert not ok
        errors = [p for p in env.popups if p["kind"] == "error"]
        assert len(errors) == 1
        assert "column K" in errors[0]["message"]


class TestDryRunAndCancelPopups:
    def test_dry_run_completion_popup_says_not_modified(self, env):
        env.grid = make_to_grid()
        ok, _ = env.run("TO", dry_run=True)
        assert ok
        final = env.popups[-1]
        assert final["title"] == "Dry run complete"
        assert "NOT modified" in final["message"]

    def test_stop_before_run_pops_cancelled(self, env):
        import threading
        env.grid = make_to_grid()
        stop = threading.Event()
        stop.set()
        ok, _ = env.run("TO", stop_event=stop)
        assert not ok
        assert any(p["title"] == "Cancelled" for p in env.popups)


class TestPopupContract:
    def test_every_popup_is_logged_to_the_run_file_format(self, env):
        """Popups carry the troubleshooting payload; they must never exist
        only on screen. _popup writes each one to the file log, so the
        payload survives after the box is dismissed."""
        env.grid = make_to_grid()
        env.fail_table = "Z50CFG_ENG_VALD"
        env.run("TO", step_popups=True)
        # contract-level check: every popup event has the fields the GUI
        # and the file log need
        for p in env.popups:
            assert p["kind"] in ("step", "info", "error")
            assert p["title"] and p["message"]
            if p["kind"] == "step":
                assert p["ack"] is not None and p["result"] is not None
            else:
                assert p["ack"] is None
