"""esa_lookup.ipynb must not drift from the .py sources it is built from.

The notebook is what the ESA developers download and run, but it is
GENERATED from pipeline.py + excel_ops.py + sap_ops.py by build_notebook.py.
Nothing forced anyone to re-run the builder, so a fix could land in the .py
files, pass the whole suite, be pushed -- and leave the only artefact anyone
actually runs on the old code. That happened. This test makes it loud.

If this fails, you edited a .py file without regenerating the notebook:

    python build_notebook.py
"""
import json
import os
import sys

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

NOTEBOOK_PATH = os.path.join(ROOT, "esa_lookup.ipynb")


@pytest.fixture(scope="module")
def rebuilt():
    """The notebook build_notebook.py would write right now."""
    import build_notebook
    return build_notebook.NOTEBOOK


@pytest.fixture(scope="module")
def committed():
    with open(NOTEBOOK_PATH, encoding="utf-8") as fh:
        return json.load(fh)


def _cell_sources(nb, kind="code"):
    return ["".join(c["source"]) for c in nb["cells"] if c["cell_type"] == kind]


class TestNotebookIsCurrent:
    def test_notebook_matches_the_sources(self, rebuilt, committed):
        assert committed == rebuilt, (
            "esa_lookup.ipynb is stale -- run: python build_notebook.py")

    def test_cell_count_matches(self, rebuilt, committed):
        assert len(committed["cells"]) == len(rebuilt["cells"])


class TestNotebookCarriesTheCode:
    """Spot-checks with names, so a failure says WHAT is missing rather than
    'two large dicts differ'."""

    @pytest.mark.parametrize("needle", [
        "def resolve_push_button(",
        "def push_button_id(",
        "def list_filter_slots(",
        "def _salvage_completed_steps(",
        "def _diagnose(",
        "def screen_snapshot(",
        "def _build_lookup(",
        "PUSH_BUTTONS = {",
    ])
    def test_source_is_inlined(self, committed, needle):
        assert any(needle in s for s in _cell_sources(committed)), needle

    def test_removed_machinery_stays_removed(self, committed):
        """The export fallback, label resolution, and file transport were
        deleted so the operator can read the utilities. If any of these
        names reappear, the bloat is creeping back."""
        for needle in ("import pandas", "describe_selection_screen",
                       "export_alv_to_file", "_read_sap_export",
                       "SAP_COLUMN_ALIASES",
                       "fill_multi_value_filter_from_file"):
            assert not any(needle in s for s in _cell_sources(committed)), needle

    def test_process_cells_carry_the_operator_interface(self, committed):
        """The ESA operator thinks in processes, not in framework parts:
        one runnable cell per process, a one-line POPUPS switch, and real
        message boxes -- the interface that persuaded them off the old
        per-cell notebook. If these disappear, the notebook has regressed
        to machinery-first."""
        code = _cell_sources(committed)
        assert any('run_process("TO"' in s for s in code)
        assert any('run_process("NOTIF"' in s for s in code)
        assert sum("POPUPS = True" in s for s in code) >= 2, \
            "each process cell needs its own one-line popup switch"
        assert any("MessageBoxW" in s for s in code), \
            "popups must be real Windows message boxes, not prints"
        assert any("diagnose=True" in s for s in code), \
            "the troubleshooting cell must expose Diagnose"

    def test_run_all_is_safe(self, committed):
        """Cell -> Run All must execute the combined flow (TO then NOTIF),
        pick the workbook once, halt on failure, and never auto-run
        Diagnose. Verified live in the session that added this; pinned
        here so it cannot regress."""
        code = _cell_sources(committed)
        setup = next(s for s in code if "def run_process(" in s)
        # failure/cancel raises so Jupyter cancels the queued cells
        assert "raise SystemExit" in setup
        assert "did not complete" in setup
        # one Browse for the whole Run All: the picked path is remembered
        assert "_last_browsed" in setup
        # the Diagnose call must be commented out, not live
        diag = code[-1]
        for line in diag.splitlines():
            if "diagnose=True" in line:
                assert line.lstrip().startswith("#"), \
                    "Diagnose must never run during a Run All"

    def test_process_boundaries_are_markdown_headings(self, committed):
        md = _cell_sources(committed, kind="markdown")
        assert any(s.startswith("# TO Number Process") for s in md)
        assert any(s.startswith("# Notification Number Process") for s in md)
        # the shared-code cell is named in the operator's vocabulary --
        # "Utility class for data pulling" -- after "Engine", "helper
        # functions", and "Setup" all drew confused questions
        assert any(s.startswith("## Utility class for data pulling")
                   for s in md)

    def test_module_qualified_refs_are_rewritten(self, committed):
        """Everything is inlined into one namespace, so `sap_ops.foo(...)`
        would raise NameError at run time. Prose inside strings is fine."""
        import re
        offenders = []
        for src in _cell_sources(committed):
            for line in src.splitlines():
                s = line.strip()
                if s.startswith("#") or ".py" in line:
                    continue
                if re.search(r"\b(sap_ops|excel_ops)\.", line):
                    offenders.append(s)
        assert not offenders, offenders
