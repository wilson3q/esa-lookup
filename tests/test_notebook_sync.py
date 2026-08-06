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
        "def describe_selection_screen(",
        "def screen_inventory(",
        "def resolve_push_button(",
        "def push_button_id(",
        "def _salvage_completed_steps(",
        "def _diagnose(",
        "CharTop",                      # the label-pairing fix
        "ESA_LOOKUP_PUSH_",             # the no-rebuild filter override
    ])
    def test_source_is_inlined(self, committed, needle):
        assert any(needle in s for s in _cell_sources(committed)), needle

    def test_old_pixel_pairing_is_gone(self, committed):
        assert not any("labels_by_top" in s for s in _cell_sources(committed))

    def test_diagnose_is_reachable_from_the_config_cell(self, committed):
        code = _cell_sources(committed)
        assert any("DIAGNOSE = False" in s for s in code), \
            "config cell must expose a DIAGNOSE switch"
        assert any("diagnose=DIAGNOSE" in s for s in code), \
            "execute cell must pass diagnose through to RunConfig"
        assert any("FILTER_OVERRIDES" in s for s in code), \
            "config cell must let a filter slot be pinned without a rebuild"

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
