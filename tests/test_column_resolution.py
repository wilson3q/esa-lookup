"""Tests for resolving SAP field names against real ALV column titles.

The regression these cover: the ESA ZTBV layout prints English SHORT LABELS
rather than technical names, so ABLAD arrives as 'Unl. Point'. _build_lookup
raised "SAP export is missing these expected columns: ['ABLAD']" even though
the field was right there in the export.
"""
import pandas as pd
import pytest

from pipeline import _build_lookup, _norm_col, _resolve_column


# The real LTAP titles from the customer's export (trimmed), including the
# repeated ones that make a title ambiguous.
ESA_LTAP_TITLES = [
    "WhN", "TO Number", "Item", "Item", "Item", "Material", "Plnt", "Batch",
    "S", "S", "Sp.Stck No", "Haz.", "Recipient", "Unl. Point", "GR Date",
    "GR Number", "Cert. No.", "Printer", "Proc.", "Typ", "Sec", "Source Bin",
    "DB", "B.pos", "Quant", "Typ", "Sec", "Dest. Bin", "DB", "B.pos",
]


def _esa_frame(rows=1):
    data = {}
    for i, t in enumerate(ESA_LTAP_TITLES):
        data.setdefault(t, []).append(i)
    # build with duplicate titles preserved
    return pd.DataFrame(
        [[f"v{i}" for i in range(len(ESA_LTAP_TITLES))] for _ in range(rows)],
        columns=ESA_LTAP_TITLES,
    )


class TestNormCol:
    @pytest.mark.parametrize("a, b", [
        ("Unl. Point", "Unl.Point"),
        ("Unl. Point", "UNL POINT"),
        ("TO Number", "to_number"),
        ("Dest. Bin", "Dest.Bin"),
    ])
    def test_punctuation_and_case_folded(self, a, b):
        assert _norm_col(a) == _norm_col(b)

    def test_distinct_names_stay_distinct(self):
        assert _norm_col("Source Bin") != _norm_col("Dest. Bin")


class TestResolveColumn:
    def test_ablad_resolves_to_short_label(self):
        """The exact ESA failure: ABLAD present only as 'Unl. Point'."""
        df = _esa_frame()
        assert _resolve_column(df, "ABLAD") == "Unl. Point"

    def test_tanum_resolves(self):
        assert _resolve_column(_esa_frame(), "TANUM") == "TO Number"

    def test_technical_name_still_preferred_when_present(self):
        df = pd.DataFrame(columns=["ABLAD", "Unl. Point"])
        assert _resolve_column(df, "ABLAD") == "ABLAD"

    def test_punctuation_variant_needs_no_new_alias(self):
        df = pd.DataFrame(columns=["WhN", "Unl.Point"])
        assert _resolve_column(df, "ABLAD") == "Unl.Point"

    def test_genuinely_absent_returns_none(self):
        assert _resolve_column(pd.DataFrame(columns=["WhN"]), "ABLAD") is None


class TestBuildLookupErrors:
    def test_missing_column_error_suggests_close_titles(self):
        """A label no alias lists yet, close enough to name in the error."""
        df = pd.DataFrame(columns=["TO Number", "Unloading Pt."])
        with pytest.raises(RuntimeError) as e:
            _build_lookup(df, ["TANUM"], ["ABLAD"])
        msg = str(e.value)
        assert "missing these expected columns" in msg
        assert "Did you mean" in msg
        assert "Unloading Pt." in msg.split("Did you mean")[1]

    def test_no_hint_when_nothing_is_close(self):
        df = pd.DataFrame(columns=["TO Number"])
        with pytest.raises(RuntimeError) as e:
            _build_lookup(df, ["TANUM"], ["QMNUM"])
        assert "Did you mean" not in str(e.value)

    def test_duplicate_title_is_rejected_not_silently_wrong(self):
        """'Item' appears 3x; resolving RSPOS onto it must fail loudly."""
        df = _esa_frame()
        with pytest.raises(RuntimeError, match="more than once"):
            _build_lookup(df, ["TANUM"], ["RSPOS"])

    def test_unambiguous_columns_build_normally(self):
        df = pd.DataFrame({"TO Number": ["0001785822"], "Unl. Point": ["BAY-12"]})
        lookup, dups, blanks = _build_lookup(df, ["TANUM"], ["ABLAD"])
        assert lookup == {"1785822": {"ABLAD": "BAY-12"}}
        assert (dups, blanks) == (0, 0)
