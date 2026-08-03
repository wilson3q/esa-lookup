"""Unit tests for the pure key-normalization / helper functions."""
import pandas as pd
import pytest

import pipeline
from pipeline import (
    _build_lookup,
    _clean_cell,
    _col_letter,
    _is_skip_key_cell,
    clean_numeric_for_sap,
    normalize_key,
)


class TestNormalizeKey:
    @pytest.mark.parametrize("raw, expected", [
        ("00123", "123"),              # leading zeros dropped
        ("123.0", "123"),              # trailing .0
        ("123.00", "123"),             # trailing .00
        ("123.", "123"),               # trailing dot
        (" 12,345 ", "12345"),         # whitespace + thousands separator
        ("'0012", "12"),               # apostrophe + zeros
        ("12 34", "1234"),        # NBSP stripped
        ("1.5E3", "1500"),             # full-string scientific notation
        ("0", "0"),                    # lone zero survives
        (123456789.0, "123456789"),    # float from Excel COM
        (None, ""),
        ("", ""),
    ])
    def test_values(self, raw, expected):
        assert normalize_key(raw) == expected

    def test_nan_is_blank(self):
        assert normalize_key(float("nan")) == ""  # Fix B

    def test_excel_error_code_is_blank(self):
        assert normalize_key(-2146826281) == ""   # Fix G: #N/A et al.

    def test_embedded_e_not_expanded(self):
        # Fix E: only FULL-string sci-notation expands
        assert normalize_key("OBJ1E5") == "OBJ1E5"


class TestCleanNumericForSap:
    def test_preserves_leading_zeros(self):
        assert clean_numeric_for_sap("00123") == "00123"
        assert clean_numeric_for_sap("00123.0") == "00123"

    def test_same_stripping_as_normalize(self):
        assert clean_numeric_for_sap(" 1,234.00 ") == "1234"


class TestSkipMarkers:
    @pytest.mark.parametrize("raw, skipped", [
        ("NOT FOUND", True),
        ("notfound", True),
        ("", True),
        (None, True),
        ("T001", False),
        (0, False),
    ])
    def test_flags(self, raw, skipped):
        assert _is_skip_key_cell(raw) is skipped


class TestColLetter:
    @pytest.mark.parametrize("idx, letter", [
        (1, "A"), (2, "B"), (26, "Z"), (27, "AA"), (52, "AZ"), (53, "BA"),
    ])
    def test_values(self, idx, letter):
        assert _col_letter(idx) == letter


class TestCleanCell:
    def test_none_nan_error(self):
        assert _clean_cell(None) == ""
        assert _clean_cell(float("nan")) == ""
        assert _clean_cell(-2146826281) == ""

    def test_bool_not_treated_as_error_code(self):
        assert _clean_cell(True) == "True"


class TestBuildLookup:
    def test_alias_resolution(self):
        # ALV export used display titles, not technical names
        df = pd.DataFrame({
            "Reservation": [100], "Item": [1], "Object Number": ["OBJ1"]})
        lookup, dups, blanks = _build_lookup(
            df, ["RSNUM", "RSPOS"], ["OBJNR"])
        assert lookup == {"100|1": {"OBJNR": "OBJ1"}}
        assert dups == 0 and blanks == 0

    def test_duplicate_keeps_first(self):
        df = pd.DataFrame({"TANUM": ["T1", "T1"], "ABLAD": ["A", "B"]})
        lookup, dups, _ = _build_lookup(df, ["TANUM"], ["ABLAD"])
        assert lookup["T1"]["ABLAD"] == "A"   # Fix C: first wins
        assert dups == 1

    def test_partial_blank_key_skipped(self):
        # Fix D: "12345|" must not become a match-all key
        df = pd.DataFrame({
            "RSNUM": [100, 200], "RSPOS": [1, None], "OBJNR": ["O1", "O2"]})
        lookup, _, blanks = _build_lookup(df, ["RSNUM", "RSPOS"], ["OBJNR"])
        assert "200|" not in lookup
        assert blanks == 1

    def test_missing_column_raises_with_hint(self):
        df = pd.DataFrame({"SOMETHING": [1]})
        with pytest.raises(RuntimeError, match="missing these expected columns"):
            _build_lookup(df, ["TANUM"], ["ABLAD"])

    def test_nan_value_becomes_empty(self):
        df = pd.DataFrame({"TANUM": ["T1"], "ABLAD": [float("nan")]})
        lookup, _, _ = _build_lookup(df, ["TANUM"], ["ABLAD"])
        assert lookup["T1"]["ABLAD"] == ""
