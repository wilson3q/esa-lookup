"""Unit tests for the pure key-normalization / helper functions."""
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
    def test_records_by_technical_name(self):
        rows = [{"RSNUM": 100, "RSPOS": 1, "OBJNR": "OBJ1"}]
        lookup, dups, blanks = _build_lookup(rows, ["RSNUM", "RSPOS"], ["OBJNR"])
        assert lookup == {"100|1": {"OBJNR": "OBJ1"}}
        assert dups == 0 and blanks == 0

    def test_duplicate_keeps_first(self):
        rows = [{"TANUM": "T1", "ABLAD": "A"}, {"TANUM": "T1", "ABLAD": "B"}]
        lookup, dups, _ = _build_lookup(rows, ["TANUM"], ["ABLAD"])
        assert lookup["T1"]["ABLAD"] == "A"   # Fix C: first wins
        assert dups == 1

    def test_partial_blank_key_skipped(self):
        rows = [{"RSNUM": 100, "RSPOS": "", "OBJNR": "X"},
                {"RSNUM": 200, "RSPOS": 2, "OBJNR": "Y"}]
        lookup, _, blanks = _build_lookup(rows, ["RSNUM", "RSPOS"], ["OBJNR"])
        assert list(lookup) == ["200|2"]      # Fix D: "100|" never matches
        assert blanks == 1

    def test_none_value_becomes_empty(self):
        rows = [{"TANUM": "T1", "ABLAD": None}]
        lookup, _, _ = _build_lookup(rows, ["TANUM"], ["ABLAD"])
        assert lookup["T1"]["ABLAD"] == ""

class TestUnloadingPointSplit:
    """The template formulas, pinned as executable truth (operator-supplied
    2026-08-07):  N: =IFERROR(VALUE(MID(M,2,9)),"")
                  O: =IFERROR(VALUE(MID(M,11,4)),"")
    """

    @staticmethod
    def _formula_n(m):     # Excel MID is 1-based; VALUE strips zeros
        s = str(m)[1:10]
        return str(int(s)) if s.isdigit() else ""

    @staticmethod
    def _formula_o(m):
        s = str(m)[10:14]
        return str(int(s)) if s.isdigit() else ""

    def test_live_workbook_samples(self):
        from pipeline import (_item_from_unloading_point,
                              _reservation_from_unloading_point)
        for ablad, n, o in [("05175455500001", "517545550", "1"),
                            ("05175456020012", "517545602", "12"),
                            ("05181461210005", "518146121", "5"),
                            ("05181461370012", "518146137", "12")]:
            assert _reservation_from_unloading_point(ablad) == n
            assert _item_from_unloading_point(ablad) == o

    def test_agrees_with_the_template_formula_on_encoded_values(self):
        from pipeline import (_item_from_unloading_point,
                              _reservation_from_unloading_point)
        # every 14-digit value starting with '0' -- the entire observed
        # population -- must give exactly what the template formula gave
        for ablad in ("05175455500001", "00000001000001", "09999999990120",
                      "05181461210005", "00000000010001"):
            assert _reservation_from_unloading_point(ablad) == self._formula_n(ablad)
            assert _item_from_unloading_point(ablad) == self._formula_o(ablad)

    def test_keeps_a_10_digit_reservation_the_formula_would_truncate(self):
        from pipeline import _reservation_from_unloading_point
        # first char nonzero: formula MID(2,9) silently drops it
        assert _reservation_from_unloading_point("15175455500001") == "1517545550"
        assert self._formula_n("15175455500001") == "517545550"  # the bug we avoid

    def test_non_encoded_values_yield_blank(self):
        from pipeline import (_item_from_unloading_point,
                              _reservation_from_unloading_point)
        for bad in ("", None, "NOT FOUND", "BAY-07", "12A4", "123"):
            assert _reservation_from_unloading_point(bad) == ""
            assert _item_from_unloading_point(bad) == ""
