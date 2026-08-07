"""Tests for reading whatever format SAP actually wrote the ALV export in.

The regression these cover: SAP names every export ".xlsx", but the `&PC`
fallback (used whenever `&XXL` is unavailable) writes delimited TEXT under
that name. pd.read_excel on it raised "Excel file format cannot be
determined, you must specify an engine manually" and killed the run.
"""
import pandas as pd
import pytest

from pipeline import _decode_sap_text, _read_sap_export, _sap_text_to_frame


TAB_EXPORT = (
    "TANUM\tABLAD\tDESCRIPT\n"
    "0001785822\tBAY-12\tGearbox housing\n"
    "0001785823\tBAY-07\tShaft, drive\n"
)

# `&PC` -> "unconverted": pipe-ruled ALV list with a report title block.
PIPE_EXPORT = (
    "Transfer Order List                                 04.08.2026\n"
    "--------------------------------------------------------------\n"
    "|TANUM      |ABLAD  |DESCRIPT         |\n"
    "|-----------+-------+-----------------|\n"
    "|0001785822 |BAY-12 |Gearbox housing  |\n"
    "|0001785823 |BAY-07 |Shaft, drive     |\n"
    "--------------------------------------------------------------\n"
    "2 record(s) selected\n"
)


class TestSapTextToFrame:
    def test_tab_delimited_spreadsheet_format(self):
        df = _sap_text_to_frame(TAB_EXPORT)
        assert list(df.columns) == ["TANUM", "ABLAD", "DESCRIPT"]
        assert len(df) == 2
        assert df.iloc[0]["DESCRIPT"] == "Gearbox housing"

    def test_pipe_ruled_list_format_drops_rulers_and_title(self):
        df = _sap_text_to_frame(PIPE_EXPORT)
        assert list(df.columns) == ["TANUM", "ABLAD", "DESCRIPT"]
        assert len(df) == 2
        assert df.iloc[1]["ABLAD"] == "BAY-07"

    def test_leading_zeros_survive(self):
        # pandas' numeric inference would turn this into 1785822 and break
        # the join against SAP's own zero-padded keys.
        df = _sap_text_to_frame(TAB_EXPORT)
        assert df.iloc[0]["TANUM"] == "0001785822"

    def test_empty_text(self):
        assert _sap_text_to_frame("").empty


class TestDecodeSapText:
    @pytest.mark.parametrize("enc", ["utf-8", "utf-16", "cp1252"])
    def test_roundtrip_with_bom_and_codepage(self, enc):
        raw = TAB_EXPORT.encode(enc)
        if enc == "utf-8":
            raw = b"\xef\xbb\xbf" + raw
        assert _decode_sap_text(raw).lstrip("﻿") == TAB_EXPORT

    def test_undecodable_bytes_do_not_raise(self):
        assert _decode_sap_text(b"A\xff\xfeB\x81") != ""


_HTML_TABLE = (
    "<html><body><table>"
    "<tr><th>TANUM</th><th>ABLAD</th></tr>"
    "<tr><td>0001785822</td><td>BAY-12</td></tr>"
    "</table></body></html>"
)

# What SAP writes where the site's Office integration is on: the HTML table
# wrapped in a MIME preamble, still named .xlsx.
MHTML_EXPORT = (
    "MIME-Version: 1.0\n"
    "Content-Location: file:///C:/temp/export.mhtml\n"
    'Content-Type: multipart/related; boundary="----=_NextPart_01"\n\n'
    "------=_NextPart_01\n"
    'Content-Type: text/html; charset="windows-1252"\n\n'
    + _HTML_TABLE
)


class TestMhtmlAndHtml:
    """The other `&XXL` failure mode: an MHTML-wrapped table named .xlsx."""

    def test_mhtml_preamble_is_stripped(self, tmp_path):
        p = tmp_path / "export.xlsx"
        p.write_bytes(MHTML_EXPORT.encode("cp1252"))
        df = _read_sap_export(str(p))
        assert list(df.columns) == ["TANUM", "ABLAD"]
        assert len(df) == 1

    def test_bare_html_table(self, tmp_path):
        p = tmp_path / "export.xlsx"
        p.write_bytes(_HTML_TABLE.encode("utf-8"))
        df = _read_sap_export(str(p))
        assert list(df.columns) == ["TANUM", "ABLAD"]

    def test_mhtml_without_html_body_raises(self, tmp_path):
        p = tmp_path / "export.xlsx"
        p.write_bytes(b"MIME-Version: 1.0\nContent-Location: f\n\nno body\n")
        with pytest.raises(RuntimeError, match="no <html> body"):
            _read_sap_export(str(p))


class TestReadSapExport:
    def test_text_file_named_xlsx(self, tmp_path):
        """The exact failure the developer hit."""
        p = tmp_path / "LTAP_TO_NUMBER_1785822548_0.xlsx"
        p.write_bytes(TAB_EXPORT.encode("cp1252"))

        with pytest.raises(ValueError, match="format cannot be determined"):
            pd.read_excel(p)          # what the old code did

        df = _read_sap_export(str(p))  # what it does now
        assert len(df) == 2
        assert list(df.columns) == ["TANUM", "ABLAD", "DESCRIPT"]

    def test_real_xlsx_still_reads_via_openpyxl(self, tmp_path):
        p = tmp_path / "real.xlsx"
        pd.DataFrame({"TANUM": [1785822], "ABLAD": ["BAY-12"]}).to_excel(
            p, index=False)
        df = _read_sap_export(str(p))
        assert list(df.columns) == ["TANUM", "ABLAD"]
        # dtype=str on purpose: numeric inference would eat the leading
        # zeros of encoded values like the ESA unloading point. Key
        # matching normalizes both sides, so strings match fine.
        assert df.iloc[0]["TANUM"] == "1785822"

    def test_logs_when_falling_back_to_text(self, tmp_path):
        p = tmp_path / "export.xlsx"
        p.write_text(TAB_EXPORT, encoding="utf-8")
        msgs = []
        _read_sap_export(str(p), log=msgs.append)
        assert any("&PC" in m for m in msgs)

    def test_unparseable_binary_raises_actionable_error(self, tmp_path):
        p = tmp_path / "junk.xlsx"
        p.write_bytes(b"\x00\x01\x02\x03\x04\x05\x06\x07")
        with pytest.raises(RuntimeError, match="Could not parse"):
            _read_sap_export(str(p))
