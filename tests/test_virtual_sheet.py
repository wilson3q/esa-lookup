"""Unit tests for the Gen 4 VirtualSheet column resolver."""
import excel_ops
from pipeline import VirtualSheet


class TestPublished:
    def test_last_row_ignores_trailing_blanks(self):
        vs = VirtualSheet(sheet=None)
        vs.publish(3, ["a", "b", "", ""])
        # values start at row 2 -> "b" sits in row 3, blanks don't count
        assert vs.last_row(3) == 3

    def test_last_row_all_blank_means_no_data(self):
        vs = VirtualSheet(sheet=None)
        vs.publish(3, ["", "", ""])
        assert vs.last_row(3) == 1   # caller treats <2 as "step skipped"

    def test_get_col_pads_and_trims(self):
        vs = VirtualSheet(sheet=None)
        vs.publish(3, ["a", "b"])
        assert vs.get_col(3, 5) == ["a", "b", "", ""]   # rows 2..5, padded
        assert vs.get_col(3, 2) == ["a"]                # trimmed to range

    def test_is_virtual(self):
        vs = VirtualSheet(sheet=None)
        assert not vs.is_virtual(3)
        vs.publish(3, ["x"])
        assert vs.is_virtual(3)


class TestExcelFallback:
    def test_unpublished_column_reads_from_excel(self, monkeypatch):
        reads = []
        monkeypatch.setattr(excel_ops, "last_row_in_column",
                            lambda sheet, col: 4)
        monkeypatch.setattr(
            excel_ops, "read_range_2d",
            lambda sheet, rng: reads.append(rng) or [["x"], ["y"], ["z"]])
        vs = VirtualSheet(sheet=object())
        assert vs.last_row(2) == 4
        assert vs.get_col(2, 4) == ["x", "y", "z"]
        assert reads == ["B2:B4"]

    def test_published_column_never_touches_excel(self, monkeypatch):
        def boom(*a, **k):
            raise AssertionError("Excel read attempted for published column")
        monkeypatch.setattr(excel_ops, "last_row_in_column", boom)
        monkeypatch.setattr(excel_ops, "read_range_2d", boom)
        vs = VirtualSheet(sheet=object())
        vs.publish(3, ["OBJ1", "OBJ2"])
        assert vs.last_row(3) == 3
        assert vs.get_col(3, 3) == ["OBJ1", "OBJ2"]
