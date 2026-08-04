"""describe_selection_screen: read the S<n> -> field mapping off the screen.

ZTBV names its select-options generically (S3, S15, S29), so the slot for a
field that is not already in PUSH_BUTTONS cannot be inferred. This walks the
live selection screen and pairs each multi-value button with the label on
its row, so the mapping is read rather than guessed.
"""
import pytest

import sap_ops


class FakeCtl:
    def __init__(self, id_, type_, top, text=""):
        self.Id = id_
        self.Type = type_
        self.Top = top
        self.Text = text
        self._kids = []

    @property
    def Children(self):
        kids = self._kids

        class C:
            Count = len(kids)

            def __call__(self, i):
                return kids[i]
        return C()

    def add(self, *c):
        self._kids.extend(c)
        return self


def _screen():
    """A ZTBV selection screen: three filters on three rows."""
    root = FakeCtl("wnd[0]/usr", "GuiUserArea", 0)
    rows = [
        (3, "Transfer Order Number", "S3"),
        (5, "Number of Reservation", "S15"),
        (7, "Notification", "S29"),
    ]
    for top, label, param in rows:
        root.add(
            FakeCtl(f"wnd[0]/usr/lbl{param}", "GuiLabel", top, label),
            FakeCtl(f"wnd[0]/usr/ctxt{param}-LOW", "GuiCTextField", top),
            FakeCtl(f"wnd[0]/usr/btn%_{param}_%_APP_%-VALU_PUSH",
                    "GuiButton", top),
        )
    return root


class FakeSession:
    def __init__(self, root):
        self.root = root

    def find(self, path):
        if path == "wnd[0]/usr":
            return self.root
        raise Exception(f"no control {path}")


class TestDescribeSelectionScreen:
    def test_maps_each_filter_to_its_label(self):
        got = sap_ops.describe_selection_screen(FakeSession(_screen()))
        by_param = {f["param"]: f["label"] for f in got}
        assert by_param == {
            "S3": "Transfer Order Number",
            "S15": "Number of Reservation",
            "S29": "Notification",
        }

    def test_push_ids_match_the_PUSH_BUTTONS_format(self):
        got = sap_ops.describe_selection_screen(FakeSession(_screen()))
        ids = {f["param"]: f["push_id"] for f in got}
        assert ids["S15"] == sap_ops.PUSH_BUTTONS[("Z50CFG_ENG_CRNT", "RSNUM")]
        assert ids["S29"] == sap_ops.PUSH_BUTTONS[("Z50CFG_ENG_CRNT", "QMNUM")]
        assert ids["S3"] == sap_ops.PUSH_BUTTONS[("LTAP", "TO_NUMBER")]

    def test_logs_each_filter(self):
        msgs: list = []
        sap_ops.describe_selection_screen(FakeSession(_screen()), log=msgs.append)
        assert any("Transfer Order Number" in m for m in msgs)

    def test_no_selection_screen_raises(self):
        class Dead:
            def find(self, path):
                raise Exception("boom")
        with pytest.raises(sap_ops.SapError, match="No selection screen"):
            sap_ops.describe_selection_screen(Dead())

    def test_screen_without_filters_raises(self):
        bare = FakeCtl("wnd[0]/usr", "GuiUserArea", 0).add(
            FakeCtl("wnd[0]/usr/lblX", "GuiLabel", 1, "nothing here"))
        with pytest.raises(sap_ops.SapError, match="No multi-value filter"):
            sap_ops.describe_selection_screen(FakeSession(bare))
