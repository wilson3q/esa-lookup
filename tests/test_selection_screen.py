"""describe_selection_screen / resolve_push_button: read the S<n> -> field
mapping off the screen.

ZTBV names its select-options generically (S3, S15, S29), so the slot for a
field that is not already in PUSH_BUTTONS cannot be inferred. This walks the
live selection screen and pairs each multi-value button with the label on
its row, so the mapping is read rather than guessed.

FIXTURE NOTE -- read before editing FakeCtl. The first version of this file
gave the label and the push button on one row the SAME `Top`, and defined no
CharTop/CharLeft at all. Real SAP GUI does neither: `Top` is a PIXEL
coordinate and a label's baseline does not line up with the button drawn
beside it, while CharTop/CharLeft are the exact dynpro row/column. That fake
made a broken pairing look correct -- the suite stayed green while every
filter on the customer's screen came back '(no label found)'. Keep the
label's pixel Top offset from the button's, and keep the character metric
populated, so this file can fail the way production did.
"""
import pytest

import sap_ops


class FakeCtl:
    """A control on wnd[0]/usr.

    Carries BOTH coordinate systems, as the real GuiVComponent does:
    CharTop/CharLeft (dynpro row/column) and Top/Left (pixels).
    """

    def __init__(self, id_, type_, row, col, text="", name="", tooltip="",
                 pixel_offset=0, char_metric=True):
        self.Id = id_
        self.Type = type_
        self.Text = text
        self.Name = name
        self.Tooltip = tooltip
        if char_metric:
            self.CharTop = row
            self.CharLeft = col
        # Pixels: a row is ~16px tall; pixel_offset is how far this control's
        # top sits from the row's baseline (labels really do differ).
        self.Top = row * 16 + pixel_offset
        self.Left = col * 8
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


# A selection screen row reads: [label] [LOW] [HIGH] [=> button]
LABEL_COL, LOW_COL, HIGH_COL, BTN_COL = 1, 25, 45, 65

ROWS = [
    (3, "Transfer Order Number", "S3", ""),
    (5, "Number of Reservation", "S15", ""),
    (7, "Notification", "S29", ""),
]


def _screen(rows=ROWS, with_labels=True, char_metric=True, tooltips=False,
            label_pixel_offset=3, label_type="GuiLabel", separator=None):
    """A ZTBV selection screen. `rows` is [(row, label, param, retained)].

    `label_type` -- the ESA system renders field names as GuiTextField, not
    GuiLabel, so the reader has to cope with both.
    `separator`  -- the range word SAP prints between LOW and HIGH ("to").
    """
    root = FakeCtl("wnd[0]/usr", "GuiUserArea", 0, 0, char_metric=char_metric)
    for row, label, param, retained in rows:
        if with_labels:
            root.add(FakeCtl(
                f"wnd[0]/usr/lbl{param}", label_type, row, LABEL_COL,
                text=label, pixel_offset=label_pixel_offset,
                char_metric=char_metric))
        if separator:
            # sits BETWEEN the two inputs -- nearest text to the button's
            # left, and the reason "nearest label" reported 'to' for
            # every filter on the real screen.
            root.add(FakeCtl(
                f"wnd[0]/usr/sep{param}", label_type, row,
                (LOW_COL + HIGH_COL) // 2, text=separator,
                char_metric=char_metric))
        root.add(
            FakeCtl(f"wnd[0]/usr/ctxt{param}-LOW", "GuiCTextField", row,
                    LOW_COL, text=retained, name=f"{param}-LOW",
                    tooltip=(f"{label} (F1)" if tooltips else ""),
                    char_metric=char_metric),
            FakeCtl(f"wnd[0]/usr/ctxt{param}-HIGH", "GuiCTextField", row,
                    HIGH_COL, name=f"{param}-HIGH", char_metric=char_metric),
            FakeCtl(f"wnd[0]/usr/btn%_{param}_%_APP_%-VALU_PUSH", "GuiButton",
                    row, BTN_COL, char_metric=char_metric),
        )
    return root


class FakeSession:
    def __init__(self, root):
        self.root = root

    def find(self, path):
        if path == "wnd[0]/usr":
            return self.root
        raise Exception(f"no control {path}")


@pytest.fixture(autouse=True)
def _restore_push_buttons():
    """resolve_push_button caches into PUSH_BUTTONS; keep tests independent."""
    saved = dict(sap_ops.PUSH_BUTTONS)
    yield
    sap_ops.PUSH_BUTTONS.clear()
    sap_ops.PUSH_BUTTONS.update(saved)


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
        bare = FakeCtl("wnd[0]/usr", "GuiUserArea", 0, 0).add(
            FakeCtl("wnd[0]/usr/lblX", "GuiLabel", 1, 1, text="nothing here"))
        with pytest.raises(sap_ops.SapError, match="No multi-value filter"):
            sap_ops.describe_selection_screen(FakeSession(bare))

    def test_label_finds_its_row_despite_a_pixel_offset(self):
        """The production bug: pairing on pixel Top left every filter
        unlabelled, because a label's Top never equals its button's."""
        for offset in (-6, -3, 0, 3, 6):
            got = sap_ops.describe_selection_screen(
                FakeSession(_screen(label_pixel_offset=offset)))
            assert [f["label"] for f in got] == [
                "Transfer Order Number", "Number of Reservation",
                "Notification"], f"offset {offset}"

    def test_pairs_in_pixel_metric_when_char_metric_is_absent(self):
        got = sap_ops.describe_selection_screen(
            FakeSession(_screen(char_metric=False)))
        assert {f["param"]: f["label"] for f in got} == {
            "S3": "Transfer Order Number",
            "S15": "Number of Reservation",
            "S29": "Notification",
        }

    def test_retained_filter_values_are_never_used_as_labels(self):
        """Select-option inputs hold values left over from a previous run.
        Treating one as a label would paste keys into the wrong filter and
        return confidently wrong data."""
        rows = [(3, "Transfer Order Number", "S3",
                 "0000000012345678901234567890")]
        got = sap_ops.describe_selection_screen(FakeSession(_screen(rows)))
        assert got[0]["label"] == "Transfer Order Number"

    def test_tooltips_are_captured_from_the_low_field(self):
        got = sap_ops.describe_selection_screen(
            FakeSession(_screen(tooltips=True)))
        assert [f["tooltip"] for f in got] == [
            "Transfer Order Number (F1)", "Number of Reservation (F1)",
            "Notification (F1)"]

    def test_low_field_id_comes_from_the_screen(self):
        got = sap_ops.describe_selection_screen(FakeSession(_screen()))
        assert got[0]["low_field"] == "wnd[0]/usr/ctxtS3-LOW"

    def test_range_separator_is_not_mistaken_for_the_field_name(self):
        """The ESA screen: field names are GuiTextField (no GuiLabel at all)
        and SAP prints 'to' between LOW and HIGH. 'to' is the nearest text to
        the left of the => button, so a nearest-label rule reported every one
        of the 39 filters as 'to'. The field name is the LEFTMOST text."""
        got = sap_ops.describe_selection_screen(
            FakeSession(_screen(label_type="GuiTextField", separator="to")))
        assert {f["param"]: f["label"] for f in got} == {
            "S3": "Transfer Order Number",
            "S15": "Number of Reservation",
            "S29": "Notification",
        }

    def test_row_texts_are_reported_for_diagnosis(self):
        got = sap_ops.describe_selection_screen(
            FakeSession(_screen(label_type="GuiTextField", separator="to")))
        assert got[0]["row_texts"] == ["Transfer Order Number", "to"]

    @pytest.mark.parametrize("sep", ["to", "bis", "à", "hasta", "-"])
    def test_separators_in_other_logon_languages(self, sep):
        got = sap_ops.describe_selection_screen(
            FakeSession(_screen(label_type="GuiTextField", separator=sep)))
        assert got[0]["label"] == "Transfer Order Number"

    def test_resolution_works_on_a_separator_screen(self):
        """End to end: the exact shape that failed in production."""
        sap_ops.PUSH_BUTTONS.pop(("Z50CFG_ENG_CRNT", "TO_NUMBER"), None)
        rows = [(3, "TO Number", "S7", ""), (5, "Notification", "S29", "")]
        got = sap_ops.resolve_push_button(
            FakeSession(_screen(rows, label_type="GuiTextField",
                                separator="to")),
            "Z50CFG_ENG_CRNT", "TO_NUMBER")
        assert got == sap_ops.push_button_id("S7")

    def test_a_row_with_only_a_separator_reports_no_label(self):
        """Better to say 'no label' than to name a filter 'to'."""
        root = FakeCtl("wnd[0]/usr", "GuiUserArea", 0, 0)
        root.add(
            FakeCtl("wnd[0]/usr/sep", "GuiTextField", 3, 30, text="to"),
            FakeCtl("wnd[0]/usr/btn%_S9_%_APP_%-VALU_PUSH", "GuiButton",
                    3, BTN_COL),
        )
        got = sap_ops.describe_selection_screen(FakeSession(root))
        assert got[0]["label"] == "(no label found)"

    def test_screen_with_no_labels_reports_them_as_missing(self):
        got = sap_ops.describe_selection_screen(
            FakeSession(_screen(with_labels=False)))
        assert {f["label"] for f in got} == {"(no label found)"}


class TestResolvePushButton:
    def test_recorded_pairs_never_read_the_screen(self, monkeypatch):
        """Step 1 (LTAP -> TO_NUMBER) is a recorded pair. It must resolve
        from PUSH_BUTTONS without touching the screen-reading code at all --
        that path is what already works in production."""
        def boom(*a, **k):
            raise AssertionError("describe_selection_screen must not be called")
        monkeypatch.setattr(sap_ops, "describe_selection_screen", boom)

        for (table, field), expected in list(sap_ops.PUSH_BUTTONS.items()):
            assert sap_ops.resolve_push_button(None, table, field) == expected

    def test_resolves_an_unrecorded_pair_by_label(self):
        sap_ops.PUSH_BUTTONS.pop(("Z50CFG_ENG_CRNT", "TO_NUMBER"), None)
        rows = [(3, "TO Number", "S7", ""), (5, "Notification", "S29", "")]
        got = sap_ops.resolve_push_button(
            FakeSession(_screen(rows)), "Z50CFG_ENG_CRNT", "TO_NUMBER")
        assert got == sap_ops.push_button_id("S7")

    def test_falls_back_to_tooltips_when_no_label_matched(self):
        sap_ops.PUSH_BUTTONS.pop(("Z50CFG_ENG_CRNT", "TO_NUMBER"), None)
        rows = [(3, "TO Number", "S7", ""), (5, "Notification", "S29", "")]
        got = sap_ops.resolve_push_button(
            FakeSession(_screen(rows, with_labels=False, tooltips=True)),
            "Z50CFG_ENG_CRNT", "TO_NUMBER")
        assert got == sap_ops.push_button_id("S7")

    def test_caches_a_resolved_pair(self):
        sap_ops.PUSH_BUTTONS.pop(("Z50CFG_ENG_CRNT", "TO_NUMBER"), None)
        rows = [(3, "TO Number", "S7", "")]
        sap_ops.resolve_push_button(
            FakeSession(_screen(rows)), "Z50CFG_ENG_CRNT", "TO_NUMBER")
        assert sap_ops.PUSH_BUTTONS[("Z50CFG_ENG_CRNT", "TO_NUMBER")] == \
            sap_ops.push_button_id("S7")

    def test_an_item_neighbour_loses_the_tie_to_the_number(self):
        """'TO Number' vs 'TO Number Item' -- the item field must never win.
        Same failure mode as RSPOS once binding to the TO item (3d8644f)."""
        sap_ops.PUSH_BUTTONS.pop(("Z50CFG_ENG_CRNT", "TO_NUMBER"), None)
        rows = [(3, "TO Number", "S7", ""), (5, "TO Number Item", "S8", "")]
        got = sap_ops.resolve_push_button(
            FakeSession(_screen(rows)), "Z50CFG_ENG_CRNT", "TO_NUMBER")
        assert got == sap_ops.push_button_id("S7")

    def test_ambiguous_labels_refuse_rather_than_guess(self):
        """Two equally plausible labels that no exclusion separates. Refusing
        is right: pasting keys into the wrong filter returns confidently
        wrong data, which is worse than stopping."""
        sap_ops.PUSH_BUTTONS.pop(("Z50CFG_ENG_CRNT", "TO_NUMBER"), None)
        rows = [(3, "TO Number", "S7", ""), (5, "Transfer Order", "S8", "")]
        with pytest.raises(sap_ops.SapError, match="2 candidate"):
            sap_ops.resolve_push_button(
                FakeSession(_screen(rows)), "Z50CFG_ENG_CRNT", "TO_NUMBER")

    def test_no_match_names_every_filter_on_the_screen(self):
        sap_ops.PUSH_BUTTONS.pop(("Z50CFG_ENG_CRNT", "TO_NUMBER"), None)
        rows = [(3, "Plant", "S2", ""), (5, "Notification", "S29", "")]
        with pytest.raises(sap_ops.SapError) as ei:
            sap_ops.resolve_push_button(
                FakeSession(_screen(rows)), "Z50CFG_ENG_CRNT", "TO_NUMBER")
        msg = str(ei.value)
        assert "S2" in msg and "S29" in msg
        # and it must say how to unblock without a rebuild
        assert "ESA_LOOKUP_PUSH_Z50CFG_ENG_CRNT_TO_NUMBER" in msg

    def test_env_override_wins_and_needs_no_screen(self, monkeypatch):
        sap_ops.PUSH_BUTTONS.pop(("Z50CFG_ENG_CRNT", "TO_NUMBER"), None)
        monkeypatch.setenv(
            "ESA_LOOKUP_PUSH_Z50CFG_ENG_CRNT_TO_NUMBER", "S11")
        assert sap_ops.resolve_push_button(
            None, "Z50CFG_ENG_CRNT", "TO_NUMBER") == \
            sap_ops.push_button_id("S11")

    def test_env_override_accepts_a_full_control_id(self, monkeypatch):
        sap_ops.PUSH_BUTTONS.pop(("Z50CFG_ENG_CRNT", "TO_NUMBER"), None)
        full = sap_ops.push_button_id("S12")
        monkeypatch.setenv("ESA_LOOKUP_PUSH_Z50CFG_ENG_CRNT_TO_NUMBER", full)
        assert sap_ops.resolve_push_button(
            None, "Z50CFG_ENG_CRNT", "TO_NUMBER") == full

    def test_env_override_can_correct_a_recorded_pair(self, monkeypatch):
        """The whole point of the override: fix a wrong mapping in the field
        without a rebuild-and-redistribute cycle."""
        monkeypatch.setenv("ESA_LOOKUP_PUSH_LTAP_TO_NUMBER", "S4")
        assert sap_ops.resolve_push_button(None, "LTAP", "TO_NUMBER") == \
            sap_ops.push_button_id("S4")

    def test_esa_screen_resolves_every_field_we_look_up(self):
        """The real ESA Z50CFG_ENG_CRNT 'Table Display' screen, transcribed
        from a screenshot: 39 select-options, S2..S40, one per field in
        screen order.

        Every field this app looks up is paired on that screen with a
        neighbour matching the same synonym -- 'Transfer Order Number' next
        to 'Transfer order item', 'Number of Reservation/Depend' next to
        'Item Number of Reservation/D', 'Notification No' next to
        'Notification Type'. Each pick must land on the NUMBER.
        """
        fields = [
            "Disposition ID", "Object number", "Disposition material",
            "Quantity Dispositioned", "Disposition Date", "Disposition time",
            "Batch Number", "Plant for batch", "Serial Number", "User Name",
            "Part Status", "Assigned Material", "Assigned Quantity",
            "Number of Reservation/Depend", "Item Number of Reservation/D",
            "Internal table flag", "Order Number", "Transfer Order Number",
            "Transfer order item", "Not More Closely Defined Are",
            "Valid flag", "Checkbox", "Transaction Code",
            "Warehouse Number / Warehouse", "Notification Type",
            "Order Number", "Purchasing Document Number", "Notification No",
            "Equipment Number", "Valuation Type", "Special Stock Indicator",
            "Serial Number", "Batch Number", "Character field, length 70",
            "Character field, length 70", "Character field, length 70",
            "Equipment rec create or chan", "PBOM Use Code",
            "Gate III Critical Part Flag",
        ]
        # slot = position + 1, i.e. the first field is S2
        rows = [(3 + i, name, f"S{i + 2}", "") for i, name in enumerate(fields)]
        screen = _screen(rows, label_type="GuiTextField", separator="to")

        assert len(fields) == 39, "the run log reported 39 filters, S2..S40"

        for field, expected_label, expected_slot in [
            ("TO_NUMBER", "Transfer Order Number", "S19"),
            ("RSNUM", "Number of Reservation/Depend", "S15"),
            ("QMNUM", "Notification No", "S29"),
        ]:
            sap_ops.PUSH_BUTTONS.pop(("Z50CFG_ENG_CRNT", field), None)
            got = sap_ops.resolve_push_button(
                FakeSession(screen), "Z50CFG_ENG_CRNT", field)
            assert got == sap_ops.push_button_id(expected_slot), (
                f"{field} resolved to {got}, expected {expected_slot} "
                f"({expected_label})")

    def test_recorded_slots_agree_with_the_screen(self):
        """The recorded RSNUM/QMNUM slots came from the original working
        notebook; the TO_NUMBER slot was derived from the screenshot by
        position. If the derivation is right, reading the screen by LABEL
        must independently produce the same three answers -- which is what
        the test above checks. This pins the recorded values themselves so a
        typo in PUSH_BUTTONS cannot slip through."""
        assert sap_ops.PUSH_BUTTONS[("Z50CFG_ENG_CRNT", "RSNUM")] == \
            sap_ops.push_button_id("S15")
        assert sap_ops.PUSH_BUTTONS[("Z50CFG_ENG_CRNT", "QMNUM")] == \
            sap_ops.push_button_id("S29")
        assert sap_ops.PUSH_BUTTONS[("Z50CFG_ENG_CRNT", "TO_NUMBER")] == \
            sap_ops.push_button_id("S19")

    def test_unknown_field_with_no_synonyms_raises(self):
        with pytest.raises(sap_ops.SapError, match="no label synonyms"):
            sap_ops.resolve_push_button(
                FakeSession(_screen()), "LTAP", "NOT_A_FIELD")

    def test_no_step_depends_on_reading_the_screen(self, monkeypatch):
        """Blast radius of the screen reader, stated as a fact about the
        workflows rather than a claim in a commit message.

        Every step whose (table, field) is recorded in PUSH_BUTTONS resolves
        without reading the screen, so a change to the reader cannot affect
        it. Since Z50CFG_ENG_CRNT/TO_NUMBER was pinned to S19 that is now
        EVERY step in both workflows -- label reading is a fallback for
        unrecorded pairs and the Diagnose run, not something a normal run
        depends on. If this list grows, a reader change has become able to
        break a step that works today, and that is worth knowing.
        """
        import pipeline

        def boom(*a, **k):
            raise AssertionError("screen read")
        monkeypatch.setattr(sap_ops, "describe_selection_screen", boom)

        needs_screen = []
        for name, steps in pipeline.WORKFLOWS.items():
            for i, step in enumerate(steps, 1):
                try:
                    sap_ops.resolve_push_button(
                        None, step.sap_table, step.push_button_field)
                except AssertionError:
                    needs_screen.append(
                        f"{name} step {i}: {step.sap_table}/"
                        f"{step.push_button_field}")
        assert needs_screen == [], needs_screen
