import drawing_markup as dm


def test_label_m2_superscript_and_plain():
    assert dm._label_m2("Consultation | Room | 14 m²") == 14.0
    assert dm._label_m2("BALCONY 7.5 m2") == 7.5


def test_label_m2_takes_largest_figure():
    assert dm._label_m2("APT 3.01 | 82 m2 | Balcony 12 m2") == 82.0


def test_label_m2_none_without_area():
    assert dm._label_m2("SWITCH ROOM") is None


def test_unit_keys_treat_excel_numbers_as_same_unit():
    keys = dm._UnitKeys()
    keys.setdefault("3.10", []).append("row")
    assert "3.1" in keys and "03.10" in keys
    assert keys["3.1"] == ["row"]


def test_unit_keys_leave_text_ids_alone():
    keys = dm._UnitKeys()
    keys.setdefault("A101", [])
    assert "A101" in keys and "a101" not in keys


def test_pages_ranges():
    assert dm._pages("3,10-12") == [3, 10, 11, 12]


def _page_with(lines):
    import fitz
    doc = fitz.open()
    page = doc.new_page(width=600, height=400)
    for (x, y), text in lines:
        page.insert_text((x, y), text, fontsize=9)
    return page


def test_unit_tag_with_type_code_and_no_area():
    # 'UNIT 2U.01' over '507': the word after UNIT is the apartment type, the number sits underneath
    page = _page_with([((100, 100), "UNIT 2U.01"), ((115, 111), "507"), ((300, 300), "507")])
    keys = dm._UnitKeys()
    keys.setdefault("507", [])
    rects = dm._unit_label_rects(page, keys)
    assert list(rects) == [dm._UnitKeys.canon("507")]
    assert rects[dm._UnitKeys.canon("507")].y1 < 150      # the tagged one, not the bare 507 lower down


def test_stacked_typical_floor_tag_marks_each_number_alone():
    page = _page_with([((100, 100), "UNIT 2A.01"), ((100, 111), "909 1009 1109")])
    keys = dm._UnitKeys()
    for u in ("909", "1109"):
        keys.setdefault(u, [])
    rects = dm._unit_label_rects(page, keys)
    assert set(rects) == {dm._UnitKeys.canon("909"), dm._UnitKeys.canon("1109")}
    assert not rects[dm._UnitKeys.canon("1109")].intersects(rects[dm._UnitKeys.canon("909")])
