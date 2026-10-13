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
