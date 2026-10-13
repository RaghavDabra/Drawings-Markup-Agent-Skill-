import drawing_markup as dm


def test_label_m2_superscript_and_plain():
    assert dm._label_m2("Consultation | Room | 14 m²") == 14.0
    assert dm._label_m2("BALCONY 7.5 m2") == 7.5


def test_label_m2_takes_largest_figure():
    assert dm._label_m2("APT 3.01 | 82 m2 | Balcony 12 m2") == 82.0


def test_label_m2_none_without_area():
    assert dm._label_m2("SWITCH ROOM") is None
