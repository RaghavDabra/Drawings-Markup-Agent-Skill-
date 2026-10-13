import pytest

import drawing_markup as dm


def test_named_colour():
    assert dm.colour("Red") == dm.COLOURS["red"]


def test_hex_colour():
    assert dm.colour("#ff0000") == (1.0, 0.0, 0.0)


def test_tuple_passes_through():
    assert dm.colour((0.1, 0.2, 0.3)) == (0.1, 0.2, 0.3)


@pytest.mark.parametrize("bad", ["magenta", "#fff", "#gg0000"])
def test_bad_colour_raises(bad):
    with pytest.raises(ValueError):
        dm.colour(bad)
