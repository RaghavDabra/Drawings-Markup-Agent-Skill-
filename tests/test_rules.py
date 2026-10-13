import pytest

import drawing_markup as dm


def test_parse_rules_order_and_ops():
    rules = dm._parse_rules(">8:blue, <8:orange, =8:pink")
    assert [(op, val) for op, val, _, _ in rules] == [(">", 8.0), ("<", 8.0), ("==", 8.0)]
    assert rules[0][2] == dm.COLOURS["blue"]


def test_parse_rules_decimal_and_hex():
    (op, val, rgb, _), = dm._parse_rules(">=7.5:#00ff00")
    assert (op, val, rgb) == (">=", 7.5, (0.0, 1.0, 0.0))


@pytest.mark.parametrize("spec", ["", "8:blue", ">x:blue", ">8:nocolour"])
def test_parse_rules_rejects_bad_specs(spec):
    with pytest.raises(ValueError):
        dm._parse_rules(spec)


def test_match_rule_first_match_wins():
    rules = dm._parse_rules(">=8:blue,>5:orange")
    assert dm._match_rule(9, rules)[1] == ">=8:blue"
    assert dm._match_rule(6, rules)[1] == ">5:orange"


def test_match_rule_no_match():
    assert dm._match_rule(3, dm._parse_rules(">5:orange")) == (None, None)
