import os

import yaml

SKILL = os.path.join(os.path.dirname(__file__), "..", "drawing-markup", "SKILL.md")


def test_skill_frontmatter_is_valid_yaml():
    text = open(SKILL, encoding="utf-8").read()
    assert text.startswith("---\n")
    meta = yaml.safe_load(text.split("---")[1])
    assert meta["name"] == "drawing-markup"
    assert "colour" in meta["description"].lower()
