import csv

import pymupdf
import pytest

import drawing_markup as dm


@pytest.fixture
def plan(tmp_path):
    """Two closed rooms drawn as black linework, each with a text label."""
    doc = pymupdf.open()
    page = doc.new_page(width=600, height=400)
    for rect, text in ((pymupdf.Rect(50, 50, 250, 250), "GYM"), (pymupdf.Rect(300, 50, 450, 200), "STORE")):
        page.draw_rect(rect, color=(0, 0, 0), width=2)
        page.insert_text((rect.x0 + 20, rect.y0 + 40), text, fontsize=12)
    path = tmp_path / "plan.pdf"
    doc.save(path)
    return path


def test_run_label_colours_the_room(plan, tmp_path):
    out = tmp_path / "out.pdf"
    log = dm.run_label(str(plan), str(out), 1, "GYM", dm.colour("green"), check_dir=str(tmp_path / "checks"))
    assert len(log) == 1 and log[0]["status"] == "coloured"
    x0, y0 = (int(v) for v in log[0]["region"].split(")")[0].strip("(").split(","))
    assert 48 <= x0 <= 56 and 48 <= y0 <= 56
    assert (tmp_path / "checks" / log[0]["check_png"]).exists()


def test_run_label_reports_missing_text(plan, tmp_path):
    log = dm.run_label(str(plan), str(tmp_path / "out.pdf"), 1, "POOL", dm.colour("red"))
    assert log == [{"page": 1, "target": "POOL", "status": "NOT found on page"}]
