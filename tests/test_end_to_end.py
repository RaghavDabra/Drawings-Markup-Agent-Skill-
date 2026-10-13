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


def test_write_log_round_trips(tmp_path):
    rows = [{"page": 1, "status": "coloured"}, {"page": 2, "status": "NOT coloured", "reason": "open"}]
    path = tmp_path / "log.csv"
    dm.write_log(rows, path)
    got = list(csv.DictReader(open(path)))
    assert [r["reason"] for r in got] == ["", "open"]


def test_markup_goes_on_its_own_layer(plan, tmp_path):
    out = tmp_path / "out.pdf"
    dm.run_label(str(plan), str(out), 1, "GYM", dm.colour("green"))
    layers = [v["name"] for v in pymupdf.open(out).get_ocgs().values()]
    assert "Debuilt markup" in layers


@pytest.fixture
def balconies(tmp_path):
    """Two balconies whose drawn sizes match their labelled m2 (same scale)."""
    doc = pymupdf.open()
    page = doc.new_page(width=600, height=400)
    for rect, m2 in ((pymupdf.Rect(50, 50, 150, 150), 10), (pymupdf.Rect(300, 220, 400, 270), 5)):
        page.draw_rect(rect, color=(0, 0, 0), width=2)
        page.insert_text((rect.x0 + 15, rect.y0 + 20), "BALCONY\n%d m2" % m2, fontsize=9)
    path = tmp_path / "balconies.pdf"
    doc.save(path)
    return path


@pytest.mark.slow
def test_area_rule_colours_by_labelled_m2(balconies, tmp_path):
    rules = dm._parse_rules(">8:blue,<8:orange")
    log = dm.run_area_rule(str(balconies), str(tmp_path / "out.pdf"), 1, "BALCONY", rules)
    got = {row["area_m2"]: (row["status"], row.get("rule")) for row in log}
    assert got == {10.0: ("coloured", ">8:blue"), 5.0: ("coloured", "<8:orange")}
