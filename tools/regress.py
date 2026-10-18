"""Re-run real projects and compare with their known-good results before uploading a new skill zip.

    python tools/regress.py tools/cases.json

cases.json lists each project (drawings stay on your machine, only paths and expected counts go in the file):
    [{"name": "Prospect St", "pdf": "...pdf", "xlsx": "...xlsx", "style": "label", "min_coloured": 60}]
A case with "expect_scan": true passes when inspect() refuses the PDF. Exit code 1 if any case fails.
"""
import json, os, sys, tempfile, time

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "drawing-markup", "scripts"))
import drawing_markup as dm  # noqa: E402


def run(case, tmp):
    info = dm.inspect(case["pdf"], case.get("xlsx"))
    if case.get("expect_scan"):
        return not info["ok"], "refused as scan" if not info["ok"] else "NOT refused"
    if not info["ok"]:
        return False, "refused as scan"
    out = os.path.join(tmp, "out.pdf")
    log = dm.run_units(case["pdf"], out, case["xlsx"], check_dir=os.path.join(tmp, "checks"),
                       style=case.get("style", "label"))
    coloured = {r["target"] for r in log if str(r.get("status", "")).startswith("coloured")}
    missing = sorted({r["target"] for r in log if str(r.get("status", "")).startswith("NOT")} - coloured)
    ok = len(coloured) >= case["min_coloured"]
    return ok, "%d/%d units coloured (need %d); not coloured: %s" % (
        len(coloured), info["rows"], case["min_coloured"], ", ".join(missing) or "none")


def main(path):
    cases = json.load(open(path))
    failed = 0
    for case in cases:
        t = time.time()
        with tempfile.TemporaryDirectory() as tmp:
            try:
                ok, msg = run(case, tmp)
            except Exception as e:  # a crash is a failure, not a stop: report every case
                ok, msg = False, "crashed: %r" % e
        failed += not ok
        print("%s  %-28s %4.0fs  %s" % ("PASS" if ok else "FAIL", case["name"], time.time() - t, msg))
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(__file__), "cases.json"))
