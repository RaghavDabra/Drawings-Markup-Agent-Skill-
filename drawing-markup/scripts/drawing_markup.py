"""
drawing_markup.py  -  colour regions of vector construction drawings (PDF).

Needs only: pymupdf (fitz), numpy, Pillow (openpyxl for Excel tables).  Works in M365 Copilot Analyst / Code Interpreter.

Four operations (all write a new PDF, a check PNG per edit, and a log table):

  label     colour the closed region that contains a text label        e.g. "1:8" on page 8
  leader    colour the region a tag's leader line points to            e.g. "ST2" on page 17
  units     colour apartment polygons from a status table (Excel/CSV)  e.g. sold=green, unsold=red
  area      colour labelled areas by the m2 written in their label     e.g. balconies >8 m2 blue, <8 orange

Examples (command line):
  python drawing_markup.py label  in.pdf out.pdf --page 8  --text "1:8" --colour red
  python drawing_markup.py leader in.pdf out.pdf --page 17 --text ST2   --colour red
  python drawing_markup.py units  in.pdf out.pdf --table sales.xlsx --pages 10-21
  python drawing_markup.py area   in.pdf out.pdf --page 9 --text "Consultation Room" --rules "==15:purple,==14:red"

Or from Python:  import drawing_markup as dm;  dm.run_label(...), dm.run_leader(...), dm.run_units(...),
                 dm.run_area_rule(...)

Design rules that make it reliable:
  * Never guesses coordinates.  Every target is found from the PDF itself (text, leader geometry, filled polygons).
  * Refuses rather than guesses: an open region (fill would leak), a tag with no leader, a unit matched to two
    polygons, or a unit whose drawn area disagrees with its labelled m2 is reported, not coloured.
  * Every edit gets a zoomed PNG so a person (or the agent) can check it before the file is used.
  * All markup goes on its own PDF layer ("Debuilt markup") so it can be switched off in Acrobat/Bluebeam.
"""

import argparse, csv, math, os, re
import numpy as np
from PIL import Image, ImageDraw

try:
    import pymupdf as fitz
except ImportError:  # older installs
    import fitz

COLOURS = {"red": (0.90, 0.10, 0.10), "green": (0.10, 0.65, 0.20), "blue": (0.10, 0.35, 0.90),
           "yellow": (1.00, 0.85, 0.00), "orange": (1.00, 0.55, 0.00), "purple": (0.55, 0.20, 0.75),
           "pink": (1.00, 0.45, 0.70), "teal": (0.00, 0.60, 0.60), "grey": (0.50, 0.50, 0.50),
           "gray": (0.50, 0.50, 0.50), "brown": (0.55, 0.35, 0.15)}
PX_PER_PT = 4.0          # raster resolution for region finding (4 px per PDF point ~ 288 dpi)
MAX_SIDE_PX = 2000       # cap on render size per side; bigger search windows are rendered coarser so they stay fast
WALL_LUM = 200           # pixel darker than this (0-255) counts as a line
COLOUR_SAT = 60          # pixel with channel spread above this counts as coloured (dimension, cloud, stamp)


def colour(c):
    """Colour name, '#RRGGBB' or an (r, g, b) tuple of 0-1 floats -> (r, g, b)."""
    if isinstance(c, (tuple, list)):
        return tuple(c)
    c = c.strip().lower()
    if c.startswith("#"):
        if not re.fullmatch(r"#[0-9a-f]{6}", c):
            raise ValueError("bad hex colour %r (use #RRGGBB)" % c)
        return tuple(int(c[i:i + 2], 16) / 255 for i in (1, 3, 5))
    if c not in COLOURS:
        raise ValueError("unknown colour %r (use one of %s, or #RRGGBB)" % (c, ", ".join(sorted(COLOURS))))
    return COLOURS[c]


# ---------------------------------------------------------------- page handling
class Work:
    """Open a PDF, work on pages in unrotated coordinates, restore rotation on save."""

    def __init__(self, path):
        self.doc = fitz.open(path)
        self.rot = {}
        try:
            self.ocg = self.doc.add_ocg("Debuilt markup", on=True)
        except Exception:   # some PDFs have broken layer tables; then draw without a layer
            self.ocg = 0
        self.log = []          # one dict per target
        self.checks = []       # (page_no, rect, png_name)

    def page(self, n):          # n is 1-based, as people say it
        p = self.doc[n - 1]
        if n not in self.rot:
            self.rot[n] = p.rotation
            p.set_rotation(0)
        return p

    def save(self, out_pdf, check_dir=None, context=60):
        for n, r in self.rot.items():
            self.doc[n - 1].set_rotation(r)
        self.doc.save(out_pdf, garbage=3, deflate=True)
        pngs = []
        if check_dir:
            os.makedirs(check_dir, exist_ok=True)
            for n, rect, name in self.checks:
                p = self.doc[n - 1]
                clip = (rect + (-context, -context, context, context)) * p.rotation_matrix
                clip &= p.rect
                zoom = min(4.0, 1600 / max(clip.width, clip.height))
                path = os.path.join(check_dir, name)
                p.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=clip).save(path)
                pngs.append(path)
        return pngs


# ---------------------------------------------------------------- finding things
def find_text(page, text):
    """Exact-word matches of `text` (case-insensitive), unrotated coords."""
    want = text.strip().lower()
    hits = [fitz.Rect(w[:4]) for w in page.get_text("words") if w[4].strip().lower() == want]
    if not hits:  # fall back to phrase search (multi-word labels)
        hits = [r for r in page.search_for(text)]
    return hits


_DRAWINGS = {}


def _drawings(page):
    """page.get_drawings(), read once per page (before any markup is added to it)."""
    key = (id(page.parent), page.number)
    if key not in _DRAWINGS:
        _DRAWINGS[key] = page.get_drawings()
    return _DRAWINGS[key]


_SEGS = {}
_BLOCKS = {}
_BOXES = {}


def _seg_index(page, cell=1.0):
    """Segments plus a grid of endpoint -> segments, so a leader can be followed without scanning the sheet."""
    key = (id(page.parent), page.number)
    if key not in _SEGS:
        segs = _segments(page)
        grid = {}
        for i, (a, b) in enumerate(segs):
            for p in (a, b):
                grid.setdefault((int(p.x // cell), int(p.y // cell)), []).append(i)
        _SEGS[key] = (segs, grid, cell)
    return _SEGS[key]


def _near(index, pt, r):
    segs, grid, cell = index
    cx, cy, k = int(pt.x // cell), int(pt.y // cell), int(r // cell) + 1
    out = set()
    for dx in range(-k, k + 1):
        for dy in range(-k, k + 1):
            out.update(grid.get((cx + dx, cy + dy), ()))
    return [segs[i] for i in sorted(out)]


def _segments(page):
    segs = []
    for d in _drawings(page):
        for it in d["items"]:
            if it[0] == "l":
                segs.append((fitz.Point(it[1]), fitz.Point(it[2])))
            elif it[0] == "c":
                segs.append((fitz.Point(it[1]), fitz.Point(it[4])))
    return segs


def _tag_shape(page, text_rect):
    """Smallest drawn shape (bubble/ellipse/box) around the tag text, else the text box itself."""
    key = (id(page.parent), page.number)
    if key not in _BOXES:
        _BOXES[key] = [tuple(d["rect"]) for d in _drawings(page)]
    t = tuple(text_rect)
    best = None
    for x0, y0, x1, y1 in _BOXES[key]:
        if x0 <= t[0] and y0 <= t[1] and x1 >= t[2] and y1 >= t[3] \
                and x1 - x0 < (t[2] - t[0]) * 6 and y1 - y0 < (t[3] - t[1]) * 6:
            if best is None or (x1 - x0) * (y1 - y0) < (best[2] - best[0]) * (best[3] - best[1]):
                best = (x0, y0, x1, y1)
    return fitz.Rect(best) if best else fitz.Rect(text_rect)


def follow_leader(page, text_rect, snap=1.5, join=0.6, arrow_len=8.0):
    """Follow the leader line that starts at a tag. Returns (tip, unit_direction, points) or None."""
    shape = _tag_shape(page, text_rect)
    zone = shape + (-snap, -snap, snap, snap)
    index = _seg_index(page)
    # starting segments: one end touches the bubble edge, the other end is outside it
    starts = []
    for a, b in _near(index, (zone.tl + zone.br) / 2, max(zone.width, zone.height) / 2 + 1):
        for p, q in ((a, b), (b, a)):
            if zone.contains(p) and not zone.contains(q) and abs(p - q) > 2:
                starts.append((p, q))
    if not starts:
        return None
    # take the longest start (bubble outlines are short arcs)
    p, q = max(starts, key=lambda s: abs(s[0] - s[1]))
    pts, used = [p, q], {(round(p.x, 2), round(p.y, 2), round(q.x, 2), round(q.y, 2))}
    cur = q
    for _ in range(200):
        nxt = []
        for a, b in _near(index, cur, join):
            for s, e in ((a, b), (b, a)):
                key = (round(s.x, 2), round(s.y, 2), round(e.x, 2), round(e.y, 2))
                rkey = (key[2], key[3], key[0], key[1])
                if abs(s - cur) < join and key not in used and rkey not in used and abs(s - e) > 0.05:
                    nxt.append((s, e, key))
        if not nxt:
            break
        if len(nxt) > 1 and all(abs(s - e) < arrow_len for s, e, _ in nxt):
            break  # arrowhead: several short strokes meet here -> this is the tip
        long_ = [n for n in nxt if abs(n[0] - n[1]) >= arrow_len] or nxt
        s, e, key = max(long_, key=lambda n: abs(n[0] - n[1]))
        used.add(key)
        pts.append(e)
        cur = e
    tip = pts[-1]
    v = pts[-1] - pts[-2]
    L = abs(v) or 1
    return tip, fitz.Point(v.x / L, v.y / L), pts


_TEXTLESS = {}  # text-free copy of each page, made once


def _render(page, clip, hide_text=True, px=None):
    """Render clip to a numpy RGB array; text removed so labels don't break the region."""
    src = page
    key = (id(page.parent), page.number)
    if hide_text and key in _TEXTLESS:
        src = _TEXTLESS[key][0]
    elif hide_text:
        tmp = fitz.open()
        tmp.insert_pdf(page.parent, from_page=page.number, to_page=page.number)
        src = tmp[0]
        src.set_rotation(0)
        for w in src.get_text("words"):
            src.add_redact_annot(fitz.Rect(w[:4]))
        src.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE, graphics=fitz.PDF_REDACT_LINE_ART_NONE)
        _TEXTLESS[key] = (src, tmp)
    pix = src.get_pixmap(matrix=fitz.Matrix(px or PX_PER_PT, px or PX_PER_PT), clip=clip, alpha=False)
    return np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, 3).copy()


def _dilate(a, r):
    """Square dilation of a boolean array by r pixels (numpy only)."""
    if r <= 0:
        return a
    out = a.copy()
    for k in range(1, r + 1):
        out[k:, :] |= a[:-k, :]; out[:-k, :] |= a[k:, :]
    b = out.copy()
    for k in range(1, r + 1):
        out[:, k:] |= b[:, :-k]; out[:, :-k] |= b[:, k:]
    return out


def _flood(wall, sx, sy):
    space = Image.fromarray(np.where(wall, 0, 255).astype(np.uint8)).copy()  # copy: PIL cannot fill a numpy-backed image
    if space.getpixel((sx, sy)) == 0:  # seed sits on a line: step to the nearest open pixel
        arr = np.array(space)
        y0, x0 = max(0, sy - 12), max(0, sx - 12)
        ys, xs = np.nonzero(arr[y0:sy + 13, x0:sx + 13])
        if len(xs) == 0:
            return None
        k = np.argmin((xs + x0 - sx) ** 2 + (ys + y0 - sy) ** 2)
        sx, sy = int(xs[k] + x0), int(ys[k] + y0)
    ImageDraw.floodfill(space, (sx, sy), 128)
    return np.array(space) == 128


ROOM = {"wall_lum": 120, "gaps": (1, 2, 4, 6, 8), "start": 250, "px": 2.0, "room": True}
# ROOM preset: for plans drawn as plain linework. Only black walls stop the fill (grey furniture, fixtures and
# hatching are ignored) and wider gaps are closed, so the fill covers a whole room or apartment interior.
def _arc_centre(a, b, c, e):
    """Centre of the circular arc a cubic bezier approximates (intersection of the end normals)."""
    t1, t2 = b - a, c - e
    n1, n2 = fitz.Point(-t1.y, t1.x), fitz.Point(-t2.y, t2.x)
    det = n1.x * (-n2.y) - n1.y * (-n2.x)
    if abs(det) < 1e-9:
        return None
    d = e - a
    u = (d.x * (-n2.y) - d.y * (-n2.x)) / det
    return fitz.Point(a.x + u * n1.x, a.y + u * n1.y)


def _dark_linework(page, clip, px, max_grey=0.45, m_per_pt=None):
    """Boolean mask of only the dark (black / near-black) vector strokes and fills inside clip.
    Furniture, fixtures and hatching drawn in grey are left out, so a room fill passes over them.
    With m_per_pt (sheet scale known), every door swing (an arc of 0.5-1.3 m radius) is closed by drawing
    its two radii, so an open doorway does not join a room to the corridor."""
    W, H = int(clip.width * px) + 1, int(clip.height * px) + 1
    im = Image.new("L", (W, H), 0)
    dr = ImageDraw.Draw(im)
    tx = lambda p: ((p.x - clip.x0) * px, (p.y - clip.y0) * px)
    dark = lambda c: c is not None and max(c) <= max_grey and max(c) - min(c) < 0.15
    for d in page.get_drawings():
        if not d["rect"].intersects(clip):
            continue
        st, fi = dark(d.get("color")), dark(d.get("fill"))
        if not (st or fi):
            continue
        w = max(1, int(round((d.get("width") or 0) * px)))
        for it in d["items"]:
            if it[0] == "l":
                pts = [tx(it[1]), tx(it[2])]
            elif it[0] == "c":
                a, b, c2, e = it[1:5]
                if m_per_pt and st:
                    hinge = _arc_centre(a, b, c2, e)
                    if hinge is not None:
                        rad = abs(hinge - a) * m_per_pt
                        if 0.5 <= rad <= 1.3 and abs(abs(hinge - e) - abs(hinge - a)) < 0.1 * abs(hinge - a):
                            dr.line([tx(a), tx(hinge), tx(e)], fill=255, width=max(w, 2))
                pts = [tx(fitz.Point((1-t)**3*a.x+3*(1-t)**2*t*b.x+3*(1-t)*t*t*c2.x+t**3*e.x,
                                     (1-t)**3*a.y+3*(1-t)**2*t*b.y+3*(1-t)*t*t*c2.y+t**3*e.y)) for t in np.linspace(0, 1, 9)]
            elif it[0] == "re":
                r = fitz.Rect(it[1]); pts = [tx(r.tl), tx(r.tr), tx(r.br), tx(r.bl), tx(r.tl)]
            elif it[0] == "qu":
                q = it[1]; pts = [tx(q.ul), tx(q.ur), tx(q.lr), tx(q.ll), tx(q.ul)]
            else:
                continue
            if fi and len(pts) > 2:
                dr.polygon(pts, fill=255)
            if st or fi:
                dr.line(pts, fill=255, width=w)
    return np.array(im)[: int(clip.height * px), : int(clip.width * px)] > 0


GAP_STEPS = (1, 3, 6, 10)  # px; lets the fill close hairline gaps where CAD lines don't quite meet


def region_at(page, seed, ignore_coloured_lines=True, start=120, max_win=500, max_area_frac=0.25,
              wall_lum=None, gaps=None, px=None, room=False, target_pt2=None, m_per_pt=None):
    """Flood-fill the closed region around `seed` (unrotated pt). Returns dict or raises ValueError.

    Runs the fill several times, each time closing slightly bigger gaps in the linework. A fill that
    escapes through a hairline gap is much bigger than the next one, so the first size that stays
    stable is taken as the real region."""
    win = start
    while True:
        PX = min(px or PX_PER_PT, MAX_SIDE_PX / (2.0 * win))
        want = fitz.Rect(seed.x - win, seed.y - win, seed.x + win, seed.y + win)
        clip = want & page.rect
        img = _render(page, clip, px=PX)
        lum = img.mean(axis=2)
        sat = img.max(axis=2).astype(int) - img.min(axis=2)
        wall = lum < (wall_lum or WALL_LUM)
        if ignore_coloured_lines:  # a pixel is coloured if its tint is strong relative to how dark it is (catches anti-aliased edges)
            wall &= (sat < COLOUR_SAT) & (sat < 0.45 * (255 - lum))
        if room:  # walls = dark vector linework only (thin black glazing lines count, grey furniture does not)
            dl = _dark_linework(page, clip, PX, m_per_pt=m_per_pt)
            h, w = wall.shape
            wall = np.zeros_like(wall)
            wall[: dl.shape[0], : dl.shape[1]] = dl[:h, :w]
        sx, sy = int((seed.x - clip.x0) * PX), int((seed.y - clip.y0) * PX)
        fills = []
        for r in (gaps or GAP_STEPS):
            m = _flood(_dilate(wall, r), sx, sy)
            if m is None:
                continue
            edge = m[0].any() or m[-1].any() or m[:, 0].any() or m[:, -1].any()
            fills.append((r, m, edge, int(m.sum())))
        if not fills:
            raise ValueError("seed point is inside solid linework")
        pick = None
        closed = [f for f in fills if not f[2]]
        if target_pt2 and closed:  # labelled area known: take the gap size whose fill is closest to it
            best = min(closed, key=lambda f: abs(math.log(max(f[3], 1) / (target_pt2 * PX * PX))))
            if 0.6 < best[3] / (target_pt2 * PX * PX) < 1.3:
                pick = best
        for i, (r, m, edge, n) in enumerate(fills if pick is None else []):
            nxt = fills[i + 1][3] if i + 1 < len(fills) else None
            if not edge and (nxt is None or n <= 1.6 * nxt + 50):
                pick = fills[i]; break
        if pick is None and any(f[2] for f in fills) and win < max_win:
            win *= 2
            continue
        if pick is None:
            raise ValueError("region is not closed (fill would leak outside the panel)")
        break
    r, m, edge, n = pick
    m = _dilate(m, r) & ~wall | (m)  # give back the pixels the gap-closing took off the edges
    # fill holes (text, leader, arrowheads inside the region)
    outside = Image.fromarray(np.where(m, 0, 255).astype(np.uint8)).copy()
    h, w = m.shape
    for x, y in ((0, 0), (w - 1, 0), (0, h - 1), (w - 1, h - 1)):
        if outside.getpixel((x, y)) == 255:
            ImageDraw.floodfill(outside, (x, y), 64)
    mask = (np.array(outside) != 64) | m
    ys, xs = np.nonzero(mask)
    box = fitz.Rect(xs.min(), ys.min(), xs.max() + 1, ys.max() + 1)
    rect = fitz.Rect(clip.x0 + box.x0 / PX, clip.y0 + box.y0 / PX,
                     clip.x0 + box.x1 / PX, clip.y0 + box.y1 / PX)
    sub = mask[int(box.y0):int(box.y1), int(box.x0):int(box.x1)]
    if rect.get_area() > max_area_frac * page.rect.get_area():
        raise ValueError("region is implausibly large (%.0f%% of sheet)" % (100 * rect.get_area() / page.rect.get_area()))
    return {"rect": rect, "mask": sub, "fill_ratio": float(sub.mean()), "gap_closed_pt": round(r / PX, 2) if r > 1 else 0, "px": PX}


def paint_region(work, page, reg, rgb, opacity=0.45):
    """Rectangular regions become a vector rectangle; other shapes an exact transparent overlay."""
    r = reg["rect"]
    if reg["fill_ratio"] > 0.97:
        sh = page.new_shape()
        sh.draw_rect(r)
        sh.finish(color=rgb, fill=rgb, fill_opacity=opacity, stroke_opacity=0.9, width=0.6, oc=work.ocg)
        sh.commit()
        return "vector rectangle"
    a = (reg["mask"] * int(255 * opacity)).astype(np.uint8)
    h, w = a.shape
    rgba = np.zeros((h, w, 4), dtype=np.uint8)
    rgba[..., 0], rgba[..., 1], rgba[..., 2] = [int(255 * v) for v in rgb]
    rgba[..., 3] = a
    tmp = Image.fromarray(rgba, "RGBA")
    import io
    buf = io.BytesIO(); tmp.save(buf, "PNG")
    page.insert_image(r, stream=buf.getvalue(), oc=work.ocg, overlay=True)
    return "exact-shape overlay"


# ---------------------------------------------------------------- operation 1: label
def run_label(pdf, out_pdf, page_no, text, rgb, which=None, check_dir=None, room=False):
    """room=True: fill the whole room/apartment around the label on a linework plan (see ROOM)."""
    work = Work(pdf)
    page = work.page(page_no)
    hits = find_text(page, text)
    for i, r in enumerate(hits, 1):
        row = {"page": page_no, "target": text, "match": f"{i} of {len(hits)}", "found_at": _fmt(r)}
        if which and i not in which:
            row.update(status="skipped (not requested)"); work.log.append(row); continue
        c = (r.tl + r.br) / 2
        try:
            reg = region_at(page, c, **(ROOM if room else {}))
            how = paint_region(work, page, reg, rgb)
            row.update(status="coloured", method=how, region=_fmt(reg["rect"]), gap_closed_pt=reg["gap_closed_pt"])
            name = f"check_p{page_no}_{_safe(text)}_{i}.png"
            work.checks.append((page_no, reg["rect"], name)); row["check_png"] = name
        except ValueError as e:
            row.update(status="NOT coloured", reason=str(e))
        work.log.append(row)
    if not hits:
        work.log.append({"page": page_no, "target": text, "status": "NOT found on page"})
    work.save(out_pdf, check_dir)
    return work.log


# ---------------------------------------------------------------- operation 2: leader
def run_leader(pdf, out_pdf, page_no, text, rgb, which=None, check_dir=None):
    work = Work(pdf)
    page = work.page(page_no)
    hits = find_text(page, text)
    for i, r in enumerate(hits, 1):
        row = {"page": page_no, "target": f"{text} leader", "match": f"{i} of {len(hits)}", "found_at": _fmt(r)}
        if which and i not in which:
            row.update(status="skipped (not requested)"); work.log.append(row); continue
        lead = follow_leader(page, r)
        if not lead or abs(lead[0] - ((r.tl + r.br) / 2)) < 10:
            row.update(status="NOT coloured", reason="no leader line (legend/schedule entry?)"); work.log.append(row); continue
        tip, d, pts = lead
        seed = fitz.Point(tip.x + 2 * d.x, tip.y + 2 * d.y)
        row["leader_tip"] = _fmt(fitz.Rect(tip, tip))
        try:
            reg = region_at(page, seed)
            how = paint_region(work, page, reg, rgb)
            row.update(status="coloured", method=how, region=_fmt(reg["rect"]), gap_closed_pt=reg["gap_closed_pt"])
            box = fitz.Rect(reg["rect"]) | r
            name = f"check_p{page_no}_{_safe(text)}_{i}.png"
            work.checks.append((page_no, box, name)); row["check_png"] = name
        except ValueError as e:
            row.update(status="NOT coloured", reason=str(e))
        work.log.append(row)
    if not hits:
        work.log.append({"page": page_no, "target": f"{text} leader", "status": "NOT found on page"})
    work.save(out_pdf, check_dir)
    return work.log


# ---------------------------------------------------------------- operation 3: units from a table
def _subpaths(d):
    """The closed outlines making up a filled drawing path, as lists of (x, y)."""
    out, poly = [], []
    def close():
        nonlocal poly
        if len(poly) > 2:
            out.append(poly)
        poly = []
    last = None
    for it in d["items"]:
        if it[0] in ("re", "qu"):
            q = fitz.Rect(it[1]).quad if it[0] == "re" else it[1]
            close(); poly = [(p.x, p.y) for p in (q.ul, q.ur, q.lr, q.ll)]; close(); last = None; continue
        a, b = it[1], it[-1]
        if last is None or abs(a - last) > 0.01:
            close(); poly = [(a.x, a.y)]
        poly.append((b.x, b.y)); last = b
    close()
    return out


def _path_area(d):
    """Shoelace area of a filled drawing path (pt^2), subpaths summed by absolute value."""
    area = 0.0
    for poly in _subpaths(d):
        s = 0.0
        for (x1, y1), (x2, y2) in zip(poly, poly[1:] + poly[:1]):
            s += x1 * y2 - x2 * y1
        area += abs(s) / 2
    return area


def _inside(d, pt):
    """Is pt inside the filled path d (even-odd rule over its outlines)?"""
    if not d["rect"].contains(pt):
        return False
    hit = False
    for poly in _subpaths(d):
        for (x1, y1), (x2, y2) in zip(poly, poly[1:] + poly[:1]):
            if (y1 > pt.y) != (y2 > pt.y) and pt.x < x1 + (pt.y - y1) * (x2 - x1) / (y2 - y1):
                hit = not hit
    return hit


def _leader_unit(page, stack, polys, max_len=150.0):
    """Index of the unit polygon a tag's leader line ends in, or None (no short leader, or it ends in no unit)."""
    lead = follow_leader(page, stack)
    if not lead:
        return None
    tip, _, pts = lead
    if sum(abs(a - b) for a, b in zip(pts, pts[1:])) > max_len:
        return None  # wandered along walls: not a real leader
    inside = [k for k, d in enumerate(polys) if _inside(d, tip)]
    return min(inside, key=lambda k: _path_area(polys[k])) if inside else None


def unit_polygons(page, min_area=150.0):
    """Coloured filled polygons (Revit/ArchiCAD colour-fill areas) that look like units."""
    out = []
    for d in _drawings(page):
        f = d.get("fill")
        if not f or max(f) - min(f) < 0.04 or d["rect"].get_area() < min_area:
            continue
        if d["rect"].get_area() > 0.08 * page.rect.get_area():
            continue  # whole-floor backgrounds
        out.append(d)
    return out


def _tag_stack(page, r, gap=4.0):
    """The label plus lines stacked directly under/over it (type, area) -> (bbox, text)."""
    key = (id(page.parent), page.number)
    if key not in _BLOCKS:
        _BLOCKS[key] = [(fitz.Rect(b[:4]), " ".join(b[4].split())) for b in page.get_text("blocks")]
    lines = _BLOCKS[key]
    box, txt = fitz.Rect(r), []
    vertical = r.height > 1.5 * r.width
    own = [t for lr, t in lines if (lr & box).get_area() > 0.5 * min(lr.get_area(), box.get_area())]
    shape = re.sub(r"\d", "9", own[0]) if own else None
    grown = True
    while grown:
        grown = False
        for lr, t in lines:
            if (lr & box).get_area() > 0.5 * min(lr.get_area(), box.get_area()):
                if t not in txt: txt.append(t)
                continue
            cx, cy = (lr.x0 + lr.x1) / 2, (lr.y0 + lr.y1) / 2
            if re.sub(r"\d", "9", t) == shape:
                continue  # another unit's label: not part of this tag
            if vertical:  # text rotated 90 degrees: lines sit side by side
                near = box.y0 - 6 <= cy <= box.y1 + 6 and (-2 <= lr.x0 - box.x1 <= gap or -2 <= box.x0 - lr.x1 <= gap)
            else:
                near = box.x0 - 6 <= cx <= box.x1 + 6 and (-2 <= lr.y0 - box.y1 <= gap or -2 <= box.y0 - lr.y1 <= gap)
            if near:
                box |= lr; txt.append(t); grown = True
    return box, " | ".join(txt)


def _label_m2(stext):
    """Unit area from its tag: the largest 'NN m2' figure (ignores balcony/terrace figures)."""
    vals = [float(v) for v in re.findall(r"(\d+(?:\.\d+)?)\s*m(?:²|2)", stext)]
    return max(vals) if vals else None


def _dist(a, b):
    dx = max(b.x0 - a.x1, a.x0 - b.x1, 0); dy = max(b.y0 - a.y1, a.y0 - b.y1, 0)
    return math.hypot(dx, dy)


def read_table(path):
    if path.lower().endswith((".xlsx", ".xlsm", ".xls")):
        import openpyxl
        ws = openpyxl.load_workbook(path, data_only=True).active
        rows = list(ws.iter_rows(values_only=True))
        head = [str(h).strip() if h is not None else "" for h in rows[0]]
        return [dict(zip(head, r)) for r in rows[1:] if any(v is not None for v in r)]
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def _pick(row, *names):
    low = {str(k).strip().lower(): v for k, v in row.items()}
    for n in names:
        n = str(n).strip().lower()
        if n in low and low[n] not in (None, ""):
            return str(low[n]).strip()
    return ""


class _UnitKeys(dict):
    """Unit lookup that treats 3.1, 3.10 and 03.10 as the same unit. Excel often stores unit numbers as
    numbers, which drops the leading and trailing zeros the drawings show."""
    @staticmethod
    def canon(k):
        k = str(k).strip()
        try:
            return repr(float(k)) if re.fullmatch(r"[0-9]+(\.[0-9]+)?", k) else k
        except ValueError:
            return k
    def __contains__(self, k): return dict.__contains__(self, self.canon(k))
    def __getitem__(self, k): return dict.__getitem__(self, self.canon(k))
    def setdefault(self, k, v=None): return dict.setdefault(self, self.canon(k), v)


def run_units(pdf, out_pdf, table, pages=None, status_colours=None, aspect=None, check_dir=None,
              unit_col=("unit", "apartment", "apt", "unit no", "lot"), status_col=("status", "sold", "sale status"),
              aspect_col=("aspect", "orientation", "facing"), tag_col=("tag contains", "building")):
    """Colour each unit's own drawn polygon by status. aspect='north' limits to rows whose aspect contains it."""
    status_colours = status_colours or {"sold": colour("green"), "unsold": colour("red")}
    rows = read_table(table)
    work = Work(pdf)
    pages = pages or range(1, len(work.doc) + 1)
    wanted = []
    for r in rows:
        u = _pick(r, *unit_col); s = _pick(r, *status_col).lower(); a = _pick(r, *aspect_col)
        if s in ("yes", "y", "true", "1", "sold", "exchanged", "settled", "unconditional"): s = "sold"
        elif s in ("no", "n", "false", "0", "unsold", "available", "for sale", ""): s = "unsold"
        keep = not aspect or aspect.lower() in a.lower()
        wanted.append({"unit": u, "status": s, "aspect": a, "tag": _pick(r, *tag_col), "keep": keep})
    by_id = _UnitKeys()
    for w in wanted:
        by_id.setdefault(w["unit"], []).append(w)
    # Two passes over the sheets. "calibrate" learns the drawn-area / labelled-m2 ratio from every unit on every
    # sheet; "paint" then checks each unit against it. A sheet with only a few units (e.g. a penthouse level)
    # can't calibrate itself: its own median would just be whichever polygon happened to sit nearest a label.
    sheet_ratios, poly_cache, global_med = {}, {}, None
    for phase, n in [("calibrate", n) for n in pages] + [("paint", n) for n in pages]:
        if phase == "paint" and global_med is None:
            allr = [r for rs in sheet_ratios.values() for r in rs]
            global_med = float(np.median(allr)) if len(allr) >= 5 else 0.0
        page = work.page(n)
        polys = poly_cache.get(n)
        words = [(fitz.Rect(w[:4]), w[4]) for w in page.get_text("words")]
        cands = []  # (dist, unit, label_rect, stack_text, poly_index)
        if polys is None and not any(t in by_id for _, t in words):
            polys = unit_polygons(page) if any(page.search_for(u) for u in by_id if " " in u) else None
        for lr, t in words:
            if t not in by_id:
                continue
            if polys is None:
                polys = poly_cache[n] = unit_polygons(page)
            if not polys:
                break
            stack, stext = _tag_stack(page, lr)
            via = _leader_unit(page, stack, polys)
            if via is not None:  # the tag's leader line ends inside this unit: that is the unit it labels
                cands.append((0.0, t, lr, stext, via))
                continue
            for k, d in enumerate(polys):
                cands.append((_dist(stack, d["rect"]), t, lr, stext, k))
        if polys == [] or (polys is None and not cands):
            if phase == "paint":
                _units_by_room(work, page, n, by_id)
            continue
        if not cands:
            continue
        # one-to-one assignment. Pass 1 pairs nearest-first to learn the sheet's drawn-area / labelled-m2 ratio;
        # pass 2 re-pairs using distance + area agreement, so a label sitting between two units picks the right one.
        def assign(cost):
            used_poly, used_label, out = set(), set(), []
            for c in sorted(cands, key=cost):
                dist, t, lr, stext, k = c
                key = (t, round(lr.x0), round(lr.y0))
                if key in used_label or k in used_poly or dist > 40:
                    continue
                used_label.add(key); used_poly.add(k); out.append((t, lr, stext, k, dist))
            return out
        area = {}
        def ratio(stext, k):
            m2 = _label_m2(stext)
            if not m2:
                return None
            if k not in area:
                area[k] = _path_area(polys[k])
            return area[k] / m2
        matches = assign(lambda c: c[0])
        ratios = [ratio(stext, k) for t, lr, stext, k, dist in matches]
        ratios = [r for r in ratios if r]
        if phase == "calibrate":
            sheet_ratios[n] = ratios
            continue
        med = float(np.median(ratios)) if ratios else None
        if global_med and (len(ratios) < 5 or not med or abs(math.log(med / global_med)) > 0.25):
            med = global_med  # too few units on this sheet (or an odd sheet): use the drawing set's ratio
        if med:
            def cost(c):
                r = ratio(c[3], c[4])
                return c[0] + (60 * abs(math.log(r / med)) if r else 0)
            matches = assign(cost)
        # pair each label on the sheet with a table row. Tagged rows ("Tag contains" = FS) claim labels whose
        # tag text contains that word; an untagged row is used only if exactly one label is left for it.
        pairs = []
        for t in sorted({mm[0] for mm in matches}):
            occ = [mm for mm in matches if mm[0] == t]
            left = []
            for mm in occ:
                tagged = [w for w in by_id[t] if w["tag"] and w["tag"].lower() in mm[2].lower()]
                if tagged:
                    pairs.append((mm, tagged[0], None))
                else:
                    left.append(mm)
            plain = [w for w in by_id[t] if not w["tag"]]
            if len(left) == 1 and len(plain) == 1:
                pairs.append((left[0], plain[0], None))
            elif left and plain:
                for mm in left:
                    pairs.append((mm, plain[0], "unit number appears %d times on this sheet; add a 'Tag contains' column (e.g. FS)" % len(occ)))
        for (t, lr, stext, k, dist), w, problem in pairs:
            if not w["keep"]:
                continue  # filtered out (aspect), but it still claimed its label so duplicates resolve correctly
            row = {"page": n, "target": t, "label_text": stext, "status_in_table": w["status"],
                   "aspect": w["aspect"], "found_at": _fmt(lr), "label_to_unit_pt": round(dist, 1)}
            r = ratio(stext, k)
            if r and med:
                dev = r / med - 1
                row["area_check"] = f"{dev:+.0%}"
                if abs(dev) > 0.15 and not problem:
                    problem = "drawn area does not match labelled m2 (wrong polygon?)"
            if problem:
                row.update(status="NOT coloured", reason=problem)
                work.log.append(row); continue
            rgb = status_colours.get(w["status"], colour("yellow"))
            d = polys[k]
            sh = page.new_shape()
            for it in d["items"]:
                if it[0] == "l": sh.draw_line(it[1], it[2])
                elif it[0] == "c": sh.draw_bezier(it[1], it[2], it[3], it[4])
                elif it[0] == "re": sh.draw_rect(it[1])
                elif it[0] == "qu": sh.draw_quad(it[1])
            sh.finish(color=rgb, fill=rgb, fill_opacity=0.55, width=0.8, closePath=d.get("closePath", True),
                      even_odd=d.get("even_odd", False), oc=work.ocg)
            sh.commit()
            row.update(status="coloured " + {"sold": "green", "unsold": "red"}.get(w["status"], "yellow (status unclear)"),
                       method="unit's own drawn polygon")
            w["done"] = True
            work.log.append(row)
        if matches:
            name = f"check_p{n}_units.png"
            box = fitz.Rect()
            for t, lr, stext, k, dist in matches:
                box |= polys[k]["rect"]
            work.checks.append((n, box, name))
            _legend(work, page, status_colours, os.path.basename(table))
    for w in wanted:
        if w["keep"] and not w.get("done") and not any(_UnitKeys.canon(r.get("target", "")) == _UnitKeys.canon(w["unit"]) for r in work.log):
            work.log.append({"target": w["unit"], "status": "NOT found on the pages searched", "status_in_table": w["status"]})
    work.save(out_pdf, check_dir, context=20)
    return work.log


def _sheet_m2_per_pt2(page):
    """Real m2 per PDF pt2 from the title block (e.g. '1:100 @ A1'), or None."""
    m = re.search(r"1\s*:\s*(\d+)\s*@\s*A(\d)", page.get_text())
    if not m:
        return None
    long_mm = {0: 1189, 1: 841, 2: 594, 3: 420, 4: 297}[int(m.group(2))]
    mm_per_pt = long_mm / max(page.rect.width, page.rect.height)
    return (mm_per_pt * int(m.group(1)) / 1000) ** 2


def _units_by_room(work, page, n, by_id):
    """Plans with no coloured unit areas: fill the room/apartment interior around each unit label."""
    k = _sheet_m2_per_pt2(page)
    for t, rows in by_id.items():
        t = rows[0]["unit"]
        hits = find_text(page, t)
        if not hits:
            continue
        for i, lr in enumerate(hits, 1):
            stack, stext = _tag_stack(page, lr)
            tagged = [w for w in rows if w["tag"] and w["tag"].lower() in stext.lower()]
            plain = [w for w in rows if not w["tag"]]
            w = tagged[0] if tagged else (plain[0] if plain else None)
            if not w or not w["keep"]:
                continue
            row = {"page": n, "target": t, "label_text": stext, "status_in_table": w["status"], "aspect": w["aspect"],
                   "found_at": _fmt(lr), "method": "room fill from label (linework plan)"}
            if len(hits) > 1 and not tagged:
                row.update(status="NOT coloured", reason="label appears %d times on this sheet; add a 'Tag contains' column" % len(hits))
                work.log.append(row); continue
            m2 = _label_m2(stext)
            try:
                reg = region_at(page, (lr.tl + lr.br) / 2, target_pt2=(m2 / k if m2 and k else None),
                                m_per_pt=(math.sqrt(k) if k else None), **ROOM)
            except ValueError as e:
                row.update(status="NOT coloured", reason=str(e)); work.log.append(row); continue
            m2 = _label_m2(stext)
            if m2 and k:
                drawn = reg["mask"].sum() / reg["px"] ** 2 * k
                row["area_check"] = f"{drawn / m2 - 1:+.0%} (rooms behind closed walls, e.g. ensuites, are not filled)"
                if drawn > 1.3 * m2 or drawn < 0.5 * m2:
                    row.update(status="NOT coloured", reason="filled area %.0f m2 vs labelled %.0f m2" % (drawn, m2))
                    work.log.append(row); continue
            rgb = {"sold": colour("green"), "unsold": colour("red")}.get(w["status"], colour("yellow"))
            paint_region(work, page, reg, rgb, opacity=0.5)
            row["status"] = "coloured " + {"sold": "green", "unsold": "red"}.get(w["status"], "yellow (status unclear)")
            w["done"] = True
            work.log.append(row)
            work.checks.append((n, reg["rect"], f"check_p{n}_{_safe(t)}_{i}.png"))


def _legend(work, page, sc, source):
    x, y = 30, 30
    h = 18 + 14 * len(sc)
    sh = page.new_shape()
    sh.draw_rect(fitz.Rect(x - 6, y - 6, x + 200, y + h))
    sh.finish(color=(0, 0, 0), fill=(1, 1, 1), width=0.5, oc=work.ocg)
    for i, (k, rgb) in enumerate(sc.items()):
        sh.draw_rect(fitz.Rect(x, y + 12 + 14 * i, x + 18, y + 22 + 14 * i))
        sh.finish(color=rgb, fill=rgb, fill_opacity=0.55, oc=work.ocg)
    sh.commit()
    page.insert_text((x, y + 4), ("Sales status: " + source)[:48], fontsize=7, oc=work.ocg)
    for i, k in enumerate(sc):
        page.insert_text((x + 24, y + 20 + 14 * i), k.capitalize(), fontsize=8, oc=work.ocg)


# ---------------------------------------------------------------- helpers / CLI
def _fmt(r):
    return f"({r.x0:.0f},{r.y0:.0f})-({r.x1:.0f},{r.y1:.0f})"


def _safe(s):
    return re.sub(r"[^A-Za-z0-9]+", "_", s).strip("_")


def write_log(log, path):
    keys = []
    for r in log:
        for k in r:
            if k not in keys: keys.append(k)
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=keys); w.writeheader(); w.writerows(log)


def _pages(s):
    out = []
    for part in s.split(","):
        a, _, b = part.partition("-")
        out += range(int(a), int(b or a) + 1)
    return out



# ---------------------------------------------------------------- operation 4: area rule
def _parse_rules(spec):
    """'>8:blue,<8:orange,=8:pink' -> [('>',8.0,(r,g,b),'>8'), ...]. First match wins."""
    rules = []
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        m = re.match(r"(>=|<=|==|=|>|<)\s*(\d+(?:\.\d+)?)\s*:\s*(.+)", part)
        if not m:
            raise ValueError("bad rule %r (use e.g. '>8:blue')" % part)
        op = "==" if m.group(1) == "=" else m.group(1)
        rules.append((op, float(m.group(2)), colour(m.group(3).strip()), part))
    if not rules:
        raise ValueError("no rules given")
    return rules


def _match_rule(area, rules):
    """First rule whose test the area passes -> (rgb, label), else (None, None)."""
    import operator
    ops = {">": operator.gt, "<": operator.lt, ">=": operator.ge, "<=": operator.le, "==": operator.eq}
    for op, val, rgb, label in rules:
        if ops[op](area, val):
            return rgb, label
    return None, None


def _calib_scale(page, labels, room):
    """pt2 per m2 for this sheet: from the title block scale, else median of flooded labelled rooms."""
    mpp2 = _sheet_m2_per_pt2(page)
    if mpp2:
        return 1.0 / mpp2, math.sqrt(mpp2)
    ratios = []
    for r, stext, area in labels:
        if not area:
            continue
        c = (r.tl + r.br) / 2
        try:
            reg = region_at(page, c, **(ROOM if room else {}))
            filled = reg["mask"].sum() / (reg["px"] ** 2)
            ratios.append(filled / area)
        except ValueError:
            continue
    if len(ratios) < 2:
        return None, None
    ratios.sort()
    return ratios[len(ratios) // 2], None   # median pt2 per m2


def _best_region(page, label_rect, target_pt2, kw):
    """Try the label centre, then a ring of seeds around it; return the fill closest to the
    labelled area (within 0.5-2x). Returns (region|None, frac|None, error_message)."""
    c = (label_rect.tl + label_rect.br) / 2
    step = math.sqrt(target_pt2) * 0.35 if target_pt2 else 20.0
    seeds = [c]
    for dx, dy in ((0, -1), (0, 1), (-1, 0), (1, 0), (-1, -1), (1, 1), (-1, 1), (1, -1)):
        seeds.append(fitz.Point(c.x + dx * step, c.y + dy * step))
    best = None; best_err = "region not found"
    for s in seeds:
        try:
            reg = region_at(page, s, **kw)
        except ValueError as e:
            best_err = str(e); continue
        if not target_pt2:
            return reg, None, ""
        filled = reg["mask"].sum() / (reg["px"] ** 2)
        frac = filled / target_pt2
        rw, rh = reg["rect"].width, reg["rect"].height
        if max(rw, rh) > 8 * max(min(rw, rh), 1e-6):   # long thin strip (wall, balustrade), not a room
            best_err = "only a thin strip found near the label"
            continue
        if 0.5 < frac < 2.0:
            score = abs(math.log(frac))
            if best is None or score < best[1]:
                best = (reg, score, frac)
    if best:
        return best[0], best[2], ""
    return None, None, "room could not be isolated (open-plan, or no fill near the labelled area)"


def _on_colour_fill(page, c):
    """True if point c sits on a small non-white, non-black vector fill (a coloured room/area polygon)."""
    for d in page.get_drawings():
        f = d.get("fill")
        if f is None or len(f) != 3 or min(f) > 0.97 or max(f) < 0.05:
            continue
        r = d["rect"]
        if r.contains(c) and r.width * r.height < 2e5:
            return True
    return False


def run_area_rule(pdf, out_pdf, page_no, text, rules, which=None, check_dir=None, room=None):
    """Colour each labelled area by the m2 figure written in its own label.

    `rules` is a list from _parse_rules: (op, value, rgb, label); the first rule a label's
    area satisfies decides its colour. The sheet scale (title block, or a median calibrated
    across the labelled rooms) tells region_at how big the room should be, so the fill snaps to
    the ROOM outline and ignores furniture. A label with no m2, or a room that can't be isolated
    (e.g. a balcony open to the room), is reported rather than guessed."""
    work = Work(pdf)
    page = work.page(page_no)
    hits = find_text(page, text)
    labels = [(r, _tag_stack(page, r)[1]) for r in hits]
    labels = [(r, st, _label_m2(st)) for (r, st) in labels]
    if room is None:   # auto: rooms already drawn as coloured fills -> flood that fill; plain linework -> room mode
        pts = [(r.tl + r.br) / 2 for r, _, a in labels if a]
        on = sum(_on_colour_fill(page, c) for c in pts)
        room = not (pts and on * 2 > len(pts))
    pt2_per_m2, m_per_pt = _calib_scale(page, labels, room)
    for i, (r, stext, area) in enumerate(labels, 1):
        row = {"page": page_no, "target": text, "match": f"{i} of {len(hits)}",
               "found_at": _fmt(r), "label_text": stext, "area_m2": area}
        if which and i not in which:
            row.update(status="skipped (not requested)"); work.log.append(row); continue
        if area is None:
            row.update(status="NOT coloured", reason="no m2 area in the label")
            work.log.append(row); continue
        rgb, hit = _match_rule(area, rules)
        row["rule"] = hit
        if rgb is None:
            row.update(status="NOT coloured", reason="area %s matches no rule" % area)
            work.log.append(row); continue
        target_pt2 = area * pt2_per_m2 if pt2_per_m2 else None
        kw = dict(ROOM) if room else {}
        if target_pt2:
            kw["target_pt2"] = target_pt2
            if m_per_pt:
                kw["m_per_pt"] = m_per_pt
        reg, best_frac, err = _best_region(page, r, target_pt2, kw)
        if reg is None:
            row.update(status="NOT coloured", reason=err)
        else:
            how = paint_region(work, page, reg, rgb)
            row.update(status="coloured", method=how, region=_fmt(reg["rect"]),
                       gap_closed_pt=reg["gap_closed_pt"])
            name = f"check_p{page_no}_{_safe(text)}_{_safe(str(area))}_{i}.png"
            work.checks.append((page_no, reg["rect"], name)); row["check_png"] = name
        work.log.append(row)
    if not hits:
        work.log.append({"page": page_no, "target": text, "status": "NOT found on page"})
    work.save(out_pdf, check_dir)
    return work.log


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("op", choices=["label", "leader", "units", "area"])
    ap.add_argument("pdf"); ap.add_argument("out")
    ap.add_argument("--page", type=int); ap.add_argument("--pages")
    ap.add_argument("--text"); ap.add_argument("--which", help="e.g. 1,3 to colour only those matches")
    ap.add_argument("--colour", default="red")
    ap.add_argument("--table"); ap.add_argument("--aspect")
    ap.add_argument("--rules", help="area rules, e.g. \">8:blue,<8:orange,=8:pink\"")
    ap.add_argument("--no-room", dest="room", action="store_false", default=None, help="area op: force coloured-fill plan (default: auto-detect)")
    ap.add_argument("--checks", default=None, help="folder for check PNGs (default: next to output)")
    a = ap.parse_args()
    checks = a.checks or os.path.splitext(a.out)[0] + "_checks"
    which = [int(x) for x in a.which.split(",")] if a.which else None
    if a.op == "label":
        log = run_label(a.pdf, a.out, a.page, a.text, colour(a.colour), which, checks)
    elif a.op == "leader":
        log = run_leader(a.pdf, a.out, a.page, a.text, colour(a.colour), which, checks)
    elif a.op == "area":
        log = run_area_rule(a.pdf, a.out, a.page, a.text, _parse_rules(a.rules), which, checks, room=a.room)
    else:
        log = run_units(a.pdf, a.out, a.table, _pages(a.pages) if a.pages else None, aspect=a.aspect, check_dir=checks)
    write_log(log, os.path.splitext(a.out)[0] + "_log.csv")
    for r in log:
        print(r)
