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

import argparse, collections, csv, math, os, re
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
        self.max_checks = None # render only the first n check images (each new sheet rendered costs seconds)

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
            lists = {}  # one display list per page: building it is the slow part on big sheets
            for n, rect, name in self.checks[:self.max_checks]:
                p = self.doc[n - 1]
                if n not in lists:
                    lists[n] = p.get_displaylist()
                clip = (rect + (-context, -context, context, context)) * p.rotation_matrix
                clip &= p.rect
                zoom = min(4.0, 1600 / max(clip.width, clip.height))
                path = os.path.join(check_dir, name)
                lists[n].get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=clip).save(path)
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
        # one page-sized redaction removes every glyph but (with the options below) keeps all linework;
        # a redaction per word is quadratic and takes minutes on dense sheets
        src.add_redact_annot(src.rect + (-10, -10, 10, 10))
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


_DARK_ITEMS = {}  # page -> list of (is_stroke, width, items) for dark drawings only, filtered once
_DARK_MASK = {}   # (page, px, m_bucket) -> full-page dark-linework mask


def _dark_items(page, max_grey=0.45):
    key = (id(page.parent), page.number)
    if key not in _DARK_ITEMS:
        dark = lambda c: c is not None and max(c) <= max_grey and max(c) - min(c) < 0.15
        out = []
        for d in _drawings(page):
            st, fi = dark(d.get("color")), dark(d.get("fill"))
            if st or fi:
                out.append((st, fi, max(1, round(d.get("width") or 0)), d["items"]))
        _DARK_ITEMS[key] = out
    return _DARK_ITEMS[key]


def _dark_linework(page, clip, px, max_grey=0.45, m_per_pt=None):
    """Boolean mask of only the dark (black / near-black) vector strokes and fills inside clip.
    Furniture, fixtures and hatching drawn in grey are left out, so a room fill passes over them.
    With m_per_pt (sheet scale known), every door swing (an arc of 0.5-1.3 m radius) is closed by drawing
    its two radii, so an open doorway does not join a room to the corridor.

    The whole-page mask is rendered once per resolution and cached; each call just crops it. Rebuilding it
    from the page's (often >100k) vector items for every seed window was what made multi-unit linework
    sheets time out."""
    pr = page.rect
    key = (id(page.parent), page.number, round(px, 3), round((m_per_pt or 0), 4))
    mask = _DARK_MASK.get(key)
    if mask is None:
        W, H = int(pr.width * px) + 1, int(pr.height * px) + 1
        im = Image.new("L", (W, H), 0)
        dr = ImageDraw.Draw(im)
        tx = lambda p: ((p.x - pr.x0) * px, (p.y - pr.y0) * px)
        for st, fi, w0, items in _dark_items(page, max_grey):
            w = max(1, int(round(w0 * px)))
            for it in items:
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
        mask = _DARK_MASK[key] = np.array(im) > 0
    x0 = int(round((clip.x0 - pr.x0) * px)); y0 = int(round((clip.y0 - pr.y0) * px))
    W, H = int(clip.width * px), int(clip.height * px)
    return mask[y0:y0 + H, x0:x0 + W]


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
    sideways = r.height > r.width                  # a short tag ("505") on a rotated sheet: too square to tell
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
            side = box.y0 - 6 <= cy <= box.y1 + 6 and (-2 <= lr.x0 - box.x1 <= gap or -2 <= box.x0 - lr.x1 <= gap)
            stack = box.x0 - 6 <= cx <= box.x1 + 6 and (-2 <= lr.y0 - box.y1 <= gap or -2 <= box.y0 - lr.y1 <= gap)
            if vertical or (sideways and page.rotation in (90, 270)):  # text rotated 90 degrees: lines sit side by side
                near = side
            else:
                near = stack
            if near:
                box |= lr; txt.append(t); grown = True
    return box, " | ".join(txt)


def _label_m2(stext):
    """Unit area from its tag: the largest 'NN m2' figure (ignores balcony/terrace figures)."""
    vals = [float(v) for v in re.findall(r"(\d+(?:\.\d+)?)\s*(?:m²|m2|sqm|sq\.?\s*m)\b", stext, re.I)]
    return max(vals) if vals else None


def _dist(a, b):
    dx = max(b.x0 - a.x1, a.x0 - b.x1, 0); dy = max(b.y0 - a.y1, a.y0 - b.y1, 0)
    return math.hypot(dx, dy)


UNIT_COLS = ("unit", "unit no", "unit no.", "apartment", "apt", "apt no", "lot", "lot no")
STATUS_COLS = ("status", "sale status", "sales status", "sold")


def read_table(path, sheet=None, unit_col=UNIT_COLS, status_col=STATUS_COLS):
    """Rows of a sales table as dicts. The header row is found automatically (sales schedules often have
    title and summary rows above it); rows after it without a unit number (totals, notes) are dropped."""
    if path.lower().endswith((".xlsx", ".xlsm", ".xls")):
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True)
        rows = list((wb[sheet] if sheet else wb.active).iter_rows(values_only=True))
    else:
        with open(path, newline="", encoding="utf-8-sig") as f:
            rows = list(csv.reader(f))
    want_u = {c.lower() for c in unit_col}; want_s = {c.lower() for c in status_col}
    hi = 0
    for i, r in enumerate(rows[:40]):
        cells = {str(v).strip().lower() for v in r if v is not None}
        if cells & want_u and cells & want_s:
            hi = i; break
    head = [str(h).strip() if h is not None else "" for h in rows[hi]]
    ucol = next((k for k, h in enumerate(head) if h.lower() in want_u), None)
    out = []
    for r in rows[hi + 1:]:
        if not any(v not in (None, "") for v in r):
            continue
        if ucol is not None:
            u = r[ucol] if ucol < len(r) else None
            if str(u or "").strip().lower().startswith("total"):
                break  # summary rows below the total are not units
            if u in (None, ""):
                continue
        out.append(dict(zip(head, r)))
    return out


def _pick(row, *names):
    low = {str(k).strip().lower(): v for k, v in row.items()}
    for n in names:
        n = str(n).strip().lower()
        if n in low and low[n] not in (None, ""):
            return str(low[n]).strip()
    return ""


_TOWER_CODE_RE = re.compile(r"([A-Za-z])([0-9]{1,2})\.([0-9]{2})")   # tower + level . unit, e.g. B3.01


class _UnitKeys(dict):
    """Unit lookup that treats 3.1, 3.10 and 03.10 as the same unit. Excel often stores unit numbers as
    numbers, which drops the leading and trailing zeros the drawings show."""
    @staticmethod
    def canon(k):
        k = str(k).strip().rstrip("+*")
        try:
            return repr(float(k)) if re.fullmatch(r"[0-9]+(\.[0-9]+)?", k) else k
        except ValueError:
            return k
    def _key(self, k):
        """Exact unit first; else a drawing code with a building prefix, T1.0501 -> 501, when the table
        has no such code itself (two towers both using 0501 then need a 'Tag contains' column)."""
        c = self.canon(k)
        if dict.__contains__(self, c):
            return c
        m = _TOWER_CODE_RE.fullmatch(str(k).strip().rstrip("+*"))
        if m:                                          # drawing B3.01 = schedule 301B (level, unit, tower)
            t = self.canon("%s%s%s" % (m.group(2), m.group(3), m.group(1).upper()))
            if dict.__contains__(self, t):
                return t
        m = re.fullmatch(r"[A-Za-z]+[0-9]*[.\-]([0-9]+)", str(k).strip().rstrip("+*"))
        return self.canon(m.group(1)) if m else c
    def __contains__(self, k): return dict.__contains__(self, self._key(k))
    def __getitem__(self, k): return dict.__getitem__(self, self._key(k))
    def setdefault(self, k, v=None): return dict.setdefault(self, self.canon(k), v)


STATUS_COLOURS = {"sold": colour("blue"), "unsold": colour("pink"), "settled": colour("grey"),
                  "reserved": colour("orange")}


def status_of(s):
    """Normalise a sales-table status: settled / reserved / sold / unsold, else the text itself (coloured yellow)."""
    s = str(s or "").strip().lower()
    if s.startswith("settle"):
        return "settled"
    if s.startswith(("reserv", "hold", "eoi")):
        return "reserved"
    if s.startswith(("unsold", "available", "for sale")) or s in ("", "no", "n", "false", "0", "none"):
        return "unsold"
    if s.startswith(("sold", "exchanged", "unconditional", "contract")) or s in ("yes", "y", "true", "1"):
        return "sold"
    return s


def _status_word(st, sc):
    rgb = sc.get(st)
    name = next((n for n, v in COLOURS.items() if v == rgb), None)
    return ("coloured " + (name or "custom")) if rgb else "coloured yellow (status unclear)"


# Apartment-label matching (the reliable "follow the labelisation" method).
# Every sheet labels each apartment with its number next to its FLOOR AREA, e.g. "1.03+  TWO BEDROOM  78.4 m2".
# The rule is deliberately strict so it marks the real apartment label and nothing else - no walls, rooms or
# geometry are involved:
#   * the figure must be an AREA (m2 / sqm / square metres), never a volume (m3) - storage boxes are labelled
#     "1.08  7.5 m3  STORE" right next to a unit number, and must not be mistaken for the apartment;
#   * the area value must be in a sane dwelling range (>= _AREA_MIN m2), which also rejects those small stores;
#   * door / window tags (1.03-01) carry no area and are skipped;
#   * each unit number is paired with its NEAREST area figure, so a number never grabs a neighbour's area.
_AREA_UNIT_RE = re.compile(r"^(m²|m2|sqm|sq\.?m|m\^?2|㎡)\.?$", re.I)      # an area unit (NOT m3)
_AREA_JOIN_RE = re.compile(r"^(\d{1,4}(?:\.\d+)?)\s*(?:m²|m2|sqm|㎡)$", re.I)  # joined form "78.4m2"
_VOL_RE = re.compile(r"^(m³|m3|m\^?3|cu|cubic)", re.I)                     # a volume unit - always rejected
_NUM_RE = re.compile(r"^\d{1,4}(?:\.\d+)?$")
_AREA_MIN, _AREA_MAX = 20.0, 3000.0   # a dwelling's floor area in m2; excludes stores/wet-area scraps and dimensions


def _is_area_unit_seq(ws, j):
    """Do the word(s) at ws[j] spell a FLOOR-AREA unit (m2 / sqm / square metre(s) / metre square)? Never m3."""
    u = ws[j][4].strip()
    if _VOL_RE.match(u):
        return False
    if _AREA_UNIT_RE.match(u):
        return True
    low = u.lower()
    nxt = ws[j + 1][4].strip().lower() if j + 1 < len(ws) else ""
    if low in ("square", "sq", "sq.") and nxt in ("metre", "metres", "meter", "meters", "m", "m2", "m²"):
        return True
    if low in ("metre", "metres", "meter", "meters") and nxt == "square":
        return True
    return False


def _area_anchors(ws):
    """Every floor-area figure on the page: a list of (rect, value_m2). A figure is a number in the dwelling
    range immediately followed by an area unit, or the joined form "78.4m2". Volumes (m3) are excluded."""
    out = []
    for i, wd in enumerate(ws):
        t = wd[4].strip()
        m = _AREA_JOIN_RE.match(t)
        if m and _AREA_MIN <= float(m.group(1)) <= _AREA_MAX:
            out.append((fitz.Rect(wd[:4]), float(m.group(1))))
            continue
        if _NUM_RE.fullmatch(t) and _AREA_MIN <= float(t) <= _AREA_MAX:
            for j in range(i + 1, min(len(ws), i + 3)):
                wj = ws[j]
                same_line = wj[5:7] == wd[5:7]                # PDF says same text line (works on rotated sheets)
                if not same_line and (abs(wj[1] - wd[1]) > 20 or (wj[0] - wd[2]) > 40 or wj[0] < wd[0] - 5):
                    continue
                if _VOL_RE.match(wj[4].strip()):
                    break
                if _is_area_unit_seq(ws, j):
                    out.append((fitz.Rect(wd[:4]) | fitz.Rect(wj[:4]), float(t)))
                    break
    return out


_UNIT_PREFIX = {"UNIT", "APARTMENT", "APT", "APT.", "LOT", "TOWNHOUSE", "VILLA", "RESIDENCE", "DWELLING", "SHOP"}
_TYPE_CODE_RE = re.compile(r"[0-9]{0,2}[A-Za-z]{1,3}[0-9]{0,2}([.\-][0-9A-Za-z]{1,3})?")   # apartment type: 2U.01, 3B, B2-1


def _tag_head(ws, i):
    """Index of the UNIT word heading the tag that unit number ws[i] belongs to, and whether the tag is a
    stack of floors, else (None, False). Accepts 'UNIT 4', a type code between them ('UNIT 2U.01' over '507')
    and a typical-floor tag listing one number per floor ('UNIT 2A.01' over '909 1009 1109 ... 1509').
    Each step back must touch the word before it, so a tag never reaches into unrelated text."""
    def touch(a, b):
        return fitz.Rect(a[:4]).intersects(fitz.Rect(b[:4]) + (-20, -20, 20, 20))
    def prefix(w):
        return w[4].strip().upper().rstrip(":") in _UNIT_PREFIX
    j = i - 1
    while j >= 0 and re.fullmatch(r"\d{2,5}", ws[j][4].strip()) and touch(ws[j], ws[j + 1]):
        j -= 1                                         # earlier floors' numbers in the same stacked tag
    stacked = j < i - 1 or (i + 1 < len(ws) and re.fullmatch(r"\d{2,5}", ws[i + 1][4].strip()) is not None
                             and touch(ws[i], ws[i + 1]))
    if j >= 0 and prefix(ws[j]) and (ws[j][5:7] == ws[j + 1][5:7] or touch(ws[j], ws[j + 1])):
        return j, stacked
    if (j >= 1 and prefix(ws[j - 1]) and _TYPE_CODE_RE.fullmatch(ws[j][4].strip())
            and touch(ws[j - 1], ws[j]) and touch(ws[j], ws[j + 1])):
        return j - 1, stacked
    return None, False


def _unit_label_rects(page, by_id):
    """Find each unit's apartment label: its number sitting beside its floor-area figure.
    Returns {unit_key: fitz.Rect}. See the module comment above for the (strict) matching rule."""
    ws = page.get_text("words")
    anchors = _area_anchors(ws)
    best = {}
    for i, wd in enumerate(ws):
        t = wd[4].strip()
        if re.search(r"-\d", t):                       # door / window tag like 1.03-01
            continue
        key = by_id._key(re.sub(r"[+*]+$", "", t)) if isinstance(by_id, _UnitKeys) else _UnitKeys.canon(re.sub(r"[+*]+$", "", t))
        if key not in by_id:
            continue
        head, stacked = _tag_head(ws, i)
        named = head is not None
        if re.fullmatch(r"\d{1,2}(\.\d)?", t) and not named:
            continue                                   # "6" / "4.0" alone is a star rating or EER, not unit 6 / 4
        if not named and not _is_unit_label(t, " ".join(w[4] for w in ws[i + 1:i + 8])):
            continue                                   # a bare dimension integer, not a unit number
        wr = fitz.Rect(wd[:4]); near = None
        for ar, _val in anchors:                       # pair with the NEAREST area figure, same line / column
            dx = max(ar.x0 - wr.x1, wr.x0 - ar.x1, 0.0)
            dy = max(ar.y0 - wr.y1, wr.y0 - ar.y1, 0.0)
            if dx > 70 or dy > 45:
                continue
            d = dx * dx + dy * dy
            if near is None or d < near[0]:
                near = (d, ar)
        if near is None:
            # a tower code (B3.01) that matched a schedule unit (301B) is a label on its own, no area needed
            coded = _TOWER_CODE_RE.fullmatch(t) is not None and key != _UnitKeys.canon(t)
            if not (named or coded):
                continue
            grp = wr                                   # "UNIT 4" / "UNIT 2U.01 507" with no area figure
            if named and not stacked:                  # a stacked tag is shared by floors: mark only this number
                for w in ws[head:i]:
                    grp |= fitz.Rect(w[:4])
            if key not in best:
                best[key] = grp
            continue
        grp = wr | near[1]
        if key not in best or grp.get_area() > best[key].get_area():
            best[key] = grp
    return best


def run_units(pdf, out_pdf, table, pages=None, status_colours=None, aspect=None, check_dir=None,
              unit_col=UNIT_COLS, status_col=STATUS_COLS,
              aspect_col=("aspect", "orientation", "facing"), tag_col=("tag contains", "building"), sheet=None,
              style="label", max_checks=None):
    """Colour each unit by its sales status. Default colours: sold blue, unsold pink, settled grey, reserved
    orange (override with status_colours={"sold": ..., ...}). aspect='north' limits to rows whose aspect contains it.
    style="label" (default) highlights each apartment's number+area label in its status colour - reliable on any
    drawing and never touches corridors or rooms. style="fill" shades each unit's drawn area (works only where the
    plan draws each unit as one fill polygon). style="rooms" fills each whole unit on linework plans (walls only:
    fitout, DA sets) by joining rooms through doors, and highlights its label too."""
    only_listed = status_colours is not None               # caller's own colour list: other statuses stay uncoloured
    status_colours = dict(status_colours or STATUS_COLOURS)
    rows = read_table(table, sheet, unit_col, status_col)
    work = Work(pdf)
    work.max_checks = max_checks
    pages = list(pages or range(1, len(work.doc) + 1))
    wanted = []
    for r in rows:
        u = _pick(r, *unit_col); s = status_of(_pick(r, *status_col)); a = _pick(r, *aspect_col)
        keep = not aspect or aspect.lower() in a.lower()
        wanted.append({"unit": u, "status": s, "aspect": a, "tag": _pick(r, *tag_col), "keep": keep})
    by_id = _UnitKeys()
    for w in wanted:
        by_id.setdefault(w["unit"], []).append(w)

    if style == "rooms":
        # linework plans (fitout / DA sets): fill each whole unit through its doors, and highlight its label too
        for n in pages:
            page = work.page(n)
            _units_by_room(work, page, n, by_id, status_colours if only_listed else None)
        log_fill = work.log; work.log = []
        style = "label"
    else:
        log_fill = []
    if style == "label":
        done = set()
        for n in pages:
            page = work.page(n)
            labels = _unit_label_rects(page, by_id)
            for key, R in labels.items():
                w = next((x for x in by_id[key] if x["keep"]), None)
                if not w or (only_listed and w["status"] not in status_colours):
                    continue
                rgb = status_colours.get(w["status"], colour("yellow"))
                an = page.add_highlight_annot(R + (-2, -2, 2, 2))
                an.set_colors(stroke=tuple(rgb)); an.set_opacity(0.5); an.update()
                unit = by_id[key][0]["unit"]
                work.checks.append((n, R, f"check_p{n}_{_safe(unit)}.png"))
                if key not in done:
                    work.log.append({"page": n, "target": unit, "status_in_table": by_id[key][0]["status"],
                                     "method": "label highlight", "status": _status_word(w["status"], status_colours)})
                    done.add(key)
            if labels:
                _legend(work, page, status_colours, os.path.basename(str(table)))
        work.log = log_fill + work.log
        for w in wanted:
            if only_listed and w["status"] not in status_colours:
                work.log.append({"target": w["unit"], "status_in_table": w["status"], "status": "left uncoloured (status not in the colour list)"})
                continue
            if w["keep"] and _UnitKeys.canon(w["unit"]) not in done:
                work.log.append({"target": w["unit"], "status_in_table": w["status"],
                                 "status": "NOT found (no unit label - number with its m2 or UNIT tag - on the sheets searched)"})
        work.save(out_pdf, check_dir, context=25)
        return work.log

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
        words = [(lr, t) for lr, t in words if t in by_id and _is_unit_label(t, _tag_stack(page, lr)[1])]
        for lr, t in words:
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
                _units_by_room(work, page, n, by_id, status_colours)
            continue
        if not cands:
            continue
        nlabels = len({_UnitKeys.canon(t) for _, t in words})
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
        if phase == "paint" and len({_UnitKeys.canon(mm[0]) for mm in matches}) < 0.5 * nlabels:
            # most labels have no drawn unit shape near them: this is a linework plan, fill rooms instead
            _units_by_room(work, page, n, by_id, status_colours)
            continue
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
            row.update(status=_status_word(w["status"], status_colours), method="unit's own drawn polygon")
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
    text = page.get_text()
    m = re.search(r"1\s*:\s*(\d+)\s*@\s*A(\d)", text)
    if m:
        long_mm = {0: 1189, 1: 841, 2: 594, 3: 420, 4: 297}[int(m.group(2))]
        mm_per_pt = long_mm / max(page.rect.width, page.rect.height)
    else:  # plain "1 : 100": assume the PDF page is the real sheet size (CAD exports are)
        scales = [int(x) for x in re.findall(r"\b1\s*:\s*(\d{2,3})\b", text) if int(x) in (50, 75, 80, 100, 150, 200, 250, 500)]
        if not scales:
            return None
        m = max(set(scales), key=scales.count)
        mm_per_pt = 25.4 / 72
        return (mm_per_pt * m / 1000) ** 2
    return (mm_per_pt * int(m.group(1)) / 1000) ** 2


_UNIT_WORDS = ("apartment", "unit", "apt", "lot", "bed", "studio", "sqm", "m²", "m2", "penthouse", "townhouse")


def _is_unit_label(token, stext):
    """A bare integer like 3008 is usually a dimension. Accept it only next to a unit word or area;
    tokens like 0906, 1.01+ or G01 are distinctive enough on their own."""
    if not re.fullmatch(r"[1-9][0-9]*", token.strip()):
        return True
    return any(w in stext.lower() for w in _UNIT_WORDS)


def _units_by_room(work, page, n, by_id, status_colours=None):
    """Linework plans (no coloured unit areas): fill each whole apartment found by the room graph below."""
    sc = status_colours or STATUS_COLOURS
    k = _sheet_m2_per_pt2(page)
    labels = _unit_label_rects(page, by_id)
    if not labels or not k:
        return
    ws = page.get_text("words")
    found = apartment_masks(page, labels, math.sqrt(k))[0]
    # the printed scale is not always the true one (sheet resized on export): calibrate on the best-filled units
    lab_m2 = {key: max([a[1] for a in _area_anchors(ws) if a[0].intersects(R)] or [0]) for key, R in labels.items()}
    rs = sorted(found[key][3] / lab_m2[key] for key in found if lab_m2.get(key))
    cal = 1.0
    if len(rs) >= 4:
        q = rs[int(0.75 * (len(rs) - 1))]
        if 0.3 < q < 3 and abs(math.log(q)) > 0.12:
            cal = q
    for key, R in labels.items():
        rows = by_id[key]; t = rows[0]["unit"]
        w = next((x for x in rows if x["keep"]), None)
        if not w or (status_colours and w["status"] not in sc):
            continue                                   # status left out of status_colours = leave uncoloured
        m2 = max([a[1] for a in _area_anchors(ws) if a[0].intersects(R)] or [0])
        row = {"page": n, "target": t, "status_in_table": w["status"], "found_at": _fmt(R),
               "method": "whole apartment fill (rooms joined through doors)"}
        if key not in found:
            row.update(status="NOT coloured", reason="label room is open to the outside or corridor")
            work.log.append(row); continue
        rect, mask, nrooms, drawn = found[key]
        drawn /= cal
        if cal != 1.0:
            row["scale_note"] = "areas calibrated x%.2f from this sheet's labels" % (1 / cal)
        if not m2 and not 20 <= drawn <= 1000:
            row.update(status="NOT coloured", reason="no area on the label and the enclosed space is %.0f m2" % drawn)
            work.log.append(row); continue
        if m2:
            row["area_check"] = "drawn %.0f m2 vs labelled %.0f m2 (%+.0f%%)" % (drawn, m2, 100 * (drawn / m2 - 1))
            lo_ok = 0.45 if cal == 1.0 else 0.6
            if drawn > (1.25 if cal == 1.0 else 1.35) * m2 or drawn < lo_ok * m2:
                row.update(status="NOT coloured", reason="filled area %.0f m2 vs labelled %.0f m2" % (drawn, m2))
                work.log.append(row); continue
        paint_region(work, page, {"rect": rect, "mask": mask, "fill_ratio": 0.5}, sc.get(w["status"], colour("yellow")), opacity=0.5)
        row["status"] = _status_word(w["status"], sc)
        w["done"] = True
        work.log.append(row)
        work.checks.append((n, rect, f"check_p{n}_{_safe(t)}.png"))


def _legend(work, page, sc, source):
    x, y = 30, 30
    h = 18 + 14 * len(sc)
    sh = page.new_shape()
    sh.draw_rect(fitz.Rect(x - 6, y - 6, x + 200, y + h))
    sh.finish(color=(0, 0, 0), fill=(1, 1, 1), width=0.5, oc=work.ocg)
    for i, (k, rgb) in enumerate(sc.items()):
        sh.draw_rect(fitz.Rect(x, y + 12 + 14 * i, x + 18, y + 22 + 14 * i))
        sh.finish(color=rgb, fill=rgb, fill_opacity=0.55, oc=work.ocg)
    # one commit for the whole legend: each commit re-scans the page's (often huge) content stream
    sh.insert_text((x, y + 4), ("Sales status: " + source)[:48], fontsize=7)
    for i, k in enumerate(sc):
        sh.insert_text((x + 24, y + 20 + 14 * i), k.capitalize(), fontsize=8)
    sh.finish(color=(0, 0, 0), fill=(0, 0, 0), oc=work.ocg)
    sh.commit()



# ---------------------------------------------------------------- whole-apartment fill on linework plans
# A fitout plan draws walls, not apartments. Room graph: seal the wall linework (CAD wall/door/window layers
# when the PDF kept them, else dark linework), label every enclosed room, link the two rooms either side of
# each door swing, grow each apartment from its label room through its own doors. A room whose doors reach two
# apartments is the corridor/lobby and is never coloured; doorless rooms (balcony, robe) join the one
# apartment they share walls with. Pure numpy, ~5-12 s per sheet.
def _erode(a, r):
    return ~_dilate(~a, r)


def _item_pts(it, n=9):
    if it[0] == "l":
        return [it[1], it[2]]
    if it[0] == "c":
        a, b, c, e = it[1:5]
        return [fitz.Point((1-t)**3*a.x + 3*(1-t)**2*t*b.x + 3*(1-t)*t*t*c.x + t**3*e.x,
                           (1-t)**3*a.y + 3*(1-t)**2*t*b.y + 3*(1-t)*t*t*c.y + t**3*e.y)
                for t in np.linspace(0, 1, n)]
    return None


def _circle3(p, q, r):
    ax, ay, bx, by, cx, cy = p.x, p.y, q.x, q.y, r.x, r.y
    d = 2 * (ax * (by - cy) + bx * (cy - ay) + cx * (ay - by))
    if abs(d) < 1e-9:
        return None
    ux = ((ax*ax + ay*ay) * (by - cy) + (bx*bx + by*by) * (cy - ay) + (cx*cx + cy*cy) * (ay - by)) / d
    uy = ((ax*ax + ay*ay) * (cx - bx) + (bx*bx + by*by) * (ax - cx) + (cx*cx + cy*cy) * (bx - ax)) / d
    return fitz.Point(ux, uy)


def _door_arcs(page, m_per_pt, rmin=0.45, rmax=1.35):
    """Door swings: runs of connected strokes that trace ~a quarter circle of door-leaf radius.
    Works whether the CAD export wrote the swing as one bezier or as many short straight segments.
    Returns [(hinge, p_start, p_end, radius_pt)]."""
    lo, hi = rmin / m_per_pt, rmax / m_per_pt
    out = []

    def test(ch):
        if len(ch) < 3:
            return
        xs = [p.x for p in ch]; ys = [p.y for p in ch]
        if not (0.5 * lo <= max(max(xs) - min(xs), max(ys) - min(ys)) <= 1.2 * hi):
            return
        h = _circle3(ch[0], ch[len(ch) // 2], ch[-1])
        if h is None:
            return
        rad = abs(h - ch[0])
        if not (lo <= rad <= hi):
            return
        if max(abs(abs(h - p) - rad) for p in ch) > 0.06 * rad:
            return
        L = sum(abs(a - b) for a, b in zip(ch, ch[1:]))
        span = L / rad
        if not (1.15 <= span <= 1.95):           # ~65-110 degrees: a door swing, not a full/half circle
            return
        out.append((h, ch[0], ch[-1], rad))

    lay = _layered(page)
    loose = []
    for d in _drawings(page):
        if not d.get("color"):
            continue
        if lay and not _DOOR_LAYER_RE.match(_short_layer(d)):
            continue
        cur = []
        for it in d["items"]:
            pts = _item_pts(it)
            if pts is None:
                test(cur); cur = []; continue
            if cur and abs(pts[0] - cur[-1]) <= 0.3:
                # split at a sharp turn (leaf line meeting the arc) so each piece is tested on its own
                v1 = cur[-1] - cur[-2] if len(cur) > 1 else None
                v2 = pts[1] - pts[0]
                if v1 is not None and abs(v1) > 1e-6 and abs(v2) > 1e-6:
                    cosang = (v1.x * v2.x + v1.y * v2.y) / (abs(v1) * abs(v2))
                    if cosang < math.cos(math.radians(30)):
                        test(cur); cur = list(pts); continue
                cur.extend(pts[1:])
            else:
                test(cur); cur = list(pts)
        test(cur)
        if len(d["items"]) == 1 and d["items"][0][0] == "l":
            a, b = d["items"][0][1], d["items"][0][2]
            if 0 < abs(a - b) <= 0.35 * hi:
                loose.append((a, b))
    # some exports write each tiny chord of a door swing as its own path: chain them end to end
    key = lambda p: (round(p.x * 2), round(p.y * 2))
    ends = {}
    for i, (a, b) in enumerate(loose):
        ends.setdefault(key(a), []).append(i); ends.setdefault(key(b), []).append(i)
    used = set()
    for i in range(len(loose)):
        if i in used:
            continue
        used.add(i)
        chain = list(loose[i])
        for side in (0, 1):
            while True:
                tip = chain[-1] if side == 0 else chain[0]
                nxt = [j for j in ends.get(key(tip), []) if j not in used]
                if len(nxt) != 1:
                    break
                j = nxt[0]; used.add(j); a, b = loose[j]
                far = b if key(a) == key(tip) else a
                if side == 0:
                    chain.append(far)
                else:
                    chain.insert(0, far)
        if len(chain) >= 5:
            # split at sharp turns, test each smooth run
            run = chain[:2]
            for p in chain[2:]:
                v1 = run[-1] - run[-2]; v2 = p - run[-1]
                if abs(v1) > 1e-6 and abs(v2) > 1e-6 and (v1.x * v2.x + v1.y * v2.y) / (abs(v1) * abs(v2)) < math.cos(math.radians(30)):
                    test(run); run = [run[-1], p]
                else:
                    run.append(p)
            test(run)
    return out


def _label_cells(free):
    """Connected components (4-connectivity) of a boolean image, numpy + run-length union-find."""
    H, W = free.shape
    pad = np.zeros((H, W + 2), np.int8)
    pad[:, 1:-1] = free
    dd = np.diff(pad, axis=1)
    sr, sc = np.nonzero(dd == 1)
    er, ec = np.nonzero(dd == -1)
    n = len(sr)
    parent = np.arange(n)

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    row_start = np.searchsorted(sr, np.arange(H + 1))
    for y in range(1, H):
        a0, a1 = row_start[y - 1], row_start[y]
        b0, b1 = row_start[y], row_start[y + 1]
        i, j = a0, b0
        while i < a1 and j < b1:
            if sc[i] < ec[j] and sc[j] < ec[i]:
                ri, rj = find(i), find(j)
                if ri != rj:
                    parent[max(ri, rj)] = min(ri, rj)
            if ec[i] < ec[j]:
                i += 1
            else:
                j += 1
    roots = np.array([find(i) for i in range(n)])
    uniq, lab = np.unique(roots, return_inverse=True)
    img = np.zeros((H, W), np.int32)
    for k in range(n):
        img[sr[k], sc[k]:ec[k]] = lab[k] + 1
    return img, len(uniq)


def _nearest_open(lab, x, y, rad):
    H, W = lab.shape
    x, y = int(round(x)), int(round(y))
    if 0 <= x < W and 0 <= y < H and lab[y, x] > 0:
        return lab[y, x]
    y0, y1, x0, x1 = max(0, y - rad), min(H, y + rad + 1), max(0, x - rad), min(W, x + rad + 1)
    sub = lab[y0:y1, x0:x1]
    ys, xs = np.nonzero(sub)
    if len(xs) == 0:
        return 0
    k = np.argmin((xs + x0 - x) ** 2 + (ys + y0 - y) ** 2)
    return sub[ys[k], xs[k]]


def apartment_masks(page, labels, m_per_pt, PX=1.5, G=6):
    """labels: {unit_key: label rect (pt)}. Returns ({unit_key: bool mask full page}, info)."""
    wall = _wall_mask_layers(page, PX) if _layered(page) else _wall_mask(page, PX)
    wc = _erode(_dilate(wall, G), G)
    lab, n = _label_cells(~wc)
    H, W = lab.shape
    area = np.bincount(lab.ravel(), minlength=n + 1)
    border = set(np.unique(np.concatenate([lab[0], lab[-1], lab[:, 0], lab[:, -1]]))) - {0}
    m2px = (m_per_pt / PX) ** 2
    bg = set(border) | {c for c in range(1, n + 1) if area[c] * m2px > 1500}
    # the space around the building inside the sheet frame: a cell spanning most of the sheet is outside, not a room
    yy, xx = np.nonzero(lab)
    ll = lab[yy, xx]
    x0s = np.full(n + 1, W); x1s = np.zeros(n + 1, int); y0s = np.full(n + 1, H); y1s = np.zeros(n + 1, int)
    np.minimum.at(x0s, ll, xx); np.maximum.at(x1s, ll, xx); np.minimum.at(y0s, ll, yy); np.maximum.at(y1s, ll, yy)
    bg |= {c for c in range(1, n + 1) if x1s[c] - x0s[c] > 0.6 * W or y1s[c] - y0s[c] > 0.6 * H}
    P = lambda pt: ((pt.x - page.rect.x0) * PX, (pt.y - page.rect.y0) * PX)
    cell = lambda pt: _nearest_open(lab, *P(pt), int(0.4 / m_per_pt * PX))
    # door links
    adj = {}
    for h, a, e, r in _door_arcs(page, m_per_pt):
        # leaf = the radius drawn as a line; opening = the other radius (mostly open pixels)
        def openness(q):
            k = 0
            for t in np.linspace(0.2, 0.8, 7):
                x, y = P(h + (q - h) * t)
                xi, yi = int(x), int(y)
                k += 0 <= xi < W and 0 <= yi < H and not wall[yi, xi]
            return k
        b = a if openness(a) >= openness(e) else e
        tip = e if b is a else a
        mid = (a + e) / 2
        bis = mid - h; bis = bis / max(abs(bis), 1e-6)
        r1 = cell(h + bis * (r + 0.25 / m_per_pt))            # room the door swings into
        hb = (h + b) / 2
        away = hb - (h + tip) / 2; away = away / max(abs(away), 1e-6)   # from leaf side across the wall line
        nrm = fitz.Point(-(b - h).y, (b - h).x); nrm = nrm / max(abs(nrm), 1e-6)
        if nrm.x * (tip - h).x + nrm.y * (tip - h).y > 0:
            nrm = -nrm
        r2 = cell(hb + nrm * (0.3 / m_per_pt))                 # room on the other side of the doorway
        if r1 and r2 and r1 != r2 and r1 not in bg and r2 not in bg:
            adj.setdefault(r1, set()).add(r2); adj.setdefault(r2, set()).add(r1)
    seeds = {}
    for key, R in labels.items():
        c = cell((R.tl + R.br) / 2)
        if c and c not in bg:
            seeds[key] = c
    seed_cells = set(seeds.values())
    common = set()
    for _ in range(6):
        dist = {}
        for key, s in seeds.items():
            d = {s: 0}; q = [s]
            while q:
                u = q.pop(0)
                if u in common:
                    continue
                for v in adj.get(u, ()):
                    if v not in d and v not in bg and (v not in seed_cells or v == s):
                        d[v] = d[u] + 1; q.append(v)
            dist[key] = d
        owner = {}
        allc = set().union(*[set(d) for d in dist.values()]) if dist else set()
        for c in allc:
            if c in common:
                continue
            ds = sorted((dist[k][c], k) for k in dist if c in dist[k])
            if len(ds) == 1 or ds[0][0] < ds[1][0]:
                owner[c] = ds[0][1]
        new = set()
        for c in allc:
            if c in seed_cells or c in common:
                continue
            ks = {owner[v] for v in adj.get(c, ()) if v in owner}
            if c in owner:
                ks.add(owner[c])
            if len(ks) >= 2:
                new.add(c)
        if not new:
            break
        common |= new
    # adopt doorless rooms (balcony behind a slider, robes, a bedroom whose door wasn't recognised).
    # Unowned rooms are grouped with the unowned rooms they share a wall with; a group is adopted only if it
    # shares walls with exactly ONE apartment and never with the lobby. Stairs and lift shafts sit against the
    # lobby, so their whole group is refused. The wall test reaches across a partition (~0.1 m) but not
    # comfortably across a party wall.
    wt = int(0.22 / m_per_pt * PX) + 1
    cand = [c for c in range(1, n + 1) if c not in owner and c not in bg and c not in common
            and area[c] * m2px >= 0.3]
    touch = {}
    for c in cand:                       # each cell's own window (bounding boxes computed once above)
        y0, y1 = max(0, y0s[c] - wt), min(H, y1s[c] + wt + 1)
        x0, x1 = max(0, x0s[c] - wt), min(W, x1s[c] + wt + 1)
        sub = lab[y0:y1, x0:x1]
        ring = _dilate(sub == c, wt) & (sub != c) & (sub > 0)
        touch[c] = {t for t in np.unique(sub[ring]).tolist() if t != c and t not in bg and area[t] * m2px >= 0.3}
    par = {c: c for c in cand}
    def f(c):
        while par[c] != c:
            par[c] = par[par[c]]; c = par[c]
        return c
    for c in cand:
        for t in touch[c]:
            if t in par:
                par[f(c)] = f(t)
    groups = {}
    for c in cand:
        groups.setdefault(f(c), []).append(c)
    for g in groups.values():
        ga = sum(area[c] for c in g) * m2px
        nb = set().union(*[touch[c] for c in g]) - set(g)
        if ga > 80 or nb & common:
            continue
        ks = {owner.get(t) for t in nb} - {None}
        if len(ks) == 1 and all(t in owner for t in nb):
            k = ks.pop()
            for c in g:
                owner[c] = k
    small = 6.0                          # enclosed pockets up to this size (joinery, islands) count as the unit's
    big_free = [c for c in range(1, n + 1) if c not in owner and c not in bg and area[c] * m2px >= small]
    r = int(0.3 / m_per_pt * PX) + 1
    masks = {}
    pad = 2 * r + 2                      # work in the unit's own window: dilate+erode by r only reach 2r out
    for key in seeds:
        cs = [c for c, k in owner.items() if k == key]
        if not cs:
            continue
        wy0, wy1 = max(0, min(y0s[c] for c in cs) - pad), min(H, max(y1s[c] for c in cs) + pad + 1)
        wx0, wx1 = max(0, min(x0s[c] for c in cs) - pad), min(W, max(x1s[c] for c in cs) + pad + 1)
        lw = lab[wy0:wy1, wx0:wx1]
        m = np.isin(lw, cs)
        others = np.isin(lw, [c for c, k in owner.items() if k != key] + list(common) + big_free)
        mm = _erode(_dilate(m, r), r) & ~others          # take in the unit's own internal walls
        ys, xs = np.nonzero(mm)
        if len(xs) == 0:
            continue
        y0, y1, x0, x1 = ys.min(), ys.max() + 1, xs.min(), xs.max() + 1
        sub = mm[y0:y1, x0:x1]
        im = Image.fromarray(np.where(np.pad(sub, 1), 0, 255).astype(np.uint8))
        ImageDraw.floodfill(im, (0, 0), 64)
        enclosed = (np.array(im)[1:-1, 1:-1] != 64) & ~sub
        hl, hn = _label_cells(enclosed)
        ha = np.bincount(hl.ravel(), minlength=hn + 1) * m2px
        keep = ha < small; keep[0] = False
        full = np.zeros_like(mm)
        full[y0:y1, x0:x1] = sub | keep[hl]
        masks[key] = (wy0, wx0, full & ~others, len(cs), float(m.sum()) * m2px)
    # a pixel claimed by two units (shared wall) goes to neither
    cnt = np.zeros((H, W), np.uint8)
    for wy0, wx0, full, _nc, _a in masks.values():
        cnt[wy0:wy0 + full.shape[0], wx0:wx0 + full.shape[1]] += full
    out = {}
    for key, (wy0, wx0, full, nc, a) in masks.items():
        full &= cnt[wy0:wy0 + full.shape[0], wx0:wx0 + full.shape[1]] == 1
        ys, xs = np.nonzero(full)
        if len(xs) == 0:
            continue
        y0, y1, x0, x1 = ys.min(), ys.max() + 1, xs.min(), xs.max() + 1
        out[key] = (fitz.Rect((x0 + wx0) / PX, (y0 + wy0) / PX, (x1 + wx0) / PX, (y1 + wy0) / PX) + (page.rect.x0, page.rect.y0, page.rect.x0, page.rect.y0),
                    full[y0:y1, x0:x1], nc, a)
    return out, {"cells": n, "doors": sum(len(v) for v in adj.values()) // 2, "common": len(common)}


def _dash_lines(page, min_gaps=15, min_span=300, dot=3.0, dot_frac=0.3):
    """Axis-aligned lines drawn as dash-dot runs of separate short strokes: grid lines, centre lines,
    hidden/overhead joinery. They are not walls, but sealing small gaps would turn them into walls that
    slice rooms apart. Returns {("h"|"v", coord)}. Facade glazing lines (regular long pieces) don't match."""
    import collections
    H = collections.defaultdict(list); V = collections.defaultdict(list)
    for d in _drawings(page):
        for it in d["items"]:
            if it[0] != "l":
                continue
            a, b = it[1], it[2]
            if abs(a.y - b.y) < 0.05 and abs(a.x - b.x) > 0.05:
                H[round(a.y, 1)].append((min(a.x, b.x), max(a.x, b.x)))
            elif abs(a.x - b.x) < 0.05 and abs(a.y - b.y) > 0.05:
                V[round(a.x, 1)].append((min(a.y, b.y), max(a.y, b.y)))
    out = set()
    for tag, D in (("h", H), ("v", V)):
        for c, segs in D.items():
            if len(segs) < min_gaps:
                continue
            segs.sort()
            span = segs[-1][1] - segs[0][0]
            gaps = sum(1 for p, q in zip(segs, segs[1:]) if q[0] - p[1] > 0.3)
            dots = sum(1 for s in segs if s[1] - s[0] < dot)
            if span >= min_span and gaps >= min_gaps and dots >= dot_frac * len(segs):
                out.add((tag, c))
    return out


def _wall_mask(page, PX, max_grey=0.45):
    """Dark vector linework as a boolean image, minus dash-dot lines (grid, centre and hidden lines)."""
    skip = _dash_lines(page)
    pr = page.rect
    W, Hh = int(pr.width * PX) + 1, int(pr.height * PX) + 1
    im = Image.new("L", (W, Hh), 0)
    dr = ImageDraw.Draw(im)
    tx = lambda p: ((p.x - pr.x0) * PX, (p.y - pr.y0) * PX)
    for st, fi, w0, items in _dark_items(page, max_grey):
        w = max(1, int(round(w0 * PX)))
        for it in items:
            if it[0] == "l":
                a, b = it[1], it[2]
                if abs(a - b) < 60 and ((abs(a.y - b.y) < 0.05 and ("h", round(a.y, 1)) in skip) or
                                        (abs(a.x - b.x) < 0.05 and ("v", round(a.x, 1)) in skip)):
                    continue
                pts = [tx(a), tx(b)]
            elif it[0] == "c":
                pts = [tx(q) for q in _item_pts(it)]
            elif it[0] == "re":
                r = fitz.Rect(it[1]); pts = [tx(r.tl), tx(r.tr), tx(r.br), tx(r.bl), tx(r.tl)]
            elif it[0] == "qu":
                q = it[1]; pts = [tx(q.ul), tx(q.ur), tx(q.lr), tx(q.ll), tx(q.ul)]
            else:
                continue
            if fi and len(pts) > 2:
                dr.polygon(pts, fill=255)
            dr.line(pts, fill=255, width=w)
    return np.array(im)[: int(pr.height * PX) + 1, : int(pr.width * PX) + 1] > 0


_WALL_LAYER_RE = re.compile(r"^(A-?WALL|I-WALL|WALL-|A-?DOOR|A-?WIND|A-?GLAZ|A-?CURT|A-?STRUCT|S-COL|A-?COL|A-?EGRESS)", re.I)
_DOOR_LAYER_RE = re.compile(r"^A-?DOOR", re.I)


def _short_layer(d):
    return re.split(r"[$|]", d.get("layer") or "")[-1].strip()


def _layered(page):
    """True if the PDF kept the architect's CAD layers and they include walls."""
    k = sum(1 for d in _drawings(page) if _WALL_LAYER_RE.match(_short_layer(d)))
    return k >= 200


def _wall_mask_layers(page, PX):
    """Walls, doors, windows, columns and stair/lift cores from their CAD layers; furniture, fixtures,
    joinery, dimensions, grid and symbols are left out whatever colour they are drawn in."""
    pr = page.rect
    W, Hh = int(pr.width * PX) + 1, int(pr.height * PX) + 1
    im = Image.new("L", (W, Hh), 0)
    dr = ImageDraw.Draw(im)
    tx = lambda p: ((p.x - pr.x0) * PX, (p.y - pr.y0) * PX)
    for d in _drawings(page):
        if not _WALL_LAYER_RE.match(_short_layer(d)):
            continue
        w = max(1, int(round((d.get("width") or 0) * PX)))
        fi = d.get("fill") is not None
        for it in d["items"]:
            if it[0] == "l":
                pts = [tx(it[1]), tx(it[2])]
            elif it[0] == "c":
                pts = [tx(q) for q in _item_pts(it)]
            elif it[0] == "re":
                r = fitz.Rect(it[1]); pts = [tx(r.tl), tx(r.tr), tx(r.br), tx(r.bl), tx(r.tl)]
            elif it[0] == "qu":
                q = it[1]; pts = [tx(q.ul), tx(q.ur), tx(q.lr), tx(q.ll), tx(q.ul)]
            else:
                continue
            if fi and len(pts) > 2:
                dr.polygon(pts, fill=255)
            dr.line(pts, fill=255, width=w)
    return np.array(im) > 0


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
    for d in _drawings(page):  # cached: called once per label, and get_drawings() is slow on big sheets
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


SCAN_MESSAGE = ("This file has no vector linework or text the script can read. "
                "Please send the PDF exported from CAD.")


def inspect(pdf, table=None, sheet=None, labels=True):
    """Setup check before any colouring. Returns a dict:
    pages, readable (vector pages with text), scanned (image-only pages), ok (False when nothing is readable:
    show SCAN_MESSAGE and stop), and with a table: rows, statuses {status: count}, and unit_labels
    {page: n} - how many of the table's units have a label the script can find on each page (labels=False skips
    that: it reads every sheet's text, the slow part)."""
    doc = fitz.open(pdf)
    readable, scanned = [], []
    for n, p in enumerate(doc, 1):
        if p.get_fonts() or p.get_text("words"):     # fonts: cheap test that the page has real text
            readable.append(n)
        else:
            covered = sum(fitz.Rect(i["bbox"]).get_area() for i in p.get_image_info())
            if covered > 0.5 * p.rect.get_area() or not p.get_drawings():
                scanned.append(n)
    info = {"pages": len(doc), "readable": readable, "scanned": scanned, "ok": bool(readable)}
    if not readable:
        info["message"] = SCAN_MESSAGE
    if table:
        rows = read_table(table, sheet)
        info["rows"] = len(rows)
        info["statuses"] = dict(collections.Counter(status_of(_pick(r, *STATUS_COLS)) for r in rows))
        by_id = _UnitKeys()
        for r in rows:
            by_id.setdefault(_pick(r, *UNIT_COLS), []).append(r)
        info["unit_labels"] = {}
        for n in (readable if labels else []):
            k = len(_unit_label_rects(doc[n - 1], by_id))
            if k:
                info["unit_labels"][n] = k
    return info


UPLOAD_DIRS = ("/mnt/data", "/mnt/user-data", ".")


DOWNLOAD_LIMIT = 14 * 1024 * 1024   # Copilot's code interpreter drops files over ~15-16 MB


def _small_copies(out, pages, limit=DOWNLOAD_LIMIT):
    """A coloured PDF over the download limit: save only its coloured sheets ('..._sheets.pdf'), split into
    parts that each fit if needed. Returns the files to hand over (just [out] when it already fits)."""
    if os.path.getsize(out) <= limit or not pages:
        return [out]
    stem = os.path.splitext(out)[0]
    def save(sel, path):
        d = fitz.open(out)
        d.select([n - 1 for n in sel])
        d.save(path, garbage=4, deflate=True, use_objstms=1)
        return os.path.getsize(path)
    one = stem + "_sheets.pdf"
    if save(pages, one) <= limit or len(pages) == 1:
        return [one]
    os.remove(one)
    parts, cur = [], []
    for n in pages:                                   # grow each part sheet by sheet until it would not fit
        path = "%s_sheets_part%d.pdf" % (stem, len(parts) + 1)
        if cur and save(cur + [n], path) > limit:
            save(cur, path)
            parts.append(path)
            cur = [n]
        else:
            cur.append(n)
    if cur:
        path = "%s_sheets_part%d.pdf" % (stem, len(parts) + 1)
        save(cur, path)
        parts.append(path)
    return parts


def _is_output(path):
    return "_colour_coded" in os.path.basename(path).lower() or os.path.basename(path) == "log.csv"


def find_all_inputs(folder=None):
    """Every attached drawing PDF (oldest first, so sheets keep the order they were attached in) and the newest
    sales table in the upload folder (M365 Copilot puts attachments in /mnt/data), skipping files this script
    wrote. Returns ([pdf, ...], table or None)."""
    for d in ([folder] if folder else UPLOAD_DIRS):
        if not os.path.isdir(d):
            continue
        files = sorted((os.path.join(d, f) for f in os.listdir(d)), key=os.path.getmtime)
        pdfs = [f for f in files if f.lower().endswith(".pdf") and not _is_output(f)]
        tables = [f for f in files if f.lower().endswith((".xlsx", ".xlsm", ".xls", ".csv")) and not _is_output(f)]
        if pdfs or tables:
            return pdfs, (tables[-1] if tables else None)
    return [], None


def find_inputs(folder=None):
    """The newest attached drawing PDF and sales table: (pdf, table), either can be None."""
    pdfs, table = find_all_inputs(folder)
    return (pdfs[-1] if pdfs else None), table


def colour_by_status(pdf=None, table=None, out=None, style="label", status_colours=None, aspect=None, folder=None,
                     checks=3):
    """Everything for sale-status colouring in ONE call, so the agent needs one code run: find the attachments
    (every attached PDF, e.g. one file per floor), refuse scans, colour, write the logs. Returns a summary:
    ok, message, table, rows, statuses, coloured {colour word: n}, not_found {"level not in this drawing set":
    [...], "no unit tag on its sheet": [...]}, files [{pdf, pages, ok, out, log_csv, unit_label_pages}],
    scanned [pdfs refused as scans], preview [up to `checks` check PNGs], checks_zip, and download_zip (all
    coloured PDFs, logs and checks in one file), downloads (exactly the files to hand over, each under the
    ~15 MB Copilot limit: a big coloured PDF becomes its coloured sheets only, split into parts if needed).
    With one PDF also out, log_csv, pdf, pages, unit_label_pages.
    pdf may be a path or a list. checks: check images to render in total (default 3, None = one per unit)."""
    pdfs, found_table = ([], None) if (pdf is not None and table is not None) else find_all_inputs(folder)
    pdfs = [pdf] if isinstance(pdf, str) else list(pdf or pdfs)
    table = table or found_table
    if not pdfs:
        return {"ok": False, "message": "No drawing PDF attached. Please attach the drawing PDF and the sales Excel."}
    if not table:
        return {"ok": False, "message": "No sales Excel attached. Please attach the sales schedule to colour units by sale status."}
    import zipfile
    rows = read_table(table)
    summary = {"ok": False, "table": table, "rows": len(rows),
               "statuses": dict(collections.Counter(status_of(_pick(r, *STATUS_COLS)) for r in rows)),
               "files": [], "scanned": []}
    base = os.path.dirname(os.path.abspath(pdfs[0]))
    check_root = os.path.join(base, "checks")
    if os.path.isdir(check_root):                # a rerun must not mix in the previous run's images
        for root, _dirs, fs in os.walk(check_root):
            for f in fs:
                if f.endswith(".png"):
                    os.remove(os.path.join(root, f))
    logs, pngs, budget = [], [], checks
    for path in pdfs:
        info = inspect(path, labels=False)
        entry = {"pdf": path, "pages": info["pages"], "ok": info["ok"]}
        summary["files"].append(entry)
        if not info["ok"]:
            summary["scanned"].append(path)
            continue
        stem = os.path.splitext(os.path.basename(path))[0]
        o = out if (out and len(pdfs) == 1) else os.path.join(base, stem + "_Colour_Coded.pdf")
        cdir = check_root if len(pdfs) == 1 else os.path.join(check_root, stem)
        log = run_units(path, o, table, status_colours=status_colours, aspect=aspect, check_dir=cdir, style=style,
                        max_checks=budget)
        log_csv = os.path.splitext(o)[0] + "_log.csv"
        write_log(log, log_csv)
        made = sorted(os.path.join(cdir, f) for f in os.listdir(cdir) if f.endswith(".png")) if os.path.isdir(cdir) else []
        pngs += made
        if budget is not None:
            budget = max(0, budget - len(made))
        lp = sorted({r["page"] for r in log if r.get("page")})
        entry.update(out=o, log_csv=log_csv, unit_label_pages=lp, downloads=_small_copies(o, lp))
        logs.append(log)
    if not logs:
        summary["message"] = SCAN_MESSAGE
        return summary
    summary["ok"] = True
    if summary["scanned"]:
        summary["message"] = "Not coloured (scanned or flattened, no readable linework): " + ", ".join(
            os.path.basename(p) for p in summary["scanned"]) + ". Please send the PDF exported from CAD."
    # a unit counts as done if ANY sheet coloured it; style="rooms" logs a fill row and a label row: count once
    coloured, done = collections.Counter(), set()
    for log in logs:
        for r in log:
            st, key = str(r.get("status", "")), _UnitKeys.canon(r.get("target"))
            if st.startswith("coloured") and key not in done:
                coloured[st.replace("coloured ", "")] += 1
                done.add(key)
    level_of = {_UnitKeys.canon(_pick(r, *UNIT_COLS)): _pick(r, "level", "floor", "lvl") for r in rows}
    found_levels = {level_of.get(u) for u in done} - {None, ""}
    nf = {"level not in this drawing set": [], "no unit tag on its sheet": []}
    for log in logs:
        for r in log:
            u = r.get("target")
            if not str(r.get("status", "")).startswith("NOT") or _UnitKeys.canon(u) in done:
                continue
            lvl = level_of.get(_UnitKeys.canon(u))
            k = "no unit tag on its sheet" if lvl and lvl in found_levels else "level not in this drawing set"
            if u not in nf[k]:
                nf[k].append(u)
    ok_files = [f for f in summary["files"] if f.get("out")]
    stem = os.path.splitext(ok_files[0]["out"])[0] if len(ok_files) == 1 else os.path.join(base, "All_Colour_Coded")
    checks_zip = stem + "_checks.zip"
    with zipfile.ZipFile(checks_zip, "w", zipfile.ZIP_DEFLATED) as z:  # one download instead of dozens of images
        for f in pngs:
            z.write(f, os.path.relpath(f, check_root))
    summary.update(coloured=dict(coloured), not_found={k: v for k, v in nf.items() if v},
                   preview=pngs[:3], checks_zip=checks_zip, check_dir=check_root)
    if len(ok_files) == 1:
        f = ok_files[0]
        summary.update(pdf=f["pdf"], pages=f["pages"], out=f["out"], log_csv=f["log_csv"],
                       unit_label_pages=f["unit_label_pages"], downloads=f["downloads"] + [f["log_csv"], checks_zip])
    else:
        download_zip = os.path.join(base, "All_Colour_Coded.zip")
        with zipfile.ZipFile(download_zip, "w", zipfile.ZIP_DEFLATED) as z:
            for f in ok_files:
                for d in f["downloads"]:
                    z.write(d, os.path.basename(d))
                z.write(f["log_csv"], os.path.basename(f["log_csv"]))
            z.write(checks_zip, os.path.basename(checks_zip))
        summary["download_zip"] = download_zip
        zips = [download_zip]
        if os.path.getsize(download_zip) > DOWNLOAD_LIMIT:   # too big for one download: one zip per drawing
            os.remove(download_zip)
            zips = []
            for f in ok_files:
                zp = os.path.splitext(f["out"])[0] + ".zip"
                with zipfile.ZipFile(zp, "w", zipfile.ZIP_DEFLATED) as z:
                    for d in f["downloads"] + [f["log_csv"]]:
                        z.write(d, os.path.basename(d))
                zips.append(zp)
            summary["download_zip"] = None
        summary["downloads"] = zips + [checks_zip]
    return summary


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("op", choices=["inspect", "label", "leader", "units", "area"])
    ap.add_argument("pdf"); ap.add_argument("out", nargs="?")
    ap.add_argument("--page", type=int); ap.add_argument("--pages")
    ap.add_argument("--text"); ap.add_argument("--which", help="e.g. 1,3 to colour only those matches")
    ap.add_argument("--colour", default="red")
    ap.add_argument("--table"); ap.add_argument("--aspect")
    ap.add_argument("--rules", help="area rules, e.g. \">8:blue,<8:orange,=8:pink\"")
    ap.add_argument("--no-room", dest="room", action="store_false", default=None, help="area op: force coloured-fill plan (default: auto-detect)")
    ap.add_argument("--checks", default=None, help="folder for check PNGs (default: next to output)")
    a = ap.parse_args()
    if a.op == "inspect":
        print(inspect(a.pdf, a.table))
        raise SystemExit(0)
    if a.out is None:
        ap.error("%s needs an output PDF path" % a.op)
    need = {"label": ("page", "text"), "leader": ("page", "text"), "area": ("page", "text", "rules"), "units": ("table",)}
    missing = ["--" + k for k in need[a.op] if getattr(a, k) is None]
    if missing:
        ap.error("%s needs %s" % (a.op, ", ".join(missing)))
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
