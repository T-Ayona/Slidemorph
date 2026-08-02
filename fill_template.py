"""
fill_template.py  (v5 - spec driven)
=====================================

Fills a PowerPoint template with AI-generated content while PRESERVING the
manual entrance animations set up in the source deck, and fixing title
wrapping / title-vs-content overlap.

All template-specific rules (slide layout, section slide indices, per-section
slide counts, placeholder character/bullet limits, and the 6-20 slide
composition table) live in templates/<name>/spec.json. Adding a new template
means adding templates/<name>/template.pptx + templates/<name>/spec.json --
no changes to this file. This engine only knows the generic shape of a deck:

    Start (1 slide) -> Wheel (fixed count) -> Bars (variable) ->
    Hexagon (variable) -> End (1 slide)

with the exact numbers and text-fit rules coming from the spec.

Pipeline (see generate_ai_deck):
  1. load_spec(template_name)           - read templates/<name>/spec.json
  2. build_slide_sequence(count, spec)  - which template slides to clone, in order
  3. build_deck(...)                    - clone slides into a fresh .pptx
  4. generate_content(topic, spec, ...) - Gemini Stage 1: outline (titles/tags/headings)
  5. fill_placeholders(...)             - Gemini Stage 2 (adaptive bullets) + write
  6. save_output(...)

Text-fit mechanics (unchanged from earlier versions):
  - KEEP EXISTING ANIMATIONS. The fill is "structure-aware": it preserves
    every paragraph an animation targets and only rewrites the runs inside it.
  - Bullet COUNT adapts to how many bullets would wrap onto a 2nd line
    (measured with Pillow), per the spec's max_bullets_* / words_per_bullet.
  - Bullet boxes grow in height (never width) to fit; titles are widened /
    given a one-line height so they never overlap the content box, without
    ever changing font size.
  - AI_content / AI_title are renamed per-slide (AI_content_<n>, AI_title_<n>)
    only by rename_unique() (used by the dummy-fill test path) so PowerPoint
    doesn't try to morph text boxes between dissimilar slides; the AI content
    pipeline intentionally leaves names shared so Bar/Hex slides morph.
"""

import argparse
import os
import re
import copy
import json
import time

from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.shapes.group import GroupShape
from pptx.parts.slide import SlidePart
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.exc import PackageNotFoundError
from PIL import ImageFont

R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

# ---------------------------------------------------------------------------
# Namespaces / engine constants (generic -- apply to any template)
# ---------------------------------------------------------------------------
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS = {"a": A, "p": P}
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

EMU = 914400                      # EMU per inch
BULLET = "\u2022 "                # "- "
CONTENT_PT = 18                   # AI_content effective size (defaultTextStyle=1800)
TITLE_PT_DEFAULT = 32             # AI_title (Impact 3200)
LINE_FACTOR = 1.2                 # single-line spacing multiplier
FIT_SLACK_IN = 0.06               # one-line width slack for titles
TITLE_WIDTH_SLACK_IN = 0.20       # extra width when widening a title to one line
GAP_IN = 0.08                     # min gap between title bottom and content top
MIN_TOP_IN = 0.10                 # keep boxes this far from slide edges
MAX_BULLET_ITERS = 3              # regen attempts before trim

# ---- Gemini config ---------------------------------------------------------
# Tried in order; on a quota/429 OR a 404/unavailable error the run transparently
# falls back to the next model.
GEMINI_MODELS = ["gemini-2.0-flash", "gemini-flash-latest", "gemini-2.5-flash",
                 "gemini-3.5-flash", "gemini-3.6-flash", "gemini-3.1-flash-lite",
                 "gemini-flash-lite-latest", "gemini-3.5-flash-lite"]
QUOTA_RETRY_CAP = 45.0                      # max seconds to wait on a rate-limit
MAX_QUOTA_WAITS = 2                         # wait+retry attempts on the last model
API_CALL_DELAY = 2.0                        # seconds between calls (rate limiting)

FONTS_DIR = r"C:\Windows\Fonts"
FONT_FILE = {
    "Impact": "impact.ttf",
    "Calibri Light": "calibril.ttf",
    "Calibri": "calibri.ttf",
    "+mj-lt": "calibril.ttf",     # theme major font = Calibri Light
    "+mn-lt": "calibri.ttf",
}


def font_path(typeface):
    return os.path.join(FONTS_DIR, FONT_FILE.get(typeface, "arial.ttf"))


_FONT_CACHE = {}


def _font(typeface, pt):
    """Load and cache a font by typeface + size. Falls back to Pillow's default
    font when the requested TTF is missing (e.g. non-Windows deploy targets)."""
    key = f"{typeface}-{pt}"
    if key not in _FONT_CACHE:
        try:
            path = font_path(typeface)
            _FONT_CACHE[key] = ImageFont.truetype(path, int(round(pt)))
        except (OSError, IOError):
            print(f"Warning: Font {typeface!r} not found. Using default font.")
            _FONT_CACHE[key] = ImageFont.load_default()
    return _FONT_CACHE[key]


def text_width_in(text, typeface, pt):
    """Rendered width of `text` in inches (font size in pt-as-px -> px==pt)."""
    return _font(typeface, pt).getlength(text) / 72.0


def bul(text):
    return BULLET + text


# ---------------------------------------------------------------------------
# Generic helpers
# ---------------------------------------------------------------------------
def iter_all_shapes(shapes):
    for sh in shapes:
        yield sh
        if isinstance(sh, GroupShape):
            yield from iter_all_shapes(sh.shapes)


def find_shape(slide, name):
    for sh in iter_all_shapes(slide.shapes):
        if sh.name == name:
            return sh
    return None


def _xfrm(shape):
    return shape._element.find(".//a:xfrm", NS)


def geom_emu(shape):
    x = _xfrm(shape)
    off, ext = x.find("a:off", NS), x.find("a:ext", NS)
    return (int(off.get("x")), int(off.get("y")),
            int(ext.get("cx")), int(ext.get("cy")))


def set_geom(shape, ox, oy, cx, cy):
    x = _xfrm(shape)
    off, ext = x.find("a:off", NS), x.find("a:ext", NS)
    off.set("x", str(int(round(ox))));  off.set("y", str(int(round(oy))))
    ext.set("cx", str(int(round(cx))));  ext.set("cy", str(int(round(cy))))


def get_ins(shape):
    """Text-frame insets in inches (defaults: l/r=0.1, t/b=0.05)."""
    b = shape._element.find(".//a:bodyPr", NS)

    def g(attr, default):
        v = b.get(attr) if b is not None else None
        return (int(v) / EMU) if v is not None else default
    return {"l": g("lIns", 0.1), "r": g("rIns", 0.1),
            "t": g("tIns", 0.05), "b": g("bIns", 0.05)}


def run_font(shape, default_typeface, default_pt):
    """Effective (typeface, pt) of the shape's first run (or endParaRPr)."""
    el = shape._element
    rpr = el.find(".//a:r/a:rPr", NS)
    if rpr is None:
        rpr = el.find(".//a:endParaRPr", NS)
    tf, pt = default_typeface, default_pt
    if rpr is not None:
        latin = rpr.find("a:latin", NS)
        if latin is not None and latin.get("typeface"):
            tf = latin.get("typeface")
        if rpr.get("sz"):
            pt = int(rpr.get("sz")) / 100.0
    return tf, pt


# ---------------------------------------------------------------------------
# Run / paragraph surgery (formatting preserving)
# ---------------------------------------------------------------------------
def _template_rpr(txBody):
    r = txBody.find(".//a:r", NS)
    if r is not None and r.find("a:rPr", NS) is not None:
        return copy.deepcopy(r.find("a:rPr", NS))
    epr = txBody.find(".//a:endParaRPr", NS)
    if epr is not None:
        rpr = copy.deepcopy(epr)
        rpr.tag = qn("a:rPr")
        return rpr
    return None


def _para_rpr(p, txBody):
    """rPr to reuse for a paragraph's new runs (its own run first)."""
    r = p.find("a:r", NS)
    if r is not None and r.find("a:rPr", NS) is not None:
        return copy.deepcopy(r.find("a:rPr", NS))
    epr = p.find("a:endParaRPr", NS)
    if epr is not None:
        rpr = copy.deepcopy(epr)
        rpr.tag = qn("a:rPr")
        return rpr
    return _template_rpr(txBody)


def _make_run(parent, rpr, text):
    r = parent.makeelement(qn("a:r"), {})
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    t = parent.makeelement(qn("a:t"), {})
    t.text = text
    if text != text.strip():
        t.set(XML_SPACE, "preserve")
    r.append(t)
    return r


def _rebuild_para(p, lines, rpr):
    """Replace a paragraph's runs/breaks with `lines` (joined by <a:br/>),
    preserving its <a:pPr> and <a:endParaRPr>."""
    for ch in list(p):
        if ch.tag in (qn("a:r"), qn("a:br")):
            p.remove(ch)
    endpr = p.find("a:endParaRPr", NS)
    new_nodes = []
    for i, line in enumerate(lines):
        if i > 0:
            br = p.makeelement(qn("a:br"), {})
            if rpr is not None:
                br.append(copy.deepcopy(rpr))
            new_nodes.append(br)
        new_nodes.append(_make_run(p, rpr, line))
    if endpr is not None:
        for node in new_nodes:
            endpr.addprevious(node)
    else:
        for node in new_nodes:
            p.append(node)


def fill_simple(shape, lines):
    """Rebuild a whole text frame as one paragraph per line (titles/tags/subtitle).
    Preserves run + paragraph formatting; never touches <a:bodyPr>."""
    txBody = shape.text_frame._txBody
    rpr = _template_rpr(txBody)
    ppr = txBody.find(".//a:pPr", NS)
    ppr = copy.deepcopy(ppr) if ppr is not None else None
    for p in txBody.findall("a:p", NS):
        txBody.remove(p)
    for line in lines:
        p = txBody.makeelement(qn("a:p"), {})
        if ppr is not None:
            p.append(copy.deepcopy(ppr))
        p.append(_make_run(p, rpr, line))
        txBody.append(p)


# ---------------------------------------------------------------------------
# Animation-structure inspection
# ---------------------------------------------------------------------------
def animated_para_indices(slide, spid):
    """Paragraph indices the slide's timing animates for shape `spid`."""
    idx = set()
    for spt in slide._element.findall(".//p:spTgt", NS):
        if spt.get("spid") != str(spid):
            continue
        for prg in spt.findall(".//p:pRg", NS):
            idx.update(range(int(prg.get("st")), int(prg.get("end")) + 1))
    return sorted(idx)


# ---------------------------------------------------------------------------
# Content fill + fit
# ---------------------------------------------------------------------------
def content_marL_emu():
    """Left margin for the hanging indent = width of the "- " prefix (EMU)."""
    return int(round(text_width_in(BULLET, "+mj-lt", CONTENT_PT) * EMU))


def count_display_lines(text, usable_in, hang_in):
    """Greedy word-wrap line count with a hanging indent: the first line has the
    full usable width, wrapped continuation lines are narrower by `hang_in`."""
    words = text.split(" ")
    lines, cur, avail = 1, "", usable_in
    for w in words:
        trial = w if not cur else cur + " " + w
        if not cur or text_width_in(trial, "+mj-lt", CONTENT_PT) <= avail:
            cur = trial
        else:
            lines += 1
            cur = w
            avail = usable_in - hang_in     # continuation lines are indented
    return lines


def _set_pPr_hanging(p, marL_emu):
    """Give paragraph `p` a hanging indent (first line flush, wraps indented)."""
    ppr = p.find("a:pPr", NS)
    if ppr is None:
        ppr = p.makeelement(qn("a:pPr"), {})
        p.insert(0, ppr)            # <a:pPr> must be the first child
    ppr.set("marL", str(marL_emu))
    ppr.set("indent", str(-marL_emu))


def make_bullet_para(txBody, rpr, ppr_base, marL_emu, text):
    """Build one <a:p> bullet paragraph with a hanging indent."""
    p = txBody.makeelement(qn("a:p"), {})
    ppr = copy.deepcopy(ppr_base) if ppr_base is not None else p.makeelement(qn("a:pPr"), {})
    ppr.set("marL", str(marL_emu))
    ppr.set("indent", str(-marL_emu))
    p.append(ppr)
    p.append(_make_run(p, rpr, text))
    if rpr is not None:
        epr = copy.deepcopy(rpr)
        epr.tag = qn("a:endParaRPr")
        p.append(epr)
    return p


def update_animation_range(slide, spid, start, count):
    """Widen every <p:pRg> that targets shape `spid` to cover `count` paragraphs
    starting at `start` -- keeps the SAME fade effect/timing/delay, just over the
    now-split bullet paragraphs. Returns 'old -> new' for reporting."""
    changed = None
    for spt in slide._element.findall(".//p:spTgt", NS):
        if spt.get("spid") != str(spid):
            continue
        for prg in spt.findall(".//p:pRg", NS):
            old = f"{prg.get('st')}-{prg.get('end')}"
            prg.set("st", str(start))
            prg.set("end", str(start + count - 1))
            changed = f"{old} -> {start}-{start + count - 1}"
    return changed or "unchanged"


def fill_content_bullets(slide, bullets, slide_h_in, top_limit_in, marL_emu):
    """Fill AI_content with `bullets` (one hanging-indent paragraph each), fitting
    the box height and preserving any animation. Grouped fades widen to cover all
    bullet paragraphs; staggered per-bullet fades are untouched.

    Returns (final_bullets, fit_note, anim_note).
    """
    shape = find_shape(slide, "AI_content")
    spid = shape.shape_id
    animated = animated_para_indices(slide, spid)
    txBody = shape.text_frame._txBody

    bullets = list(bullets)
    staggered = len(animated) > 1
    if staggered:                   # one bullet per animated paragraph
        bullets = bullets[:len(animated)]
    hang_in = marL_emu / EMU

    # --- fit check (grow height; drop bullets only for grouped fades) ---------
    ox, oy, cx, cy = geom_emu(shape)
    ins = get_ins(shape)
    usable = cx / EMU - ins["l"] - ins["r"]
    lh = LINE_FACTOR * CONTENT_PT / 72.0

    def needed_in(bl):
        return sum(count_display_lines(bul(b), usable, hang_in) for b in bl) * lh \
            + ins["t"] + ins["b"]

    fit_note = "fits"
    if needed_in(bullets) > cy / EMU + 1e-6:
        center = oy / EMU + cy / EMU / 2.0
        top_lim = max(MIN_TOP_IN, top_limit_in)
        bot_lim = slide_h_in - MIN_TOP_IN
        new_h = min(needed_in(bullets), bot_lim - top_lim)
        new_oy = max(top_lim, min(center - new_h / 2.0, bot_lim - new_h))
        set_geom(shape, ox, new_oy * EMU, cx, new_h * EMU)
        cy = int(new_h * EMU)
        fit_note = f"grew h->{new_h:.2f}in"
        while (not staggered) and needed_in(bullets) > cy / EMU + 1e-6 and len(bullets) > 1:
            bullets.pop()
            fit_note += f", dropped to {len(bullets)}"

    # --- write bullets (one paragraph each) + hanging indent ------------------
    paras = txBody.findall("a:p", NS)
    anim_note = "unchanged (staggered)"
    if staggered:                                   # fill existing animated paras
        for i, slot in enumerate(animated):
            p = paras[slot]
            rpr = _para_rpr(p, txBody)
            _set_pPr_hanging(p, marL_emu)
            _rebuild_para(p, [bul(bullets[i])], rpr)
    else:                                           # rebuild box with N bullets
        # Take run/paragraph formatting from the first run-bearing paragraph, then
        # DROP every existing paragraph (including any template placeholder bullets)
        # so only our bullets remain.
        tmpl = next((p for p in paras if p.find("a:r", NS) is not None),
                    paras[0] if paras else None)
        if tmpl is not None:
            rpr = _para_rpr(tmpl, txBody)
            ppr_base = tmpl.find("a:pPr", NS)
            ppr_base = copy.deepcopy(ppr_base) if ppr_base is not None else None
        else:
            rpr, ppr_base = _template_rpr(txBody), None
        for p in paras:
            txBody.remove(p)
        for b in bullets:
            txBody.append(make_bullet_para(txBody, rpr, ppr_base, marL_emu, bul(b)))
        if not bullets:                             # keep a valid, empty txBody
            txBody.append(txBody.makeelement(qn("a:p"), {}))
        if animated:
            anim_note = update_animation_range(slide, spid, 0, len(bullets))

    return bullets, fit_note, anim_note


# ---------------------------------------------------------------------------
# Titles: widen to one line, avoid content overlap (font size never changes)
# ---------------------------------------------------------------------------
def fit_title_width(slide, typeface, pt, slide_w_in):
    """Widen the AI_title box horizontally so its (single-line) heading fits
    without wrapping to a second row. Grows to the RIGHT, shifting left only if
    the box would run off the slide; never shrinks the box or the font. Height
    and vertical position are left for fix_title_geometry.

    Returns {'widened_in', 'box_w_in', 'need_in', 'fit'} or None if no AI_title.
    """
    title = find_shape(slide, "AI_title")
    if title is None:
        return None
    text = title.text_frame.text.replace("\x0b", " ").strip()
    ins = get_ins(title)
    ox, oy, cx, cy = geom_emu(title)
    need_usable = text_width_in(text, typeface, pt) + TITLE_WIDTH_SLACK_IN
    cur_usable = cx / EMU - ins["l"] - ins["r"]
    if need_usable <= cur_usable:
        return {"widened_in": 0.0, "box_w_in": cx / EMU,
                "need_in": need_usable, "fit": True}

    margin = MIN_TOP_IN
    max_w = slide_w_in - 2 * margin
    new_cx_in = min(need_usable + ins["l"] + ins["r"], max_w)
    ox_in = ox / EMU
    if ox_in + new_cx_in > slide_w_in - margin:          # would overflow right
        ox_in = max(margin, slide_w_in - margin - new_cx_in)
    set_geom(title, ox_in * EMU, oy, new_cx_in * EMU, cy)
    fit = (new_cx_in - ins["l"] - ins["r"]) + 1e-6 >= need_usable
    return {"widened_in": new_cx_in - cx / EMU, "box_w_in": new_cx_in,
            "need_in": need_usable, "fit": fit}


def fix_title_geometry(slide, typeface, pt):
    """Set the AI_title box to a one-line height and nudge it up if it would
    overlap the AI_content box. Width and left position are never changed."""
    title = find_shape(slide, "AI_title")
    content = find_shape(slide, "AI_content")
    if title is None or content is None:
        return None
    ins = get_ins(title)
    one_line = LINE_FACTOR * pt / 72.0 + ins["t"] + ins["b"]      # inches
    tox, toy, tcx, tcy = geom_emu(title)
    _, coy, _, _ = geom_emu(content)
    new_tcy = one_line * EMU
    new_toy = toy
    title_bottom = toy + new_tcy
    content_top = coy
    moved = 0.0
    if title_bottom > content_top - GAP_IN * EMU:                 # overlap -> up
        new_toy = content_top - GAP_IN * EMU - new_tcy
        new_toy = max(new_toy, MIN_TOP_IN * EMU)
        moved = (toy - new_toy) / EMU
    set_geom(title, tox, new_toy, tcx, new_tcy)
    return {"one_line_in": one_line, "moved_up_in": moved,
            "gap_in": (coy - (new_toy + new_tcy)) / EMU}


# ---------------------------------------------------------------------------
# Unique names (prevent morphing of the text boxes) -- used by the dummy-fill
# test path (run_build_tests); the AI pipeline leaves names shared on purpose
# so Bar/Hex slides morph into each other.
# ---------------------------------------------------------------------------
def rename_unique(prs):
    n = 0
    for sidx, slide in enumerate(prs.slides, start=1):
        for sh in iter_all_shapes(slide.shapes):
            if sh.name == "AI_content":
                sh.name = f"AI_content_{sidx}"; n += 1
            elif sh.name == "AI_title":
                sh.name = f"AI_title_{sidx}"; n += 1
            # AI_tag_*, AI_subtitle, USER_info, decorations: left as-is
    return n


def delete_slides(prs, indices_to_remove, verbose=True):
    sldIdLst = prs.slides._sldIdLst
    sldIds = list(sldIdLst)
    for idx in sorted(indices_to_remove, reverse=True):
        if not 0 <= idx < len(sldIds):
            print(f"  delete_slides: index {idx} out of range, skipping")
            continue
        rId = sldIds[idx].get(qn("r:id"))
        prs.part.drop_rel(rId)
        sldIdLst.remove(sldIds[idx])
        if verbose:
            print(f"  deleted slide at index {idx} (rId={rId})")
    return prs


# ---------------------------------------------------------------------------
# Spec loading
# ---------------------------------------------------------------------------
def load_spec(template_name):
    """Load templates/<template_name>/spec.json. Adds '_dir' and
    '_template_path' (the resolved template.pptx path) for convenience."""
    folder = os.path.dirname(os.path.abspath(__file__))
    spec_dir = os.path.join(folder, "templates", template_name)
    spec_path = os.path.join(spec_dir, "spec.json")
    if not os.path.exists(spec_path):
        raise FileNotFoundError(f"No spec.json for template {template_name!r} "
                                 f"(expected {spec_path})")
    with open(spec_path, encoding="utf-8") as f:
        spec = json.load(f)
    spec["_dir"] = spec_dir
    spec["_template_path"] = os.path.join(spec_dir, spec["file"])
    if not os.path.exists(spec["_template_path"]):
        raise FileNotFoundError(f"Template file not found: {spec['_template_path']}")
    return spec


# ---------------------------------------------------------------------------
# Slide-count selection rules (Start + Wheel + Bars + Hexagon + End)
# ---------------------------------------------------------------------------
def _extra_block_size(remaining, min_size, max_size):
    """Size of the next extra-section block (looping phase beyond total_slides)."""
    if remaining >= min_size + max_size:
        return max_size
    if remaining >= min_size:
        return min_size
    return remaining          # < min_size: take what's left


def _fill_extra(extra, n_bars, n_hex, block_size_range):
    """Extra slides beyond spec['total_slides'], inserted between the last
    Hexagon slide and End. Alternating Bars/Hexagon blocks starting with Bars;
    bars and hexagons each advance through their own continuous counter,
    looping back through the template's Bar/Hexagon slides."""
    min_size, max_size = block_size_range
    labels = []
    remaining, want, b, h = extra, "B", 0, 0
    while remaining > 0:
        size = _extra_block_size(remaining, min_size, max_size)
        for _ in range(size):
            if want == "B":
                labels.append(f"B{b % n_bars + 1}"); b += 1
            else:
                labels.append(f"H{h % n_hex + 1}"); h += 1
        remaining -= size
        want = "H" if want == "B" else "B"
    return labels


def build_slide_sequence(count, spec):
    """Ordered variant sequence for a `count`-slide deck, driven entirely by
    `spec` (sections, slide_rules, loop_rules). Returns (kind, template_index,
    label) tuples. Raises ValueError if `count` is out of the spec's range."""
    min_slides = spec.get("min_slides", 6)
    max_slides = spec.get("max_slides", 30)
    if count < min_slides:
        raise ValueError(f"Minimum {min_slides} slides required.")
    if count > max_slides:
        raise ValueError(f"Maximum {max_slides} slides allowed.")

    sections = spec["sections"]
    start_idx = sections["start"]["slide"]
    end_idx = sections["end"]["slide"]
    wheel_idx, n_wheel = sections["wheel"]["slides"], sections["wheel"]["count"]
    bar_idx, n_bars = sections["bars"]["slides"], sections["bars"]["count"]
    hex_idx, n_hex = sections["hexagon"]["slides"], sections["hexagon"]["count"]
    total_slides = spec["total_slides"]

    labels = ["S"] + [f"W{i + 1}" for i in range(n_wheel)]
    if count <= total_slides:
        rule = spec["slide_rules"].get(str(count))
        if rule is None:
            raise ValueError(f"No slide_rules entry for count={count}")
        labels += [f"B{i + 1}" for i in range(rule["bars"])]
        labels += [f"H{i + 1}" for i in range(rule["hex"])]
    else:                                            # beyond total_slides: loop extras
        labels += [f"B{i + 1}" for i in range(n_bars)]
        labels += [f"H{i + 1}" for i in range(n_hex)]
        labels += _fill_extra(count - total_slides, n_bars, n_hex,
                               spec["loop_rules"]["block_size_range"])
    labels += ["E"]

    idx = {"S": start_idx, "E": end_idx}
    for i in range(n_wheel):
        idx[f"W{i + 1}"] = wheel_idx[i]
    for i in range(n_bars):
        idx[f"B{i + 1}"] = bar_idx[i]
    for i in range(n_hex):
        idx[f"H{i + 1}"] = hex_idx[i]
    kind = {"S": "Start", "W": "Wheel", "B": "Bars", "H": "Hexagon", "E": "End"}

    seq = [(kind[lab[0]], idx[lab], lab) for lab in labels]
    assert len(seq) == count, f"sequence length {len(seq)} != {count}"
    return seq


def describe_sequence(count, spec):
    """Print the full labeled variant sequence, e.g. '20 slides: S W1 ... E'."""
    seq = build_slide_sequence(count, spec)
    print(f"{count} slides: {' '.join(t[2] for t in seq)}")
    return seq


# ---------------------------------------------------------------------------
# Slide cloning + deck assembly (loops bars/hex by duplicating template slides)
# ---------------------------------------------------------------------------
def clone_slide(prs, src_slide):
    """Append a full copy of `src_slide` to `prs`. Shares media parts, remaps
    relationship ids in the copied XML, drops the per-slide notes rel, and keeps
    the morph transition. Returns the new SlidePart."""
    package = prs.part.package
    src = src_slide.part
    partname = package.next_partname("/ppt/slides/slide%d.xml")
    new_el = copy.deepcopy(src._element)
    new_part = SlidePart(partname, src.content_type, package, new_el)

    rid_map = {}
    for rId, rel in src.rels.items():
        if rel.reltype.endswith("notesSlide"):
            continue                            # notes belong to one slide
        if rel.is_external:
            new_rId = new_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
        else:
            new_rId = new_part.relate_to(rel.target_part, rel.reltype)
        rid_map[rId] = new_rId
    for el in new_el.iter():                    # point copied r:embed/r:id at new rIds
        for attr, val in list(el.attrib.items()):
            if attr.startswith(R_NS) and val in rid_map:
                el.set(attr, rid_map[val])

    rId = prs.part.relate_to(new_part, RT.SLIDE)
    prs.slides._sldIdLst.add_sldId(rId)
    return new_part


def open_template(src_path):
    """Open the template, hydrating a OneDrive cloud-only placeholder if needed.

    A dehydrated OneDrive file is a reparse point that python-pptx can't open
    (PackageNotFoundError) and that Python's own read can't copy (PermissionError).
    The Windows shell copy engine (Copy-Item) does trigger the on-demand download,
    so fall back to that and open the resulting local copy.
    """
    try:
        return Presentation(src_path)
    except (PackageNotFoundError, PermissionError):
        import subprocess, tempfile
        local = os.path.join(tempfile.gettempdir(), "slidemorph_template_local.pptx")
        subprocess.run(
            ["powershell", "-NoProfile", "-Command",
             f"Copy-Item -LiteralPath \"{src_path}\" -Destination \"{local}\" -Force"],
            check=True, capture_output=True,
        )
        print(f"  [note] template was cloud-only; hydrated to local copy: {local}")
        return Presentation(local)


def build_deck(count, src_path, out_path, spec):
    """Build a `count`-slide deck from the template per build_slide_sequence and
    save it to out_path. Returns the resulting slide count."""
    seq = build_slide_sequence(count, spec)
    prs = open_template(src_path)
    n_orig = len(prs.slides._sldIdLst)
    for _variant, tmpl_idx, _label in seq:      # clone in final order
        clone_slide(prs, prs.slides[tmpl_idx - 1])
    delete_slides(prs, list(range(n_orig)), verbose=False)  # drop the originals
    prs.save(out_path)
    return len(prs.slides._sldIdLst)


def compress_images(prs, max_dim_px=1600, jpeg_quality=80):
    """Placeholder -- not implemented yet."""
    return prs


# ---------------------------------------------------------------------------
# Reporting
# ---------------------------------------------------------------------------
def print_inserted(prs):
    print("\n" + "=" * 62)
    print("FULL TEXT INSERTED (by slide / shape)")
    print("=" * 62)
    for sidx, slide in enumerate(prs.slides, start=1):
        for sh in slide.shapes:
            nm = sh.name
            if not nm.startswith("AI_"):
                continue
            if not sh.has_text_frame:
                continue
            raw = sh.text_frame.text.replace("\x0b", "\n")
            lines = [ln for ln in raw.split("\n") if ln.strip()]
            print(f"\n[Slide {sidx}] {nm}")
            for ln in lines:
                print(f"    {ln}")
    ui = find_shape(prs.slides[0], "USER_info")
    ui_txt = ui.text_frame.text.replace("\n", " | ") if ui else "(missing)"
    print(f"\n[Slide 1] USER_info  ->  UNTOUCHED: {ui_txt!r}")


def cleanup_numbered_outputs(folder):
    """Delete previously generated output_<number>.pptx files (leaves the
    output_test_*.pptx / output_<template>.pptx deliverables alone)."""
    removed = []
    for f in os.listdir(folder):
        stem = f[len("output_"):-len(".pptx")] if (
            f.startswith("output_") and f.endswith(".pptx")) else ""
        if stem.isdigit():
            os.remove(os.path.join(folder, f))
            removed.append(f)
    return sorted(removed)


def run_build_tests(template_name="neon"):
    """Print compositions/sequences and build output_[count].pptx for a set of
    counts; confirm out-of-range counts are rejected."""
    folder = os.path.dirname(os.path.abspath(__file__))
    spec = load_spec(template_name)
    src = spec["_template_path"]
    max_slides = spec.get("max_slides", 30)
    min_slides = spec.get("min_slides", 6)
    valid = [6, 9, 13, 17, 20, 23, 25, 27, 30]
    valid = [c for c in valid if min_slides <= c <= max_slides]
    reject = [min_slides - 1, max_slides + 1]

    print("=" * 64)
    print("FULL SEQUENCES")
    print("=" * 64)
    for c in valid:
        describe_sequence(c, spec)
        print()

    print("=" * 64)
    print("RANGE VALIDATION")
    print("=" * 64)
    for c in reject:
        try:
            build_slide_sequence(c, spec)
            print(f"  {c}: ERROR - was NOT rejected!")
        except ValueError as e:
            print(f"  {c}: rejected -> \"{e}\"")

    print("\n" + "=" * 64)
    print("BUILDING output_[count].pptx")
    print("=" * 64)
    removed = cleanup_numbered_outputs(folder)
    print(f"Cleaned {len(removed)} old file(s): {', '.join(removed) or '(none)'}")
    print(f"Source: {src}")
    for c in valid:
        out = os.path.join(folder, f"output_{c}.pptx")
        try:
            n = build_deck(c, src, out, spec)
            status = "OK" if n == c else f"MISMATCH (got {n})"
            print(f"  output_{c}.pptx  -> {n} slides  [{status}]")
        except Exception as e:  # noqa: BLE001 - report and continue
            print(f"  output_{c}.pptx  -> ERROR: {type(e).__name__}: {e}")


# ---------------------------------------------------------------------------
# Gemini AI content generation
# ---------------------------------------------------------------------------
STAGE1_PROMPT = """You are a presentation content writer. Given a topic and a slide count, create an outline.

Topic: {TOPIC}
Total content slides: {content_slide_count} (this is TARGET_SLIDE_COUNT minus 2, since Start and End are not content)

The content slides are divided into sections by PRESENTATION ROLE (this is about structure, not about the subject):
- WHEEL slides: introduce and define the main concepts (shorter boxes).
- BAR slides: the main body -- detailed explanation, one subtopic per slide (bigger boxes).
- HEXAGON slides: applications, examples, or summary/conclusions.

For this presentation there are:
- 3 Wheel slides
- {bar_count} Bar slides
- {hex_count} Hexagon slides

Do NOT assume anything about the subject's content type. ANY subject may need formulas,
equations, code, calculations, or quotes on a given slide -- business has ROI and
break-even formulas, biology has growth equations, economics has calculations,
programming has code, etc. -- and any subject may also be purely descriptive on a given
slide. That decision is made per point in the next stage, from the actual content, not
from the subject. Here, just plan the topics.

For each section, set "needs_long_content": true ONLY if that slide's point genuinely
needs a large block (a multi-line code block, a multi-line formula derivation, a long
paragraph, or a long quote). Otherwise set it false. This is optional and rare.

Return ONLY a JSON object, no markdown, no explanation:
{
  "title": "short title, max 20 chars",
  "subtitle": "subtitle, max 30 chars",
  "wheel_tags": ["tag1", "tag2", "tag3"],
  "sections": [
    {"type": "wheel", "variant": 1, "heading": "max 14 chars", "topic": "what this slide covers", "needs_long_content": false},
    {"type": "bar", "variant": 1, "heading": "max 25 chars", "topic": "what this slide covers", "needs_long_content": false},
    {"type": "hexagon", "variant": 1, "heading": "max 25 chars", "topic": "what this slide covers", "needs_long_content": false}
  ]
}

Rules:
- title max 20 characters
- subtitle max 30 characters
- bar and hexagon headings max 25 characters (must fit on ONE line)
- wheel_tags are the 3 labels on the wheel, max 18 chars each
- sections must have exactly {content_slide_count} entries
- Keep all headings short enough to never wrap to a second line"""

STAGE2_PROMPT = """Write the on-slide content for ONE presentation slide.

Topic of this slide: {section.topic}
Heading: {section.heading}
Slide role: {role_note}

For each point, choose the clearest way to present it. Use bullets for descriptive
information. If a specific point is best explained with a formula, equation, code
snippet, calculation, or short example, include that directly instead of just describing
it. This applies to ANY subject -- business, biology, economics, programming, chemistry,
etc. Do not assume a subject has or lacks technical content; decide from the actual point
being made.

The box is the ONLY hard constraint:
- Each line must fit on ONE line in the box: at most about {max_chars} characters.
- The box holds about {capacity} lines. FILL {target_min}-{target_max} of them
  (70-85% of the box). Do not leave the box nearly empty, and do not overflow it.
- A formula, code line, or calculation counts toward the line budget exactly like a
  bullet does (one line each, unless it is short enough to share).
- Never drop a major point. If space is tight, SHORTEN the wording of a point rather
  than removing the point.

Style:
- Descriptive lines: clean concrete phrases, start with a capital, no trailing period.
- Formulas / equations / code / calculations: write them in their natural form --
  symbols, operators, and numbers are expected (e.g. "ROI = (Gain - Cost) / Cost * 100").

Return ONLY a JSON object, no markdown, no explanation:
{
  "lines": ["line 1", "line 2", "line 3"]
}"""


def _read_env_value(key, env_path):
    """Read a single KEY=value from a .env file (value is never printed)."""
    if not os.path.exists(env_path):
        return None
    for line in open(env_path, encoding="utf-8"):
        line = line.strip()
        if line.startswith("#") or "=" not in line:
            continue
        k, v = line.split("=", 1)
        if k.strip() == key:
            return v.strip().strip('"').strip("'")
    return None


def _is_quota_error(exc):
    """True if `exc` looks like a Gemini quota / rate-limit (429) failure."""
    s = f"{type(exc).__name__}: {exc}".lower()
    return ("resourceexhausted" in s or "429" in s or "quota" in s
            or "exhausted" in s or "rate limit" in s)


def _is_unavailable_error(exc):
    """True if `exc` means the model itself is gone/unusable (404, retired,
    "no longer available", "not found") -- a reason to skip to the next model."""
    s = f"{type(exc).__name__}: {exc}".lower()
    return ("notfound" in s or "404" in s or "not found" in s
            or "no longer available" in s or "not supported" in s)


def _retry_seconds(exc):
    """Seconds the API suggests waiting before retrying (from a 429), or None."""
    s = f"{exc}"
    m = re.search(r"retry in ([\d.]+)s", s) or re.search(r"seconds:\s*(\d+)", s)
    return float(m.group(1)) if m else None


class _FallbackModel:
    """Wraps a chain of Gemini models. On a quota/429 error it switches to the
    next model in the chain (for this call and every later one) and retries,
    so a mid-run quota exhaustion never loses progress. Any non-quota error, or
    running out of fallbacks, propagates normally."""

    def __init__(self, genai, model_names):
        self._genai = genai
        self._names = list(model_names)
        self._i = 0
        self._model = genai.GenerativeModel(self._names[0])

    @property
    def name(self):
        return self._names[self._i]

    def generate_content(self, prompt):
        waits = 0
        while True:
            try:
                return self._model.generate_content(prompt)
            except Exception as exc:  # noqa: BLE001 - inspected, then re-raised
                skip = _is_quota_error(exc) or _is_unavailable_error(exc)
                if skip and self._i + 1 < len(self._names):     # switch models
                    old = self._names[self._i]
                    reason = "quota" if _is_quota_error(exc) else "unavailable"
                    self._i += 1
                    print(f"    [{reason}] {old} -> falling back to "
                          f"{self._names[self._i]}")
                    self._model = self._genai.GenerativeModel(self._names[self._i])
                    time.sleep(API_CALL_DELAY)
                    waits = 0
                    continue
                if _is_quota_error(exc) and waits < MAX_QUOTA_WAITS:  # last model
                    delay = min(_retry_seconds(exc) or QUOTA_RETRY_CAP,
                                QUOTA_RETRY_CAP)
                    waits += 1
                    print(f"    [quota] {self.name} rate-limited; waiting "
                          f"{delay:.0f}s then retrying ({waits}/{MAX_QUOTA_WAITS})")
                    time.sleep(delay)
                    continue
                raise


def load_gemini(folder):
    """Configure Gemini from GEMINI_API_KEY (env var first, then .env);
    return a model handle."""
    import google.generativeai as genai  # lazy: keeps the deprecation warning out of other runs
    key = os.getenv("GEMINI_API_KEY") or _read_env_value(
        "GEMINI_API_KEY", os.path.join(folder, ".env"))
    if not key:
        raise RuntimeError("GEMINI_API_KEY not found in env or .env")
    genai.configure(api_key=key)
    print(f"Gemini ready (models={GEMINI_MODELS}, key loaded: {len(key)} chars)")
    return _FallbackModel(genai, GEMINI_MODELS)


def _resp_text(resp):
    """Extract text from a Gemini response, tolerating blocked/empty results."""
    try:
        if resp.text:
            return resp.text
    except Exception:
        pass
    parts = []
    for c in getattr(resp, "candidates", None) or []:
        for p in getattr(getattr(c, "content", None), "parts", None) or []:
            parts.append(getattr(p, "text", "") or "")
    return "".join(parts)


def _strip_fences(text):
    """Drop ```json ... ``` markdown fences Gemini sometimes adds."""
    t = text.strip()
    if t.startswith("```"):
        nl = t.find("\n")
        t = t[nl + 1:] if nl != -1 else t[3:]
        if "```" in t:
            t = t[:t.rfind("```")]
    return t.strip()


def gemini_json(model, prompt, retries=1):
    """Call Gemini, strip fences, parse JSON. Retry once on invalid JSON."""
    last = None
    for attempt in range(retries + 1):
        raw = _resp_text(model.generate_content(prompt))
        try:
            return json.loads(_strip_fences(raw))
        except (json.JSONDecodeError, ValueError) as e:
            last = e
            if attempt < retries:
                print("    [warn] invalid JSON from Gemini, retrying once...")
                time.sleep(API_CALL_DELAY)
    raise ValueError(f"invalid JSON after retry: {last}")


def stage2_bullets(model, section, max_chars, capacity, target_min, target_max, rn):
    """Ask Gemini for the slide's content lines. Each returned line is whatever
    best communicates that point -- a bullet, a formula, a code line, or a short
    calculation -- constrained only by the box (max_chars per line, ~capacity
    lines, fill target_min..target_max). Returns a list of raw line strings."""
    prompt = (STAGE2_PROMPT
              .replace("{section.topic}", str(section.get("topic", "")))
              .replace("{section.heading}", str(section.get("heading", "")))
              .replace("{role_note}", str(rn))
              .replace("{max_chars}", str(max_chars))
              .replace("{capacity}", str(capacity))
              .replace("{target_min}", str(target_min))
              .replace("{target_max}", str(target_max)))
    data = gemini_json(model, prompt)
    if not isinstance(data, dict):
        return []
    return data.get("lines") or data.get("bullets") or []


def truncate(text, n, ellipsis=False):
    """Clip `text` to `n` chars. With ellipsis=True an over-limit string is cut to
    n-3 chars and gets a trailing '...' so the RESULT is still <= n chars."""
    text = (text or "").strip()
    if len(text) <= n:
        return text
    if ellipsis and n > 3:
        return text[:n - 3].rstrip() + "..."
    return text[:n].rstrip()


# ---------------------------------------------------------------------------
# Adaptive content generation, driven ONLY by box size (Pillow-measured).
#
# No subject assumptions and no per-type "bullets-only" rule: each returned line
# is whatever best communicates its point (bullet, formula, code, calculation),
# and the sole hard constraint is the box -- ~max_chars per line and a line
# capacity computed from the box height. We aim to FILL 70-85% of the box.
# ---------------------------------------------------------------------------
CONTENT_LINE_H_IN = LINE_FACTOR * CONTENT_PT / 72.0   # one content line's height
FILL_TARGET_LO = 0.70      # aim to fill at least this fraction of the box
FILL_TARGET_HI = 0.85      # ... and at most this fraction
FILL_UNDERFULL = 0.60      # below this -> regenerate asking for more detail

_CHARS_SAMPLE = ("The quick brown fox jumps over the lazy dog while counting "
                 "1234567890 items and notes")


def count_wraps(bullets, usable_in, hang_in):
    """How many `bullets` are too wide to fit on ONE line in the content box
    (measured with Pillow via count_display_lines, including the '- ' hang)."""
    return sum(1 for b in bullets
               if count_display_lines(bul(b), usable_in, hang_in) > 1)


def total_display_lines(items, usable_in, hang_in):
    """Total rendered line count for `items` (each may wrap), measured w/ Pillow."""
    return sum(count_display_lines(bul(x), usable_in, hang_in) for x in items)


def approx_chars_per_line(usable_in):
    """Roughly how many content-font characters fit on one usable-width line.
    Measured with Pillow against a mixed-case sample (real gate stays the
    per-line wrap measurement; this is just guidance passed to the model)."""
    avg = text_width_in(_CHARS_SAMPLE, "+mj-lt", CONTENT_PT) / len(_CHARS_SAMPLE)
    return max(10, int(usable_in / avg))


def box_line_capacity(cy_emu, ins):
    """How many content lines fit in a box of height `cy_emu` (EMU) given insets."""
    usable_h = cy_emu / EMU - ins["t"] - ins["b"]
    return max(1, int(usable_h / CONTENT_LINE_H_IN))


def role_note(stype):
    """Structural (NOT subject) role blurb passed to the Stage 2 prompt."""
    return {
        "wheel":   "Wheel slide -- introduce/define a core concept (shorter box).",
        "bar":     "Bar slide -- main body; explain this subtopic in detail.",
        "bars":    "Bar slide -- main body; explain this subtopic in detail.",
        "hexagon": "Hexagon slide -- applications, examples, or summary.",
    }.get((stype or "").lower(), "Content slide -- explain this point clearly.")


def _looks_technical(line):
    """Heuristic marker for the printout: does this line read as a formula,
    equation, calculation, or code (rather than a descriptive phrase)? Used only
    for reporting so a human can verify technical content appeared where needed."""
    s = line.strip()
    has_math = bool(re.search(r"[=×÷≈≤≥∑∏√^]", s)) or \
        bool(re.search(r"\d\s*[-+*/×÷]\s*\d", s))
    has_code = bool(re.search(r"[A-Za-z_]\w*\([^)]*\)", s)) or "()" in s or \
        s.endswith(":") or bool(re.search(r"[{}<>]|::|=>|->", s))
    return has_math or has_code


def adaptive_bullets(model, section, usable_in, hang_in, ph_content_spec, capacity):
    """Generate the slide's content lines, filling 70-85% of the box.

    Requests content sized to the box (max_chars/line + fill target derived from
    `capacity`), measures the rendered line total with Pillow, and regenerates
    once or twice if the box is under-full (<60%) or overflowing (>100%). Points
    are never dropped here -- the model is asked to shorten instead. Returns
    (lines, total_lines, log). `ph_content_spec` is accepted for signature
    compatibility but the box now drives everything.
    """
    max_chars = approx_chars_per_line(usable_in)
    target_min = max(1, round(FILL_TARGET_LO * capacity))
    target_max = max(target_min, int(FILL_TARGET_HI * capacity))
    low_water = max(1, int(FILL_UNDERFULL * capacity))
    rn = role_note(section.get("type"))
    log = []
    last = []
    for attempt in range(MAX_BULLET_ITERS):
        try:
            raw = stage2_bullets(model, section, max_chars, capacity,
                                 target_min, target_max, rn)
        except Exception as exc:  # noqa: BLE001 - reported, then fall back
            log.append(f"try{attempt + 1}: gen FAILED: {exc}")
            time.sleep(API_CALL_DELAY)
            break
        time.sleep(API_CALL_DELAY)
        items = [x.strip() for x in raw if isinstance(x, str) and x.strip()]
        total = total_display_lines(items, usable_in, hang_in)
        pct = (100.0 * total / capacity) if capacity else 0.0
        log.append(f"try{attempt + 1}: {len(items)} items -> {total} lines "
                   f"(cap {capacity}, {pct:.0f}% full; target {target_min}-{target_max})")
        last = items
        if total > capacity:                    # overflow -> ask to shorten/fewer
            target_max = max(target_min, capacity - 1)
            continue
        if total < low_water:                   # under-full -> ask for more detail
            target_min = min(capacity, target_max + 1)
            target_max = capacity
            continue
        return items, total, log                # 60-100% full -> accept

    # did not converge: keep the last set. fill_content_bullets grows the box to
    # fit, and only drops as an absolute last resort.
    return last, total_display_lines(last, usable_in, hang_in), log


def _fill_shape_text(slide, name, lines):
    sh = find_shape(slide, name)
    if sh is not None and sh.has_text_frame:
        fill_simple(sh, lines)


# ---------------------------------------------------------------------------
# Orchestration: generate_content (Gemini Stage 1) / fill_placeholders
# (Gemini Stage 2 + write) / save_output
# ---------------------------------------------------------------------------
def generate_content(topic, spec, sequence, model):
    """Gemini Stage 1: one outline call for the whole deck -- title, subtitle,
    wheel tags, and a heading+topic for every Wheel/Bar/Hexagon slide. Returns
    a content dict consumed by fill_placeholders."""
    bar_count = sum(1 for t in sequence if t[0] == "Bars")
    hex_count = sum(1 for t in sequence if t[0] == "Hexagon")
    wheel_count = sum(1 for t in sequence if t[0] == "Wheel")
    content_slide_count = len(sequence) - 2   # minus Start + End

    ph_start = spec["placeholders"]["start"]
    max_title = ph_start["AI_title"]["max_chars"]
    max_subtitle = ph_start["AI_subtitle"]["max_chars"]
    max_tag = spec["placeholders"]["wheel"]["AI_tag_1"]["max_chars"]

    print(f"Topic: {topic!r} | {len(sequence)} slides = "
          f"1 Start + {wheel_count} Wheel + {bar_count} Bar + "
          f"{hex_count} Hexagon + 1 End")

    print("\n=== STAGE 1: outline ===")
    p1 = (STAGE1_PROMPT.replace("{TOPIC}", topic)
          .replace("{content_slide_count}", str(content_slide_count))
          .replace("{bar_count}", str(bar_count))
          .replace("{hex_count}", str(hex_count)))
    outline = gemini_json(model, p1)
    time.sleep(API_CALL_DELAY)

    # One-line headings: truncate over-limit text with a trailing "..." .
    title = truncate(outline.get("title", ""), max_title, ellipsis=True)
    subtitle = truncate(outline.get("subtitle", ""), max_subtitle, ellipsis=True)
    wheel_tags = [truncate(t, max_tag) for t in
                  (outline.get("wheel_tags") or [])][:wheel_count]
    sections = outline.get("sections") or []
    print(f"  title     : {title!r}  ({len(title)} chars, <= {max_title})")
    print(f"  subtitle  : {subtitle!r}  ({len(subtitle)} chars, <= {max_subtitle})")
    print(f"  wheel_tags: {wheel_tags}")
    print(f"  sections  : {len(sections)} (need {content_slide_count})")
    for s in sections:
        lc = "  [needs_long_content]" if s.get("needs_long_content") else ""
        print(f"     - {s.get('type'):7} heading={s.get('heading')!r} "
              f"topic={s.get('topic')!r}{lc}")

    # Queue per section kind, keyed to match spec["placeholders"] ("bar" -> "bars").
    queues = {"wheel": [], "bars": [], "hexagon": []}
    for sec in sections:
        stype = (sec.get("type") or "").lower()
        key = "bars" if stype == "bar" else stype
        queues.setdefault(key, []).append(sec)

    return {"title": title, "subtitle": subtitle, "wheel_tags": wheel_tags,
            "queues": queues}


def fill_placeholders(prs, content, spec, sequence, model):
    """Gemini Stage 2 (adaptive, per-slide bullet generation measured against
    each slide's actual content box) + write everything (title/subtitle/tags/
    bullets) into `prs`. Returns a summary list of (slide#, kind, n_bullets,
    n_wraps)."""
    slide_h_in = prs.slide_height / EMU
    slide_w_in = prs.slide_width / EMU
    marL = content_marL_emu()
    hang = marL / EMU
    queues = content["queues"]

    print("\n=== STAGE 2 + FILL (per-point format, box-only constraint, fill 70-85%) ===")
    summary = []
    long_flags = []
    for i, (variant, _idx, _label) in enumerate(sequence):
        slide = prs.slides[i]

        if variant == "Start":
            _fill_shape_text(slide, "AI_title", [content["title"]])
            _fill_shape_text(slide, "AI_subtitle", [content["subtitle"]])
            print(f"\nSlide {i + 1:2} Start")
            print(f"    title    : {content['title']!r}")
            print(f"    subtitle : {content['subtitle']!r}   (USER_info untouched)")
            continue
        if variant == "End":
            continue

        key = {"Wheel": "wheel", "Bars": "bars", "Hexagon": "hexagon"}[variant]
        ph = spec["placeholders"][key]
        default_sec = {"type": "bar" if key == "bars" else key,
                       "heading": "", "topic": ""}
        sec = queues[key].pop(0) if queues[key] else default_sec
        if sec.get("needs_long_content"):
            long_flags.append((i + 1, variant, sec.get("heading", "")))

        content_shape = find_shape(slide, "AI_content")
        _, _, cx, cy = geom_emu(content_shape)
        ins = get_ins(content_shape)
        usable = cx / EMU - ins["l"] - ins["r"]
        capacity = box_line_capacity(cy, ins)

        # box-driven content generation (per-point format, fill 70-85% of the box)
        bullets, _tot, log = adaptive_bullets(
            model, sec, usable, hang, ph["AI_content"], capacity)

        # headings / wheel tags
        if variant == "Wheel":
            for ti, tag in enumerate(content["wheel_tags"], start=1):
                _fill_shape_text(slide, f"AI_tag_{ti}", [tag])
            head_line = f"tags={content['wheel_tags']}"
        else:
            max_heading = ph["AI_title"]["max_chars"]
            heading = truncate(sec.get("heading", ""), max_heading, ellipsis=True)
            _fill_shape_text(slide, "AI_title", [heading])
            th = find_shape(slide, "AI_title")
            wnote = ""
            if th is not None:
                tf, pt = run_font(th, "Impact", TITLE_PT_DEFAULT)
                winfo = fit_title_width(slide, tf, pt, slide_w_in)   # widen to 1 line
                fix_title_geometry(slide, tf, pt)                    # 1-line height
                if winfo:
                    wnote = (f"  [title box W->{winfo['box_w_in']:.2f}in, "
                             f"+{winfo['widened_in']:.2f}in, "
                             f"one-line={'yes' if winfo['fit'] else 'NO'}]")
            head_line = (f"heading={heading!r} ({len(heading)} chars, "
                         f"<= {max_heading}){wnote}")

        # fill content; height-fit may drop lines -> recount on the FINAL set
        final_bullets, fit_note, _anim = fill_content_bullets(
            slide, bullets, slide_h_in, MIN_TOP_IN, marL)
        final_lines = total_display_lines(final_bullets, usable, hang)
        final_wraps = count_wraps(final_bullets, usable, hang)
        n_tech = sum(1 for b in final_bullets if _looks_technical(b))
        fill_pct = (100.0 * final_lines / capacity) if capacity else 0.0

        print(f"\nSlide {i + 1:2} {variant}"
              f"{'  [needs_long_content]' if sec.get('needs_long_content') else ''}")
        print(f"    {head_line}")
        for step in log:
            print(f"    - {step}")
        print(f"    -> FINAL items={len(final_bullets)} lines={final_lines} "
              f"({fill_pct:.0f}% of {capacity}-line box)  wraps={final_wraps}  "
              f"tech-lines={n_tech}  [{fit_note}]  (usable={usable:.2f}in)")
        for b in final_bullets:
            mark = "  <-- formula/code" if _looks_technical(b) else ""
            print(f"        - {b}{mark}")
        summary.append((i + 1, variant, len(final_bullets), final_lines,
                        capacity, n_tech))

    print("\n=== SUMMARY: fill % and technical-line counts ===")
    for sidx, variant, nb, nlines, cap, ntech in summary:
        pct = (100.0 * nlines / cap) if cap else 0.0
        print(f"   Slide {sidx:2} {variant:8}: items={nb} lines={nlines}/{cap} "
              f"({pct:.0f}% full)  formula/code lines={ntech}")

    if long_flags:
        print("\n=== needs_long_content flagged (NOT routed) ===")
        print("   The neon template has no long-content layout (slides 21-22 do "
              "not exist), so these were filled in the normal box instead:")
        for sidx, variant, heading in long_flags:
            print(f"   Slide {sidx:2} {variant:8}: {heading!r}")
    return summary


def save_output(prs, output_path):
    prs.save(output_path)
    print(f"\nSaved: {output_path}")
    return output_path


def generate_ai_deck(template_name, topic, count, output_path=None):
    """End-to-end: load spec -> build the cloned deck -> Gemini outline
    (Stage 1) -> per-slide adaptive bullets + write (Stage 2) -> save."""
    folder = os.path.dirname(os.path.abspath(__file__))
    spec = load_spec(template_name)
    src = spec["_template_path"]
    out = output_path or os.path.join(folder, f"output_{template_name}.pptx")

    seq = build_slide_sequence(count, spec)     # validates the count against spec

    model = load_gemini(folder)
    content = generate_content(topic, spec, seq, model)

    build_deck(count, src, out, spec)
    prs = Presentation(out)

    fill_placeholders(prs, content, spec, seq, model)

    print_inserted(prs)
    save_output(prs, out)
    return out


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def parse_args():
    p = argparse.ArgumentParser(
        description="Fill a spec-driven PowerPoint template with AI-generated content.")
    p.add_argument("--template", required=True,
                    help="Template name (folder under templates/, e.g. 'neon')")
    p.add_argument("--topic", required=True, help="Presentation topic")
    p.add_argument("--count", type=int, required=True, help="Target slide count")
    p.add_argument("--output", default=None,
                    help="Output .pptx path (default: output_<template>.pptx)")
    return p.parse_args()


if __name__ == "__main__":
    args = parse_args()
    try:
        generate_ai_deck(args.template, args.topic, args.count, args.output)
    except Exception as exc:  # noqa: BLE001 - surface the failure clearly
        import traceback
        print(f"\n[ERROR] {type(exc).__name__}: {exc}")
        traceback.print_exc()
