#!/usr/bin/env python3
"""Make the skill's assets/template.pptx: the lab's template plus sample slides for every theme.

Source: tools/template/lab-template.pptx (the lab's own template, as handed out) and the lab's
logos in tools/template/logos/. The lab's cover is kept as it is for the classic theme; every
other sample slide is redrawn here from the lab template's elements — the navy band with its
radar rings, the blue-to-gold rule, the dot grid — to the geometry in THEMES in build_deck.py.

Each sample slide is named "<theme>/<kind>" in its <p:cSld name>; build_deck.py finds them by
that name, so their order doesn't matter:

    <theme>/cover            title, date and name (also the closing slide); paper has an "Image slot"
    <theme>/section          divider: a title and one optional line
    <theme>/content          white title-only page for every slide that carries content

Content slides carry two footer text boxes the builder fills: "Footer label" (the current part of
the deck) and "Slide number" (a slide-number field).

Usage:
    python tools/make_template.py            write skills/pairs-group-meeting/assets/template.pptx
    python tools/make_template.py --art DIR  also save the background art as PNGs in DIR
Then run python tools/sync_codex.py to copy it to the Codex version.
"""

import argparse
import io
import random
import sys
from copy import deepcopy
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.util import Emu

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "tools" / "template"
OUT = ROOT / "skills" / "pairs-group-meeting" / "assets" / "template.pptx"
sys.path.insert(0, str(ROOT / "skills" / "pairs-group-meeting" / "scripts"))
import build_deck as bd  # noqa: E402  (THEMES: the geometry and colors the art is drawn to)

W, H = 3840, 2160                   # 13.333" x 7.5" at 288 px per inch
PPI = 288
RNG = np.random.default_rng(7)
EMU = 914400

LOGO = {name: SRC / "logos" / f"{name}.png" for name in ("pairs-lab-white", "pairs-lab-color",
                                                          "hkustgz-white", "hkustgz-color")}

# Colors from the lab template's art.
BAND = ("12214E", "1B326C")         # header band, left to right
RULE = ("40AAE5", "8AACA1", "D5AF5C")  # the blue-to-gold rule
GOLD = "D6B05C"
BODY = "FCFCFE"


def px(inches):
    return int(round(inches * PPI))


def emu(inches):
    return int(round(inches * EMU))


def rgb(h):
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


# ---------------------------------------------------------------------------
# Drawing


def gradient(w, h, stops, angle=0.0, dither=True):
    yy, xx = np.mgrid[0:h, 0:w].astype(np.float32)
    a = np.deg2rad(angle)
    t = xx * np.cos(a) + yy * np.sin(a)
    t = (t - t.min()) / max(1e-6, t.max() - t.min())
    out = np.zeros((h, w, 3), np.float32)
    for c in range(3):
        out[..., c] = np.interp(t, [s[0] for s in stops], [rgb(s[1])[c] for s in stops])
    if dither:
        out += RNG.uniform(-0.8, 0.8, out.shape).astype(np.float32)
    return Image.fromarray(out.clip(0, 255).astype(np.uint8), "RGB").convert("RGBA")


def layer(w, h, draw_fn, ss=2):
    """Draw at ss x resolution on a transparent layer, then downsample (anti-aliasing)."""
    big = Image.new("RGBA", (w * ss, h * ss), (0, 0, 0, 0))
    draw_fn(ImageDraw.Draw(big), ss)
    return big.resize((w, h), Image.LANCZOS)


def rings(w, h, cx, cy, radii, alphas, width, color="FFFFFF"):
    r, g, b = rgb(color)

    def fn(d, s):
        for rad, a in zip(radii, alphas):
            R = rad * s
            d.ellipse([cx * s - R, cy * s - R, cx * s + R, cy * s + R],
                      outline=(r, g, b, int(255 * a)), width=max(1, round(width * s)))
    return layer(w, h, fn)


def radial(w, h, cx, cy, radius, color, alpha0, power=2.0):
    yy, xx = np.mgrid[0:h, 0:w].astype(np.float32)
    t = np.clip(1 - np.sqrt((xx - cx) ** 2 + (yy - cy) ** 2) / radius, 0, 1) ** power
    arr = np.zeros((h, w, 4), np.uint8)
    arr[..., :3] = rgb(color)
    arr[..., 3] = (t * alpha0 * 255).astype(np.uint8)
    return Image.fromarray(arr, "RGBA")


def dots(w, h, x0, y0, x1, y1, step, radius, color, alpha_fn):
    r, g, b = rgb(color)

    def fn(d, s):
        for y in np.arange(y0, y1, step):
            for x in np.arange(x0, x1, step):
                a = alpha_fn(x, y)
                if a > 0.01:
                    R = radius * s
                    d.ellipse([x * s - R, y * s - R, x * s + R, y * s + R], fill=(r, g, b, int(255 * a)))
    return layer(w, h, fn)


def rule(w, h):
    return gradient(w, h, [(0, RULE[0]), (0.5, RULE[1]), (1, RULE[2])])


def fading_rings(w, h, cx, cy, r0, step, n, a0, a1, width, color):
    radii = [r0 + step * i for i in range(n)]
    alphas = [a0 + (a1 - a0) * i / max(1, n - 1) for i in range(n)]
    return rings(w, h, cx, cy, radii, alphas, width, color)


# ---------------------------------------------------------------------------
# Art for each sample slide (3840 x 2160)


def band_end(lab_content):
    """First row under the lab's header band and rule."""
    for y in range(lab_content.height // 3):
        if lab_content.getpixel((W // 2, y))[:3] == rgb(BODY):
            return y
    raise SystemExit("can't find the header band in the lab template's content background")


def classic_content(lab):
    """The lab's content page, minus the dot grid at the bottom right: the content area now runs
    down there (the logos moved into the band), and the dots would show through charts."""
    img = lab.copy()
    ImageDraw.Draw(img).rectangle([px(10.2), px(5.45), W, px(7.03)], fill=rgb(BODY) + (255,))
    return img


def classic_section():
    """Divider in the lab's colors: the cover's navy, its rings large on the right, its gold mark."""
    img = gradient(W, H, [(0, "101D48"), (0.6, "142659"), (1, "1B336F")], angle=12)
    cx, cy = px(11.3), px(3.9)
    img.alpha_composite(radial(W, H, cx, cy, px(6.5), "2E95EE", 0.18))
    img.alpha_composite(fading_rings(W, H, cx, cy, px(0.55), px(0.30), 30, 0.20, 0.02, 2.0, "FFFFFF"))

    def fade(x, y):
        t = 1 - np.hypot((x - px(0.35)) / px(4.2), (y - px(0.3)) / px(1.6))
        return max(0.0, t) ** 1.3 * 0.30
    img.alpha_composite(dots(W, H, px(0.35), px(0.3), px(4.6), px(1.9), px(0.075), 2.4, "C7D3E6", fade))
    ImageDraw.Draw(img).rectangle([px(0.825), px(2.72), px(0.892), px(3.67)], fill=rgb(GOLD) + (255,))
    img.alpha_composite(rule(W, px(0.045)), (0, H - px(0.045)))
    return img


def paper_rings(cx, cy, n=26, a0=0.42, a1=0.04):
    img = Image.new("RGBA", (W, H), (255, 255, 255, 255))
    img.alpha_composite(radial(W, H, cx, cy, px(7.5), "EEF1F6", 0.9, power=1.4))
    img.alpha_composite(fading_rings(W, H, cx, cy, px(0.5), px(0.30), n, a0, a1, 1.8, "9AA6BC"))
    return img


def paper_cover():
    return paper_rings(px(11.4), px(3.75))


def paper_section():
    return paper_rings(px(12.1), px(4.1), n=24, a0=0.34)


# ---------------------------------------------------------------------------
# Slides

R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
P14_CREATION_ID = "{http://schemas.microsoft.com/office/powerpoint/2010/main}creationId"


def clone_slide(prs, src):
    """A copy of `src` at the end of the deck, its relationships re-made.

    The copy's shapes go into the shape tree python-pptx made for the new slide (it caches that tree
    on the slide object), so the shapes can still be edited through `slide.shapes` afterwards.
    """
    dst = prs.slides.add_slide(src.slide_layout)
    new_rid = {}
    for rid, rel in src.part.rels.items():
        if rel.reltype in (RT.SLIDE_LAYOUT, RT.NOTES_SLIDE):
            continue
        new_rid[rid] = dst.part.relate_to(rel.target_ref if rel.is_external else rel.target_part, rel.reltype,
                                          is_external=rel.is_external)
    tree = dst.shapes._spTree
    for child in list(tree):
        tree.remove(child)
    for child in src.shapes._spTree:
        tree.append(deepcopy(child))
    bg = src._element.find(qn("p:cSld")).find(qn("p:bg"))
    if bg is not None:
        dst._element.find(qn("p:cSld")).insert(0, deepcopy(bg))
    sld = dst._element
    for el in sld.iter("*"):
        for attr, value in el.attrib.items():
            if attr.startswith(R_NS) and value in new_rid:
                el.set(attr, new_rid[value])
    for el in sld.iter(P14_CREATION_ID):
        el.set("val", str(random.randint(1, 2**31 - 1)))
    return dst


def prune_rels(slide):
    """Drop the slide's image relationships nothing on it refers to any more."""
    used = {v for el in slide._element.iter("*") for a, v in el.attrib.items() if a.startswith(R_NS)}
    for rid, rel in list(slide.part.rels.items()):
        if rel.reltype == RT.IMAGE and rid not in used:
            slide.part.drop_rel(rid)


def set_background(slide, art=None, solid=None):
    """Slide-level background: a picture (the template's way) or a solid color."""
    c_sld = slide._element.find(qn("p:cSld"))
    old = c_sld.find(qn("p:bg"))
    if old is not None:
        c_sld.remove(old)
    if art is not None:
        buf = io.BytesIO()
        art.convert("RGB").save(buf, "JPEG", quality=90, optimize=True)
        buf.seek(0)
        _, rid = slide.part.get_or_add_image_part(buf)
        fill = (f'<a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="{rid}"/><a:srcRect/>'
                '<a:stretch><a:fillRect/></a:stretch></a:blipFill>')
    else:
        fill = f'<a:solidFill><a:srgbClr val="{solid}"/></a:solidFill>'
    c_sld.insert(0, parse_xml(f'<p:bg {nsdecls("p", "a", "r")}><p:bgPr>{fill}<a:effectLst/></p:bgPr></p:bg>'))


def drop_groups(slide):
    """Remove the lab's logo groups."""
    tree = slide.shapes._spTree
    for grp in tree.findall(qn("p:grpSp")):
        tree.remove(grp)


def add_logo(slide, name, x, y, h, descr):
    pic = slide.shapes.add_picture(str(LOGO[name]), Emu(emu(x)), Emu(emu(y)), height=Emu(emu(h)))
    pic.name = descr
    pic._element.nvPicPr.cNvPr.set("descr", descr)
    return pic


def add_logo_right(slide, name, right, y, h, descr):
    with Image.open(LOGO[name]) as im:
        w = h * im.width / im.height
    return add_logo(slide, name, right - w, y, h, descr)


def run_props(size, color, font, bold=False, italic=False):
    meta = bd.FONT_META.get(font, "")
    i = ' i="1"' if italic else ""
    return (f'<a:rPr lang="en-US" altLang="zh-CN" sz="{size * 100}" b="{1 if bold else 0}"{i} dirty="0">'
            f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
            f'<a:latin typeface="{font}" {meta}/>'
            '<a:ea typeface="等线" panose="02010600030101010101" pitchFamily="2" charset="-122"/>'
            f'<a:cs typeface="{font}" {meta}/></a:rPr>')


def para(lines, size, color, font, bold=False, italic=False, algn="l", space_before=0):
    out = []
    for k, text in enumerate(lines):
        before = f'<a:spcBef><a:spcPts val="{space_before if k else 0}"/></a:spcBef>'
        rpr = run_props(size, color, font, bold, italic)
        out.append(f'<a:p {nsdecls("a")}><a:pPr marL="0" indent="0" algn="{algn}"><a:lnSpc><a:spcPct val="100000"/>'
                   f'</a:lnSpc>{before}<a:buNone/></a:pPr><a:r>{rpr}<a:t>{text}</a:t></a:r>'
                   f'{rpr.replace("<a:rPr", "<a:endParaRPr").replace("</a:rPr>", "</a:endParaRPr>")}</a:p>')
    return out


def set_text(shape, box, lines, size, color, font, bold=False, italic=False, anchor="ctr", algn="l", space_before=0):
    """Place a shape and give it the sample text, formatted on every run as the lab template does."""
    x, y, w, h = box
    shape.left, shape.top, shape.width, shape.height = (Emu(emu(v)) for v in box)
    tx = shape.text_frame._txBody
    body_pr = tx.find(qn("a:bodyPr"))
    for child in list(body_pr):
        body_pr.remove(child)
    for k, v in {"wrap": "square", "lIns": "0", "tIns": "0", "rIns": "0", "bIns": "0", "anchor": anchor}.items():
        body_pr.set(k, v)
    body_pr.append(parse_xml(f'<a:normAutofit {nsdecls("a")}/>'))
    lst = tx.find(qn("a:lstStyle"))
    if lst is None:
        body_pr.addnext(parse_xml(f'<a:lstStyle {nsdecls("a")}/>'))
    for p in tx.findall(qn("a:p")):
        tx.remove(p)
    for p in para(lines, size, color, font, bold, italic, algn, space_before):
        tx.append(parse_xml(p))


def add_box(slide, name, box, lines, size, color, font, **kw):
    shape = slide.shapes.add_textbox(Emu(0), Emu(0), Emu(1), Emu(1))
    shape.name = name
    set_text(shape, box, lines, size, color, font, **kw)
    return shape


def add_slide_number(slide, box, size, color, font, bold=False, algn="r"):
    shape = add_box(slide, "Slide number", box, ["0"], size, color, font, bold=bold, algn=algn)
    p = shape.text_frame._txBody.find(qn("a:p"))
    run = p.find(qn("a:r"))
    fld = parse_xml(f'<a:fld {nsdecls("a")} id="{{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}}" type="slidenum">'
                    f'{run_props(size, color, font, bold)}<a:t>‹#›</a:t></a:fld>')
    run.addprevious(fld)
    p.remove(run)


def add_slot(slide, shape_type, box, color):
    """Where an optional picture goes; the builder swaps it for the picture or removes it."""
    shape = slide.shapes.add_shape(shape_type, *(Emu(emu(v)) for v in box))
    shape.name = "Image slot"
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor.from_string(color)
    shape.line.fill.background()
    tf = shape.text_frame
    tf.text = "Optional image"
    return shape


def name_slide(slide, name):
    slide._element.find(qn("p:cSld")).set("name", name)


def tokens(theme):
    return bd.THEMES[theme]["palette"]


def footer(slide, theme, number_box, label_right, label_y, number_bold=True):
    pal = tokens(theme)
    font = bd.THEMES[theme]["fonts"]["body"]
    add_box(slide, "Footer label", (label_right - 4.4, label_y, 4.4, 0.26), ["XXX section"], 12, pal["muted"], font,
            algn="r")
    add_slide_number(slide, number_box, 12, pal["strong"], font, bold=number_bold,
                     algn="l" if theme == "classic" else "r")


def content_title(slide, theme, box):
    family, bold, italic = bd.THEMES[theme]["fonts"]["title"]
    size = bd.THEMES[theme]["title"]["sizes"][0]
    color = "FFFFFF" if theme == "classic" else tokens(theme)["strong"]
    set_text(slide.shapes.title, box, ["XXX results"], size, color, family, bold, italic)


def build(art_dir=None):
    prs = Presentation(str(SRC / "lab-template.pptx"))
    lab = {s.part.partname.filename: s for s in prs.slides}
    cover, content = lab["slide1.xml"], lab["slide3.xml"]
    section_layout = next(lo for lo in prs.slide_layouts if lo.name == "节标题")
    lab_bg = Image.open(io.BytesIO(content.part.related_part(
        content._element.find(".//" + qn("p:bg")).find(".//" + qn("a:blip")).get(qn("r:embed"))).blob)).convert("RGBA")
    originals = list(prs.slides._sldIdLst)
    arts = {}

    def art(name, image):
        arts[name] = image
        return image

    # A photo that bleeds off the slide's edge starts right under the band: it must not cover the band's rule.
    band = band_end(lab_bg) / PPI
    if bd.THEMES["classic"]["panel_top"] / EMU < band - 1e-3:
        raise SystemExit(f"classic panel_top in build_deck.py is above the band's lower edge ({band:.3f}\")")

    # classic ---------------------------------------------------------------
    s = clone_slide(prs, cover)
    name_slide(s, "classic/cover")

    s = clone_slide(prs, content)
    name_slide(s, "classic/content")
    set_background(s, art("classic-content", classic_content(lab_bg)))
    drop_groups(s)
    add_logo_right(s, "pairs-lab-white", 12.95, 0.32, 0.40, "PAIRS Lab logo")
    footer(s, "classic", (12.26, 6.98, 0.7, 0.28), 12.18, 6.80)

    s = prs.slides.add_slide(section_layout)
    name_slide(s, "classic/section")
    set_background(s, art("classic-section", classic_section()))
    set_text(s.shapes.title, (1.11, 2.72, 10.6, 0.95), ["XXX section"], 44, "FFFFFF", "Calibri", bold=True)
    body = next(sh for sh in s.placeholders if sh.placeholder_format.idx == 1)
    set_text(body, (1.11, 3.86, 10.6, 0.9), ["XXX one line on what this part covers"], 24, "B8C5DA", "Calibri",
             anchor="t")
    add_logo_right(s, "pairs-lab-white", 12.95, 0.32, 0.40, "PAIRS Lab logo")

    # paper -----------------------------------------------------------------
    s = clone_slide(prs, cover)
    name_slide(s, "paper/cover")
    set_background(s, art("paper-cover", paper_cover()))
    drop_groups(s)
    add_logo(s, "hkustgz-color", 0.83, 0.62, 0.50, "The Hong Kong University of Science and Technology (Guangzhou)")
    add_logo(s, "pairs-lab-color", 3.75, 0.62, 0.50, "PAIRS Lab logo")
    set_text(s.shapes.title, (0.83, 2.45, 7.0, 1.2), ["Group Meeting"], 60, "111F4A", "Georgia", anchor="b")
    sub = next(sh for sh in s.placeholders if sh.placeholder_format.idx == 1)
    set_text(sub, (0.83, 3.98, 6.0, 1.0), ["2026.XX.XX", "Name"], 24, "5A6578", "Calibri", anchor="t", space_before=600)
    add_slot(s, MSO_SHAPE.RECTANGLE, (8.15, 0.0, 13.333 - 8.15, 7.5), "E6EAF0")

    s = clone_slide(prs, content)
    name_slide(s, "paper/content")
    set_background(s, solid="FFFFFF")
    drop_groups(s)
    content_title(s, "paper", (0.6, 0.40, 9.9, 0.75))
    add_logo_right(s, "pairs-lab-color", 12.73, 0.60, 0.34, "PAIRS Lab logo")
    footer(s, "paper", (12.23, 6.95, 0.5, 0.26), 12.12, 6.95, number_bold=False)

    s = prs.slides.add_slide(section_layout)
    name_slide(s, "paper/section")
    set_background(s, art("paper-section", paper_section()))
    set_text(s.shapes.title, (0.83, 2.62, 8.2, 1.1), ["XXX section"], 48, "111F4A", "Georgia", anchor="b")
    body = next(sh for sh in s.placeholders if sh.placeholder_format.idx == 1)
    set_text(body, (0.83, 3.92, 8.2, 0.9), ["XXX one line on what this part covers"], 24, "5A6578", "Calibri",
             anchor="t")
    add_logo_right(s, "pairs-lab-color", 12.73, 0.60, 0.34, "PAIRS Lab logo")

    # Drop the lab's own four slides; their art stays only where a sample still uses it.
    for sld_id in originals:
        prs.slides._sldIdLst.remove(sld_id)
        prs.part.drop_rel(sld_id.rId)
    for slide in prs.slides:
        prune_rels(slide)
    # Number the slide files 1..N in deck order, as PowerPoint does; python-pptx assumes it when it opens one.
    prs.part.rename_slide_parts([sld_id.rId for sld_id in prs.slides._sldIdLst])
    OUT.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(OUT))
    print(f"Wrote {OUT} ({len(prs.slides)} sample slides)")
    for slide in prs.slides:
        print(f"  {slide.part.partname.filename:12s} {slide.name}")
    if art_dir:
        art_dir.mkdir(parents=True, exist_ok=True)
        for name, image in arts.items():
            image.convert("RGB").resize((1600, 900), Image.LANCZOS).save(art_dir / f"{name}.png")
        print(f"Art previews in {art_dir}")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--art", type=Path, help="also save the background art as PNGs here, for a look")
    build(ap.parse_args().art)


if __name__ == "__main__":
    main()
