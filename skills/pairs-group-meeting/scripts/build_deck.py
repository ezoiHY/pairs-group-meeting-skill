#!/usr/bin/env python3
"""Build a PAIRS Lab group-meeting deck from a JSON spec.

deck_structure.py makes the deck's structure: the template's sample slides,
cloned in spec order, without the unused originals. python-pptx then fills each
clone, so every slide keeps the template's background, logos and title styling.
deck_structure.py also checks the finished deck. The Claude and Codex versions
of the skill share this file and differ only in deck_structure.py.

Usage:
    python build_deck.py spec.json [-o out.pptx] [--no-validate] [--pptx-skill DIR (Claude version)]

The spec format is documented in references/spec-format.md.
"""

import argparse
import datetime
import difflib
import json
import os
import re
import subprocess
import sys
import tempfile
from copy import deepcopy
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import deck_structure  # noqa: E402
from xml.sax.saxutils import escape

from PIL import Image, ImageDraw, ImageFont, ImageOps, UnidentifiedImageError
from pptx import Presentation
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.util import Emu

SKILL_DIR = Path(__file__).resolve().parent.parent
TEMPLATE = SKILL_DIR / "assets" / "template.pptx"
CONFIG = SKILL_DIR / "config.json"
# The user's own settings (presenter, output folder) live outside the skill, so updates never touch them.
USER_CONFIG = Path(os.environ.get("XDG_CONFIG_HOME") or Path.home() / ".config") / "pairs-group-meeting" / "config.json"

EMU_PER_PT = 12700
EMU_PER_IN = 914400
SLIDE_W, SLIDE_H = 12192000, 6858000
GAP = int(0.2 * EMU_PER_IN)
COL_GAP = int(0.35 * EMU_PER_IN)
WIDE_FIGURE = 2.6                      # width/height from which a figure is larger under the text than beside it


def _in(inches):
    return int(round(inches * EMU_PER_IN))


# ---------------------------------------------------------------------------
# Themes
#
# Each theme is a set of sample slides in the template, named "<theme>/<kind>" (cover, content,
# section), plus the tokens below for everything the builder writes itself. The art (backgrounds,
# logos) lives on the sample slides; tools/make_template.py draws it to the geometry given here.
# Every page that carries content is white: only covers and section dividers have art.
#
# Palette keys: bg (what the text sits on, for contrast checks), text, strong (takeaways, project
# names, headings), muted (captions, footnotes, list letters, agenda items not being covered),
# bullet, number (agenda numbers), accent (highlights: "ours", the best bar, the current
# milestone), mark (the lab's gold bar and rules; decoration, never text), done / ongoing (status;
# "scheme:accent6" is a theme color), table colors (head_fill, head_text, body_fill, stripe, rule,
# hair; "none" = no fill), chart (series colors), chart_muted (a quiet series, e.g. "before"),
# grid (chart gridlines and axis lines).

THEMES = {
    "classic": {
        "about": "the lab template, for group meetings: navy header band with radar rings, white pages",
        "fonts": {"body": "Calibri", "title": ("Calibri", True, False), "subtitle": ("Calibri", True, False),
                  "cover": ("Calibri", True, False), "section": ("Calibri", True, False),
                  "heading": ("Calibri", True, False), "statement": ("Calibri", True, False),
                  "number": ("Calibri", True, False)},
        "margin": 349250,                                   # the template's title and body left edge
        "title": {"w": 7620000, "max_w": _in(9.70), "sizes": (36, 24)},
        "cover_title": {"w": 6096000, "sizes": (48, 32)},
        "section_title": {"w": _in(10.6), "sizes": (44, 30)},
        "content": (1219200, _in(6.60)),                    # top, bottom of the content area
        "panel_top": _in(1.09),                             # a bleeding photo starts right under the band
        "footnote": (349250, _in(6.74), _in(7.4), _in(0.32)),  # x, y, w, h
        "table": "rules",
        "mark": "bar",                                      # the gold bar the lab's cover puts beside its title
        "palette": {"bg": "FCFCFE", "text": "002147", "strong": "111F4A", "muted": "4A5878", "bullet": "2F6DB5",
                    "number": "2F6DB5", "accent": "C9A04A", "mark": "D6B05C",
                    "done": "scheme:accent6", "ongoing": "980000",
                    "head_fill": "111F4A", "head_text": "FFFFFF", "body_fill": "FFFFFF", "stripe": "F2F5FA",
                    "rule": "D5DBE6", "hair": "D5DBE6",
                    "chart": ["111F4A", "5B8DB8", "C9A04A", "980000", "8A97B0"], "chart_muted": "A3AEC2",
                    "grid": "D5DBE6"},
    },
    "paper": {
        "about": "editorial, for formal talks: white pages, serif titles, ruled tables, the lab's rings on cover "
                 "and sections",
        "fonts": {"body": "Calibri", "title": ("Georgia", False, False), "subtitle": ("Georgia", False, True),
                  "cover": ("Georgia", False, False), "section": ("Georgia", False, False),
                  "heading": ("Georgia", False, False), "statement": ("Georgia", False, False),
                  "number": ("Georgia", False, False)},
        "margin": _in(0.6),
        "title": {"w": _in(9.9), "max_w": _in(9.9), "sizes": (36, 26)},
        "cover_title": {"w": _in(7.0), "sizes": (60, 40)},
        "section_title": {"w": _in(8.2), "sizes": (48, 32)},
        "content": (_in(1.42), _in(6.62)),
        "panel_top": _in(1.42),
        "footnote": (_in(0.6), _in(6.90), _in(7.0), _in(0.30)),
        "table": "booktabs",
        "mark": "rule",                                     # a short rule, as booktabs tables draw them
        "palette": {"bg": "FFFFFF", "text": "1E2635", "strong": "111F4A", "muted": "5A6578", "bullet": "B28A2E",
                    "number": "B28A2E", "accent": "B28A2E", "mark": "B28A2E",
                    "done": "2E7D32", "ongoing": "A3201C",
                    "head_fill": "none", "head_text": "111F4A", "body_fill": "none", "stripe": "none",
                    "rule": "111F4A", "hair": "CBD2DD",
                    "chart": ["1F2E52", "8E9BB3", "B28A2E", "A3201C", "5F6E89"], "chart_muted": "B9C2D1",
                    "grid": "E1E5EC"},
    },
}
DEFAULT_THEME = "classic"
# Words a spec may use for a color instead of hex: they follow the slide's palette.
COLOR_WORDS = {"primary": lambda p: p["chart"][0], "secondary": lambda p: p["chart"][1],
               "accent": lambda p: p["accent"], "muted": lambda p: p["chart_muted"],
               "text": lambda p: p["text"], "strong": lambda p: p["strong"]}

# The current theme, its palette and its geometry (EMU), set by use_theme.
THEME = None
P = None
THEME_NAME = None
MARGIN_X = CONTENT_TOP = CONTENT_BOTTOM = FULL_W = LIST_W = 0
FOOTNOTE_BOX = (0, 0, 0, 0)
BODY_FONT = "Calibri"


# Themes that were dropped, and the one a spec or config naming them gets instead.
RETIRED_THEMES = {"midnight": "classic"}


def use_theme(name):
    global THEME, THEME_NAME, P, MARGIN_X, CONTENT_TOP, CONTENT_BOTTOM, FULL_W, LIST_W, FOOTNOTE_BOX, BODY_FONT
    if name in RETIRED_THEMES:
        warn(f"the {name} theme is gone (content pages are white in every theme now) — built in "
             f"{RETIRED_THEMES[name]}")
        name = RETIRED_THEMES[name]
    if name not in THEMES:
        near = difflib.get_close_matches(str(name), THEMES, n=1)
        hint = f" — did you mean {near[0]!r}?" if near else ""
        raise SystemExit(f"theme {name!r} is not one of {', '.join(THEMES)}{hint}")
    THEME, THEME_NAME = THEMES[name], name
    P = THEME["palette"]
    MARGIN_X = THEME["margin"]
    CONTENT_TOP, CONTENT_BOTTOM = THEME["content"]
    FULL_W = SLIDE_W - 2 * MARGIN_X
    LIST_W = min(8890000, FULL_W)        # the lab template's Done/Todo text box width
    FOOTNOTE_BOX = THEME["footnote"]
    BODY_FONT = THEME["fonts"]["body"]


def resolve_color(value):
    """A spec color — hex, or a palette word like "accent" — as hex for the current slide."""
    if value is None:
        return None
    word = str(value).strip().lower()
    return COLOR_WORDS[word](P) if word in COLOR_WORDS else str(value)


# Latin font attributes PowerPoint writes for the fonts the themes use.
FONT_META = {
    "Calibri": 'panose="020F0502020204030204" pitchFamily="34" charset="0"',
    "Georgia": 'panose="02040502050405020303" pitchFamily="18" charset="0"',
}


def fonts_xml(family=None):
    """Latin and complex-script font for `family` (default: the theme's body font), DengXian for Chinese."""
    family = family or BODY_FONT
    meta = FONT_META.get(family, "")
    return (f'<a:latin typeface="{escape(family)}" {meta}/>'
            '<a:ea typeface="等线" panose="02010600030101010101" pitchFamily="2" charset="-122"/>'
            f'<a:cs typeface="{escape(family)}" {meta}/>')

DONE_WORDS = {"done", "finished", "complete", "completed", "完成", "已完成"}
ONGOING_WORDS = {"on-going", "ongoing", "in progress", "in-progress", "wip", "进行中"}

EMBEDDABLE = {"PNG", "JPEG", "GIF", "BMP", "TIFF"}     # image formats python-pptx embeds, as Pillow names them
RENDER_EXTS = {".pdf", ".svg"}
VIDEO_MIME = {".mp4": "video/mp4", ".m4v": "video/x-m4v", ".mov": "video/quicktime",
              ".avi": "video/x-msvideo", ".wmv": "video/x-ms-wmv"}

WARNINGS = []

# House style; config.json "style" overrides these.
STYLE = {"language": "en", "min_font_size": 22, "max_bullets_per_slide": 3, "max_words_per_bullet": 14,
         "max_words_per_item": 10, "max_compare_papers": 5, "max_compare_columns": 5}


def warn(msg):
    WARNINGS.append(msg)


def show_warnings():
    for w in WARNINGS:
        print(f"WARNING: {w}")
    WARNINGS.clear()


def font_sizes(top):
    """Sizes to try for one kind of text: from `top` down to the house minimum, in 2pt steps."""
    floor = STYLE["min_font_size"]
    return list(range(max(top, floor), floor, -2)) + [floor]


# ---------------------------------------------------------------------------
# Text measurement

CJK = "\u1100-\u11ff\u2e80-\u9fff\uac00-\ud7af\uf900-\ufaff\ufe30-\ufe4f\uff00-\uffef"
CJK_RE = re.compile(f"[{CJK}]")
SEGMENT_RE = re.compile(f"[{CJK}]+|[^{CJK}]+")
# Whitespace a line may break at: any but the no-break space that ties a number to its unit.
BREAK = "[^\\S ]"
TOKEN_RE = re.compile(f"[{CJK}]|(?:[^\\s{CJK}]| )+{BREAK}*|{BREAK}+")
# "1 ms", "60 Hz", "9.6 °": the space between a number and its unit becomes a no-break space, so a line
# never ends on the number and starts the next one with the unit.
UNIT_GAP_RE = re.compile(r"(?<=\d) (?=(?:ns|µs|us|ms|s|min|h|Hz|kHz|MHz|GHz|nm|µm|mm|cm|m|km|mg|g|kg|mV|V|mA|"
                         r"W|kW|dB|px|fps|FPS|KB|MB|GB|TB|K|M|B|×|%|°|°C)(?!\w))")

FONT_DIRS = [
    Path("/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts"),
    Path("/System/Library/Fonts/Supplemental"),
    Path("/Library/Fonts"),
    Path.home() / "Library" / "Fonts",
    Path("C:/Windows/Fonts"),
    Path("/usr/share/fonts/truetype/crosextra"),
]
# (family, bold, italic) -> file names to look for. A missing italic falls back to the upright face.
FONT_FILES = {
    ("Calibri", False, False): ("Calibri.ttf", "calibri.ttf", "Carlito-Regular.ttf"),
    ("Calibri", True, False): ("Calibrib.ttf", "calibrib.ttf", "Carlito-Bold.ttf"),
    ("Calibri", False, True): ("Calibrii.ttf", "calibrii.ttf", "Carlito-Italic.ttf"),
    ("Georgia", False, False): ("Georgia.ttf", "georgia.ttf"),
    ("Georgia", True, False): ("Georgia Bold.ttf", "georgiab.ttf"),
    ("Georgia", False, True): ("Georgia Italic.ttf", "georgiai.ttf"),
    ("DengXian", False, False): ("Deng.ttf", "deng.ttf"),
    ("DengXian", True, False): ("Dengb.ttf", "dengb.ttf", "Deng.ttf"),
}
# Width of a font relative to Calibri, for estimating when it is not installed.
FONT_WIDTH_GUESS = {"Georgia": 1.13}
# Single line spacing as a multiple of the size, as PowerPoint renders it.
LINE_FACTOR = {"Calibri": 1.2, "Georgia": 1.14}


class Metrics:
    """Text widths in points, from the deck's own fonts when they are installed."""

    def __init__(self):
        self.fonts = {}
        self.missing = set()

    def font(self, family, bold, italic):
        key = (family, bold, italic)
        if key not in self.fonts:
            names = FONT_FILES.get(key) or FONT_FILES.get((family, bold, False)) or ()
            hit = next((folder / n for folder in FONT_DIRS for n in names if (folder / n).exists()), None)
            self.fonts[key] = ImageFont.truetype(str(hit), 1000) if hit else None
            if not hit and family != "DengXian":
                self.missing.add(family)
        return self.fonts[key]

    def width(self, text, size, bold=False, italic=False, family=None):
        family = family or BODY_FONT
        total = 0.0
        for seg in SEGMENT_RE.findall(text):
            cjk = bool(CJK_RE.match(seg))
            font = self.font("DengXian" if cjk else family, bold, False if cjk else italic)
            if font:
                total += font.getlength(seg) / 1000 * size
            else:
                guess = 1.0 if cjk else FONT_WIDTH_GUESS.get(family, 1.0)
                total += sum(_guess_em(c, cjk) for c in seg) * size * guess * (1.04 if bold else 1.0)
        return total


def _guess_em(c, cjk):
    if cjk:
        return 1.0
    if c.isspace():
        return 0.23
    if c.isupper():
        return 0.6
    if c.isalnum():
        return 0.5
    return 0.4


METRICS = None


def line_count(runs, width_emu, size):
    """Greedy word wrap, the way PowerPoint breaks lines."""
    limit = width_emu / EMU_PER_PT * 0.97
    lines, x = 1, 0.0
    for run in runs:
        for tok in TOKEN_RE.findall(run.text):
            w = METRICS.width(tok, size, run.bold, run.italic, run.font)
            ink = METRICS.width(tok.rstrip(), size, run.bold, run.italic, run.font)
            if x > 0 and x + ink > limit:
                lines += 1
                x = 0.0
                if tok.isspace():
                    continue
            if ink > limit:
                lines += int((x + ink) // limit)
                x = (x + w) % limit
                continue
            x += w
    return lines


def last_line(runs, width_emu, size):
    """(lines, words on the last line) for a paragraph, broken the way line_count breaks it."""
    limit = width_emu / EMU_PER_PT * 0.97
    lines, x, words_on_line = 1, 0.0, 0
    for run in runs:
        for tok in TOKEN_RE.findall(run.text):
            w = METRICS.width(tok, size, run.bold, run.italic, run.font)
            ink = METRICS.width(tok.rstrip(), size, run.bold, run.italic, run.font)
            if x > 0 and x + ink > limit:
                lines, x, words_on_line = lines + 1, 0.0, 0
                if tok.isspace():
                    continue
            words_on_line += 0 if tok.isspace() else 1
            x += w
    return lines, words_on_line


def widowed(paras, width_emu, size):
    """True when a paragraph that wraps leaves a single word alone on its last line."""
    for p in paras:
        if p.runs:
            lines, last = last_line(p.runs, width_emu - p.marL, size)
            if lines > 1 and last == 1:
                return True
    return False


def prefer_unwidowed(sizes, paras_by_box, width_emu):
    """The first size in `sizes` (all of which fit) that leaves no word alone on a last line, else the first."""
    return next((size for size in sizes if not any(widowed(ps, width_emu, size) for ps in paras_by_box)),
                sizes[0] if sizes else None)


def line_height(runs, size):
    # Measured in PowerPoint renders: single spacing is ~1.2x the size for Calibri and DengXian alike.
    has_cjk = any(CJK_RE.search(r.text) for r in runs)
    factor = max([LINE_FACTOR.get(r.font or BODY_FONT, 1.2) for r in runs] or [LINE_FACTOR.get(BODY_FONT, 1.2)])
    return size * max(factor, 1.24 if has_cjk else 0)


# ---------------------------------------------------------------------------
# Paragraph model -> DrawingML


class Run:
    """A run of text. color None = the slide palette's text color; a "scheme:<name>" color is a theme color;
    font None = the theme's body font."""

    def __init__(self, text, bold=False, color=None, scheme=None, italic=False, font=None):
        if color and str(color).startswith("scheme:"):
            color, scheme = None, color[len("scheme:"):]
        self.text, self.bold, self.scheme, self.italic, self.font = text, bold, scheme, italic, font
        self._color = color

    @property
    def color(self):
        return self._color or P["text"]


class Para:
    """One <a:p>. Spacing is given as a fraction of the font size so it scales with fitting."""

    def __init__(self, runs, marL=0, indent=0, before=0.0, bullet=None, algn="l", color=None, bullet_color=None):
        self.runs, self.marL, self.indent = runs, marL, indent
        self.before, self.bullet, self.algn = before, bullet, algn
        self._color, self._bullet_color = color, bullet_color

    @property
    def color(self):
        return self._color or P["text"]

    def height(self, width_emu, size):
        if not self.runs:
            return line_height([], size) * EMU_PER_PT
        lines = line_count(self.runs, width_emu - self.marL, size)
        return (lines * line_height(self.runs, size) + self.before * size) * EMU_PER_PT

    def xml(self, size):
        if self.bullet:
            bu = (f'<a:buClr><a:srgbClr val="{self._bullet_color or P["bullet"]}"/></a:buClr><a:buSzPct val="100000"/>'
                  f'<a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/>'
                  f'<a:buChar char="{escape(self.bullet)}"/>')
        else:
            bu = "<a:buNone/>"
        ppr = (f'<a:pPr marL="{self.marL}" indent="{self.indent}" algn="{self.algn}">'
               f'<a:lnSpc><a:spcPct val="100000"/></a:lnSpc>'
               f'<a:spcBef><a:spcPts val="{round(self.before * size * 100)}"/></a:spcBef>'
               f'<a:spcAft><a:spcPts val="0"/></a:spcAft>{bu}</a:pPr>')
        runs = "".join(run_xml(r, size) for r in self.runs)
        font = self.runs[-1].font if self.runs else None
        end = (f'<a:endParaRPr lang="en-US" altLang="zh-CN" sz="{round(size * 100)}" dirty="0">'
               f'<a:solidFill><a:srgbClr val="{self.color}"/></a:solidFill>{fonts_xml(font)}</a:endParaRPr>')
        return f'<a:p {nsdecls("a")}>{ppr}{runs}{end}</a:p>'


def run_xml(run, size):
    if run.scheme:
        fill = f'<a:solidFill><a:schemeClr val="{run.scheme}"/></a:solidFill>'
    else:
        fill = f'<a:solidFill><a:srgbClr val="{run.color}"/></a:solidFill>'
    attrs = f'lang="en-US" altLang="zh-CN" sz="{round(size * 100)}" b="{1 if run.bold else 0}"'
    if run.italic:
        attrs += ' i="1"'
    return f'<a:r><a:rPr {attrs} dirty="0">{fill}{fonts_xml(run.font)}</a:rPr><a:t>{escape(run.text)}</a:t></a:r>'


def inline_runs(text, bold=False, color=None, italic=False, font=None):
    """Split **bold** markup into runs (and keep each number on the line of its unit)."""
    runs = []
    for i, part in enumerate(UNIT_GAP_RE.sub(" ", str(text)).split("**")):
        if part:
            runs.append(Run(part, bold=bold or i % 2 == 1, color=color, italic=italic, font=font))
    return runs


def stack_height(paras, width_emu, size):
    return sum(p.height(width_emu, size) for p in paras)


def write_paragraphs(tx_body, paras, size):
    for p in tx_body.findall(qn("a:p")):
        tx_body.remove(p)
    for p in paras:
        tx_body.append(parse_xml(p.xml(size)))
    if not paras:
        tx_body.append(parse_xml(f'<a:p {nsdecls("a")}/>'))   # PowerPoint rejects a text body with no paragraph


def add_textbox(slide, box, name, anchor="t"):
    x, y, w, h = (int(v) for v in box)
    shape = slide.shapes.add_textbox(Emu(x), Emu(y), Emu(w), Emu(h))
    shape.name = name
    body_pr = shape.text_frame._txBody.find(qn("a:bodyPr"))
    for child in list(body_pr):
        body_pr.remove(child)
    for k, v in {"wrap": "square", "lIns": "0", "tIns": "0", "rIns": "0", "bIns": "0", "anchor": anchor}.items():
        body_pr.set(k, v)
    body_pr.append(parse_xml(f'<a:normAutofit {nsdecls("a")}/>'))
    return shape


def fit_size(paras, width_emu, height_emu, sizes):
    for size in sizes:
        if stack_height(paras, width_emu, size) <= height_emu:
            return size, True
    return sizes[-1], False


def set_single_text(shape, text, size):
    """Replace a placeholder's text but keep the template's first run formatting."""
    tx_body = shape.text_frame._txBody
    paras = tx_body.findall(qn("a:p"))
    for extra in paras[1:]:
        tx_body.remove(extra)
    p = paras[0]
    runs = p.findall(qn("a:r"))
    for extra in runs[1:]:
        p.remove(extra)
    runs[0].find(qn("a:t")).text = text
    for rpr in (runs[0].find(qn("a:rPr")), p.find(qn("a:endParaRPr"))):
        if rpr is not None:
            if size:
                rpr.set("sz", str(round(size * 100)))
            ensure_cjk_font(rpr)


def ensure_cjk_font(rpr):
    """Template runs name only a Latin font; give CJK text DengXian instead of the theme's SimSun."""
    if rpr.find(qn("a:ea")) is not None:
        return
    ea = parse_xml(f'<a:ea {nsdecls("a")} typeface="等线" panose="02010600030101010101" pitchFamily="2" charset="-122"/>')
    latin = rpr.find(qn("a:latin"))
    if latin is not None:
        latin.addnext(ea)
        return
    later = [rpr.find(qn(t)) for t in ("a:cs", "a:sym", "a:hlinkClick", "a:hlinkMouseOver", "a:rtl", "a:extLst")]
    later = [e for e in later if e is not None]
    if later:
        later[0].addprevious(ea)
    else:
        rpr.append(ea)


# ---------------------------------------------------------------------------
# Slide fillers


def fit_one_line(text, font, width, sizes):
    """(size, fits): the largest size from sizes[0] down to sizes[1], in 2pt steps, at which text is one line."""
    family, bold, italic = font
    runs = [Run(text, bold=bold, italic=italic, font=family)]
    size = sizes[0]
    while size > sizes[1] and line_count(runs, width, size) > 1:
        size -= 2
    return size, line_count(runs, width, size) == 1


def set_title(slide, text):
    shape = slide.shapes.title
    spec = THEME["title"]
    width = spec["w"]
    size, ok = fit_one_line(text, THEME["fonts"]["title"], width, (spec["sizes"][0], spec["sizes"][0]))
    if not ok:
        width = spec["max_w"]
        size, ok = fit_one_line(text, THEME["fonts"]["title"], width, spec["sizes"])
        if not ok:
            warn(f"title too long for one line even at {size}pt: {text!r}")
    shape.width = Emu(width)
    set_single_text(shape, text, size)


def fill_cover(slide, s, base_dir, work):
    lines = [s["date"], s.get("name", "")] + list(s.get("extra_lines", []))
    fill_title_page(slide, s.get("title") or "Group Meeting", lines, s.get("image"), base_dir, work, "cover")


def fill_closing(slide, s, base_dir, work):
    """The last slide of a talk, on the cover's page: "Thank you" and a line or two (questions, contact)."""
    fill_title_page(slide, s.get("title") or "Thank you", list(s.get("lines", [])), s.get("image"), base_dir, work,
                    "closing slide")


def fill_title_page(slide, title, lines, image, base_dir, work, what):
    """The cover's page: its title, the lines under it (date and name, or a closing line), its photo."""
    shape = slide.shapes.title
    spec = THEME["cover_title"]
    font = THEME["fonts"]["cover"]
    size, ok = fit_one_line(title, font, spec["w"], spec["sizes"])
    if not ok:
        # A long talk title takes two lines rather than shrinking further. It may widen only where no photo
        # panel sits beside it, and it grows upwards, so the lines under it stay where they are.
        photo = bool(image) and any(sh.name == "Image slot" for sh in slide.shapes)
        width = spec["w"] if photo else spec["w"] + 1000000
        runs = [Run(title, bold=font[1], italic=font[2], font=font[0])]
        size = next((sz for sz in range(spec["sizes"][0], spec["sizes"][1] - 1, -2)
                     if line_count(runs, width, sz) <= 2), spec["sizes"][1])
        title_lines = line_count(runs, width, size)
        if title_lines > 2:
            warn(f"{what} title runs past two lines even at {size}pt — shorten it: {title!r}")
        height = int(title_lines * size * LINE_FACTOR.get(font[0], 1.2) * EMU_PER_PT + 0.1 * EMU_PER_IN)
        if height > shape.height:
            shape.top, shape.height = Emu(shape.top + shape.height - height), Emu(height)
        shape.width = Emu(width)
    set_single_text(shape, title, size)

    sub = next(sh for sh in slide.placeholders if sh.placeholder_format.idx == 1)
    lines = [line for line in lines if line]
    if not lines:
        sub._element.getparent().remove(sub._element)
        fill_image_slot(slide, image, base_dir, work, what)
        return
    tx_body = sub.text_frame._txBody
    paras = tx_body.findall(qn("a:p"))
    while len(paras) < len(lines):
        clone = deepcopy(paras[-1])
        tx_body.append(clone)
        paras.append(clone)
    for extra in paras[len(lines):]:
        tx_body.remove(extra)
    for p, text in zip(paras, lines):
        runs = p.findall(qn("a:r"))
        for extra in runs[1:]:
            p.remove(extra)
        runs[0].find(qn("a:t")).text = text
        ensure_cjk_font(runs[0].find(qn("a:rPr")))
    if len(lines) > 2:
        sub.height = Emu(sub.height + (len(lines) - 2) * int(0.42 * EMU_PER_IN))
    fill_image_slot(slide, image, base_dir, work, what)


def fill_image_slot(slide, image, base_dir, work, what):
    """Put `image` where the template's "Image slot" shape is, cropped to its shape; drop the slot either way."""
    slot = next((sh for sh in slide.shapes if sh.name == "Image slot"), None)
    if not image:
        if slot is not None:
            slot._element.getparent().remove(slot._element)
        return
    if slot is None:
        warn(f"the {THEME_NAME} theme's {what} has no place for an image — left out (classic covers have none)")
        return
    item = prepare_media(image, base_dir, work)
    if item["kind"] != "image":
        raise SystemExit(f"{what} image: use a picture (png, jpg…), not a {item['kind']}")
    geometry = slot._element.spPr.find(qn("a:prstGeom")).get("prst")
    x, y, w, h = slot.left, slot.top, slot.width, slot.height
    pic = slide.shapes.add_picture(str(item["file"]), x, y, w, h)
    pic.name = f"{what.capitalize()} image"
    pic._element.nvPicPr.cNvPr.set("descr", item.get("caption") or item["src"].name)
    crop_to_fill(pic, w / h, item)
    pic._element.spPr.find(qn("a:prstGeom")).set("prst", geometry)
    slot._element.addprevious(pic._element)          # keep the slot's place in the stacking order
    slot._element.getparent().remove(slot._element)


def crop_to_fill(pic, box_aspect, item):
    """Crop a picture or video to `box_aspect` (width/height) around the item's focus point (default: the
    middle), so it fills its box instead of sitting in it with bars."""
    focus = item.get("focus") or [0.5, 0.5]
    try:
        fx, fy = (min(1.0, max(0.0, float(v))) for v in focus[:2])
    except (TypeError, ValueError):
        warn(f"focus {focus!r} should be two numbers from 0 to 1 ([x, y] from the top left) — used the middle")
        fx = fy = 0.5
    aspect = item["aspect"]
    if aspect > box_aspect:
        cut = 1 - box_aspect / aspect
        pic.crop_left, pic.crop_right = cut * fx, cut * (1 - fx)
    elif aspect < box_aspect:
        cut = 1 - aspect / box_aspect
        pic.crop_top, pic.crop_bottom = cut * fy, cut * (1 - fy)


def fill_section(slide, s):
    """A divider between parts of the deck: the part's name, and one optional line under it."""
    shape = slide.shapes.title
    spec = THEME["section_title"]
    size, ok = fit_one_line(s["title"], THEME["fonts"]["section"], spec["w"], spec["sizes"])
    if not ok:
        warn(f"section title too long for one line even at {size}pt: {s['title']!r} — shorten it")
    set_single_text(shape, s["title"], size)
    sub = next((sh for sh in slide.placeholders if sh.placeholder_format.idx == 1), None)
    if sub is None:
        return
    if s.get("subtitle"):
        set_single_text(sub, s["subtitle"], None)
    else:
        sub._element.getparent().remove(sub._element)


def fill_footer(slide, label, number):
    """The footer the theme's sample slide carries: the current part of the deck, and the slide number."""
    for shape in list(slide.shapes):
        if shape.name == "Footer label":
            if label:
                set_single_text(shape, label, None)
            else:
                shape._element.getparent().remove(shape._element)
        elif shape.name == "Slide number":
            # A slide-number field: PowerPoint keeps it right when slides move; this is its cached text.
            for t in shape._element.iter(qn("a:t")):
                t.text = str(number)


def normalize_item(item):
    if isinstance(item, str):
        return {"text": item}
    return item


def status_run(status):
    key = status.strip().lower()
    if key in DONE_WORDS:
        return Run("done", color=P["done"])
    if key in ONGOING_WORDS:
        return Run(status if key == "in progress" else "on-going", color=P["ongoing"])
    return Run(status, color=P["ongoing"])


def progress_paras(projects, kind):
    """Paragraphs of a Done/Todo list, styled like the template's sample text."""
    paras = []
    for pi, proj in enumerate(projects):
        label = proj.get("label", "Project: ")
        head = [Run(label, bold=True, color=P["strong"])] if label else []
        head += inline_runs(proj.get("name", ""), bold=True, color=P["strong"])
        paras.append(Para(head, before=0 if pi == 0 else 18 / 28))
        # A project continued from the previous slide keeps counting where it left off.
        for ii, item in enumerate(proj.get("items", []), proj.get("_first", 0)):
            item = normalize_item(item)
            runs = [Run(f"{chr(ord('a') + ii % 26)}. ", color=P["muted"])] + inline_runs(item["text"])
            if kind == "done" and item.get("status"):
                runs += [Run(" — ", color=P["muted"]), status_run(item["status"])]
            paras.append(Para(runs, marL=457200, before=3 / 28))
    return paras


def plan_progress(s, kind):
    """Pick font size and width for a Done/Todo slide; split it when it cannot fit at the minimum size."""
    height = CONTENT_BOTTOM - CONTENT_TOP
    paras = progress_paras(s["projects"], kind)
    for size in font_sizes(28):
        for width in (LIST_W, FULL_W):
            if stack_height(paras, width, size) <= height:
                return [dict(s, _size=size, _width=width)]
    split_size = STYLE["min_font_size"]

    chunks, current = [], []
    for proj in s["projects"]:
        items = [normalize_item(i) for i in proj.get("items", [])]
        start = 0
        while True:
            name = proj.get("name", "") + (" (cont.)" if start else "")
            taken = []
            for item in items[start:]:
                trial = current + [dict(proj, name=name, items=taken + [item], _first=start)]
                if stack_height(progress_paras(trial, kind), FULL_W, split_size) > height:
                    break
                taken.append(item)
            if not taken and items[start:] and current:
                chunks.append(current)       # retry this project on a fresh slide
                current = []
                continue
            if not taken and items[start:]:
                taken = [items[start]]       # a single item taller than a slide: keep it anyway
                warn(f"'{s['title']}': an item is too long for a slide even on its own — shorten it: "
                     f"{taken[0]['text'][:60]!r}")
            current.append(dict(proj, name=name, items=taken, _first=start))
            start += len(taken)
            if start >= len(items):
                break
            chunks.append(current)
            current = []
    if current:
        chunks.append(current)
    warn(f"'{s['title']}' did not fit on one slide; split into {len(chunks)} slides at {split_size}pt")
    out = []
    for i, chunk in enumerate(chunks):
        title = s["title"] if i == 0 else f"{s['title']} (cont.)"
        out.append(dict(s, title=title, projects=chunk, _size=split_size, _width=FULL_W))
    return out


def fill_progress(slide, s, kind):
    set_title(slide, s["title"])
    body = add_textbox(slide, (MARGIN_X, CONTENT_TOP, s["_width"], CONTENT_BOTTOM - CONTENT_TOP),
                       "Done list" if kind == "done" else "Todo list")
    write_paragraphs(body.text_frame._txBody, progress_paras(s["projects"], kind), s["_size"])


def bullet_paras(items, color=None):
    paras = []
    for i, item in enumerate(items):
        item = normalize_item(item)
        level = int(item.get("level", 0))
        plain = item.get("bullet") is False
        runs = inline_runs(item["text"], bold=item.get("bold", False),
                           color=resolve_color(item.get("color")) or color)
        before = 0 if i == 0 else (0.15 if level else 0.35)
        if plain:
            paras.append(Para(runs, marL=level * 342900, before=before, color=color))
        elif level:
            paras.append(Para(runs, marL=628650, indent=-228600, before=before, bullet="–", color=color))
        else:
            paras.append(Para(runs, marL=285750, indent=-285750, before=before, bullet="•", color=color))
    return paras


def place_text(slide, paras, box, sizes, name, what):
    x, y, w, h = box
    size, ok = fit_size(paras, w, h, sizes)
    if not ok:
        warn(f"{what}: text overflows its area even at {size}pt — cut words or split the slide")
    used = min(stack_height(paras, w, size), h) if ok else h
    shape = add_textbox(slide, (x, y, w, used + 0.1 * EMU_PER_IN), name)
    write_paragraphs(shape.text_frame._txBody, paras, size)
    return used, size


def write_box(slide, box, paras, size, name, anchor="t"):
    shape = add_textbox(slide, box, name, anchor)
    write_paragraphs(shape.text_frame._txBody, paras, size)
    return shape


def fit_lines(paras, width, sizes, max_lines):
    """(size, fits): the largest size at which every paragraph wraps to at most max_lines lines."""
    for size in sizes:
        if all(line_count(p.runs, width - p.marL, size) <= max_lines for p in paras if p.runs):
            return size, True
    return sizes[-1], False


def place_subtitle(slide, s, x0, y0, width, label):
    """The slide's takeaway under its title, as on content slides; returns where the body starts."""
    if not s.get("subtitle"):
        return y0
    family, bold, italic = THEME["fonts"]["subtitle"]
    paras = [Para(inline_runs(s["subtitle"], bold=bold, color=P["strong"], italic=italic, font=family),
                  color=P["strong"])]
    used, _ = place_text(slide, paras, (x0, y0, width, 1.0 * EMU_PER_IN), font_sizes(24), "Subtitle",
                         f"{label} subtitle")
    return y0 + used + GAP


def place_footnote(slide, s, label):
    if s.get("footnote"):
        paras = [Para(inline_runs(s["footnote"], color=P["muted"]), color=P["muted"])]
        place_text(slide, paras, FOOTNOTE_BOX, [12, 11, 10], "Footnote", f"{label} footnote")


# ---------------------------------------------------------------------------
# Lines, bars and dots for the page forms. They carry only the fill and outline set here: python-pptx's
# default shape style would add the Office theme's outline and shadow.


def bare(shape):
    style = shape._element.find(qn("p:style"))
    if style is not None:
        shape._element.remove(style)
    return shape


def add_bar(slide, box, color, name="Mark"):
    """A filled rectangle: the lab's gold bar, or a short rule."""
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    x, y, w, h = (int(round(v)) for v in box)
    shape = bare(slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(x), Emu(y), Emu(max(w, 1)), Emu(max(h, 1))))
    shape.name = name
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor.from_string(color)
    shape.line.fill.background()
    return shape


def add_rule(slide, x1, y1, x2, y2, color, width_pt=1.0, name="Rule"):
    """A straight hairline from (x1, y1) to (x2, y2)."""
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_CONNECTOR
    from pptx.util import Pt
    line = bare(slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, *(Emu(int(round(v))) for v in (x1, y1, x2, y2))))
    line.name = name
    line.line.color.rgb = RGBColor.from_string(color)
    line.line.width = Pt(width_pt)
    return line


def add_dot(slide, cx, cy, d, fill, outline=None, outline_pt=1.5, name="Marker"):
    """A circle of diameter d centered on (cx, cy)."""
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.util import Pt
    box = (Emu(int(round(v))) for v in (cx - d / 2, cy - d / 2, d, d))
    shape = bare(slide.shapes.add_shape(MSO_SHAPE.OVAL, *box))
    shape.name = name
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor.from_string(fill)
    if outline:
        shape.line.color.rgb = RGBColor.from_string(outline)
        shape.line.width = Pt(outline_pt)
    else:
        shape.line.fill.background()
    return shape


# ---------------------------------------------------------------------------
# Page forms beyond the content slide: agenda, statement, columns, timeline. All are drawn on the
# theme's white content page, from typography and hairlines rather than boxes.


def fill_agenda(slide, s):
    """The talk's parts, numbered, one per row between hairlines. With `current`, the part about to start
    is set in full color and the others are muted, so the same agenda can open every part."""
    set_title(slide, s["title"])
    label = f"agenda '{s['title']}'"
    x0, x1 = MARGIN_X, SLIDE_W - MARGIN_X
    y0, y1 = place_subtitle(slide, s, x0, CONTENT_TOP, x1 - x0, label), CONTENT_BOTTOM
    items, current = s["items"], s.get("_current")
    hfam, hbold, hital = THEME["fonts"]["heading"]
    nfam, nbold, nital = THEME["fonts"]["number"]
    bar = THEME["mark"] == "bar"
    num_x = x0 + (_in(0.3) if bar else 0)       # leave room for the gold bar beside the current part
    num_w = _in(1.0)
    text_x = num_x + num_w
    text_w = x1 - text_x
    size, ok = fit_lines([Para(inline_runs(t, bold=True, italic=hital, font=hfam)) for t in items], text_w,
                         font_sizes(32), 1)
    if not ok:
        warn(f"{label}: an item runs past one line even at {size}pt — shorten it")
    row_h = min((y1 - y0) / len(items), _in(1.0))
    top = y0 + (y1 - y0 - row_h * len(items)) / 2
    for i, text in enumerate(items):
        y = top + i * row_h
        here = current == i
        on = current is None or here
        color, num_color = (P["strong"], P["number"]) if on else (P["muted"], P["muted"])
        number = Run(f"{i + 1:02d}", bold=nbold or here, italic=nital, color=num_color, font=nfam)
        write_box(slide, (num_x, y, num_w, row_h), [Para([number], color=num_color)], size,
                  f"Agenda number {i + 1}", "ctr")
        runs = inline_runs(text, bold=hbold or here, italic=hital, color=color, font=hfam)
        write_box(slide, (text_x, y, text_w, row_h), [Para(runs, color=color)], size, f"Agenda item {i + 1}", "ctr")
        if here and bar:
            add_bar(slide, (x0, y + row_h * 0.24, _in(0.07), row_h * 0.52), P["mark"], "Agenda mark")
        if i:
            add_rule(slide, x0, y, x1, y, P["hair"], 1.0, f"Agenda rule {i}")
    if THEME["table"] == "booktabs":
        for name, y in (("top", top), ("bottom", top + row_h * len(items))):
            add_rule(slide, x0, y, x1, y, P["rule"], 1.5, f"Agenda rule {name}")


def fill_statement(slide, s):
    """One message set large: a sentence, the number it is about if there is one, a line of detail under it."""
    set_title(slide, s["title"])
    label = f"statement '{s['title']}'"
    x0, x1, y0, y1 = MARGIN_X, SLIDE_W - MARGIN_X, CONTENT_TOP, CONTENT_BOTTOM
    family, bold, italic = THEME["fonts"]["statement"]
    bar = THEME["mark"] == "bar"
    number = s.get("number")
    if number:
        nfont = THEME["fonts"]["number"]
        num_w = _in(4.4)
        nsize, ok = fit_one_line(number, nfont, num_w, (132, 72))
        if not ok:
            warn(f"{label}: the number is too wide even at {nsize}pt — keep it to a few characters")
        text_x = x0 + num_w + _in(0.6)
        sizes = font_sizes(36)
    else:
        text_x = x0 + (_in(0.45) if bar else 0)
        sizes = font_sizes(44)
    text_w = min(x1 - text_x, _in(10.2))
    paras = [Para(inline_runs(s["text"], bold=bold, italic=italic, color=P["strong"], font=family), color=P["strong"])]
    fitting = [sz for sz in sizes if fit_lines(paras, text_w, [sz], 3)[1]]
    size = prefer_unwidowed(fitting, [paras], text_w) or sizes[-1]
    if not fitting:
        warn(f"{label}: the statement runs past three lines even at {size}pt — shorten it")
    text_h = stack_height(paras, text_w, size)
    detail = [Para(inline_runs(s["detail"], color=P["muted"]), color=P["muted"])] if s.get("detail") else []
    dsize = fit_lines(detail, text_w, font_sizes(24), 2)[0] if detail else 0
    detail_h = stack_height(detail, text_w, dsize) if detail else 0
    gap = _in(0.3) if detail else 0
    block_h = text_h + gap + detail_h
    mid = y0 + (y1 - y0) / 2
    top = mid - block_h / 2
    write_box(slide, (text_x, top, text_w, text_h + _in(0.1)), paras, size, "Statement")
    if detail:
        write_box(slide, (text_x, top + text_h + gap, text_w, detail_h + _in(0.1)), detail, dsize, "Statement detail")
    if number:
        nfam, nbold, nital = nfont
        num_h = nsize * LINE_FACTOR.get(nfam, 1.2) * EMU_PER_PT
        run = Run(number, bold=nbold, italic=nital, color=P["strong"], font=nfam)
        write_box(slide, (x0, mid - num_h / 2, num_w, num_h), [Para([run], color=P["strong"])], nsize,
                  "Statement number", "ctr")
        sep_x, span = x0 + num_w + _in(0.3), max(block_h, num_h * 0.6)
        if bar:
            add_bar(slide, (sep_x - _in(0.035), mid - span / 2, _in(0.07), span), P["mark"])
        else:
            add_rule(slide, sep_x, mid - span / 2, sep_x, mid + span / 2, P["hair"], 1.0, "Statement rule")
    elif bar:
        add_bar(slide, (x0, top, _in(0.07), text_h), P["mark"])
    else:
        add_bar(slide, (x0, top - _in(0.45), _in(1.0), _in(0.035)), P["mark"])
    place_footnote(slide, s, label)


def fill_columns(slide, s, base_dir, work):
    """Two to four points side by side — a heading, a line or two, optionally a visual on top — with
    hairlines between the columns instead of boxes around them."""
    set_title(slide, s["title"])
    label = f"slide '{s['title']}'"
    x0, x1 = MARGIN_X, SLIDE_W - MARGIN_X
    y0, y1 = place_subtitle(slide, s, x0, CONTENT_TOP, x1 - x0, label), CONTENT_BOTTOM
    cols = s["columns"]
    gap = _in(0.7)
    col_w = (x1 - x0 - gap * (len(cols) - 1)) / len(cols)
    hfam, hbold, hital = THEME["fonts"]["heading"]
    heads = [[Para(inline_runs(c.get("heading", ""), bold=hbold, italic=hital, color=P["strong"], font=hfam),
                   color=P["strong"])] for c in cols]
    texts = [[Para(inline_runs(t), before=0 if k == 0 else 0.6) for k, t in enumerate(as_list(c.get("text")))]
             for c in cols]
    media = []
    for c in cols:
        item = prepare_media(c["media"], base_dir, work) if c.get("media") else None
        if item and item.get("caption"):
            warn(f"{label}: a column's media caption is left out — the column's heading says what it shows")
        media.append(item)
    text_only = not any(media)
    hsize, ok = fit_lines([ps[0] for ps in heads], col_w, font_sizes(32 if text_only else 28), 2)
    if not ok:
        warn(f"{label}: a column heading runs past two lines even at {hsize}pt — shorten it")
    head_h = max(stack_height(ps, col_w, hsize) for ps in heads)
    mark_h = _in(0.32) if THEME["mark"] == "bar" else 0
    avail = y1 - y0
    media_h = min(avail * 0.5, max(col_w / m["aspect"] for m in media if m)) if any(media) else 0
    after_media = GAP * 1.5 if media_h else 0
    room = avail - media_h - after_media - mark_h - head_h - _in(0.16)
    fitting = [size for size in font_sizes(28 if text_only else 24)
               if max(stack_height(ps, col_w, size) for ps in texts) <= room]
    tsize = prefer_unwidowed(fitting, texts, col_w)
    if tsize is None:
        tsize = STYLE["min_font_size"]
        warn(f"{label}: column text overflows even at {tsize}pt — cut words")
    text_h = max(stack_height(ps, col_w, tsize) for ps in texts)
    block_h = media_h + after_media + mark_h + head_h + _in(0.16) + text_h
    # Columns with visuals start under the subtitle; text-only ones sit in the middle of the space left.
    top = y0 if media_h else y0 + max(0, (avail - block_h) / 2)
    for i, c in enumerate(cols):
        cx, y = x0 + i * (col_w + gap), top
        if media_h:
            if media[i]:
                add_media_item(slide, media[i], contain(media[i]["aspect"], (cx, y, col_w, media_h)), i + 1)
            y += media_h + after_media
        if mark_h:
            add_bar(slide, (cx, y, _in(0.6), _in(0.06)), P["mark"], f"Column mark {i + 1}")
            y += mark_h
        write_box(slide, (cx, y, col_w, head_h + _in(0.1)), heads[i], hsize, f"Column heading {i + 1}")
        y += head_h + _in(0.16)
        write_box(slide, (cx, y, col_w, text_h + _in(0.1)), texts[i], tsize, f"Column text {i + 1}")
        if i:
            add_rule(slide, cx - gap / 2, top, cx - gap / 2, top + block_h, P["hair"], 1.0, f"Column rule {i}")
    place_footnote(slide, s, label)


# Words a timeline step's status may use, and how it is drawn.
TIMELINE_STATUS = {**{w: "done" for w in DONE_WORDS},
                   **{w: "now" for w in ONGOING_WORDS | {"now", "current"}},
                   **{w: "next" for w in ("next", "planned", "plan", "todo", "to do", "upcoming", "later")}}


def fill_timeline(slide, s):
    """Milestones or steps along one line: each one's date or name above its marker, a line or two under
    it. Done steps are filled, the current one is gold, the ones still ahead are open circles."""
    set_title(slide, s["title"])
    label = f"timeline '{s['title']}'"
    x0, x1 = MARGIN_X, SLIDE_W - MARGIN_X
    y0, y1 = place_subtitle(slide, s, x0, CONTENT_TOP, x1 - x0, label), CONTENT_BOTTOM
    points = s["points"]
    statuses = [p.get("_status") for p in points]
    step = (x1 - x0) / len(points)
    cell_w = step - _in(0.25)
    centers = [x0 + step * (i + 0.5) for i in range(len(points))]
    hfam, hbold, hital = THEME["fonts"]["heading"]
    heads = [[Para(inline_runs(p.get("label", ""), bold=hbold, italic=hital, color=P["strong"], font=hfam),
                   algn="ctr", color=P["strong"])] for p in points]
    texts = []
    for p, st in zip(points, statuses):
        color = P["muted"] if st == "next" else None
        texts.append([Para(inline_runs(p.get("text", ""), color=color), algn="ctr", color=color)])
    lsize, ok = fit_lines([ps[0] for ps in heads], cell_w, font_sizes(26), 2)
    fitting = [sz for sz in font_sizes(24) if fit_lines([ps[0] for ps in texts], cell_w, [sz], 3)[1]]
    ok2 = bool(fitting)
    tsize = prefer_unwidowed(fitting, texts, cell_w) or STYLE["min_font_size"]
    if not (ok and ok2):
        warn(f"{label}: a step's {'label' if not ok else 'text'} runs long even at the smallest size — shorten it")
    label_h = max(stack_height(ps, cell_w, lsize) for ps in heads)
    text_h = max(stack_height(ps, cell_w, tsize) for ps in texts)
    d, above, below = _in(0.28), _in(0.22), _in(0.28)
    block = label_h + above + d + below + text_h
    top = y0 + max(0, (y1 - y0 - block) / 2)
    axis_y = top + label_h + above + d / 2
    reached = max((i for i, st in enumerate(statuses) if st in ("done", "now")), default=0)
    add_rule(slide, centers[0], axis_y, centers[-1], axis_y, P["hair"], 2.5, "Timeline axis")
    if reached:
        add_rule(slide, centers[0], axis_y, centers[reached], axis_y, P["strong"], 2.5, "Timeline progress")
    for i, (cx, st) in enumerate(zip(centers, statuses), 1):
        if st == "now":
            add_dot(slide, cx, axis_y, d * 1.9, P["bg"], P["accent"], 1.5, f"Marker {i} ring")
            add_dot(slide, cx, axis_y, d, P["accent"], name=f"Marker {i}")
        elif st == "next":
            add_dot(slide, cx, axis_y, d, P["bg"], P["strong"], 1.75, f"Marker {i}")
        else:
            add_dot(slide, cx, axis_y, d, P["strong"], name=f"Marker {i}")
        write_box(slide, (cx - cell_w / 2, top, cell_w, label_h), heads[i - 1], lsize, f"Step label {i}", "b")
        write_box(slide, (cx - cell_w / 2, axis_y + d / 2 + below, cell_w, text_h + _in(0.1)), texts[i - 1], tsize,
                  f"Step text {i}")
    place_footnote(slide, s, label)


def fill_split(slide, s, item, layout):
    """One photo or video filling half the slide out to its edge; the takeaway and bullets beside it."""
    label = f"slide '{s['title']}'"
    half = SLIDE_W // 2
    top, bottom = THEME["panel_top"], CONTENT_BOTTOM
    if layout == "split":
        panel, tx0, tx1 = (half, top, SLIDE_W - half, bottom - top), MARGIN_X, half - COL_GAP
    else:
        panel, tx0, tx1 = (0, top, half, bottom - top), half + COL_GAP, SLIDE_W - MARGIN_X
    # A photo is cropped to fill the half; a video keeps its whole frame (a demo's every corner may matter),
    # as wide as the half allows and against the slide's edge. "crop" in the media item overrides either.
    if item.get("crop", item["kind"] == "image"):
        crop_to_fill(add_media_item(slide, item, panel, 1), panel[2] / panel[3], item)
    else:
        px, py, pw, ph = panel
        if item["aspect"] >= pw / ph:
            box = (px, py + (ph - pw / item["aspect"]) / 2, pw, pw / item["aspect"])
        else:
            w = ph * item["aspect"]
            box = (SLIDE_W - w if layout == "split" else 0, py, w, ph)
        add_media_item(slide, item, box, 1)
    width = tx1 - tx0
    family, bold, italic = THEME["fonts"]["subtitle"]
    sub = [Para(inline_runs(s["subtitle"], bold=bold, color=P["strong"], italic=italic, font=family),
                color=P["strong"])] if s.get("subtitle") else []
    bullets = bullet_paras(s.get("bullets", []))
    cap = [Para(inline_runs(item["caption"], color=P["muted"]), color=P["muted"])] if item.get("caption") else []
    cap_h = stack_height(cap, width, STYLE["min_font_size"]) if cap else 0
    room = bottom - CONTENT_TOP - (cap_h + GAP if cap else 0)
    ssize = fit_lines(sub, width, font_sizes(28), 3)[0] if sub else 0
    sub_h = stack_height(sub, width, ssize) if sub else 0
    between = _in(0.32) if sub and bullets else 0
    bsize, ok = fit_size(bullets, width, room - sub_h - between, font_sizes(26)) if bullets else (0, True)
    if not ok:
        warn(f"{label} bullets: text overflows its column even at {bsize}pt — cut words or split the slide")
    bul_h = stack_height(bullets, width, bsize) if bullets else 0
    y = CONTENT_TOP + max(0, (room - sub_h - between - bul_h) / 2)
    if sub:
        write_box(slide, (tx0, y, width, sub_h + _in(0.1)), sub, ssize, "Subtitle")
        y += sub_h + between
    if bullets:
        write_box(slide, (tx0, y, width, bul_h + _in(0.1)), bullets, bsize, "Bullets")
    if cap:
        write_box(slide, (tx0, bottom - cap_h, width, cap_h + _in(0.05)), cap, STYLE["min_font_size"], "Caption 1")


# ---------------------------------------------------------------------------
# Media


def spec_file(value, base_dir, what):
    """A file the spec names: ~ expanded, a relative path taken from the spec's folder."""
    path = Path(os.path.expanduser(str(value)))
    if not path.is_absolute():
        path = (base_dir / path).resolve()
    if not path.exists():
        raise SystemExit(f"{what} not found: {path}")
    return path


def prepare_media(item, base_dir, work):
    item = {"path": item} if isinstance(item, str) else dict(item)
    if "chart" in item:
        item.update(kind="chart", aspect=float(item.get("aspect", 1.6)))
        return item
    path = spec_file(item["path"], base_dir, "media")
    ext = path.suffix.lower()
    item["src"] = path
    if ext in VIDEO_MIME:
        poster = spec_file(item["poster"], base_dir, "poster") if item.get("poster") else video_poster(path, work)
        poster, aspect = embeddable_image(poster, work)
        item.update(kind="video", poster=poster, mime=VIDEO_MIME[ext], aspect=aspect)
        return item
    if ext in RENDER_EXTS:
        import pymupdf
        doc = pymupdf.open(path)
        number = int(item.get("page", 1))
        if not 1 <= number <= len(doc):
            # Page 1 instead would put the wrong figure (often a paper's title page) on the slide.
            raise SystemExit(f"{path.name} has no page {number} (it has {len(doc)}) — fix the page number")
        out = work / f"{path.stem}-{len(list(work.iterdir()))}.png"
        doc[number - 1].get_pixmap(dpi=220, alpha=False).save(out)
        path = out
    path, aspect = embeddable_image(path, work)
    item.update(kind="image", file=path, aspect=aspect)
    return item


def embeddable_image(path, work):
    """(file, aspect) of an upright image python-pptx can embed.

    The format is read from the file, not its name. A file goes in as it is unless it has to be
    converted or turned upright, so a JPEG stays a JPEG and a GIF keeps its animation.
    """
    try:
        im = Image.open(path)
    except UnidentifiedImageError:
        converted = sips_jpeg(path, work)
        if not converted:
            raise SystemExit(f"can't use {path.name}: it is neither an image this builder reads nor a video "
                             f"PowerPoint plays ({', '.join(sorted(VIDEO_MIME))}) — convert it to .png, .jpg "
                             "or .mp4") from None
        path, im = converted, Image.open(converted)
    with im:
        rotated = im.getexif().get(EXIF_ORIENTATION, 1) != 1
        if im.format in EMBEDDABLE and not rotated:
            return path, im.width / im.height
        upright = ImageOps.exif_transpose(im) if rotated else im
        if im.format in ("JPEG", "MPO"):
            # A photo stays a JPEG. MPO is a phone JPEG with a depth or HDR map attached; its main image is kept.
            rgb = upright.mode in ("RGB", "L")
            out = work / f"{path.stem}-{len(list(work.iterdir()))}.jpg"
            (upright if rgb else upright.convert("RGB")).save(
                out, "JPEG", quality=92, icc_profile=im.info.get("icc_profile") if rgb else None)
        else:
            out = work / f"{path.stem}-{len(list(work.iterdir()))}.png"
            png_ready(upright).save(out)
        return out, upright.width / upright.height


def sips_jpeg(path, work):
    """A JPEG copy made by macOS's sips, which reads formats Pillow doesn't (HEIC photos); None elsewhere."""
    out = work / f"{path.stem}-{len(list(work.iterdir()))}.jpg"
    try:
        subprocess.run(["sips", "-s", "format", "jpeg", "-s", "formatOptions", "high", str(path), "--out", str(out)],
                       capture_output=True, timeout=120)
    except (OSError, subprocess.TimeoutExpired):
        return None
    return out if out.exists() else None


EXIF_ORIENTATION = 0x0112


def png_ready(im):
    """The image in a mode PNG can store (CMYK and the like become RGB)."""
    return im if im.mode in ("1", "L", "LA", "P", "RGB", "RGBA", "I", "I;16") else im.convert("RGB")


def video_poster(path, work):
    """First-frame poster via macOS Quick Look; a plain dark frame when that is unavailable."""
    out_dir = work / f"poster-{len(list(work.iterdir()))}"
    out_dir.mkdir()
    try:
        subprocess.run(["qlmanage", "-t", "-s", "1280", "-o", str(out_dir), str(path)],
                       capture_output=True, timeout=60)
        made = list(out_dir.glob("*.png"))
        if made:
            return made[0]
    except (OSError, subprocess.TimeoutExpired):
        pass
    warn(f"no poster frame for {path.name}; pass \"poster\" to use your own")
    poster = out_dir / "poster.png"
    im = Image.new("RGB", (1280, 720), (17, 31, 74))
    d = ImageDraw.Draw(im)
    d.polygon([(560, 260), (560, 460), (740, 360)], fill=(255, 255, 255))
    im.save(poster)
    return poster


def caption_height(text, width_emu, size=None):
    if not text:
        return 0
    size = size or STYLE["min_font_size"]
    runs = [Run(text)]
    return line_count(runs, width_emu, size) * line_height(runs, size) * EMU_PER_PT + 0.06 * EMU_PER_IN


def media_layout(items, box, rows=None, valign="top", what="media"):
    """Justified rows: each row spans the box width; rows are scaled together to fit its height.

    Each item owns a cell as wide as the item at full scale. When the rows shrink to fit the height,
    an item shrinks inside its cell and its caption keeps the cell's width, so a narrow (portrait)
    item doesn't wrap its caption onto more and more lines and shrink itself further.
    """
    x, y, w, h = box
    n = len(items)
    best = None
    for r in ([min(int(rows), n)] if rows else range(1, n + 1)):
        counts = [n // r + (1 if i < n % r else 0) for i in range(r)]
        rows_spec, idx = [], 0
        for c in counts:
            group = items[idx:idx + c]
            row_h = (w - GAP * (c - 1)) / sum(it["aspect"] for it in group)
            cap = max(caption_height(it.get("caption"), it["aspect"] * row_h) for it in group)
            rows_spec.append((idx, c, row_h, cap))
            idx += c
        fixed = GAP * (r - 1) + sum(cap for *_, cap in rows_spec)
        scale = min(1.0, (h - fixed) / sum(row_h for _, _, row_h, _ in rows_spec))
        if scale <= 0:
            continue
        area = sum(it["aspect"] * (row_h * scale) ** 2 for i0, c, row_h, _ in rows_spec for it in items[i0:i0 + c])
        if best is None or area > best[0]:
            best = (area, rows_spec, scale)
    if best is None:
        raise SystemExit(f"{what}: the captions leave no room for the media — shorten the captions or use "
                         "fewer media")
    _, rows_spec, scale = best

    placed, cy = [], y
    for i0, c, row_h, cap in rows_spec:
        rh = row_h * scale
        cx = x
        for it in items[i0:i0 + c]:
            cell, iw = it["aspect"] * row_h, it["aspect"] * rh
            placed.append((it, (cx + (cell - iw) / 2, cy, iw, rh), (cx, cell), cap))
            cx += cell + GAP
        cy += rh + cap + GAP
    used = cy - GAP - y
    if used > h + EMU_PER_PT:
        warn(f"{what}: figures and captions run {(used - h) / EMU_PER_IN:.2f}\" past the content area — "
             "shorten the captions or use fewer media")
    shift = max(0, (h - used) / 2) if valign == "middle" else 0
    return [(it, (bx, by + shift, bw, bh), cell, cap) for it, (bx, by, bw, bh), cell, cap in placed], used


def add_media(slide, items, box, rows=None, valign="top", what="media"):
    placed, used = media_layout(items, box, rows, valign, what)
    for k, (it, (x, y, w, h), (cell_x, cell_w), cap) in enumerate(placed, 1):
        add_media_item(slide, it, (x, y, w, h), k)
        if it.get("caption"):
            tb = add_textbox(slide, (cell_x - GAP / 2, y + h + 0.04 * EMU_PER_IN, cell_w + GAP, cap), f"Caption {k}")
            write_paragraphs(tb.text_frame._txBody, [Para(inline_runs(it["caption"], color=P["muted"]), algn="ctr",
                                                          color=P["muted"])], STYLE["min_font_size"])
    return used


def add_media_item(slide, it, box, k):
    """One media item filling `box`: a native chart, an embedded video or a picture. Returns the video or
    picture shape (None for a chart)."""
    x, y, w, h = (int(v) for v in box)
    if it["kind"] == "chart":
        add_chart(slide, it["chart"], (x, y, w, h), f"Chart {k}")
        return None
    if it["kind"] == "video":
        shape = slide.shapes.add_movie(str(it["src"]), Emu(x), Emu(y), Emu(w), Emu(h),
                                       poster_frame_image=str(it["poster"]), mime_type=it["mime"])
        shape.name = f"Video {k}"
    else:
        shape = slide.shapes.add_picture(str(it["file"]), Emu(x), Emu(y), Emu(w), Emu(h))
        shape.name = f"Figure {k}"
    shape._element.xpath("./*[1]/p:cNvPr")[0].set("descr", it.get("caption") or it["src"].name)
    return shape


def contain(aspect, box):
    """The largest (x, y, w, h) of this aspect inside box, centered in it."""
    x, y, w, h = box
    if aspect >= w / h:
        fw, fh = w, w / aspect
    else:
        fw, fh = h * aspect, h
    return x + (w - fw) / 2, y + (h - fh) / 2, fw, fh


def add_chart(slide, spec, box, name):
    """Native PowerPoint chart (stays editable), text at the house minimum size."""
    from pptx.chart.data import CategoryChartData
    from pptx.dml.color import RGBColor
    from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
    from pptx.util import Pt

    kinds = {"bar": XL_CHART_TYPE.COLUMN_CLUSTERED, "hbar": XL_CHART_TYPE.BAR_CLUSTERED,
             "line": XL_CHART_TYPE.LINE_MARKERS}
    kind = spec.get("type", "bar")
    if kind not in kinds:
        raise SystemExit(f"chart type {kind!r}: use bar, hbar or line")
    number_format = spec.get("number_format", "General")
    data = CategoryChartData(number_format=number_format)
    data.categories = [str(c) for c in spec["categories"]]
    for ser in spec["series"]:
        data.add_series(ser["name"], ser["values"])
    x, y, w, h = (int(v) for v in box)
    frame = slide.shapes.add_chart(kinds[kind], Emu(x), Emu(y), Emu(w), Emu(h), data)
    frame.name = name
    chart = frame.chart
    size = Pt(STYLE["min_font_size"])
    chart.font.size = size
    chart.font.name = BODY_FONT
    chart.font.color.rgb = RGBColor.from_string(P["text"])
    chart.has_title = False
    multi = len(spec["series"]) > 1
    chart.has_legend = multi
    if multi:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False

    plot = chart.plots[0]
    if kind != "line":
        plot.gap_width = 70
        plot.overlap = -10 if multi else 0
    highlight = {str(c) for c in spec.get("highlight", [])}
    palette = P["chart"]
    free = 0                        # series without a color of their own take the palette in order
    for i, series in enumerate(plot.series):
        given = resolve_color(spec["series"][i].get("color"))
        if not given:
            given, free = palette[free % len(palette)], free + 1
        color = RGBColor.from_string(given)
        if kind == "line":
            series.smooth = False
            series.format.line.color.rgb = color
            series.format.line.width = Pt(3)
            series.marker.size = 9
            series.marker.format.fill.solid()
            series.marker.format.fill.fore_color.rgb = color
            series.marker.format.line.color.rgb = color
        else:
            series.format.fill.solid()
            series.format.fill.fore_color.rgb = color
            if not multi:
                for j, cat in enumerate(data.categories):
                    if cat.label in highlight:
                        point = series.points[j].format.fill
                        point.solid()
                        point.fore_color.rgb = RGBColor.from_string(P["accent"])

    labels = spec.get("labels", kind != "line")
    value_axis = chart.value_axis
    if labels:
        plot.has_data_labels = True
        dl = plot.data_labels
        dl.font.size = size
        dl.font.bold = True
        dl.font.color.rgb = RGBColor.from_string(P["text"])
        dl.number_format = number_format
        dl.number_format_is_linked = False
        dl.position = XL_LABEL_POSITION.ABOVE if kind == "line" else XL_LABEL_POSITION.OUTSIDE_END
        # A narrow chart would otherwise wrap a label like "14.9°" after the number.
        dl._element.find(qn("c:txPr")).find(qn("a:bodyPr")).set("wrap", "none")
        value_axis.visible = False          # the labels carry the numbers
        value_axis.has_major_gridlines = False
    else:
        value_axis.has_major_gridlines = True
        value_axis.major_gridlines.format.line.color.rgb = RGBColor.from_string(P["grid"])
        value_axis.format.line.fill.background()
        value_axis.tick_labels.font.size = size
    for key, attr in (("y_min", "minimum_scale"), ("y_max", "maximum_scale")):
        if spec.get(key) is not None:
            setattr(value_axis, attr, float(spec[key]))
    category_axis = chart.category_axis
    category_axis.tick_labels.font.size = size
    category_axis.format.line.color.rgb = RGBColor.from_string(P["grid"])
    category_axis.has_major_gridlines = False
    if kind == "hbar":
        # PowerPoint draws horizontal bars bottom-up; reversed, the first category is at the top, as written.
        category_axis.reverse_order = True


# ---------------------------------------------------------------------------
# Tables


def cell_spec(value):
    if isinstance(value, dict):
        return {"text": str(value.get("text", "")), "bold": value.get("bold", False),
                "color": value.get("color"), "fill": value.get("fill")}
    return {"text": "" if value is None else str(value), "bold": False, "color": None, "fill": None}


CELL_PAD_X = int(0.08 * EMU_PER_IN)
CELL_PAD_Y = int(0.045 * EMU_PER_IN)


ALIGN = {"left": "l", "center": "ctr", "right": "r"}


def table_align(align, ncols, what):
    """One alignment code (or None for the default) per column."""
    if align is None:
        return [None] * ncols
    if isinstance(align, str):
        align = [align] * ncols
    if len(align) != ncols:
        warn(f"{what}: align has {len(align)} entries for {ncols} columns — the rest use the default")
    codes = []
    for j in range(ncols):
        a = align[j] if j < len(align) else None
        if a is not None and a not in ALIGN:
            warn(f"{what}: align {a!r} is not left, center or right — used the default")
            a = None
        codes.append(ALIGN.get(a))
    return codes


def plan_table(t, width, max_h, max_size=24, what="table"):
    """Pick the largest font that fits the height without wrapping cells more than the minimum size does."""
    header = [cell_spec(v) for v in t.get("header", [])]
    rows = [[cell_spec(v) for v in r] for r in t["rows"]]
    grid = ([header] if header else []) + rows
    ncols = max(len(r) for r in grid)
    for r in grid:
        r.extend(cell_spec("") for _ in range(ncols - len(r)))
    col_widths = t.get("col_widths")
    if col_widths and (len(col_widths) != ncols or not all(isinstance(v, (int, float)) and v > 0 for v in col_widths)):
        warn(f"{what}: col_widths needs {ncols} positive numbers, one per column — sized the columns to fit instead")
        col_widths = None
    font_size = t.get("font_size")
    if font_size and float(font_size) < STYLE["min_font_size"]:
        warn(f"{what}: font_size {font_size} is below the {STYLE['min_font_size']}pt minimum — used "
             f"{STYLE['min_font_size']}pt; shorten or split the table instead")
        font_size = STYLE["min_font_size"]

    def layout(size):
        table_w = width
        if col_widths:
            ratios = [float(v) for v in col_widths]
        else:
            pad = 2 * CELL_PAD_X / EMU_PER_PT + 6
            natural = [max(METRICS.width(r[j]["text"], size, True) for r in grid) + pad for j in range(ncols)]
            cap = width / EMU_PER_PT * 0.45
            ratios = [min(max(v, 3 * size), cap) for v in natural]
            # A narrow table keeps a comfortable width instead of stretching across the slide.
            table_w = min(width, int(sum(ratios) * 1.5 * EMU_PER_PT))
        total = sum(ratios)
        widths = [int(table_w * v / total) for v in ratios]
        widths[-1] = table_w - sum(widths[:-1])
        heights, wrapped = [], 0
        for i, r in enumerate(grid):
            bold_row = bool(header) and i == 0
            lines = [line_count(inline_runs(c["text"], bold=bold_row or c["bold"]), widths[j] - 2 * CELL_PAD_X, size)
                     for j, c in enumerate(r)]
            wrapped += sum(lines)
            lh = max(line_height(inline_runs(c["text"]), size) for c in r)
            heights.append(int(max(lines) * lh * EMU_PER_PT + 2 * CELL_PAD_Y))
        return {"widths": widths, "heights": heights, "size": size, "lines": wrapped}

    floor = layout(STYLE["min_font_size"])
    if font_size:
        plan = layout(float(font_size))
        if plan["lines"] > floor["lines"]:
            warn(f"table at {plan['size']:g}pt wraps cells that fit on one line at smaller sizes — "
                 "drop font_size or shorten the cells")
    else:
        plans = [layout(s) for s in font_sizes(24) if s <= max(max_size, STYLE["min_font_size"])]
        fits = [p for p in plans if sum(p["heights"]) <= max_h]
        plan = next((p for p in fits if p["lines"] <= floor["lines"]), fits[-1] if fits else floor)
    if sum(plan["heights"]) > max_h:
        warn(f"{what}: table with {len(rows)} rows is taller than the space left at {plan['size']:g}pt — split it")
    plan.update(grid=grid, header=bool(header), align=table_align(t.get("align"), ncols, what))
    return plan


def add_table(slide, plan, x, y):
    grid, widths, heights = plan["grid"], plan["widths"], plan["heights"]
    shape = slide.shapes.add_table(len(grid), len(widths), Emu(x), Emu(y), Emu(sum(widths)), Emu(sum(heights)))
    shape.name = "Table"
    table = shape.table
    table.first_row = plan["header"]
    table.horz_banding = False
    for j, w in enumerate(widths):
        table.columns[j].width = Emu(w)
    for i, h in enumerate(heights):
        table.rows[i].height = Emu(h)
    style = THEME["table"]
    last = len(grid) - 1
    for i, r in enumerate(grid):
        is_head = plan["header"] and i == 0
        body_i = i - 1 if plan["header"] else i
        if style == "booktabs":
            # Rules only: a heavy rule above the table and below it, a light one under the header.
            top = (19050, P["rule"]) if i == 0 else (9525, P["rule"]) if plan["header"] and i == 1 else None
            bottom = (19050, P["rule"]) if i == last else (9525, P["rule"]) if is_head else None
        else:
            # A filled header, then horizontal hairlines between rows; no vertical lines.
            top = None if is_head or (plan["header"] and i == 1) else (6350, P["hair"])
            bottom = None if is_head else (6350, P["hair"])
        for j, c in enumerate(r):
            cell = table.cell(i, j)
            algn = plan["align"][j] or ("ctr" if is_head or j > 0 else "l")
            color = P["head_text"] if is_head else (resolve_color(c["color"]) or P["text"])
            runs = inline_runs(c["text"], bold=is_head or c["bold"], color=color)
            write_paragraphs(cell.text_frame._txBody, [Para(runs, algn=algn, color=color)], plan["size"])
            stripe = P["stripe"] if body_i % 2 else P["body_fill"]
            fill = resolve_color(c["fill"]) or (P["head_fill"] if is_head else stripe)
            style_cell(cell, fill, top, bottom)


def line_xml(tag, line):
    """A cell border: (width EMU, hex color), or None for no line."""
    if not line:
        return f'<a:{tag} {nsdecls("a")} w="0"><a:noFill/></a:{tag}>'
    w, color = line
    return (f'<a:{tag} {nsdecls("a")} w="{w}" cap="flat" cmpd="sng" algn="ctr">'
            f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:prstDash val="solid"/>'
            f'<a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/>'
            f'</a:{tag}>')


def style_cell(cell, fill, top=None, bottom=None):
    tc_pr = cell._tc.get_or_add_tcPr()
    for child in list(tc_pr):
        tc_pr.remove(child)
    tc_pr.set("marL", str(CELL_PAD_X))
    tc_pr.set("marR", str(CELL_PAD_X))
    tc_pr.set("marT", str(CELL_PAD_Y))
    tc_pr.set("marB", str(CELL_PAD_Y))
    tc_pr.set("anchor", "ctr")
    for tag, line in (("lnL", None), ("lnR", None), ("lnT", top), ("lnB", bottom)):
        tc_pr.append(parse_xml(line_xml(tag, line)))
    if fill and fill != "none":
        tc_pr.append(parse_xml(f'<a:solidFill {nsdecls("a")}><a:srgbClr val="{fill}"/></a:solidFill>'))
    else:
        tc_pr.append(parse_xml(f'<a:noFill {nsdecls("a")}/>'))


# ---------------------------------------------------------------------------
# Content slide


def fill_content(slide, s, base_dir, work):
    set_title(slide, s["title"])
    x0, x1 = MARGIN_X, SLIDE_W - MARGIN_X
    y0, y1 = CONTENT_TOP, CONTENT_BOTTOM
    label = f"slide '{s['title']}'"
    media = [prepare_media(m, base_dir, work) for m in s.get("media", [])]
    table = s.get("table")
    if table and not table.get("rows"):
        warn(f"{label}: table has no rows — left out")
        table = None
    layout = s.get("layout", "auto")

    if layout in ("split", "split-left"):
        shown = [m for m in media if m["kind"] != "chart"]
        if shown and not table:
            if len(media) > 1:
                warn(f"{label}: the {layout} layout shows one photo or video — used the first, left out the rest")
            fill_split(slide, s, shown[0], layout)
            place_footnote(slide, s, label)
            return
        warn(f"{label}: the {layout} layout needs a photo or video and no table — used the default layout")
        layout = "auto"

    y0 = place_subtitle(slide, s, x0, y0, x1 - x0, label)
    bullets = bullet_paras(s.get("bullets", []))
    if layout == "fit":
        # Figure-led slide: a wide figure gets the full width under the text, a squarer one sits beside it.
        layout = "stack" if media and all(m["aspect"] >= WIDE_FIGURE for m in media) else "auto"
    if layout == "auto":
        layout = "text-left" if bullets and media and not table else "stack"

    if layout in ("text-left", "text-right") and bullets:
        share = float(s.get("text_width", 0.38))
        if not 0.15 <= share <= 0.75:
            warn(f"{label}: text_width {share:g} is the bullet column's share of the width, between 0.15 and "
                 "0.75 — used 0.38")
            share = 0.38
        tw = int((x1 - x0) * share)
        if layout == "text-left":
            text_box, rest = (x0, y0, tw, y1 - y0), (x0 + tw + COL_GAP, y0, x1, y1)
        else:
            text_box, rest = (x1 - tw, y0, tw, y1 - y0), (x0, y0, x1 - tw - COL_GAP, y1)
        place_text(slide, bullets, text_box, font_sizes(24), "Bullets", f"{label} bullets")
        bullets = []
        x0, y0, x1, y1 = rest

    table_max = 24
    if bullets:
        alone = not media and not table
        room = (y1 - y0) if alone else (y1 - y0) * 0.45
        sizes = font_sizes(28) if alone else font_sizes(24)
        used, table_max = place_text(slide, bullets, (x0, y0, x1 - x0, room), sizes, "Bullets", f"{label} bullets")
        y0 += used + GAP

    if table:
        room = (y1 - y0) if not media else (y1 - y0) * 0.55
        plan = plan_table(table, x1 - x0, room, table_max, f"{label} table")
        add_table(slide, plan, x0 + (x1 - x0 - sum(plan["widths"])) // 2, y0)
        y0 += sum(plan["heights"]) + GAP

    if media:
        if y1 - y0 < 0.8 * EMU_PER_IN:
            raise SystemExit(f"{label}: no room left for media — move text or the table to another slide")
        only_media = not s.get("bullets") and not table
        add_media(slide, media, (x0, y0, x1 - x0, y1 - y0), s.get("media_rows"),
                  "middle" if only_media else "top", f"{label} media")

    if s.get("valign") == "middle":
        center_vertically([sh for sh in slide.shapes if sh.shape_id != slide.shapes.title.shape_id])

    place_footnote(slide, s, label)


def center_vertically(shapes):
    """Move everything below the header down so the block sits in the middle of the content area."""
    shapes = [sh for sh in shapes if sh.top >= CONTENT_TOP - EMU_PER_PT and sh.top < CONTENT_BOTTOM]
    if not shapes:
        return
    top = min(sh.top for sh in shapes)
    bottom = max(sh.top + sh.height for sh in shapes)
    shift = CONTENT_TOP + (CONTENT_BOTTOM - CONTENT_TOP - (bottom - top)) // 2 - top
    if shift > 0:
        for sh in shapes:
            sh.top = Emu(sh.top + shift)


# ---------------------------------------------------------------------------
# Spec handling


def load_config():
    """The skill's config.json (lab defaults), overridden key by key by the user's own; `style` merges."""
    cfg = {}
    for path in (CONFIG, USER_CONFIG):
        if path.exists():
            data = json.loads(path.read_text(encoding="utf-8"))
            style = {**cfg.get("style", {}), **data.get("style", {})}
            cfg.update(data)
            cfg["style"] = style
    return cfg


def parse_date(date):
    """'2026.10.29', '2026-10-29', '20261029' or '261029' -> datetime.date (None if unreadable)."""
    digits = re.findall(r"\d+", str(date))
    if len(digits) == 1 and len(digits[0]) in (6, 8):
        v = digits[0]
        digits = [v[:-4], v[-4:-2], v[-2:]]
    if len(digits) < 3:
        return None
    y, m, d = (int(v) for v in digits[:3])
    try:
        return datetime.date(y + 2000 if y < 100 else y, m, d)
    except ValueError:
        return None


def slide_texts(s):
    texts = [s.get("title"), s.get("subtitle"), s.get("name"), s.get("footnote")] + list(s.get("extra_lines", []))
    texts += [s.get("text"), s.get("detail"), s.get("number")] + list(s.get("lines", []))
    if s.get("type") == "agenda":
        texts += [str(t) for t in s.get("items", [])]
    for col in s.get("columns", []):
        texts += [col.get("heading")] + [str(t) for t in as_list(col.get("text"))]
    for point in s.get("points", []):
        texts += [point.get("label"), point.get("text")]
    texts += [normalize_item(b)["text"] for b in s.get("bullets", [])]
    for proj in s.get("projects", []):
        texts.append(proj.get("name"))
        for item in map(normalize_item, proj.get("items", [])):
            # Known statuses (完成, 进行中…) are written in English; any other status is shown as given.
            status = item.get("status", "")
            known = status.strip().lower() in DONE_WORDS | ONGOING_WORDS
            texts += [item["text"], None if known else status]
    table = s.get("table") or {}
    for row in [table.get("header", [])] + list(table.get("rows", [])):
        texts += [cell_spec(c)["text"] for c in row]
    for m in s.get("media", []):
        if isinstance(m, dict):
            chart = m.get("chart") or {}
            texts += [m.get("caption")] + [str(c) for c in chart.get("categories", [])]
            texts += [ser.get("name") for ser in chart.get("series", [])]
    return [t for t in texts if t]


LABEL_RE = re.compile(r"^\*\*[^*]{1,30}:\*\*\s*")


def words(text):
    """Words in a line, not counting a leading **Label:** such as a paper slide's **Problem:**."""
    return len(LABEL_RE.sub("", text).replace("**", "").split())


# Keys each spec object may have. Anything else is most likely a typo, which would otherwise be
# ignored without a word.
SPEC_KEYS = {
    "spec": {"slides", "output", "cover", "theme"},
    "cover": {"type", "title", "date", "name", "extra_lines", "image", "notes"},
    "closing": {"type", "title", "lines", "image", "notes"},
    "image": {"path", "focus"},
    "section": {"type", "title", "subtitle", "notes"},
    "agenda": {"type", "title", "subtitle", "items", "current", "notes"},
    "done": {"type", "title", "projects", "notes"},
    "todo": {"type", "title", "projects", "notes"},
    "content": {"type", "title", "subtitle", "bullets", "table", "media", "media_rows", "layout", "text_width",
                "footnote", "text_only", "valign", "notes"},
    "statement": {"type", "title", "text", "number", "detail", "footnote", "notes"},
    "columns": {"type", "title", "subtitle", "columns", "footnote", "notes"},
    "column": {"heading", "text", "media"},
    "timeline": {"type", "title", "subtitle", "points", "footnote", "notes"},
    "point": {"label", "text", "status"},
    "paper": {"type", "title", "name", "full_title", "venue", "pipeline", "pipeline_page", "pipeline_caption", "problem",
              "approach", "novelty", "limitations", "takeaway", "notes", "layout", "text_width"},
    "compare": {"type", "theme", "title", "methods", "results", "notes"},
    "project": {"name", "label", "items"},
    "item": {"text", "status"},
    "bullet": {"text", "level", "bullet", "bold", "color"},
    "table": {"header", "rows", "col_widths", "align", "font_size"},
    "methods": {"header", "rows", "col_widths", "align", "font_size", "takeaway"},
    "results": {"header", "rows", "col_widths", "align", "font_size", "takeaway", "note", "better"},
    "cell": {"text", "bold", "color", "fill"},
    "media": {"path", "caption", "poster", "page", "chart", "aspect", "focus", "crop"},
    "chart": {"type", "categories", "series", "number_format", "labels", "highlight", "y_min", "y_max"},
    "series": {"name", "values", "color"},
}
SLIDE_TYPES = ("cover", "agenda", "section", "done", "todo", "content", "statement", "columns", "timeline",
               "paper", "compare", "closing")
SPEC_VALUES = {"layout": {"auto", "text-left", "text-right", "stack", "fit", "split", "split-left"},
               "valign": {"top", "middle"}}
LIST_KEYS = {"bullets", "media", "extra_lines", "projects", "items", "highlight", "categories", "header", "lines",
             "columns", "points"}
# Keys that hold one line of text: a number there is used as text (a year as a series name), a list is
# joined into one line.
TEXT_KEYS = {"title", "subtitle", "name", "label", "text", "status", "caption", "footnote", "date", "theme",
             "takeaway", "note", "full_title", "venue", "problem", "approach", "pipeline_caption", "heading",
             "detail", "number"}
HEX_RE = re.compile(r"[0-9A-Fa-f]{6}")


def check_keys(obj, kind, where):
    """Warn about keys and values the builder would ignore or choke on, and tidy the ones it can."""
    if not isinstance(obj, dict):
        return
    for key, value in list(obj.items()):
        if key not in SPEC_KEYS[kind]:
            near = difflib.get_close_matches(key, SPEC_KEYS[kind], n=1)
            hint = f" — did you mean {near[0]!r}?" if near else ""
            warn(f"{where}: unknown key {key!r} is ignored{hint}")
        elif value is None:
            del obj[key]                                     # null means "not given"
        elif key in LIST_KEYS and isinstance(value, str):
            warn(f"{where}: {key} should be a list — used {value!r} as its only entry")
            obj[key] = [value]
        elif key in TEXT_KEYS and isinstance(value, (int, float)) and not isinstance(value, bool):
            obj[key] = str(value)
        elif key in TEXT_KEYS and isinstance(value, list) and not (kind == "column" and key == "text"):
            warn(f"{where}: {key} should be one line of text — joined the list into one")
            obj[key] = " ".join(map(str, value))
        elif key in SPEC_VALUES and (not isinstance(value, str) or value not in SPEC_VALUES[key]):
            warn(f"{where}: {key} {value!r} is not one of {sorted(SPEC_VALUES[key])} — used the default")
            del obj[key]
        elif key in ("color", "fill"):
            hexa = str(value).strip().lstrip("#")
            if HEX_RE.fullmatch(hexa):
                obj[key] = hexa.upper()
            elif hexa.lower() in COLOR_WORDS:
                obj[key] = hexa.lower()
            else:
                warn(f"{where}: {key} {value!r} is not a color like \"980000\" or one of "
                     f"{', '.join(COLOR_WORDS)} — used the default")
                del obj[key]


def need(obj, keys, message):
    """Stop on an object that has none of `keys`; building would only crash on it later."""
    if isinstance(obj, dict) and not any(k in obj for k in keys):
        raise SystemExit(message)


def check_spec(spec):
    """Warn about every key the builder would ignore or choke on, before anything is built."""
    check_keys(spec, "spec", "spec")
    if isinstance(spec.get("cover"), dict):
        cover = spec.pop("cover")
        if any(isinstance(s, dict) and s.get("type") == "cover" for s in spec.get("slides", [])):
            warn("spec: \"cover\" is only true or false — its fields are ignored, the cover slide is used")
        else:
            warn("spec: \"cover\" is only true or false — made its fields into a cover slide")
            spec["slides"].insert(0, {"type": "cover", **cover})
    for i, s in enumerate(spec.get("slides", []), 1):
        kind = s.get("type", "content") if isinstance(s, dict) else None
        if kind not in SLIDE_TYPES:
            continue
        where = f"slide {i} ({s.get('title') or s.get('theme') or kind})"
        if "tone" in s:
            warn(f"{where}: \"tone\" is gone — every content page is white now; ignored")
            del s["tone"]
        check_keys(s, kind, where)
        if kind == "paper" and isinstance(s.get("pipeline"), dict):
            check_keys(s["pipeline"], "media", f"{where} pipeline")
            need(s["pipeline"], ("path",), f"{where}: \"pipeline\" needs a \"path\"")
        if kind in ("cover", "closing") and isinstance(s.get("image"), dict):
            check_keys(s["image"], "image", f"{where} image")
            need(s["image"], ("path",), f"{where}: \"image\" needs a \"path\"")
        for col in s.get("columns", []) if kind == "columns" else []:
            check_keys(col, "column", where)
            need(col, ("heading", "text"), f"{where}: every column needs a \"heading\" or a \"text\"")
            if isinstance(col, dict) and isinstance(col.get("media"), (dict, str)):
                check_keys(col["media"], "media", f"{where} column media")
                need(col["media"], ("path", "chart"), f"{where}: a column's media needs a \"path\" (or a \"chart\")")
        for point in s.get("points", []) if kind == "timeline" else []:
            check_keys(point, "point", where)
            need(point, ("label", "text"), f"{where}: every point needs a \"label\" or a \"text\"")
        for proj in s.get("projects", []):
            check_keys(proj, "project", where)
            for item in proj.get("items", []) if isinstance(proj, dict) else []:
                check_keys(item, "item", where)
                need(item, ("text",), f"{where}: every item needs a \"text\"")
                if kind == "todo" and isinstance(item, dict) and item.get("status"):
                    warn(f"{where}: Todo items have no status — {item['status']!r} is not shown")
        for b in s.get("bullets", []):
            check_keys(b, "bullet", where)
            need(b, ("text",), f"{where}: every bullet needs a \"text\"")
        tables = [("table", s.get("table"))] if kind == "content" else []
        tables += [(k, s.get(k)) for k in ("methods", "results")] if kind == "compare" else []
        for table_kind, t in tables:
            check_keys(t, table_kind, f"{where} {table_kind}")
            rows = t.get("rows") if isinstance(t, dict) else None
            for j, row in enumerate(rows if isinstance(rows, list) else []):
                if not isinstance(row, list):
                    warn(f"{where} {table_kind}: a row should be a list of cells — used {row!r} as a one-cell row")
                    rows[j] = [row]
            for row in ([t.get("header", [])] + list(t.get("rows", []))) if isinstance(t, dict) else []:
                for cell in row:
                    check_keys(cell, "cell", f"{where} {table_kind}")
        for m in s.get("media", []):
            check_keys(m, "media", where)
            need(m, ("path", "chart"), f"{where}: every media item needs a \"path\" (or a \"chart\")")
            chart = m.get("chart") if isinstance(m, dict) else None
            check_keys(chart, "chart", where)
            need(chart, ("categories",), f"{where}: a chart needs \"categories\"")
            need(chart, ("series",), f"{where}: a chart needs \"series\"")
            for ser in chart.get("series", []) if isinstance(chart, dict) else []:
                check_keys(ser, "series", where)
                need(ser, ("values",), f"{where}: every chart series needs \"values\"")


def lint(s):
    """House-style checks: English only, a few short key sentences, a visual on every content slide."""
    where = f"slide '{s.get('title') or 'Group Meeting'}'"
    if STYLE["language"] == "en":
        foreign = next((t for t in slide_texts(s) if CJK_RE.search(t)), None)
        if foreign:
            warn(f"{where}: non-English text {foreign!r} — the house style is all English")
    limit = STYLE["max_words_per_bullet"]
    if s["type"] == "content":
        bullets = [normalize_item(b)["text"] for b in s.get("bullets", [])]
        if len(bullets) > STYLE["max_bullets_per_slide"]:
            warn(f"{where}: {len(bullets)} bullets — keep to {STYLE['max_bullets_per_slide']} key sentences "
                 "or split the slide")
        long = next((b for b in bullets if words(b) > limit), None)
        if long:
            warn(f"{where}: bullet over {limit} words — {long!r}")
        if not s.get("media") and not s.get("text_only"):
            warn(f"{where}: no figure, demo video or chart — add a visual (or set \"text_only\": true)")
    if s["type"] == "section" and s.get("subtitle") and words(s["subtitle"]) > limit:
        warn(f"{where}: section subtitle over {limit} words — one short line on what this part covers")
    if s["type"] == "statement":
        if words(s["text"]) > limit:
            warn(f"{where}: statement over {limit} words — one sentence people remember")
        if s.get("detail") and words(s["detail"]) > limit:
            warn(f"{where}: detail over {limit} words — one supporting line")
    if s["type"] == "columns":
        for col in s["columns"]:
            lines = as_list(col.get("text"))
            if len(lines) > 2:
                warn(f"{where}: a column has {len(lines)} lines — keep it to a heading and two lines")
            long = next((t for t in lines if words(str(t)) > limit), None)
            if long:
                warn(f"{where}: column line over {limit} words — {long!r}")
            if col.get("heading") and words(col["heading"]) > 6:
                warn(f"{where}: column heading over 6 words — {col['heading']!r}")
    if s["type"] == "timeline":
        item_limit = STYLE["max_words_per_item"]
        long = next((p["text"] for p in s["points"] if p.get("text") and words(p["text"]) > item_limit), None)
        if long:
            warn(f"{where}: step text over {item_limit} words — {long!r}")
    if s["type"] == "agenda" and len(s["items"]) > 6:
        warn(f"{where}: {len(s['items'])} parts — an agenda reads best with six or fewer")
    for proj in s.get("projects", []):
        item_limit = STYLE["max_words_per_item"]
        long = next((normalize_item(i)["text"] for i in proj.get("items", [])
                     if words(normalize_item(i)["text"]) > item_limit), None)
        if long:
            warn(f"{where}: item over {item_limit} words — {long!r}")


def as_list(value):
    if value is None or value == "":
        return []
    return [value] if isinstance(value, str) else list(value)


def paper_slides(s):
    """One paper -> two content slides: pipeline + problem/approach, then novelty | limitations."""
    name = s.get("title") or s.get("name")
    if not name:
        raise SystemExit("every paper slide needs a short \"title\" (the paper's name)")
    where = f"paper '{name}'"
    problem, approach = s.get("problem"), s.get("approach")
    novelty, limits = as_list(s.get("novelty")), as_list(s.get("limitations"))
    pipeline = s.get("pipeline")
    figure = dict(pipeline) if isinstance(pipeline, dict) else {"path": pipeline} if pipeline else None
    if figure and s.get("pipeline_page"):
        figure["page"] = s["pipeline_page"]          # the page of a PDF the figure is on
    needed = (("the original pipeline figure", figure), ("the problem it solves", problem),
              ("how it solves it", approach), ("the novelty", novelty), ("the limitations", limits))
    missing = [label for label, value in needed if not value]
    if missing:
        warn(f"{where}: missing {', '.join(missing)} — ask the user for them, don't fill them in yourself")
    notes = as_list(s.get("notes"))

    bullets = [f"**{label}:** {text}" for label, text in (("Problem", problem), ("Approach", approach)) if text]
    first = {"type": "content", "title": f"Paper reading: {name}", "bullets": bullets, "layout": "fit",
             "text_only": not figure}
    if s.get("full_title"):
        first["subtitle"] = s["full_title"]
    if figure:
        figure.setdefault("caption", s.get("pipeline_caption", ""))
        first["media"] = [figure]
    source = [t for t in (s.get("venue"), "pipeline figure from the paper" if figure else None) if t]
    if source:
        first["footnote"] = " · ".join(source)
    first.update({k: s[k] for k in ("layout", "text_width") if k in s})
    if notes:
        first["notes"] = notes[0]
    out = [first]

    columns = [(head, items) for head, items in (("Novelty", novelty), ("Limitations", limits)) if items]
    if columns:
        for head, items in columns:
            if len(items) > 2:
                warn(f"{where}: {len(items)} {head.lower()} — keep the two that matter most")
        long = next((t for _, items in columns for t in items if words(t) > STYLE["max_words_per_bullet"]), None)
        if long:
            warn(f"{where}: point over {STYLE['max_words_per_bullet']} words — {long!r}")
        second = {"type": "columns", "title": f"{name}: novelty and limits",
                  "columns": [{"heading": head, "text": [str(t) for t in items]} for head, items in columns]}
        if s.get("takeaway"):
            second["subtitle"] = s["takeaway"]
        if len(notes) > 1:
            second["notes"] = notes[1]
        out.append(second)
    return out


# A number as results tables write it: −3.2 (Unicode minus), .85, 1,234, scientific notation (1.2e-3,
# 3.4×10^5), and an optional K/M/B/G/T unit. A minus right after a letter or digit is a hyphen
# (ResNet-50), not a sign.
NUMBER_RE = re.compile(r"((?<!\w)[-−])?(\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?|\.\d+)"
                       r"(?:[eE]([-+−]?\d+)|\s?[×x]\s?10\^?([-+−]?\d+))?"
                       r"\s?([KkMBGT](?![A-Za-z]))?")
UNITS = {"K": 1e3, "k": 1e3, "M": 1e6, "B": 1e9, "G": 1e9, "T": 1e12}


def cell_number(text):
    """The first number in a cell, or None."""
    hit = NUMBER_RE.search(text)
    if not hit:
        return None
    sign, digits, exp_e, exp_x, unit = hit.groups()
    exponent = int((exp_e or exp_x or "0").replace("−", "-"))
    value = float(digits.replace(",", "")) * 10.0 ** exponent * UNITS.get(unit or "", 1)
    return -value if sign else value


def bold_best(table):
    """Rows of a results table with the best value of each column bold (from ↓/↑ in the header, or `better`)."""
    header = [cell_spec(c)["text"] for c in table.get("header", [])]
    rows = [list(r) for r in table["rows"]]
    ncols = max(len(r) for r in rows)
    better = table.get("better") or []
    if isinstance(better, str):                      # one rule for every column after the paper names
        better = [None] + [better] * (ncols - 1)
    for j in range(ncols):
        rule = better[j] if j < len(better) else None
        if rule is None and j < len(header):
            rule = "lower" if "↓" in header[j] else "higher" if "↑" in header[j] else None
        if rule not in ("lower", "higher"):
            continue
        found = []
        for i, r in enumerate(rows):
            value = cell_number(cell_spec(r[j])["text"]) if j < len(r) else None
            if value is not None:
                found.append((value, i))
        if not found:
            continue
        best = (min if rule == "lower" else max)(v for v, _ in found)
        for v, i in found:
            if v == best:
                rows[i][j] = dict(cell_spec(rows[i][j]), bold=True)
    return rows


def compare_slides(s):
    """One theme's papers -> a methods table slide, then a results table slide. Tables only, no figures."""
    theme = s.get("theme") or s.get("title")
    if not theme:
        raise SystemExit("every compare slide needs a \"theme\" (the group of papers it compares)")
    where = f"comparison '{theme}'"
    methods, results = s.get("methods"), s.get("results")
    if not methods or not methods.get("rows"):
        raise SystemExit(f"{where}: needs a \"methods\" table — compare how the papers work before their results")
    if not results or not results.get("rows"):
        warn(f"{where}: no results table — ask the user for the numbers on a shared benchmark")
    notes = as_list(s.get("notes"))

    out = []
    for kind, t in (("methods", methods), ("results", results)):
        if not t or not t.get("rows"):
            continue
        ncols = max(len(r) for r in ([t.get("header", [])] + list(t["rows"])))
        if len(t["rows"]) > STYLE["max_compare_papers"]:
            warn(f"{where} {kind}: {len(t['rows'])} papers — keep {STYLE['max_compare_papers']} or fewer per "
                 "theme, or split the theme")
        if ncols > STYLE["max_compare_columns"]:
            warn(f"{where} {kind}: {ncols} columns — keep {STYLE['max_compare_columns']} or fewer "
                 "(paper name + the few dimensions that matter)")
        table = {k: t[k] for k in ("header", "col_widths", "font_size") if k in t}
        table["rows"] = bold_best(t) if kind == "results" else t["rows"]
        table["align"] = t.get("align") or (["left"] * ncols if kind == "methods" else ["left"] + ["center"] * (ncols - 1))
        slide = {"type": "content", "title": f"Compare: {theme} — {kind}", "table": table, "text_only": True}
        if t.get("takeaway"):
            slide["subtitle"] = t["takeaway"]
        if kind == "results":
            if t.get("note"):
                slide["footnote"] = t["note"]
            else:
                warn(f"{where} results: no \"note\" — say which benchmark and setting the numbers come from")
        if len(notes) > len(out):
            slide["notes"] = notes[len(out)]
        out.append(slide)
    return out


def expand_slides(spec, cfg):
    check_spec(spec)
    slides = []
    for s in spec["slides"]:
        kind = s.get("type", "content")
        slides.extend(paper_slides(s) if kind == "paper" else compare_slides(s) if kind == "compare" else [dict(s)])
    if not any(s.get("type") == "cover" for s in slides) and spec.get("cover", True) is not False:
        slides.insert(0, {"type": "cover"})
    sections = [str(s["title"]) for s in slides if s.get("type") == "section" and s.get("title")]
    planned = []
    for s in slides:
        kind = s.get("type", "content")
        if kind not in SLIDE_TYPES:
            raise SystemExit(f"unknown slide type {kind!r}; use {', '.join(SLIDE_TYPES)}")
        s["type"] = kind
        if kind == "closing":
            s.setdefault("title", "Thank you")
            lint(s)
            planned.append(s)
            continue
        if kind == "agenda":
            s.setdefault("title", "Outline")
            items = [str(t.get("title") or t.get("text") or "") if isinstance(t, dict) else str(t)
                     for t in as_list(s.get("items"))] or sections
            if not items:
                warn(f"agenda '{s['title']}' has no \"items\" and the deck has no section slides to list — left out")
                continue
            s["items"] = items
            current = s.get("current")
            if current is not None:
                if isinstance(current, int) and not isinstance(current, bool):
                    index = current - 1 if 1 <= current <= len(items) else None
                else:
                    key = str(current).strip().lower()
                    index = next((k for k, t in enumerate(items) if t.strip().lower() == key), None)
                if index is None:
                    warn(f"agenda '{s['title']}': current {current!r} is neither an item's number (1–{len(items)}) "
                         "nor one of its titles — nothing is highlighted")
                s["_current"] = index
            lint(s)
            planned.append(s)
            continue
        if kind == "section":
            if not s.get("title"):
                raise SystemExit("every section slide needs a title (the name of the part it opens)")
            lint(s)
            planned.append(s)
        elif kind == "cover":
            day = parse_date(s["date"]) if s.get("date") else datetime.date.today()
            if day:
                s["date"] = day.strftime("%Y.%m.%d")
                s["_day"] = day
            else:
                warn(f"cover date {s['date']!r} is not a date; used as written")
            s.setdefault("name", cfg.get("presenter", ""))
            if not s["name"]:
                warn(f"cover has no presenter name — set \"name\" in the spec or \"presenter\" in {USER_CONFIG}")
            lint(s)
            planned.append(s)
        elif kind in ("done", "todo"):
            s.setdefault("title", cfg.get(f"{kind}_title") or ("Done last two weeks" if kind == "done" else "Todo"))
            if not s.get("projects"):
                warn(f"'{s['title']}' has no \"projects\" — the slide is empty")
                s["projects"] = []
            lint(s)
            planned.extend(plan_progress(s, kind))
        else:
            if not s.get("title"):
                raise SystemExit(f"every {kind} slide needs a title")
            if kind == "statement" and not s.get("text"):
                raise SystemExit(f"statement '{s['title']}' needs a \"text\": the sentence it shows")
            if kind == "columns":
                cols = [c if isinstance(c, dict) else {"text": c} for c in as_list(s.get("columns"))]
                if len(cols) < 2:
                    raise SystemExit(f"columns '{s['title']}' needs 2 to 4 \"columns\"")
                if len(cols) > 4:
                    warn(f"columns '{s['title']}': {len(cols)} columns — kept the first 4; put the rest on another slide")
                    cols = cols[:4]
                s["columns"] = cols
            if kind == "timeline":
                points = [p if isinstance(p, dict) else {"label": str(p)} for p in as_list(s.get("points"))]
                if len(points) < 2:
                    raise SystemExit(f"timeline '{s['title']}' needs at least 2 \"points\"")
                if len(points) > 6:
                    warn(f"timeline '{s['title']}': {len(points)} points — kept the first 6; split the plan")
                    points = points[:6]
                for p in points:
                    status = str(p.get("status") or "").strip().lower()
                    p["_status"] = TIMELINE_STATUS.get(status)
                    if status and not p["_status"]:
                        warn(f"timeline '{s['title']}': status {p['status']!r} is not done, now or next — drawn as "
                             "a plain step")
                s["points"] = points
            lint(s)
            planned.append(s)
    return planned


def template_samples():
    """{"<theme>/<kind>": slide file name} for the template's sample slides, read from the file itself
    (python-pptx renumbers slide parts when it opens a deck, so its names can differ from the file's)."""
    import zipfile
    with zipfile.ZipFile(TEMPLATE) as zf:
        pres = zf.read("ppt/presentation.xml").decode("utf-8")
        rels = zf.read("ppt/_rels/presentation.xml.rels").decode("utf-8")
        targets = {re.search(r'Id="([^"]+)"', r).group(1): re.search(r'Target="([^"]+)"', r).group(1)
                   for r in re.findall(r"<Relationship\b[^>]*>", rels)}
        samples = {}
        for rid in re.findall(r'<p:sldId\b[^>]*r:id="([^"]+)"', pres):
            name = targets[rid].rsplit("/", 1)[-1]
            head = zf.read(f"ppt/slides/{name}").decode("utf-8")[:4000]
            hit = re.search(r'<p:cSld\b[^>]*\bname="([^"]*)"', head)
            if hit and hit.group(1):
                samples[hit.group(1)] = name
    return samples


def sample_of(s):
    """The template slide a spec slide is cloned from: the cover's page for the cover and the closing slide,
    the divider for sections, the white content page for everything else."""
    kind = s["type"]
    if kind in ("cover", "closing"):
        return f"{THEME_NAME}/cover"
    if kind == "section":
        return f"{THEME_NAME}/section"
    return f"{THEME_NAME}/content"


def luminance(hex_color):
    def channel(v):
        v /= 255
        return v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4
    r, g, b = (int(hex_color[i:i + 2], 16) for i in (0, 2, 4))
    return 0.2126 * channel(r) + 0.7152 * channel(g) + 0.0722 * channel(b)


def contrast(a, b):
    la, lb = sorted((luminance(a), luminance(b)), reverse=True)
    return (la + 0.05) / (lb + 0.05)


def check_colors(s):
    """Warn about colors a spec fixes that would barely show on this theme's background."""
    where = f"slide '{s.get('title')}'"
    bg = P["bg"]
    for m in s.get("media", []):
        for ser in ((m.get("chart") or {}).get("series", []) if isinstance(m, dict) else []):
            color = resolve_color(ser.get("color"))
            if color and HEX_RE.fullmatch(color) and contrast(color, bg) < 1.5:
                warn(f"{where}: series {ser.get('name')!r} color {color} barely shows on the {THEME_NAME} theme's "
                     "white page — drop \"color\" to use the theme's palette, or use \"muted\" / \"accent\"")
    texts = [b for b in s.get("bullets", []) if isinstance(b, dict)]
    table = s.get("table") or {}
    texts += [c for row in [table.get("header", [])] + list(table.get("rows", [])) for c in row if isinstance(c, dict)]
    for item in texts:
        color = resolve_color(item.get("color"))
        back = resolve_color(item.get("fill")) or bg
        if color and HEX_RE.fullmatch(color) and HEX_RE.fullmatch(back) and contrast(color, back) < 3:
            warn(f"{where}: text color {color} is hard to read on {back} in the {THEME_NAME} theme — "
                 "drop \"color\" or pick a stronger one")


def default_output(spec_path, cfg, cover):
    """A spec inside a YYMMDDGroupMeeting/ folder builds into that folder; otherwise the cover date picks it."""
    day = (cover or {}).get("_day") or datetime.date.today()
    folder = spec_path.parent.name
    if re.fullmatch(r"\d{6}GroupMeeting", folder):
        if cover and cover.get("_day") and day.strftime("%y%m%d") != folder[:6]:
            warn(f"the cover says {day:%Y.%m.%d} but the spec sits in {folder}/ — the deck goes next to the spec; "
                 "check which date is right")
        return spec_path.parent / f"{folder}.pptx"
    code = day.strftime("%y%m%d")
    root = Path(os.path.expanduser(cfg.get("output_root", "~/Documents/PairsLab")))
    return root / f"{code}GroupMeeting" / f"{code}GroupMeeting.pptx"


def main():
    global METRICS
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("spec", help="deck spec (JSON)")
    ap.add_argument("-o", "--output", help="output .pptx (default: spec 'output', else <output_root>/<YYMMDD>GroupMeeting/)")
    ap.add_argument("--theme", help=f"build in this theme ({', '.join(THEMES)}), whatever the spec and config say")
    ap.add_argument("--no-validate", action="store_true", help="skip the final check of the written deck")
    deck_structure.add_arguments(ap)
    args = ap.parse_args()

    METRICS = Metrics()
    spec_path = Path(args.spec).resolve()
    spec = json.loads(spec_path.read_text(encoding="utf-8"))
    cfg = load_config()
    STYLE.update(cfg.get("style", {}))
    use_theme(args.theme or spec.get("theme") or cfg.get("theme") or DEFAULT_THEME)
    planned = expand_slides(spec, cfg)
    cover = next((s for s in planned if s["type"] == "cover"), None)

    if args.output:
        out = Path(os.path.expanduser(args.output)).resolve()
    elif spec.get("output"):
        out = Path(os.path.expanduser(spec["output"]))
        out = out if out.is_absolute() else (spec_path.parent / out).resolve()
    else:
        out = default_output(spec_path, cfg, cover)
    out.parent.mkdir(parents=True, exist_ok=True)

    samples = template_samples()
    missing = sorted({sample_of(s) for s in planned} - set(samples))
    if missing:
        raise SystemExit(f"the template has no {', '.join(missing)} slide — regenerate it with tools/make_template.py")

    # Build beside the deck and move it into place only once it is filled and checked, so a build that
    # stops halfway never replaces the last good deck.
    partial = out.with_name(f".{out.stem}.building{out.suffix}")
    try:
        deck_structure.build(args, TEMPLATE, [samples[sample_of(s)] for s in planned], tuple(samples.values()),
                             partial)
        prs = Presentation(str(partial))
        with tempfile.TemporaryDirectory(prefix="gm_media_") as tmp:
            work = Path(tmp)
            label = None                  # the part of the deck a slide belongs to, from the last section slide
            for number, (slide, s) in enumerate(zip(prs.slides, planned), 1):
                kind = s["type"]
                if kind == "cover":
                    fill_cover(slide, s, spec_path.parent, work)
                elif kind == "closing":
                    fill_closing(slide, s, spec_path.parent, work)
                elif kind == "section":
                    fill_section(slide, s)
                    label = s["title"]
                elif kind in ("done", "todo", "agenda"):
                    label = None              # these frame the whole deck rather than belong to one part
                    fill_agenda(slide, s) if kind == "agenda" else fill_progress(slide, s, kind)
                elif kind == "statement":
                    fill_statement(slide, s)
                elif kind == "columns":
                    fill_columns(slide, s, spec_path.parent, work)
                elif kind == "timeline":
                    fill_timeline(slide, s)
                else:
                    check_colors(s)
                    fill_content(slide, s, spec_path.parent, work)
                fill_footer(slide, label, number)
                if s.get("notes"):
                    notes = s["notes"]
                    slide.notes_slide.notes_text_frame.text = (
                        "\n".join(map(str, notes)) if isinstance(notes, list) else str(notes))
            prs.save(str(partial))
        ok, report = (True, "") if args.no_validate else deck_structure.validate(args, partial, TEMPLATE)
        if ok:
            os.replace(partial, out)
    finally:
        partial.unlink(missing_ok=True)

    for family in sorted(METRICS.missing):
        warn(f"font {family} not found locally — text fitting used estimated widths; check the render")
    if ok:
        print(f"Wrote {out} ({len(planned)} slides, {THEME_NAME} theme)")
    else:
        kept = "; the deck already there is unchanged" if out.exists() else ""
        print(f"Did not write {out}: the built deck failed the check below{kept}")
    for i, s in enumerate(planned, 1):
        layout = f", {s['layout']}" if s["type"] == "content" and s.get("layout", "auto") != "auto" else ""
        print(f"  {i:2d}. [{s['type']}{layout}] {s.get('title') or 'Group Meeting'}")
    show_warnings()
    if report:
        print(report)
    if not ok:
        sys.exit(1)


if __name__ == "__main__":
    try:
        main()
    finally:
        # A failure often follows from a typo a warning names ("did you mean 'path'?"), so always show them.
        show_warnings()
