#!/usr/bin/env python3
"""Bring the Codex version of the skill up to date with the Claude version.

The Claude version (skills/pairs-group-meeting) is the source. SHARED files are
copied unchanged. The Codex SKILL.md is generated from the Claude one: the parts
about the pptx skill and Claude's tools are swapped for the Codex text below, and
everything else (content to gather, deck plan, house style, spec, warnings) is
kept word for word. Files only one version has (deck_structure.py, slide_text.py,
config.json) are written by hand in each version.

Usage:
    python tools/sync_codex.py           update the Codex version
    python tools/sync_codex.py --check   exit 1 if it is out of date
"""

import filecmp
import re
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CLAUDE = ROOT / "skills" / "pairs-group-meeting"
CODEX = ROOT / "codex" / "pairs-group-meeting"
SHARED = [
    "scripts/build_deck.py",
    "scripts/render.py",
    "references/spec-format.md",
    "references/template-anatomy.md",
    "assets/template.pptx",
    "requirements.txt",
]

INTRO = """# PAIRS Lab group-meeting deck (Codex version)

This skill knows the lab's template and conventions, and makes the file itself with python-pptx —
no other skill is needed. `scripts/build_deck.py` turns a JSON spec into a deck whose every slide
is a clone of one of the template's sample slides, so every deck looks like the template without
hand-editing XML.

```
assets/template.pptx           the lab template (never edit it)
scripts/build_deck.py          spec.json → .pptx
scripts/deck_structure.py      clones the template's slides and checks the written file
scripts/slide_text.py          prints a deck's text, slide by slide
scripts/render.py              .pptx → slide PNGs + overview, via PowerPoint (or LibreOffice)
references/spec-format.md      the spec: slide types, every field, examples
references/template-anatomy.md geometry, colors, fonts, XML — for hand edits only
config.json                    presenter name, output folder, list titles, house style limits
requirements.txt               Python packages the scripts need
```

"""

STEP0 = """## Step 0 — Set up

1. **Python**: the scripts need Python ≥ 3.10 with python-pptx, lxml, defusedxml, Pillow and PyMuPDF
   (`pip install -r <this skill>/requirements.txt`). `$SKILL_PY` below is that interpreter. If the
   packages are missing, install them in a venv — `python3 -m venv ~/.venvs/pairs-gm`, then
   `~/.venvs/pairs-gm/bin/pip install -r <this skill>/requirements.txt` — and if the sandbox blocks
   the network, ask the user to approve the install.
2. **Shell variables don't survive between commands**, so start every command by setting
   `SKILL_PY="<python>"; …`.
3. **Config**: read `config.json`. If `presenter` is empty, take the name from the user's message
   or ask for it, save it there, and tell the user you saved it — later decks then won't ask.

"""

STEP5 = """## Step 5 — QA

1. **Content**:
   `$SKILL_PY <this skill>/scripts/slide_text.py deck.pptx` — check order, typos, and that no
   template text survived: `… | grep -nE "xxx|XXX|2026\\.XX|^Name$"` must print nothing.
2. **File**: the build already checked that the deck reopens with every relationship intact.
   Rebuild after any change to the spec.
3. **Visual**:
   ```bash
   $SKILL_PY <this skill>/scripts/render.py deck.pptx
   ```
   On macOS this opens the deck in PowerPoint for a few seconds (up to a minute with large videos)
   and exports exact renders; elsewhere it uses LibreOffice. PowerPoint is driven through
   AppleScript, which a sandbox may block — if the render fails with a permission or sandbox error,
   ask the user to approve running it outside the sandbox. It prints one PNG per slide and an
   overview sheet. Look at every slide image with fresh eyes, in this order:
   - text overflowing its box or the slide, or cut off at an edge
   - overlaps — text over a figure, a caption over the logos, a table over the footnote
   - crowding — elements closer than ~0.3", a slide that feels full
   - figures or video frames too small to read, or squeezed out of shape
   - low contrast, and leftover template text

   Fix by editing the spec and rebuilding, re-render, stop when clean.

"""

STEP6 = """## Step 6 — Slides the builder can't make

For a layout the spec can't express (a 2×2 comparison grid, an annotated diagram, a custom
timeline), build everything else first, with a placeholder `content` slide where it goes — only the
title and `"text_only": true` — so it keeps the template's background, header band and logos. Then
open the built deck with python-pptx in a short script and add the shapes to that slide, with the
tokens in `references/template-anatomy.md` (colors, Calibri + 等线, safe area, sizes); keep the
title placeholder as is and only change its text. Never add a slide from a layout: the template's
background and logos live on the slides, not the layouts. Save, re-render, inspect.

Rebuilding from the spec overwrites hand edits, so make them last (or keep the script and rerun it
after each rebuild), and tell the user which slides were hand-edited.

"""

# (Claude text, Codex text) — exact replacements inside the parts both versions share.
EDITS = [
    ("It drives the pptx skill for the actual file generation and QA.",
     "It builds and checks the file itself with python-pptx; no other skill is needed."),
    ('`$SKILL_PY "$PPTX_SKILL/scripts/slide_text.py" <previous deck>`',
     "`$SKILL_PY <this skill>/scripts/slide_text.py <previous deck>`"),
    ('build_deck.py <meeting folder>/deck_spec.json --pptx-skill "$PPTX_SKILL"',
     "build_deck.py <meeting folder>/deck_spec.json"),
    ("It clones the template's sample slides in spec order with the pptx skill's `add_slide.py`, drops\n"
     "the originals with `clean.py`, fills every slide, and runs the pptx skill's `validate.py` against the\n"
     "template.",
     "It clones the template's sample slides in spec order, drops the originals, fills every slide, and\n"
     "checks that the written deck reopens with every relationship intact."),
]
SECTIONS = {"## Step 0": STEP0, "## Step 5": STEP5, "## Step 6": STEP6}


def codex_skill_md():
    text = (CLAUDE / "SKILL.md").read_text(encoding="utf-8")
    for old, new in EDITS:
        if old not in text:
            raise SystemExit(f"skills/pairs-group-meeting/SKILL.md no longer contains {old[:70]!r} — "
                             "update EDITS in tools/sync_codex.py")
        text = text.replace(old, new)
    head, *sections = re.split(r"(?m)^(?=## )", text)
    frontmatter = head[:head.index("\n---\n", 4) + 5]
    out = frontmatter + "\n" + INTRO
    seen = set()
    for section in sections:
        key = next((k for k in SECTIONS if section.startswith(k)), None)
        out += SECTIONS[key] if key else section
        seen.add(key)
    missing = set(SECTIONS) - seen
    if missing:
        raise SystemExit(f"skills/pairs-group-meeting/SKILL.md has no {sorted(missing)} section — "
                         "update SECTIONS in tools/sync_codex.py")
    return out.rstrip("\n") + "\n"


def main():
    check = "--check" in sys.argv[1:]
    stale = [f for f in SHARED if not (CODEX / f).exists() or not filecmp.cmp(CLAUDE / f, CODEX / f, shallow=False)]
    skill_md = codex_skill_md()
    target = CODEX / "SKILL.md"
    if not target.exists() or target.read_text(encoding="utf-8") != skill_md:
        stale.append("SKILL.md")
    if check:
        for f in stale:
            print(f"out of date: {f}")
        print("Codex version in sync." if not stale else "Run python tools/sync_codex.py to update it.")
        sys.exit(1 if stale else 0)
    for f in stale:
        (CODEX / f).parent.mkdir(parents=True, exist_ok=True)
        if f == "SKILL.md":
            target.write_text(skill_md, encoding="utf-8")
        else:
            shutil.copy2(CLAUDE / f, CODEX / f)
        print(f"updated: {f}")
    if not stale:
        print("Codex version already in sync.")


if __name__ == "__main__":
    main()
