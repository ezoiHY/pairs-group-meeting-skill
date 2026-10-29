---
name: pairs-group-meeting
description: Make PAIRS Lab (HKUST-GZ) group-meeting slides — 组会PPT / 组会汇报 — as a .pptx built on the lab's own template, with the cover, "Done last two weeks" progress list with done/on-going status, results slides with figures, demo videos and tables, paper-reading slides (the paper's own pipeline figure plus problem, approach, novelty, limitations), multi-paper comparison tables, and the Todo list. Use this whenever the user wants slides for a group meeting, lab meeting, weekly or biweekly progress report, or 组会, or asks to turn progress notes, experiment results, plots or videos into meeting slides, even if they don't mention the template or PowerPoint. It builds and checks the file itself with python-pptx; no other skill is needed.
---

# PAIRS Lab group-meeting deck (Codex version)

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

## Step 0 — Set up

1. **Python**: the scripts need Python ≥ 3.10 with python-pptx, lxml, defusedxml, Pillow and PyMuPDF
   (`pip install -r <this skill>/requirements.txt`). `$SKILL_PY` below is that interpreter. If the
   packages are missing, install them in a venv — `python3 -m venv ~/.venvs/pairs-gm`, then
   `~/.venvs/pairs-gm/bin/pip install -r <this skill>/requirements.txt` — and if the sandbox blocks
   the network, ask the user to approve the install.
2. **Shell variables don't survive between commands**, so start every command by setting
   `SKILL_PY="<python>"; …`.
3. **Config**: read `config.json`. If `presenter` is empty, take the name from the user's message
   or ask for it, save it there, and tell the user you saved it — later decks then won't ask.

## Step 1 — Gather the content

What a group-meeting deck needs, and what to ask for if the user hasn't said:

- **Date** of the meeting (default: today) — it names the output folder `YYMMDDGroupMeeting/`.
  If the user gives both a weekday and a date ("下周三 11.12") and they disagree, ask which is right.
- **Done**: what was finished or advanced since the last meeting, grouped by project, each item
  marked done or on-going.
- **Results**: the figures, plots and demo videos to show — with file paths — and the one takeaway
  each should land. Visuals carry these decks, so for every result ask for a plot, screenshot or
  demo video if the user didn't give one. Numbers alone can become a native `chart`.
- **Todo**: next steps per project, with deadlines if known.
- **Papers they read** (optional; one block per paper). A paper gets slides only when the user
  supplies all of this — ask for whatever is missing, and never fill a gap from the abstract or
  your own knowledge, least of all the limitations:
  1. Short name, full title, venue and year.
  2. The paper's **original pipeline figure**: the authors' own overview/method figure, as an image
     or as a PDF plus page number. Don't redraw it or swap in a diagram you made.
  3. **Problem** — what the paper sets out to solve.
  4. **Approach** — how it solves it.
  5. **Novelty** — what is new compared with earlier work.
  6. **Limitations** — where it falls short. This is the presenter's own judgment; if they have
     none yet, ask what would break it (extra sensors, one dataset, calibration, speed…).

  Answers may come loosely, in 中文. Compress each to short English — Problem and Approach one
  line each, Novelty and Limitations at most two points each, every line ≤ 14 words — and put the
  fuller explanation in the slide's speaker `notes`.
- **Comparing papers** (when they read two or more): group the papers into themes of 3–5 and
  ask for, per theme: a theme name; the few method dimensions that actually separate the papers
  (input or data, key idea, supervision, extra cost…); and each paper's numbers on one shared
  benchmark and metric, with where they come from (as reported in the paper, or re-run). Compare
  like with like: if two papers report on different benchmarks or settings, say so or leave the
  cell as `—`. Never invent a number.

**Look at the previous deck.** Find the newest earlier deck in `output_root` (from `config.json`):
usually a `YYMMDDGroupMeeting/` folder, but older decks may be named differently (`260930组会.pptx`).
Read it with `$SKILL_PY <this skill>/scripts/slide_text.py <previous deck>` and find its Todo
slide — older decks may title it differently ("To-Do List for the Next 7 Days"). Last meeting's
Todo is this meeting's Done candidate list: if the user hasn't said what they did, propose it and
let them mark each item done or on-going; if they have, use it only to ask about items they left
out. Its date is also the start of the Done period.

Ask everything missing in one message rather than one question at a time. If the user hands you
loose notes (often in 中文), extract the structure yourself, write it in English, and confirm only
what is ambiguous.

## Step 2 — Plan the deck

The lab's standard order, which the template is built around:

1. **Cover** — "Group Meeting", date, name.
2. **Done** — `Done last two weeks`. Add the period, `Done last two weeks (10.15–10.29)`, only when
   you know the previous meeting's date (from the previous deck or the user); don't invent one.
3. **Results / progress** — one message per slide, as many as the work needs.
4. **Paper reading** (optional) — one `paper` entry per paper, which becomes two slides:
   `Paper reading: <short name>` (full title as subtitle, the original pipeline figure, Problem and
   Approach, venue as footnote) and `<short name>: novelty and limits` (two columns).
5. **Paper comparison** (only with two or more papers, after the papers themselves) — one `compare`
   entry per theme, which becomes `Compare: <theme> — methods`, then `Compare: <theme> — results`.
   Methods first, results second, tables only.
6. **Todo**.

## House style (the lab's defaults — follow them)

`config.json` → `style` holds the numbers below; the builder enforces the font floor and warns on
the rest.

- **All English.** Every title, item, bullet, subtitle, caption, table cell and chart label — even
  when the user writes their notes in Chinese. Translate into plain, natural English; keep names
  and technical terms (MAE, TCN, HaMeR) as they are. The build warns on any Chinese text.
- **Body text is at least 22pt** (`min_font_size`). The builder never shrinks text below it; when
  something doesn't fit it warns instead — then cut words or split the slide, never shrink. Only
  titles (36pt) and the small source/citation footnote are outside this rule.
- **Few words, key sentences — but self-explanatory.** Someone who missed the talk should still
  get each slide:
  - `subtitle` = the takeaway as one plain sentence: "Fusion cuts occluded error by 36%".
  - At most **3 bullets** (`max_bullets_per_slide`), each one short line of **≤ 14 words**
    (`max_words_per_bullet`), saying *why* or *how* — not repeating the figure.
  - No paragraphs, no filler ("We can see that…", "As shown in the figure…").
- **Show it.** Every results slide carries a figure, demo video or chart; prefer two visuals and two
  bullets over one visual and four bullets.
  - Numbers → a native `chart` (bar / line, values labeled, the winner highlighted) rather than a
    table; keep tables for at most ~4 rows of exact values.
  - Demos → embedded video with a caption saying what to watch for: "Watch the index finger during
    the grasp".
  - Figure captions (≤ 8 words, also 22pt) say what to look at or what the axes are — never repeat
    the title. Skip them when the title already says it.
- **Light slides.** One message per slide. If a slide needs more than 3 bullets, or more than two
  visuals plus text, split it into two slides.
- **Done/Todo items**: start with a verb, ≤ 10 words, one line each; `**bold**` one key number if
  it matters. Detail belongs on a results slide.
- **Titles**: short, no period: `Occlusion results`, `Demo: HaMeR vs. WiLoR`, `Acquisition pipeline`.
- **Paper slides**: the figure is the authors' original, credited in the footnote. Problem and
  Approach are the only bullets beside it; Novelty and Limitations get their own slide, at most two
  points each. A very wide figure (≥ 2.6:1) takes the full width under the bullets.
- **Comparison slides** are the exception to "show it": tables only, no figures or charts. At most
  5 papers (`max_compare_papers`) and 5 columns (`max_compare_columns`) per theme — paper name in
  the first column, then only the dimensions that separate the papers. Results cells hold just the
  number; put ↓ or ↑ in the column header and the best value is bolded for you. The `note` under
  the results table names the benchmark and whether the numbers are as reported.
- A project with no visuals appears only on the Done and Todo slides — no text-only slide for it.
- Roughly 1.5–2 minutes per slide; a typical deck is 8–14 light slides.

## Step 3 — Write the spec

Write `deck_spec.json` in the meeting folder (`<output_root>/<YYMMDD>GroupMeeting/`), so the user
can tweak and rebuild it later. Read `references/spec-format.md` for every field. In short:

```json
{"slides": [
  {"type": "cover", "date": "2026.10.29"},
  {"type": "done", "projects": [{"name": "Vision-EMG", "items": [
     {"text": "Collected 1h+ synchronized data", "status": "done"},
     {"text": "Fusion model training", "status": "on-going"}]}]},
  {"type": "content", "title": "Occlusion results", "subtitle": "Fusion cuts occluded error by 36%",
   "bullets": ["Biggest gain where vision fails", "Non-occluded accuracy also improves"],
   "media": [{"chart": {"type": "bar", "categories": ["Occluded", "Non-occluded"],
                        "series": [{"name": "Vision only", "values": [14.9, 8.1]},
                                   {"name": "Vision + EMG", "values": [9.6, 7.1]}],
                        "number_format": "0.0\"°\""},
              "caption": "Joint angle MAE, lower is better"}]},
  {"type": "content", "title": "Demo: fusion under occlusion",
   "subtitle": "The fused hand keeps all five fingers when vision loses them",
   "media": [{"path": "/abs/path/demo.mp4", "caption": "Watch the index finger during the grasp"}]},
  {"type": "todo", "projects": [{"name": "Vision-EMG", "items": ["Multi-subject collection by 11.05"]}]}
]}
```

Papers and their comparison use two dedicated slide types; the builder expands each into the
slides described in step 2 (all fields in `references/spec-format.md`):

```json
{"type": "paper", "title": "Paper A", "venue": "CVPR 2025",
 "full_title": "Paper A: Occlusion-Aware Hand Pose Estimation",
 "pipeline": "/abs/path/paper_a_fig2.png",
 "problem": "Hand pose breaks when fingers are occluded",
 "approach": "Carries the previous pose forward as a prior",
 "novelty": ["Temporal prior gated by joint visibility"],
 "limitations": ["Drifts after long occlusions", "One camera setup only"],
 "notes": ["Fuller spoken explanation for the first slide", "…for the second"]},
{"type": "compare", "theme": "Occlusion-robust hand pose",
 "methods": {"header": ["Paper", "Input", "Key idea"],
             "rows": [["Paper A", "RGB video", "Temporal prior"], ["Paper B", "RGB + EMG", "Per-joint fusion"]],
             "takeaway": "Methods differ mainly in the extra signal they bring"},
 "results": {"header": ["Paper", "MAE ↓", "Occluded MAE ↓"],
             "rows": [["Paper A", "8.1°", "12.4°"], ["Paper B", "7.1°", "9.6°"]],
             "takeaway": "Extra sensing helps most when fingers are hidden",
             "note": "Occluded subset of <benchmark>; numbers as reported"}}
```

Use absolute media paths unless the files sit next to the spec. Don't copy the user's media
around; the builder embeds them into the .pptx. A slide whose block is short sits at the top of
the content area; give it `"valign": "middle"` when that leaves a large empty band below.

## Step 4 — Build

```bash
$SKILL_PY <this skill>/scripts/build_deck.py <meeting folder>/deck_spec.json
```

It clones the template's sample slides in spec order, drops the originals, fills every slide, and
checks that the written deck reopens with every relationship intact. It prints the slide list, then warnings. **Act on every warning**:

| Warning | What to do |
|---|---|
| list split into N slides | Usually items are too long: shorten them, or move detail to a results slide, and rebuild. Keep the split only if the content really is that long. |
| text overflows its area | Cut words or split the slide — the font never goes below 22pt. |
| table taller than the space | Turn it into a chart, or split the rows across two slides. |
| non-English text | Translate it. |
| N bullets / bullet over N words / item over N words | Cut to the key sentence, move detail to the speaker `notes`, or split the slide. |
| no figure, demo video or chart | Ask the user for a visual or make a chart from the numbers; `"text_only": true` only for a genuine discussion slide. |
| paper: missing the pipeline figure / problem / approach / novelty / limitations | Ask the user for it — don't write it yourself. |
| paper: N novelty or limitations / point over N words | Keep the two points that matter most, or shorten. |
| comparison: N papers / N columns | Split the theme, or drop columns that don't separate the papers. |
| comparison: no results table / no `note` | Ask for numbers on a shared benchmark, and for which benchmark and setting they come from. |
| no presenter name | Ask for it; set it in `config.json`. |
| no poster frame | Give the video a `"poster"` image. |

## Step 5 — QA

1. **Content**:
   `$SKILL_PY <this skill>/scripts/slide_text.py deck.pptx` — check order, typos, and that no
   template text survived: `… | grep -nE "xxx|XXX|2026\.XX|^Name$"` must print nothing.
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

## Step 6 — Slides the builder can't make

For a layout the spec can't express (a 2×2 comparison grid, an annotated diagram, a custom
timeline), build everything else first, with a placeholder `content` slide where it goes — only the
title and `"text_only": true` — so it keeps the template's background, header band and logos. Then
open the built deck with python-pptx in a short script and add the shapes to that slide, with the
tokens in `references/template-anatomy.md` (colors, Calibri + 等线, safe area, sizes); keep the
title placeholder as is and only change its text. Never add a slide from a layout: the template's
background and logos live on the slides, not the layouts. Save, re-render, inspect.

Rebuilding from the spec overwrites hand edits, so make them last (or keep the script and rerun it
after each rebuild), and tell the user which slides were hand-edited.

## Step 7 — Deliver

Tell the user where the deck is, list the slides in one line each, and mention anything you
assumed (date, statuses, which figure went where) so they can correct it. Point out that
`deck_spec.json` sits next to the deck for quick edits and rebuilds.
