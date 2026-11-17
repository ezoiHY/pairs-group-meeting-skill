# Deck spec format

`scripts/build_deck.py` reads one JSON file. Relative paths in it (media, `output`) resolve
against the spec file's own folder, so keep the spec next to the deck it builds.

## Top level

| Key | Required | Meaning |
|---|---|---|
| `slides` | yes | Ordered list of slides (below). |
| `theme` | no | The deck's look: `classic` (the lab template, for group meetings) or `paper` (for formal talks) — see [Themes](#themes). Default: `theme` in the user's config, else `config.json` (`classic`). `--theme` on the command line wins over all of them. |
| `output` | no | Where to write the .pptx. Default: next to the spec as `<folder>.pptx` when the spec sits in a `YYMMDDGroupMeeting/` folder (the build warns if the cover date disagrees); otherwise `<output_root>/<YYMMDD>GroupMeeting/<YYMMDD>GroupMeeting.pptx`, with `YYMMDD` from the cover date and `output_root` from `config.json`. `-o` on the command line wins over both. |
| `cover` | no | `false` to build without a cover. Otherwise a cover is added automatically if `slides` has none. |

Every slide takes an optional `"notes"`: speaker notes, plain text.

## Themes

Same spec, two looks. Each is a set of sample slides in `assets/template.pptx` (art, logos,
title style) plus colors and fonts in `build_deck.py` → `THEMES`; nothing else in the spec changes.
In both, every page that carries content is white; only the cover (and closing slide) and the
section dividers have art.

| Theme | Pages | Type | For |
|---|---|---|---|
| `classic` | The lab template: navy header band with radar rings, white pages; sections on navy | Calibri | Group meetings — the default |
| `paper` | White editorial pages, the logo small at the top, ruled (booktabs) tables; the lab's rings on cover and sections, a photo panel on the cover | Georgia titles and headings, Calibri text | Formal talks: defenses, conference and invited talks, reviews |

A spec or config that still names the retired `midnight` theme builds in `classic` with a
warning, and a slide's old `"tone"` key is ignored with a warning.

Content slides carry a footer: the slide number, and the name of the part of the deck they belong
to — the title of the last `section` slide (cleared again by a Done, Todo or agenda slide).

The build warns about every key it doesn't know — anywhere in the spec, with the closest valid
name — and about `layout` / `valign` values outside the lists below. Such keys would otherwise be
ignored, so a typo like `"bulets"` never silently drops content. Also: `null` means "not given"; a
string where a list belongs (`"bullets": "…"`) is used as a one-item list, and a list where one line
of text belongs (`"problem": ["…", "…"]`) is joined into one line, both with a warning; a number
where text belongs (a year as a series name) is used as text; a table row given as a string is a
one-cell row, with a warning; colors are six hex digits, with or without a leading `#`, or a word
that follows the slide's palette — `primary`, `secondary`, `accent` (gold), `muted`, `text`,
`strong` — so the color stays right in every theme; a fixed color that would barely show
on the slide's background is warned about; and `"cover": {…}` with a date or name becomes a cover
slide. A media item, chart, series, item or
bullet missing the key it can't do without (`path`, `categories`, `values`, `text`…) stops the
build. When the build stops on an error, the warnings so far are printed too, and the deck
already at the output path is left as it was.

## `cover`

```json
{"type": "cover", "date": "2026.10.29", "name": "Your Name"}
```

| Key | Default | Notes |
|---|---|---|
| `title` | `Group Meeting` | Shrinks to fit one line, from 48pt (paper: 60pt). |
| `date` | today | Any of `2026.10.29`, `2026-10-29`, `20261029`, `261029`; written as `2026.10.29`. |
| `name` | `config.json` → `presenter` | Empty name → the line is dropped and the build warns. |
| `extra_lines` | `[]` | More subtitle lines, e.g. `["PAIRS Lab"]`. |
| `image` | none | A photo for the cover — the robot, the setup, a demo frame: a path, or `{"path", "focus": [x, y]}` with the point to keep in view (0–1 from the top left, default the middle). `paper` shows it as a full-height panel on the right; `classic` keeps the lab's cover and leaves it out with a warning. Pictures only. |

For a formal talk give the talk's own `title` and put the event or affiliation in `extra_lines`.

## `agenda`

The talk's parts, numbered, one per row between hairlines. Put it after the cover; to show where
the talk is, repeat it before each part with `current` set.

```json
{"type": "agenda", "current": 2}
```

| Key | Default | Notes |
|---|---|---|
| `items` | the deck's `section` titles, in order | The parts, as strings, one line each (32pt down to 22pt). Keep to six or fewer. Without items and without sections the slide is left out with a warning. |
| `current` | none | The part about to start: its number (1-based) or its title. It is set bold, with the gold bar in `classic`; the others are muted. |
| `title` | `Outline` | |
| `subtitle` | none | As on `content`. |

## `section`

A divider that opens one part of the deck — a project, or the paper reading. Its title also
becomes the footer label of every slide after it.

```json
{"type": "section", "title": "Vision-EMG Hand Pose", "subtitle": "Muscle signals fill in what the camera cannot see"}
```

| Key | Notes |
|---|---|
| `title` | Required. The part's name, one line (44pt down to 30pt; `paper` 48pt down to 32pt). |
| `subtitle` | Optional. One line on what the part covers, ≤ 14 words. |

## `done` and `todo`

The template's progress lists: bold `Project: <name>` headers, lettered items (`a.`, `b.` …),
and for `done` a colored status after an em dash.

```json
{"type": "done", "title": "Done last two weeks (10.15–10.29)", "projects": [
  {"name": "Vision-EMG Hand Pose", "items": [
    {"text": "Collected **1h+** synchronized RGB + sEMG data", "status": "done"},
    {"text": "Occlusion-aware fusion of vision and EMG", "status": "on-going"}
  ]}
]}
```

| Key | Default | Notes |
|---|---|---|
| `title` | `config.json` → `done_title` / `todo_title` (`Done last two weeks` / `Todo`) | |
| `projects[].name` | — | `**bold**` allowed. |
| `projects[].label` | `Project: ` | Prefix before the name; `""` to drop it. |
| `projects[].items[]` | — | A string, or `{"text", "status"}`. |
| `status` (done only) | none | `done` / `finished` / `complete(d)` / `完成` / `已完成` → green "done" (classic: theme accent6); `on-going` / `ongoing` / `wip` / `进行中` → red "on-going" (classic: `980000`), `in progress` as written. Any other text is shown as written in red, and must be English. |

Fitting: 28pt in the template's text box; if that overflows, the box widens to the full
slide width, then the font steps down to the house minimum (`style.min_font_size`, 22pt). Past
that the list is split across slides (`… (cont.)`) at the minimum and the build prints a warning —
usually a sign items should be shorter.

## `content`

The template's title-only slide, with any mix of subtitle, bullets, table, media (images, videos,
charts) and a footnote. All body text — subtitle, bullets, table, captions, chart labels — is at
least `style.min_font_size` (22pt); only the title (36pt) and the footnote (12pt) are outside it.

```json
{"type": "content", "title": "Occlusion results",
 "subtitle": "Fusion cuts occluded error by 36%",
 "bullets": ["Biggest gain where vision fails", "Non-occluded accuracy also improves"],
 "media": [{"chart": {"type": "bar", "categories": ["Occluded", "Non-occluded"],
                      "series": [{"name": "Vision only", "values": [14.9, 8.1]},
                                 {"name": "Vision + EMG", "values": [9.6, 7.1]}],
                      "number_format": "0.0\"°\""},
            "caption": "Joint angle MAE, lower is better"},
           {"path": "videos/demo.mp4", "caption": "Watch the index finger"}],
 "footnote": "Wang et al., CVPR 2025",
 "layout": "auto"}
```

| Key | Notes |
|---|---|
| `title` | Required. classic: 36pt white in the header band; widens, then shrinks to 24pt, to stay on one line. paper: 36pt Georgia at the top left, shrinking to 26pt. |
| `subtitle` | One line under the title (24pt, down to 22pt) — the slide's takeaway as a plain sentence, or a paper's full title. Bold in classic, Georgia italic in paper. |
| `bullets` | Strings or `{"text", "level": 0/1, "bullet": false, "bold": true, "color": "980000"}`. `bullet: false` makes a plain paragraph. `**bold**` inline. 28pt down to 22pt when alone, 24pt down to 22pt next to a visual. The build warns above `style.max_bullets_per_slide` (3) bullets or `style.max_words_per_bullet` (14) words. A number and its unit (`1 ms`, `60 Hz`, `9.6 °`) are never split across lines, here or anywhere else. |
| `table` | `header` (optional), `rows`, optional `col_widths` (ratios), `align` (`left`/`center`/`right` per column), `font_size`. Cells are strings or `{"text", "bold", "color", "fill"}`. Font: 24pt or 22pt — the larger one when it fits without wrapping more cells, and never larger than the slide's bullets. At 22pt only ~4–6 short rows fit; for more numbers use a `chart`. Columns are sized for that font; a narrow table keeps its natural width and is centered. `font_size` forces a size (columns still sized for it) and warns if cells then wrap; it is never below the house minimum. `align` may also be one word for every column; `align` and `col_widths` should have one entry per column (the build warns otherwise). A table with no `rows` is left out with a warning. classic draws a filled header and hairlines between rows, paper draws booktabs rules (above, under the header, below); no vertical lines in either. |
| `media` | Strings (paths), `{"path", "caption", "poster", "page", "focus", "crop"}`, or `{"chart": {…}, "caption", "aspect"}`. `focus` and `crop` matter only in the `split` layouts. Images: png, jpg, gif, bmp, tif, webp, and on macOS heic; the format is read from the file, and a rotated phone photo is turned upright (and stays a JPEG). `pdf` and `svg` are rasterized at 220 dpi (`page` picks a PDF page, default 1). Videos: mp4, m4v, mov, avi, wmv — embedded and playable in PowerPoint; the poster frame comes from macOS Quick Look unless `poster` names an image. Captions are 22pt, centered under the item — keep them to ≤ 8 words. |
| `media[].chart` | A native, editable PowerPoint chart, laid out like an image (default `aspect` 1.6). `type`: `bar` (columns), `hbar` (first category on top), `line`. `categories`: x labels. `series`: `[{"name", "values", "color"?}]` — a legend appears for 2+ series; leave `color` out to use the theme's palette, or give `"muted"` for a baseline / "before" series. `number_format`: Excel format for labels, e.g. `"0.0"`, `"0%"`, `"0.0\"°\""`. `labels` (default on for bars, off for lines): value labels, which also hide the value axis. `highlight`: category names drawn in the theme's accent, gold (single-series bars) — use it for "ours". `y_min` / `y_max`: axis range. |
| `media_rows` | Force the number of media rows. Default: the row count that makes the media largest. |
| `layout` | `auto` (default), `text-left`, `text-right`, `stack`, `fit`, `split`, `split-left`. `auto` = `text-left` when there are bullets and media but no table, else `stack`. `fit` = `stack` when every media item is wide (width/height ≥ 2.6), else `auto` — for figure-led slides such as a paper's pipeline. `split` / `split-left`: one photo or video fills the right / left half of the slide out to its edge, under the header, and the subtitle and bullets sit beside it, centered — for the robot, the setup, a demo. A photo is cropped to fill the half around its `focus` (`[x, y]`, 0–1 from the top left); a video keeps its whole frame, as wide as the half; `"crop": true` / `false` in the media item switches either. Only the first photo or video is used; a chart or table falls back to `auto` with a warning. A caption goes at the bottom of the text column. |
| `text_width` | Bullet column share for `text-left` / `text-right`, between `0.15` and `0.75` (default `0.38`). |
| `footnote` | Small muted 12pt line at the bottom left, in the footer — citations, data sources only. |
| `text_only` | `true` silences the "no figure, demo video or chart" warning, for a genuine discussion slide. |
| `valign` | `top` (default) or `middle`: center the whole block (subtitle, bullets, table, media) in the content area. For short slides without media that would otherwise leave an empty band at the bottom. The footnote stays put. |

How `stack` places things, top to bottom: subtitle → bullets (≤45% of the height when other
blocks follow) → table (≤55% when media follow) → media in the remaining space. Media alone are
centered vertically; otherwise they start right under the block above. In `text-left` /
`text-right`, bullets get their own column and the table and media stack in the other column.

## `statement`

One message set large — the key result, the research question, the takeaway — with an optional
number it is about and one line of detail.

```json
{"type": "statement", "title": "Key result", "number": "36%",
 "text": "Lower occluded error when EMG joins the camera",
 "detail": "Joint-angle MAE 14.9° → 9.6°, three subjects"}
```

| Key | Notes |
|---|---|
| `title` | Required, as on `content`. |
| `text` | Required. One sentence of ≤ 14 words, up to three lines: 44pt down to 22pt alone, 36pt beside a number. |
| `number` | Optional. A short figure that is the message (`36%`, `9.6°`, `3×`): 132pt down to 72pt on the left, with the gold bar (`classic`) or a hairline (`paper`) between it and the text. |
| `detail` | Optional supporting line under the text, 24pt muted. |
| `footnote` | As on `content`. |

Without a number, `classic` puts the lab's gold bar beside the text and `paper` a short rule above
it. Use a statement for the one or two messages a talk must leave behind, not for every result.

## `columns`

Two to four points side by side — contributions, challenges, the parts of a method, the
conditions of an experiment — each a heading and a line or two, optionally under a visual.
Hairlines separate the columns; there are no boxes.

```json
{"type": "columns", "title": "What makes fusion work", "subtitle": "Three choices carry most of the gain",
 "columns": [
   {"heading": "Per-joint weights", "text": "Each joint trusts the source that sees it best"},
   {"heading": "Visibility gate", "text": "Camera confidence decides when EMG takes over"},
   {"heading": "One clock", "text": "A hardware trigger keeps video and EMG within 1 ms"}]}
```

| Key | Notes |
|---|---|
| `columns` | Required, 2–4 (more are cut with a warning). Each: `heading` (≤ 6 words), `text` (a string or a list of up to two lines, ≤ 14 words each), `media` (optional: one image, video or chart, as in `content`; its caption is left out — the heading says what it shows). |
| `title`, `subtitle`, `footnote` | As on `content`. |

All columns share one heading size and one text size. Text-only columns sit in the middle of the
space under the subtitle, at up to 32pt / 28pt; columns with visuals start right under it, the
visuals on top in one row. A size that would leave a single word alone on a line is avoided.

## `timeline`

Milestones or steps along one line: a date or step name above each marker, a line or two under it.

```json
{"type": "timeline", "title": "Road to the paper",
 "subtitle": "Two subjects and the ablations stand between us and a draft",
 "points": [{"label": "11.05", "text": "Data from 3 subjects", "status": "done"},
            {"label": "12.03", "text": "Last 2 subjects recorded", "status": "now"},
            {"label": "01.07", "text": "Full paper draft", "status": "next"}]}
```

| Key | Notes |
|---|---|
| `points` | Required, 2–6 (more are cut with a warning). Each: `label` (a date or a step's name), `text` (≤ `style.max_words_per_item` words), `status` (optional). A plain string is a point with only a label. |
| `status` | `done` (filled marker; `完成` and the other Done words work too), `now` (gold, ringed; also `current`, `on-going`), `next` (open circle, muted text; also `planned`, `todo`). The line is drawn in full up to the last `done` / `now` point. Without statuses every marker is filled. |
| `title`, `subtitle`, `footnote` | As on `content`. |

## `paper`

One paper, expanded by the builder into two slides: a `content` slide with the paper's pipeline
figure and its problem and approach, then a two-column `columns` slide, Novelty | Limitations.

```json
{"type": "paper", "title": "Paper A", "venue": "CVPR 2025",
 "full_title": "Paper A: Occlusion-Aware Hand Pose Estimation with Temporal Priors",
 "pipeline": "figs/paper_a_fig2.png",
 "problem": "Hand pose breaks when fingers are occluded",
 "approach": "Carries the previous pose forward as a prior",
 "novelty": ["Temporal prior gated by joint visibility", "Needs no extra sensors"],
 "limitations": ["Drifts after long occlusions", "One camera setup only"],
 "takeaway": "A cheap fix for short occlusions, not long ones",
 "notes": ["Spoken explanation for slide 1", "Spoken explanation for slide 2"]}
```

| Key | Notes |
|---|---|
| `title` | Required. The paper's short name; slide titles are `Paper reading: <title>` and `<title>: novelty and limits`. |
| `full_title` | Subtitle of the first slide. |
| `venue` | Footnote of the first slide, followed by `· pipeline figure from the paper`. |
| `pipeline` | The authors' original pipeline figure: any `media` path (image, or `pdf` / `svg` rasterized). Required — the build warns without it. |
| `pipeline_page` | For a PDF `pipeline`, the page the figure is on (default 1). `pipeline` may also be a `media` object, `{"path", "page"}`. |
| `pipeline_caption` | Optional caption, ≤ 8 words. |
| `problem`, `approach` | One short line each, shown as `**Problem:** …` and `**Approach:** …` beside (or above) the figure. Required. |
| `novelty`, `limitations` | A string or a list; at most 2 points of ≤ `style.max_words_per_bullet` words each. Each fills one column of the second slide, under its heading. Required; a missing one drops its column. |
| `takeaway` | Optional subtitle of the second slide. |
| `notes` | A string (first slide) or a list of up to two strings (one per slide). |
| `layout`, `text_width` | Override the first slide's layout (default `fit`, see `content`). |

Missing fields are never invented: the build warns `missing …`, and the slides come out without them.

## `compare`

One theme's papers, expanded into two `content` slides that hold only tables: how the methods
differ, then how the results differ. Methods come first; compare them before the numbers.

```json
{"type": "compare", "theme": "Occlusion-robust hand pose",
 "methods": {"header": ["Paper", "Input", "Key idea", "Extra cost"],
             "rows": [["Paper A", "RGB video", "Temporal prior", "None"],
                      ["Paper B", "RGB + EMG", "Per-joint fusion", "Armband"]],
             "takeaway": "Methods differ mainly in the extra signal they bring"},
 "results": {"header": ["Paper", "MAE ↓", "Occluded MAE ↓", "FPS ↑"],
             "rows": [["Paper A", "8.1°", "12.4°", "60"], ["Paper B", "7.1°", "9.6°", "45"]],
             "takeaway": "Extra sensing helps most when fingers are hidden",
             "note": "Occluded subset of <benchmark>; numbers as reported in each paper"},
 "notes": ["Spoken, methods slide", "Spoken, results slide"]}
```

| Key | Notes |
|---|---|
| `theme` | Required. Slide titles are `Compare: <theme> — methods` and `Compare: <theme> — results`. |
| `methods`, `results` | Each a `content` table — `header`, `rows`, optional `col_widths`, `align`, `font_size` (see `content`) — plus `takeaway` (the slide's subtitle). `methods` is required; without `results` the build warns. One row per paper, the paper's name in the first column. |
| `results.note` | The slide's footnote: the benchmark and setting, and whether numbers are as reported. The build warns without it. |
| `results.better` | Optional per-column `"lower"` / `"higher"` / `null`, overriding the header's arrows — or one of those words for every column after the paper names. By default a `↓` in the header means lower is better, `↑` higher. |
| `notes` | A string (methods slide) or a list of up to two strings (methods, results). |

The best value of each column that has a `↓` / `↑` (or a `better` entry) is bolded automatically,
ties included; a cell is read by its first number (`7.1°` → 7.1), and cells with no number are
skipped. Methods cells are left-aligned; results cells are centered except the first column.
The build warns above `style.max_compare_papers` rows (default 5) or `style.max_compare_columns`
columns (default 5) — split the theme or drop columns. There are no figures here; use `content`
for those.

## `closing`

The last slide of a formal talk, on the cover's page.

```json
{"type": "closing", "lines": ["Questions welcome", "yu.han@example.edu"]}
```

| Key | Default | Notes |
|---|---|---|
| `title` | `Thank you` | Fitted like the cover title. |
| `lines` | none | Lines under the title, like the cover's date and name. |
| `image` | none | As on `cover` (`paper` only). |

## A complete example

```json
{
  "theme": "classic",
  "slides": [
    {"type": "cover", "date": "2026.10.29", "name": "Your Name"},
    {"type": "done", "projects": [
      {"name": "Vision-EMG", "items": [
        {"text": "Collected 1h+ synchronized data", "status": "done"},
        {"text": "Fusion model training", "status": "on-going"}]}]},
    {"type": "section", "title": "Vision-EMG Hand Pose", "subtitle": "Muscle signals fill in occluded joints"},
    {"type": "content", "title": "Vision-EMG results",
     "subtitle": "Fusion beats vision-only under occlusion",
     "table": {"header": ["Model", "MAE ↓", "Occluded MAE ↓"],
               "rows": [["Vision", "8.10°", "14.9°"], ["Vision + EMG", "**7.08°**", "**9.6°**"]]},
     "media": [{"path": "figs/mae_curve.png", "caption": "Validation MAE per epoch"}]},
    {"type": "content", "title": "Demo: occluded grasp", "layout": "split",
     "subtitle": "The fused hand keeps all five fingers", "bullets": ["Watch the index finger"],
     "media": ["videos/fusion_demo.mp4"]},
    {"type": "paper", "title": "Paper A", "venue": "CVPR 2025",
     "full_title": "Paper A: Occlusion-Aware Hand Pose Estimation with Temporal Priors",
     "pipeline": "figs/paper_a_fig2.png",
     "problem": "Hand pose breaks when fingers are occluded",
     "approach": "Carries the previous pose forward as a prior",
     "novelty": ["Temporal prior gated by joint visibility"],
     "limitations": ["Drifts after long occlusions"]},
    {"type": "todo", "projects": [
      {"name": "Vision-EMG", "items": ["Multi-subject collection by 11.05", "First paper draft"]}]}
  ]
}
```
