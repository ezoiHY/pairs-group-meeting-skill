# Deck spec format

`scripts/build_deck.py` reads one JSON file. Relative paths in it (media, `output`) resolve
against the spec file's own folder, so keep the spec next to the deck it builds.

## Top level

| Key | Required | Meaning |
|---|---|---|
| `slides` | yes | Ordered list of slides (below). |
| `output` | no | Where to write the .pptx. Default: `<output_root>/<YYMMDD>GroupMeeting/<YYMMDD>GroupMeeting.pptx`, with `YYMMDD` from the cover date and `output_root` from `config.json`. `-o` on the command line wins over both. |
| `cover` | no | `false` to build without a cover. Otherwise a cover is added automatically if `slides` has none. |

Every slide takes an optional `"notes"`: speaker notes, plain text.

## `cover`

```json
{"type": "cover", "date": "2026.10.29", "name": "Your Name"}
```

| Key | Default | Notes |
|---|---|---|
| `title` | `Group Meeting` | Shrinks from 48pt to fit one line. |
| `date` | today | Any of `2026.10.29`, `2026-10-29`, `20261029`, `261029`; written as `2026.10.29`. |
| `name` | `config.json` → `presenter` | Empty name → the line is dropped and the build warns. |
| `extra_lines` | `[]` | More subtitle lines, e.g. `["PAIRS Lab"]`. |

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
| `status` (done only) | none | `done` / `finished` / `完成` → green (theme accent6); `on-going` / `ongoing` / `in progress` / `wip` / `进行中` → dark red `980000`. Any other text is shown as written in dark red. |

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
| `title` | Required. 36pt white in the header band; widens, then shrinks to 24pt, to stay on one line. |
| `subtitle` | One bold line under the header (24pt, down to 22pt) — the slide's takeaway as a plain sentence, or a paper's full title. |
| `bullets` | Strings or `{"text", "level": 0/1, "bullet": false, "bold": true, "color": "980000"}`. `bullet: false` makes a plain paragraph. `**bold**` inline. 28pt down to 22pt when alone, 24pt down to 22pt next to a visual. The build warns above `style.max_bullets_per_slide` (3) bullets or `style.max_words_per_bullet` (14) words. |
| `table` | `header` (optional), `rows`, optional `col_widths` (ratios), `align` (`left`/`center`/`right` per column), `font_size`. Cells are strings or `{"text", "bold", "color", "fill"}`. Font: 24pt or 22pt — the larger one when it fits without wrapping more cells, and never larger than the slide's bullets. At 22pt only ~4–6 short rows fit; for more numbers use a `chart`. Columns are sized for that font; a narrow table keeps its natural width and is centered. `font_size` forces a size (columns still sized for it) and warns if cells then wrap. |
| `media` | Strings (paths), `{"path", "caption", "poster", "page"}`, or `{"chart": {…}, "caption", "aspect"}`. Images: png, jpg, gif, bmp, tif, webp. `pdf` and `svg` are rasterized at 220 dpi (`page` picks a PDF page, default 1). Videos: mp4, m4v, mov, avi, wmv — embedded and playable in PowerPoint; the poster frame comes from macOS Quick Look unless `poster` names an image. Captions are 22pt, centered under the item — keep them to ≤ 8 words. |
| `media[].chart` | A native, editable PowerPoint chart, laid out like an image (default `aspect` 1.6). `type`: `bar` (columns), `hbar`, `line`. `categories`: x labels. `series`: `[{"name", "values", "color"?}]` — a legend appears for 2+ series. `number_format`: Excel format for labels, e.g. `"0.0"`, `"0%"`, `"0.0\"°\""`. `labels` (default on for bars, off for lines): value labels, which also hide the value axis. `highlight`: category names drawn in the template's gold (single-series bars) — use it for "ours". `y_min` / `y_max`: axis range. |
| `media_rows` | Force the number of media rows. Default: the row count that makes the media largest. |
| `layout` | `auto` (default), `text-left`, `text-right`, `stack`, `fit`. `auto` = `text-left` when there are bullets and media but no table, else `stack`. `fit` = `stack` when every media item is wide (width/height ≥ 2.6), else `auto` — for figure-led slides such as a paper's pipeline. |
| `text_width` | Bullet column share for `text-left` / `text-right` (default `0.38`). |
| `footnote` | Small gray 12pt line at bottom left, beside the logos — citations, data sources only. |
| `text_only` | `true` silences the "no figure, demo video or chart" warning, for a genuine discussion slide. |
| `valign` | `top` (default) or `middle`: center the whole block (subtitle, bullets, table, media) in the content area. For short slides without media that would otherwise leave an empty band at the bottom. The footnote stays put. |

How `stack` places things, top to bottom: subtitle → bullets (≤45% of the height when other
blocks follow) → table (≤55% when media follow) → media in the remaining space. Media alone are
centered vertically; otherwise they start right under the block above. In `text-left` /
`text-right`, bullets get their own column and the table and media stack in the other column.

## `paper`

One paper, expanded by the builder into two `content` slides: the paper's pipeline figure with
its problem and approach, then its novelty and limitations side by side.

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
| `pipeline_caption` | Optional caption, ≤ 8 words. |
| `problem`, `approach` | One short line each, shown as `**Problem:** …` and `**Approach:** …` beside (or above) the figure. Required. |
| `novelty`, `limitations` | A string or a list; at most 2 points of ≤ `style.max_words_per_bullet` words each. They fill the two columns of the second slide, row by row. Required; a missing one drops its column. |
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
| `results.better` | Optional per-column `"lower"` / `"higher"` / `null`, overriding the header's arrows. By default a `↓` in the header means lower is better, `↑` higher. |
| `notes` | A string (methods slide) or a list of up to two strings (methods, results). |

The best value of each column that has a `↓` / `↑` (or a `better` entry) is bolded automatically,
ties included; a cell is read by its first number (`7.1°` → 7.1), and cells with no number are
skipped. Methods cells are left-aligned; results cells are centered except the first column.
The build warns above `style.max_compare_papers` rows (default 5) or `style.max_compare_columns`
columns (default 5) — split the theme or drop columns. There are no figures here; use `content`
for those.

## A complete example

```json
{
  "slides": [
    {"type": "cover", "date": "2026.10.29", "name": "Your Name"},
    {"type": "done", "projects": [
      {"name": "Vision-EMG", "items": [
        {"text": "Collected 1h+ synchronized data", "status": "done"},
        {"text": "Fusion model training", "status": "on-going"}]}]},
    {"type": "content", "title": "Vision-EMG results",
     "subtitle": "Fusion beats vision-only under occlusion",
     "table": {"header": ["Model", "MAE ↓", "Occluded MAE ↓"],
               "rows": [["Vision", "8.10°", "14.9°"], ["Vision + EMG", "**7.08°**", "**9.6°**"]]},
     "media": [{"path": "figs/mae_curve.png", "caption": "Validation MAE per epoch"}]},
    {"type": "content", "title": "Demo", "media": [
      {"path": "videos/fusion_demo.mp4", "caption": "Occluded grasp, fusion output"}]},
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
