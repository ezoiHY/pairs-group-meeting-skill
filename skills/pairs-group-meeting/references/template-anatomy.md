# Template anatomy (`assets/template.pptx`)

Read this before editing slide XML by hand, i.e. for anything `build_deck.py` cannot express.
16:9, 12192000 × 6858000 EMU (13.333" × 7.5"). 1" = 914400 EMU, 1pt = 12700 EMU.

The template holds one set of sample slides per theme. It is made from the lab's own template by
the repository's `tools/make_template.py` (which keeps the lab's original in
`tools/template/lab-template.pptx`); regenerate it there rather than editing this file.

## Sample slides

Each sample is named `<theme>/<kind>` in its `<p:cSld name="…">`; `build_deck.py` finds them by
that name.

| Name | Spec types cloned from it | Contents |
|---|---|---|
| `classic/cover` | `cover`, `closing` | The lab's own cover, unchanged: light background with concentric circles right, gold bar left of the title, navy band at the bottom, both logos top right. `ctrTitle` "Group Meeting" 48pt bold `111F4A`; `subTitle` idx 1: date, then name, 24pt `4A5878`. |
| `classic/content` | `content`, `done`, `todo`, `agenda`, `statement`, `columns`, `timeline` | The lab's content page, white: navy header band to y≈984000 with the blue→gold rule under it (ending at 1.09"), the footer hairline at y 7.12". Title placeholder at (349250, 95250) 7620000×762000, 36pt bold white. White PAIRS Lab logo at the band's right. |
| `classic/section` | `section` | Navy page, rings on the right, gold bar beside the title. Title 44pt bold white at (1.11", 2.72"); body placeholder idx 1 24pt `B8C5DA` under it. |
| `paper/cover` | `cover`, `closing` | White, the lab's rings in light gray on the right, both color logos top left. Title Georgia 60pt `111F4A`, date and name Calibri 24pt `5A6578`. "Image slot" rectangle: the right panel, full height. |
| `paper/content` | the same as `classic/content` | White. Title Georgia 36pt `111F4A` at (0.6", 0.40"). Color PAIRS Lab logo top right. |
| `paper/section` | `section` | White with the rings; title Georgia 48pt, one line under it in Calibri 24pt. |

Every page that carries content is white in both themes; only covers and dividers have art.

The backgrounds are slide-level `<p:bg>` pictures, and the logos are pictures on each slide —
neither lives on the layouts. So **always start a new slide by cloning one of these slides**
(`build_deck.py` does this for every spec slide), never from a layout, or the slide loses its
background and logos.

Shapes the builder looks for by name:

| Shape | On | Builder does |
|---|---|---|
| `Footer label` | content samples | Writes the current section's title, or deletes the box |
| `Slide number` | content samples | A slide-number field; the builder sets its cached text |
| `Image slot` | `paper/cover` | Swaps it for the cover's (or closing slide's) `image`, cropped to the slot's shape, or deletes it |

## Safe areas (EMU, or inches)

| Area | classic | paper |
|---|---|---|
| Left/right margin | 349250 (0.38") | 0.6" |
| Content top | 1219200 (1.33", under the band) | 1.42" |
| Content bottom | 6.60" | 6.62" |
| A `split` photo's half | from 1.09" (flush under the band's rule) to 6.60", out to the slide edge | from 1.42" to 6.62", out to the slide edge |
| Footnote | x 0.38", y 6.74"–7.06", up to 7.4" wide | x 0.6", y 6.90"–7.20", up to 7.0" wide |
| Footer label (right-aligned) | ends at 12.18", y 6.80" | ends at 12.12", y 6.95" |
| Slide number | 12.26"–12.96", at the footer rule's height | 12.23"–12.73", y 6.95" |
| Logo | in the band, right edge 12.95" | top right, right edge 12.73", y ≈ 0.6" |

## Colors and fonts

Every color the builder writes comes from the theme's palette in `build_deck.py` → `THEMES`:
`text` (body), `strong` (takeaways, project names, headings, table header text on light rules),
`muted` (captions, footnotes, list letters), `bullet`, `number` (agenda numbers), `accent`
(highlights, the current milestone), `mark` (the gold bar and short rules — decoration, never
text), `done` / `ongoing`, the table colors and the chart series. A hand-made slide takes the same
tokens. The classic palette, from the lab template:

| Token | Hex | Use |
|---|---|---|
| Body navy | `002147` | All body text |
| Deep navy | `111F4A` | Cover title, takeaways, table header fill |
| Slate | `4A5878` | Cover date/name, captions, footnotes |
| Gold | `C9A04A` | Highlights ("ours"), the gold of the band's rule |
| Bar gold | `D6B05C` | The gold bar beside a title or statement, the short rule over a column heading |
| Blue | `2F6DB5` | Bullets, agenda numbers |
| Done green | `schemeClr accent6` (`4EA72E`) | "done" status |
| On-going red | `980000` | "on-going" status |
| Rule gray | `D5DBE6` | Table hairlines, chart axes |
| Stripe | `F2F5FA` | Alternate table rows |

Fonts: body text is Calibri in both themes; titles, headings, statements and big numbers are
Calibri bold (classic) or Georgia (paper). Every run also names 等线 (DengXian, ships with Office)
for Chinese:

```xml
<a:rPr lang="en-US" altLang="zh-CN" sz="2800" b="0" dirty="0">
  <a:solidFill><a:srgbClr val="002147"/></a:solidFill>
  <a:latin typeface="Calibri" panose="020F0502020204030204" pitchFamily="34" charset="0"/>
  <a:ea typeface="等线" panose="02010600030101010101" pitchFamily="2" charset="-122"/>
  <a:cs typeface="Calibri" panose="020F0502020204030204" pitchFamily="34" charset="0"/>
</a:rPr>
```

Without the `<a:ea>`, Chinese falls back to the theme's 宋体 (decks are English by house style, but
names or symbols can still need it). Sizes used: title 36, Done/Todo 28 down to 22, bullets 28/24
down to 22, subtitle 24–22, table 24–22, captions and chart text 22, agenda items 32–22, statement
44 (36 beside a number) down to 22 and its number 132–72, column headings 32/28 and text 28/24
down to 22, timeline labels 26 and text 24 down to 22, footnote and footer 12. **No body text
below 22pt** — the house rule; hand-made slides follow it too.

## Done/Todo paragraph shapes (from the lab template)

Project header — `marL="0"`, bold, `spcBef` 0 for the first project and 18pt for later ones:

```xml
<a:p><a:pPr marL="0" indent="0" algn="l"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc>
  <a:spcBef><a:spcPts val="1800"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft><a:buNone/></a:pPr>
  <a:r><a:rPr … b="1">…</a:rPr><a:t>Project: </a:t></a:r>
  <a:r><a:rPr … b="1">…</a:rPr><a:t>Vision-EMG</a:t></a:r></a:p>
```

Item — `marL="457200"`, `spcBef` 3pt, the letter in the muted color, status run after ` — `:

```xml
<a:p><a:pPr marL="457200" indent="0" algn="l">…<a:spcBef><a:spcPts val="300"/></a:spcBef>…<a:buNone/></a:pPr>
  <a:r><a:rPr …><a:solidFill><a:srgbClr val="4A5878"/></a:solidFill>…</a:rPr><a:t>a. </a:t></a:r>
  <a:r><a:rPr …>…</a:rPr><a:t>Collected 1h+ data</a:t></a:r>
  <a:r><a:rPr …><a:solidFill><a:srgbClr val="4A5878"/></a:solidFill>…</a:rPr><a:t> — </a:t></a:r>
  <a:r><a:rPr …><a:solidFill><a:schemeClr val="accent6"/></a:solidFill>…</a:rPr><a:t>done</a:t></a:r></a:p>
```

## Shape names the builder writes

`Done list` / `Todo list`, `Subtitle`, `Bullets`, `Table`, `Figure N`, `Video N`, `Chart N`,
`Caption N`, `Footnote`, `Cover image` / `Closing slide image`, `Agenda number N` / `Agenda item N`
/ `Agenda rule …` / `Agenda mark`, `Statement` / `Statement number` / `Statement detail` / `Mark`,
`Column heading N` / `Column text N` / `Column rule N` / `Column mark N`, `Step label N` / `Step text
N` / `Marker N` / `Timeline axis` / `Timeline progress` — handy for finding a shape when patching XML
after a build. Rules and markers are plain shapes without the Office theme's style, so they carry
no shadow or outline of their own.
