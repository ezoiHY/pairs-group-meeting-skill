# Template anatomy (`assets/template.pptx`)

Read this before editing slide XML by hand, i.e. for anything `build_deck.py` cannot express.
16:9, 12192000 × 6858000 EMU (13.333" × 7.5"). 1" = 914400 EMU, 1pt = 12700 EMU.

## The four sample slides

| File | Layout | Spec type | Contents |
|---|---|---|---|
| `slide1.xml` | `slideLayout1` (标题幻灯片) | `cover` | Background `bg_cover` (light, concentric circles right, gold bar left of title, navy band at bottom). `ctrTitle` "Group Meeting" at (1016000, 2466975) 6096000×733425, 48pt bold `111F4A`. `subTitle` idx 1 at (1016000, 4178300) 5080000×762000: date, then name, 24pt `4A5878`. Logo group top right. |
| `slide2.xml` | `slideLayout2` | `done` | Background `bg_content` (navy header band to y≈980000 with a blue→gold rule under it, short navy rule bottom left). Title placeholder at (349250, 95250) 7620000×762000, 36pt bold white. Text box "Google Shape;146;p31" at (349250, 1219200) 8890000×5080000 with the Done list. Small logo group bottom right. |
| `slide3.xml` | `slideLayout2` | `content` | Same background, title ("XXX results") and logo group; nothing else. Base for every free-form slide. |
| `slide4.xml` | `slideLayout2` | `todo` | Same as slide 2, items without status. |

The backgrounds are slide-level `<p:bg>` pictures, and the logos are a group on each slide —
neither lives on the layouts. So **always start a new slide by cloning one of these slides**
(`add_slide.py unpacked/ slide3.xml`), never from a layout, or the slide loses its background and logos.

## Safe areas (EMU)

| Area | Value |
|---|---|
| Left/right margin | 349250 (matches the title's left edge) |
| Content top | 1219200 (just under the header band) |
| Content bottom | 5740000 (0.2" above the logo group) |
| Logo group (content slides) | x 8128524…12532658, y 5922049…6541880 |
| Footnote band | y 5960000…6400000, x 349250…~7900000 (left of the logos) |

## Colors

| Token | Hex | Use |
|---|---|---|
| Body navy | `002147` | All body text, bullets |
| Deep navy | `111F4A` | Cover title, subtitles, table header fill |
| Slate | `4A5878` | Cover date/name, captions, footnotes |
| Done green | `schemeClr accent6` (`4EA72E`) | "done" status |
| On-going red | `980000` | "on-going" status, warnings |
| White | `FFFFFF` | Titles on the header band |
| Rule gray | `D5DBE6` | Table borders |
| Stripe | `F2F5FA` | Alternate table rows |

## Fonts

Every run names Calibri for Latin and 等线 (DengXian, ships with Office) for Chinese:

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
down to 22, subtitle 24–22, table 24–22, captions and chart text 22, footnote 12. **No body text
below 22pt** — the house rule; hand-made slides follow it too.

## Done/Todo paragraph shapes (from the template)

Project header — `marL="0"`, bold, `spcBef` 0 for the first project and 18pt for later ones:

```xml
<a:p><a:pPr marL="0" indent="0" algn="l"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc>
  <a:spcBef><a:spcPts val="1800"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft><a:buNone/></a:pPr>
  <a:r><a:rPr … b="1">…</a:rPr><a:t>Project: </a:t></a:r>
  <a:r><a:rPr … b="1">…</a:rPr><a:t>Vision-EMG</a:t></a:r></a:p>
```

Item — `marL="457200"`, `spcBef` 3pt, literal letter prefix, status run after ` — `:

```xml
<a:p><a:pPr marL="457200" indent="0" algn="l">…<a:spcBef><a:spcPts val="300"/></a:spcBef>…<a:buNone/></a:pPr>
  <a:r><a:rPr …>…</a:rPr><a:t>a. Collected 1h+ data — </a:t></a:r>
  <a:r><a:rPr …><a:solidFill><a:schemeClr val="accent6"/></a:solidFill>…</a:rPr><a:t>done</a:t></a:r></a:p>
```

## Shape names the builder writes

`Done list` / `Todo list`, `Subtitle`, `Bullets`, `Table`, `Figure N`, `Video N`, `Caption N`,
`Footnote` — handy for finding a shape when patching XML after a build.
