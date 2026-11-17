# PAIRS Lab Group Meeting Skill

A Claude Code (and Codex) skill that turns your progress notes, figures, demo videos and numbers
into **PAIRS Lab (HKUST-GZ) slides** on the lab's own template: group-meeting decks (组会 PPT) in
the lab's **classic** look, and formal talks — defenses, conference and invited talks — in the
editorial **paper** look. The actual `.pptx` generation and QA go through Anthropic's **pptx
skill** (the Codex version builds the file itself).

一个 Claude Code / Codex skill：把你两周的进展、结果图、demo 视频和实验数字，直接做成基于实验室模板的 PPT。
组会用实验室原版风格（classic），答辩、会议报告这类正式场合用论文风（paper）。
真正生成 .pptx 和质量检查交给 Anthropic 官方的 pptx skill（Codex 版自己生成）。

<img src="docs/images/two-looks.jpg" alt="The same slides in the classic and paper looks: cover, results chart, key message, timeline">

<sub>The same spec in both looks — cover, a results chart, a key message, a timeline. Content pages are white in
both; only covers and section dividers carry art. / 同一份 spec 的两种风格：内容页都是白底，只有封面和分节页有底图。</sub>

## 用哪个版本 / Which version

仓库里有两个版本，做出来的 PPT 一样，spec 格式也一样。按你用的工具选一个：

| 你用的 | 装这个文件夹 | 区别 |
|---|---|---|
| **Claude Code** | [`skills/pairs-group-meeting/`](skills/pairs-group-meeting)（原版） | 生成和检查交给 Anthropic 的 pptx skill |
| **Codex** | [`codex/pairs-group-meeting/`](codex/pairs-group-meeting)（Codex 版） | 自带生成和检查的代码，不需要 pptx skill |

**Claude 用户用原版即可，Codex 用户用 Codex 版。** 不要把原版装进 Codex：原版依赖的 pptx skill
按其许可条款只能在 Anthropic 的服务里使用。

The two versions make the same slides from the same spec (the demo deck renders pixel-identical in
PowerPoint). Claude Code users install the original in `skills/`; Codex users install
`codex/pairs-group-meeting/`, which builds the file with python-pptx alone.

## Screenshots / 效果图

### With real research results / 真实实验结果

Slides built by the skill from real Vision-EMG hand-pose work: measured results, the fusion
framework diagram, and two demo videos. The videos are embedded and play in PowerPoint; here
you see their poster frames. Spec: [`examples/real/spec.json`](examples/real/spec.json)
(the media files it points to are not in this repo).

<table>
  <tr>
    <td width="50%"><img src="docs/images/real-1-results.jpg" alt="Results chart: more training data"><br>
      <sub><b>Results as native charts</b> — before/after numbers, exact values labeled</sub></td>
    <td width="50%"><img src="docs/images/real-2-framework.jpg" alt="Vision-EMG fusion framework"><br>
      <sub><b>Method figure + key sentences</b> — the diagram explained in three short lines</sub></td>
  </tr>
  <tr>
    <td width="50%"><img src="docs/images/real-3-hamer-wilor.jpg" alt="Demo video: HaMeR vs. WiLoR"><br>
      <sub><b>Demo video</b> — HaMeR vs. WiLoR on the same recording</sub></td>
    <td width="50%"><img src="docs/images/real-4-wilor-mano.jpg" alt="Demo video: WiLoR to MANO"><br>
      <sub><b>Demo video</b> — multi-view MANO fit beside the synced camera view</sub></td>
  </tr>
</table>

### Full demo deck / 完整示例

Rendered by PowerPoint from [`examples/demo/spec.json`](examples/demo/spec.json) (synthetic data),
in the classic look.

<table>
  <tr>
    <td width="50%"><img src="docs/images/01-cover.jpg" alt="Cover slide"><br>
      <sub><b>Cover</b> — date and presenter on the template cover</sub></td>
    <td width="50%"><img src="docs/images/02-done.jpg" alt="Done last two weeks slide"><br>
      <sub><b>Done last two weeks</b> — grouped by project, done / on-going in color</sub></td>
  </tr>
  <tr>
    <td width="50%"><img src="docs/images/03-figure.jpg" alt="Figure with key sentences"><br>
      <sub><b>Figure + key sentences</b> — a takeaway line, ≤ 3 bullets, the figure sized to fit</sub></td>
    <td width="50%"><img src="docs/images/04-chart.jpg" alt="Native chart slide"><br>
      <sub><b>Native chart</b> — values labeled, still editable in PowerPoint</sub></td>
  </tr>
  <tr>
    <td width="50%"><img src="docs/images/05-two-charts.jpg" alt="Two charts side by side"><br>
      <sub><b>Two visuals</b> — laid out automatically, the best result in gold</sub></td>
    <td width="50%"><img src="docs/images/06-todo.jpg" alt="Todo slide"><br>
      <sub><b>Todo</b> — next steps per project</sub></td>
  </tr>
</table>

### Paper reading and comparison / 论文汇报与对比

Rendered by PowerPoint from [`examples/papers/spec.json`](examples/papers/spec.json) (synthetic papers
and numbers), in the classic look.

<table>
  <tr>
    <td width="50%"><img src="docs/images/paper-1-pipeline.jpg" alt="Paper reading: pipeline figure with problem and approach"><br>
      <sub><b>Paper reading</b> — the paper's own pipeline figure, its problem and approach</sub></td>
    <td width="50%"><img src="docs/images/paper-2-limits.jpg" alt="Paper novelty and limitations"><br>
      <sub><b>Novelty and limits</b> — side by side, from the presenter</sub></td>
  </tr>
  <tr>
    <td width="50%"><img src="docs/images/paper-3-methods.jpg" alt="Comparison of methods"><br>
      <sub><b>Compare methods</b> — one theme's papers, the dimensions that separate them</sub></td>
    <td width="50%"><img src="docs/images/paper-4-results.jpg" alt="Comparison of results"><br>
      <sub><b>Then results</b> — one benchmark, the best value bolded, the source in the footnote</sub></td>
  </tr>
</table>

### More page forms, two looks / 更多页面形式，两种风格

Rendered by PowerPoint from [`examples/layouts/spec.json`](examples/layouts/spec.json) (synthetic
numbers), in `paper` (formal talks) and `classic` (group meetings).

<table>
  <tr>
    <td width="50%"><img src="docs/images/forms-1-statement-paper.jpg" alt="Key result statement, paper look"><br>
      <sub><b>Key message</b> (paper) — the number and the sentence it stands for</sub></td>
    <td width="50%"><img src="docs/images/forms-2-statement-classic.jpg" alt="Key result statement, classic look"><br>
      <sub><b>The same slide</b> (classic) — the lab's gold bar instead of a hairline</sub></td>
  </tr>
  <tr>
    <td width="50%"><img src="docs/images/forms-3-columns-charts.jpg" alt="Three columns with a chart each"><br>
      <sub><b>Columns</b> (classic) — one result under three conditions, a chart each</sub></td>
    <td width="50%"><img src="docs/images/forms-4-columns-text.jpg" alt="Three text columns"><br>
      <sub><b>Columns</b> (paper) — three points, hairlines instead of boxes</sub></td>
  </tr>
  <tr>
    <td width="50%"><img src="docs/images/forms-5-timeline.jpg" alt="Timeline of milestones"><br>
      <sub><b>Timeline</b> (classic) — done, now, next</sub></td>
    <td width="50%"><img src="docs/images/forms-6-agenda.jpg" alt="Outline slide"><br>
      <sub><b>Outline</b> (paper) — read from the deck's own sections</sub></td>
  </tr>
</table>

A photo or demo video can also fill half a slide out to its edge (`"layout": "split"`), and a
formal talk ends on a `closing` slide.

## What it makes

| Slide | What goes on it |
|---|---|
| Cover | "Group Meeting" (or the talk's title), date, presenter — and, in the paper look, a photo of your robot or setup |
| Outline | The deck's parts, numbered; repeated before each part, it shows where the talk is |
| Section | A divider that opens each project or the paper reading; the slides after it show its name in the footer |
| Done last two weeks | Items grouped by project, lettered `a. b. c.`, status in green (**done**) or red (**on-going**) |
| Results | A one-line takeaway, up to 3 short bullets, and figures / embedded demo videos / native PowerPoint charts — or one photo or video filling half the slide out to its edge (`split`) |
| Key message | One sentence set large, with the number it is about (`36%`) |
| Columns | Two to four parallel points — contributions, conditions, method parts — each a heading, a line and optionally a small chart, with hairlines instead of boxes |
| Timeline | Milestones or steps along a line, done / now / next |
| Paper reading | Per paper: the authors' original pipeline figure with its problem and approach, then novelty and limitations side by side — all supplied by you, never guessed |
| Paper comparison | Per theme of 3–5 papers: a methods table, then a results table (best value bolded) — tables only, no figures. Example: [`examples/papers/spec.json`](examples/papers/spec.json) (synthetic) |
| Todo | Next steps per project |
| Closing | "Thank you" and a line or two, on the cover's page — for formal talks |

Every new page form is shown in [`examples/layouts/spec.json`](examples/layouts/spec.json).

Every slide is cloned from the template, so the background, header band and HKUST(GZ) / PAIRS Lab
logos are always right. Figures are laid out automatically, videos play in PowerPoint, charts stay
editable, long Done lists split across slides by themselves, and every content slide carries its
slide number.

### Two looks / 两种风格

The same spec builds in either theme — set `"theme"` in the spec or in your config, pass
`--theme` to try the other, or just ask Claude for the deck "in paper too". Content pages are white
in both; only covers and section dividers carry art, and variety comes from the page forms above
rather than from page colors.

| Theme | Look | For |
|---|---|---|
| `classic` (default) | The lab template: navy header band with radar rings, white pages, sections on navy | Group meetings |
| `paper` | Editorial: white pages, Georgia titles and headings, ruled tables, the lab's rings on cover and sections, a photo panel on the cover | Formal talks: defenses, conference and invited talks |

### House style

Set in `skills/pairs-group-meeting/config.json` → `style`; the builder enforces the font floor and
warns about the rest:

- **All English** — notes written in Chinese are translated.
- **Body text ≥ 22pt** — text is never shrunk below it; the slide gets cut or split instead.
- **Few words, key sentences** — a one-line takeaway plus ≤ 3 bullets of ≤ 14 words.
- **Show it** — every results slide has a figure, demo video or chart (paper-comparison slides are the exception: tables only).

## Requirements

- [Claude Code](https://claude.com/claude-code) (original version) or [Codex](https://developers.openai.com/codex) (Codex version)
- Original version only: the **pptx skill** — built into Claude apps as `anthropic-skills:pptx`; in
  Claude Code you can also get it from Anthropic's skills marketplace:
  `/plugin marketplace add anthropics/skills`, then `/plugin install document-skills@anthropic-agent-skills`.
  The Codex version doesn't need it.
- Python ≥ 3.10 with `pip install -r skills/pairs-group-meeting/requirements.txt`
  (python-pptx, lxml, defusedxml, Pillow, PyMuPDF)
- For visual QA: **Microsoft PowerPoint for Mac** (exact renders through AppleScript) or LibreOffice.
  PowerPoint renders through one folder per project, `.office-render`, so it asks for file access
  only once per project (see [office-render-skill](https://github.com/ezoiHY/office-render-skill))

## Install

### Codex

```bash
git clone https://github.com/ezoiHY/pairs-group-meeting-skill.git
mkdir -p ~/.agents/skills
ln -s "$PWD/pairs-group-meeting-skill/codex/pairs-group-meeting" ~/.agents/skills/pairs-group-meeting
```

`~/.agents/skills/` is Codex's user skills folder. Codex also reads `~/.codex/skills/` (where
`$skill-installer` puts skills), so linking there works too. Codex asks your name the first time
and saves it in `~/.config/pairs-group-meeting/config.json` (see below).

### Claude Code

**As a plugin** (in Claude Code):

```
/plugin marketplace add ezoiHY/pairs-group-meeting-skill
/plugin install pairs-group-meeting@pairs-lab
```

**Or as a personal skill**, which is easier to edit:

```bash
git clone https://github.com/ezoiHY/pairs-group-meeting-skill.git
ln -s "$PWD/pairs-group-meeting-skill/skills/pairs-group-meeting" ~/.claude/skills/pairs-group-meeting
```

### Your settings (both versions)

Claude or Codex asks your name the first time and saves it in your own config file,
`~/.config/pairs-group-meeting/config.json`, outside the skill — so updating the skill never
overwrites it, and both versions share it. You can also write it yourself:

```json
{
  "presenter": "Your Name",
  "output_root": "~/Documents/PairsLab",
  "theme": "classic"
}
```

Any key here overrides the skill's `config.json`, which holds the lab defaults (theme, list titles
and the house-style limits).

Decks are written to `<output_root>/<YYMMDD>GroupMeeting/<YYMMDD>GroupMeeting.pptx`, next to the
`deck_spec.json` they were built from, so you can tweak the spec and rebuild.

## Use

Just ask Claude Code or Codex, in Chinese or English (in Codex you can also call it by name,
`$pairs-group-meeting …`):

> 帮我做 11.12 组会的 PPT。这两周 Vision-EMG 采完了 3 个人的数据，融合模型遮挡下 MAE 从 14.9° 降到 9.6°，
> 论文 method 部分在写。结果图在 ~/figs/mae.png，demo 视频 ~/videos/occlusion.mp4。下两周采完剩下 2 人、写完初稿。

Claude collects what's missing, plans the slides, writes a deck spec, builds it, renders every slide
and checks it before handing you the `.pptx`.

For a formal talk, say what the occasion is; it takes the paper look, with the talk's title on the
cover, an outline, sections and a closing slide:

> 下周三在 ICRA 做 15 分钟口头报告，题目 Seeing Hands Through Occlusion，用 paper 风格。
> 主要讲三件事：多模态采集系统、融合模型、遮挡下的结果，材料都在 ~/talk/ 里。

You can also build a spec by hand:

```bash
python skills/pairs-group-meeting/scripts/build_deck.py examples/demo/spec.json
python skills/pairs-group-meeting/scripts/build_deck.py examples/demo/spec.json --theme paper -o build/demo/paper.pptx
python skills/pairs-group-meeting/scripts/render.py build/demo/261029GroupMeeting.pptx
```

The first render in a project exits with code 3 and asks you to run
`python3 skills/pairs-group-meeting/scripts/office_render.py --grant` once. PowerPoint then shows
"Grant File Access": click Select…, select the hidden `.office-render` folder (⌘⇧. shows it) or the
project folder, then click Grant Access. It never asks again for that project.

With the Codex version, use `codex/pairs-group-meeting/scripts/` instead of
`skills/pairs-group-meeting/scripts/`.

The spec format is documented in [`references/spec-format.md`](skills/pairs-group-meeting/references/spec-format.md);
[`examples/demo/spec.json`](examples/demo/spec.json) builds the deck shown above.

## 组会汇报怎么做

这一节说明：用这个 skill 做一次组会汇报（或者一场正式报告），你要准备什么、PPT 怎么排、论文怎么讲。
PPT 本身是全英文的，你对 Claude 说中文、贴中文笔记都可以，它会翻译。
下面说的"Claude"，用 Codex 版时就是 Codex，做法完全一样。

### 1. 整体流程

1. 对 Claude Code 或 Codex 说一句话：会议日期、这两周做了什么、结果图和 demo 视频的路径（见上面 Use 里的例子）。
2. Claude 一次性问清缺的东西：哪些结果还没有图、论文还缺哪几项、上次的 Todo 哪些做完了。
3. 它写出 `deck_spec.json`，构建 `.pptx`，用 PowerPoint 渲染每一页并检查，然后交给你。
   第一次在某个项目里渲染时，PowerPoint 会弹一次"授予文件访问权限"：点"选择…"，选中项目里的
   `.office-render` 文件夹（隐藏文件夹，按 ⌘⇧. 显示），或者直接选项目文件夹本身，再点"授予访问权限"。
   之后这个项目不会再问，渲染时 PowerPoint 也不会跳到前台。
4. 想改的话，直接改 `deck_spec.json` 再重新构建：

   ```bash
   python skills/pairs-group-meeting/scripts/build_deck.py <output_root>/<YYMMDD>GroupMeeting/deck_spec.json
   ```

   用原版时，如果提示找不到 pptx skill，加上 `--pptx-skill <pptx skill 的目录>`。
   用 Codex 版时，把路径换成 `codex/pairs-group-meeting/scripts/build_deck.py`。

### 2. 一场组会的 PPT 怎么排

| 顺序 | 页 | 内容 |
|---|---|---|
| 1 | Cover | 日期和你的名字 |
| 2 | Done last two weeks | 按项目分组，每条标 done 或 on-going |
| 3 | Results | 一页一个结论，配图 / 图表 / demo 视频，数量按工作量定。最想让大家记住的那个结果可以单独做一页核心结论（statement），实验装置照片或 demo 可以做半页照片页（split） |
| 4 | Paper reading | 可选。每篇论文 2 页 |
| 5 | Paper comparison | 可选。读了 2 篇以上才有，放在论文之后 |
| 6 | Todo | 下两周要做的事；带日期的里程碑可以做成时间线（timeline） |

一般 8–14 页，每页讲 1.5–2 分钟。讲两个以上项目、或者有论文部分时，每部分前面加一页分节页（section），
后面每页的页脚会显示这一部分的名字；有三个或更多部分时，封面后加一页目录（agenda），它会自动列出各个分节。

### 正式场合怎么排（paper 风格）

答辩、会议报告、受邀报告用 paper 风格，顺序一般是：

| 顺序 | 页 | 内容 |
|---|---|---|
| 1 | Cover | 报告题目（不是 Group Meeting）、日期、名字，再加一行场合或单位；右侧可以放一张照片 |
| 2 | Outline | 目录，自动列出各个分节 |
| 3 | 各部分 | 每部分前一页分节页；长报告可以在每部分前再放一次目录，高亮当前讲到的那一节 |
| 4 | Key message | 全场最重要的一两个结论，一句话放大，配上数字 |
| 5 | Columns | 贡献点、方法的几个模块、几种实验条件，每栏一个小标题加一行字 |
| 6 | Timeline | 研究计划或路线，可选 |
| 7 | Closing | Thank you / Questions，加一两行联系方式 |

页数按时长定，每页讲 1.5–2 分钟，比如 15 分钟的会议报告大约 8–10 页。规矩和组会一样：全英文、正文不小于 22pt、一页一个结论。

### 风格

同一份 spec 可以换两种风格（主题），内容不用改。两种风格的内容页都是白底，只有封面和分节页有底图：

- **classic**（默认）：实验室模板，深蓝标题栏 + 雷达环，白色内容页，分节页是深蓝底。用于组会
- **paper**：杂志/论文风，白底、衬线标题（Georgia）、三线表，封面和分节页用实验室的同心环，封面可以放照片。用于更正式的场合：答辩、会议报告、受邀报告

在 spec 里写 `"theme": "paper"`，或者在 `~/.config/pairs-group-meeting/config.json` 里设成你的默认。
想先比较一下，就对 Claude 说"用 paper 再出一版"。

让 PPT 不单调，靠的是换页面形式，而不是换底色：

- **目录**（`agenda`）：自动列出各个分节；正式报告里每节前重复一次，高亮当前讲到的那一节
- **核心结论**（`statement`）：一句话放大，配上它说的那个数字，比如"36%"。留给最想让大家记住的一两个结果
- **并列要点**（`columns`）：2–4 栏，每栏一个小标题加一行字，可以各配一张小图表；适合贡献点、几种条件、方法的几个部分
- **照片和 demo**（`"layout": "split"`）：照片或视频占满半页、一直延伸到页面边缘，文字在另一边
- **时间线**（`timeline`）：里程碑或步骤，分完成 / 进行中 / 计划中
- **结尾页**（`closing`）：Thank you + 一两行，正式报告用
- paper 的封面可以放一张照片（机器人、实验装置、demo 截图）：cover 里加 `"image"`
- 图表颜色交给主题；要把基线/"before" 压暗写 `"color": "muted"`，突出自己的结果用 `highlight`

### 3. 你需要准备什么

| 部分 | 要给 Claude 的东西 |
|---|---|
| Done | 上次组会以来做了什么，按项目分组，哪些完成、哪些还在做。Claude 会读上一次 PPT 的 Todo，当作这次 Done 的候选 |
| Results | 每个结果的图、截图或 demo 视频，加一句你想让大家记住的结论。只有数字也行，会做成图表 |
| 照片（可选） | 实验装置、机器人或 demo 截图：可以做成半页照片页，paper 风格的封面也能放。只用你给的照片，不会去网上找图 |
| Todo | 下两周的计划，有截止日期就写上；带日期的计划可以做成时间线 |
| 论文 | 见下面两节 |

### 4. 讲一篇论文

每篇论文你要给齐五样东西，缺哪样 Claude 就问哪样，**不会替你补**：

1. **原版 pipeline 图**：论文自己的方法总览图（图片，或 PDF 加页码），不要自己重画
2. **要解决什么问题**（Problem）
3. **怎么解决的**（Approach）
4. **创新点**（Novelty）
5. **不足**（Limitations）：这是你自己的判断，不要让 Claude 替你写。想不出来就想想"它在什么情况下会失效"：要额外传感器、只测了一个数据集、需要逐人标定……

每篇论文生成 2 页：

- 第 1 页：pipeline 图，旁边两句 Problem 和 Approach
- 第 2 页：Novelty 和 Limitations 左右两栏，每栏最多 2 条

你口头讲的详细解释会放进这两页的 speaker notes，页面上只留每行不超过 14 个词的短句。

讲的时候的建议（来自常见的 journal club / reading group 做法）：

- 先讲问题，再顺着 pipeline 图从输入讲到输出，不要把论文的图一张张过一遍
- 创新点只讲和前人相比真正新的那一两点
- 以不足和你自己的看法收尾，最好给大家抛一个具体的问题

### 5. 对比几篇论文

读了 2 篇以上时，按**主题**分组，每组 3–5 篇，每组生成两页，**只有表格，没有图**：

1. **方法表**（先比方法）：每行一篇论文，第一列是论文名，后面只放能区分这几篇的维度，比如输入、核心想法、额外成本
2. **结果表**（再比结果）：每行一篇论文，数字放在同一个基准数据集上。表头写 ↓ 或 ↑，每列最好的值会自动加粗

注意：

- 表格下方的脚注要写清楚基准数据集和设置，以及数字是不是论文自己报告的（as reported）。设置不同的数字不要硬比，对不上的格子写 `—`
- 数字必须来自你，Claude 不会编
- 一组最多 5 篇、5 列，超过会警告，这时拆成两个主题，或者删掉区分度不大的列

### 6. 写 PPT 的规矩

- **全英文**：标题、要点、表格、图注都是英文，专有名词保持原样（MAE、TCN、HaMeR）
- **正文不小于 22pt**：放不下就删字或拆页，不缩字号
- **少字**：一页一个结论，一行副标题讲结论，最多 3 条要点，每条不超过 14 个词
- **结果页要有图**：图、图表或 demo 视频；论文对比页是例外，只放表格
- **内容页白底**：两种风格的内容页都是白底；想让 PPT 不单调，就换页面形式（目录、核心结论、并列、时间线、半页照片），不换底色
- **细节放 notes**：想多讲的话写在 speaker notes 里，不要塞进页面

### 7. 构建时的 WARNING

构建结束会打印 WARNING，每一条都要处理，Claude 会照着改：缺图缺字段会回来问你，要点太长会让你缩短或拆页，
论文对比超过 5 篇会建议拆主题。完整的警告和处理办法见 [`SKILL.md`](skills/pairs-group-meeting/SKILL.md)。

### 8. 一个完整的例子

> 帮我做 11.12 组会的 PPT。这两周 Vision-EMG 采完了 3 个人的数据，融合模型遮挡下 MAE 从 14.9° 降到 9.6°，
> 结果图在 ~/figs/mae.png。读了两篇论文：
> A（CVPR 2025）要解决遮挡下手部姿态估计失效，用上一帧姿态做先验，创新是按关节可见性给先验加权，不足是长时间遮挡会漂移，pipeline 图在 ~/papers/a_fig2.png；
> B（ICRA 2025）用 EMG 和视觉融合，创新是每个关节单独学融合权重，不足是每个人都要标定，pipeline 图在 ~/papers/b_fig2.png。
> 两篇一起比一下，MAE 和遮挡下 MAE 的数字在 ~/papers/results.md。下两周采完剩下 2 人。

示例 spec 见 [`examples/papers/spec.json`](examples/papers/spec.json)（数据是虚构的）。

## Layout

```
.claude-plugin/              plugin + marketplace manifests (Claude Code)
skills/pairs-group-meeting/  the original version, for Claude Code
  SKILL.md                   workflow and house style Claude follows
  assets/template.pptx       the lab template, with sample slides for every theme
  scripts/build_deck.py      spec.json → .pptx
  scripts/deck_structure.py  clones template slides with the pptx skill's add_slide/clean/validate
  scripts/slide_text.py      prints a deck's text, slide by slide
  scripts/render.py          .pptx → slide PNGs + overview (PowerPoint or LibreOffice)
  scripts/office_render.py   PowerPoint through one once-granted folder per project (copy of office-render-skill)
  assets/blank.pptx          the deck office_render.py --grant exports
  references/                spec format, template geometry/colors/fonts
  config.json                presenter, output folder, theme, house-style limits
codex/pairs-group-meeting/   the Codex version
  SKILL.md                   generated from the original by tools/sync_codex.py
  scripts/deck_structure.py  clones template slides with python-pptx alone
  …                          everything else is a copy of the original's
tools/sync_codex.py          brings the Codex version up to date with the original
tools/make_template.py       makes assets/template.pptx: the lab template + every theme's sample slides
tools/template/              its sources: the lab's original template and the lab's logos
examples/demo/               a demo spec with synthetic data
examples/papers/             paper-reading and paper-comparison slides, synthetic data
examples/layouts/            outline, key message, columns, timeline and closing slides, synthetic data
```

**Changing the skill:** edit the original in `skills/pairs-group-meeting/`, then run
`python tools/sync_codex.py` to update the Codex version. `--check` exits non-zero if the Codex
version is out of date. Only `deck_structure.py` is edited by hand in `codex/`. To change a theme's
art or add a theme, edit `tools/make_template.py` (and `THEMES` in `build_deck.py`), run it, then
`tools/sync_codex.py`; it needs numpy on top of the skill's requirements.

The template, its backgrounds and the HKUST(GZ) and PAIRS Lab logos belong to their owners.
