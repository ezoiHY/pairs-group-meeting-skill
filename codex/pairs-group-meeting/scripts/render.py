#!/usr/bin/env python3
"""Render a .pptx to one PNG per slide, plus a contact sheet, for visual QA.

Engines, in the order `auto` tries them:
  powerpoint   Microsoft PowerPoint on macOS, driven by AppleScript. Renders exactly what the
               audience will see (real Calibri / DengXian, real video posters). PowerPoint works only
               inside the project's render folder, <project>/.office-render, through office_render.py
               (a copy of https://github.com/ezoiHY/office-render-skill). It is granted that folder once
               per project, so decks in new folders never make it ask again.
  libreoffice  soffice, through the pptx skill's wrapper when --pptx-skill is given.

Usage:
    python render.py deck.pptx [--outdir DIR] [--dpi 110] [--engine auto|powerpoint|libreoffice]
                     [--pptx-skill DIR] [--project DIR]

Prints the absolute path of every slide image and of the contact sheet. Exits 3 when PowerPoint has not
been granted the render folder yet and 4 when it is waiting on a dialog. Both need the user (the message
says what to run), so neither falls back to LibreOffice.
"""

import argparse
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

import office_render

POWERPOINT_APP = office_render.PP_APP


def via_powerpoint(deck, pdf, project):
    # Copies the deck into the render folder, exports it there and moves the PDF to `pdf`. PowerPoint
    # stays in the background, and a dialog it waits on stops the render within seconds (exit 4).
    office_render.export_pdf(deck, pdf, project=project)


def via_libreoffice(deck, pdf, pptx_skill):
    outdir = pdf.parent
    if pptx_skill:
        office = Path(pptx_skill).expanduser() / "scripts" / "office"
        sys.path.insert(0, str(office.parent))
        from office.soffice import run_soffice
        res = run_soffice(["--headless", "--convert-to", "pdf", "--outdir", str(outdir), str(deck)],
                          capture_output=True, text=True, timeout=300)
    else:
        exe = shutil.which("soffice") or "/Applications/LibreOffice.app/Contents/MacOS/soffice"
        if not Path(exe).exists():
            raise RuntimeError("LibreOffice (soffice) not found")
        with tempfile.TemporaryDirectory(prefix="lo_profile_") as profile:
            res = subprocess.run([exe, f"-env:UserInstallation={Path(profile).as_uri()}", "--headless",
                                  "--convert-to", "pdf", "--outdir", str(outdir), str(deck)],
                                 capture_output=True, text=True, timeout=300)
    made = outdir / (deck.stem + ".pdf")
    if res.returncode != 0 or not made.exists():
        raise RuntimeError(f"LibreOffice export failed: {res.stderr.strip()}")
    if made != pdf:
        made.replace(pdf)


def pdf_to_pngs(pdf, outdir, stem, dpi):
    try:
        import pymupdf
    except ImportError:
        pymupdf = None
    paths = []
    if pymupdf:
        doc = pymupdf.open(pdf)
        width = len(str(doc.page_count))
        for i, page in enumerate(doc, 1):
            p = outdir / f"{stem}-slide-{i:0{width}d}.png"
            page.get_pixmap(dpi=dpi, alpha=False).save(p)
            paths.append(p)
        return paths
    if not shutil.which("pdftoppm"):
        raise RuntimeError("need PyMuPDF (pip install pymupdf) or pdftoppm to rasterize the PDF")
    subprocess.run(["pdftoppm", "-png", "-r", str(dpi), str(pdf), str(outdir / f"{stem}-slide")], check=True)
    return sorted(outdir.glob(f"{stem}-slide-*.png"))


def contact_sheet(paths, out, cols=3, thumb_w=480):
    thumbs = []
    for p in paths:
        with Image.open(p) as im:
            thumbs.append(im.resize((thumb_w, round(im.height * thumb_w / im.width))))
    th = thumbs[0].height
    pad, label = 16, 26
    rows = (len(thumbs) + cols - 1) // cols
    sheet = Image.new("RGB", (cols * (thumb_w + pad) + pad, rows * (th + label + pad) + pad), "white")
    draw = ImageDraw.Draw(sheet)
    try:
        font = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", 18)
    except OSError:
        font = ImageFont.load_default()
    for i, t in enumerate(thumbs):
        x = pad + (i % cols) * (thumb_w + pad)
        y = pad + (i // cols) * (th + label + pad)
        draw.text((x, y), f"Slide {i + 1}", fill=(40, 40, 40), font=font)
        sheet.paste(t, (x, y + label))
        draw.rectangle([x - 1, y + label - 1, x + thumb_w, y + label + th], outline=(190, 190, 190))
    sheet.save(out, quality=90)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("deck")
    ap.add_argument("--outdir", help="where to write images (default: <deck dir>/render)")
    ap.add_argument("--dpi", type=int, default=110)
    ap.add_argument("--engine", choices=["auto", "powerpoint", "libreoffice"], default="auto")
    ap.add_argument("--pptx-skill", help="pptx skill directory, for its LibreOffice wrapper")
    ap.add_argument("--project", help="project whose render folder PowerPoint uses (default: $CLAUDE_PROJECT_DIR, "
                                      "else the current folder; inside a git repo, its main checkout)")
    args = ap.parse_args()

    deck = Path(args.deck).expanduser().resolve()
    outdir = Path(args.outdir).expanduser().resolve() if args.outdir else deck.parent / "render"
    outdir.mkdir(parents=True, exist_ok=True)
    # Export beside the last render and replace it only once the new one is there, so a render that
    # fails (PowerPoint stuck on a dialog) leaves the last good images in place.
    fresh = outdir / f".{deck.stem}.rendering.pdf"
    fresh.unlink(missing_ok=True)

    engines = [args.engine] if args.engine != "auto" else (
        ["powerpoint", "libreoffice"] if sys.platform == "darwin" and POWERPOINT_APP.exists() else ["libreoffice"])
    errors = []
    for engine in engines:
        try:
            if engine == "powerpoint":
                via_powerpoint(deck, fresh, args.project)
            else:
                via_libreoffice(deck, fresh, args.pptx_skill)
            break
        except office_render.Stop as e:
            # 3: the render folder isn't granted yet; 4: PowerPoint waits on a dialog. Both need the user.
            if e.code in (3, 4):
                fresh.unlink(missing_ok=True)
                print(e, file=sys.stderr)
                sys.exit(e.code)
            errors.append(f"{engine}: {e}")
        except (RuntimeError, subprocess.TimeoutExpired, OSError) as e:
            errors.append(f"{engine}: {e}")
    else:
        fresh.unlink(missing_ok=True)
        sys.exit("could not render:\n  " + "\n  ".join(errors))

    for stale in outdir.glob(f"{deck.stem}-slide-*.png"):
        stale.unlink()
    pdf = outdir / f"{deck.stem}.pdf"
    fresh.replace(pdf)
    paths = pdf_to_pngs(pdf, outdir, deck.stem, args.dpi)
    sheet = outdir / f"{deck.stem}-overview.jpg"
    contact_sheet(paths, sheet)
    print(f"Rendered with {engine}: {len(paths)} slides")
    for p in paths:
        print(p)
    print(f"Overview: {sheet}")


if __name__ == "__main__":
    main()
