#!/usr/bin/env python3
"""Render a .pptx to one PNG per slide, plus a contact sheet, for visual QA.

Engines, in the order `auto` tries them:
  powerpoint   Microsoft PowerPoint on macOS, driven by AppleScript. Renders exactly what the
               audience will see (real Calibri / DengXian, real video posters).
  libreoffice  soffice, through the pptx skill's wrapper when --pptx-skill is given.

Usage:
    python render.py deck.pptx [--outdir DIR] [--dpi 110] [--engine auto|powerpoint|libreoffice]
                     [--pptx-skill DIR]

Prints the absolute path of every slide image and of the contact sheet.
"""

import argparse
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

POWERPOINT_APP = Path("/Applications/Microsoft PowerPoint.app")
# A large deck (embedded video) is not ready the moment `open` returns, so wait for it and
# retry the export. A deck the user already has open is exported but never closed.
EXPORT_SCRIPT = """
on run argv
	set inPath to item 1 of argv
	set outPath to item 2 of argv
	set deckName to item 3 of argv
	tell application "Microsoft PowerPoint"
		set wasOpen to false
		repeat with i from 1 to (count of presentations)
			if name of presentation i is deckName then set wasOpen to true
		end repeat
		if not wasOpen then open (POSIX file inPath)
		set pres to missing value
		repeat 120 times
			repeat with i from 1 to (count of presentations)
				if name of presentation i is deckName then set pres to presentation i
			end repeat
			if pres is not missing value then exit repeat
			delay 0.5
		end repeat
		if pres is missing value then error "deck did not open"
		set didSave to false
		set lastErr to ""
		repeat 30 times
			try
				save pres in (POSIX file outPath) as save as PDF
				set didSave to true
				exit repeat
			on error errMsg
				set lastErr to errMsg
				delay 1
			end try
		end repeat
		if not wasOpen then close pres saving no
		if not didSave then error "save as PDF failed: " & lastErr
	end tell
end run
"""


def via_powerpoint(deck, pdf):
    with tempfile.NamedTemporaryFile("w", suffix=".applescript", delete=False) as f:
        f.write(EXPORT_SCRIPT)
    try:
        res = subprocess.run(["osascript", f.name, str(deck), str(pdf), deck.name],
                             capture_output=True, text=True, timeout=300)
    finally:
        Path(f.name).unlink(missing_ok=True)
    if res.returncode != 0 or not pdf.exists():
        raise RuntimeError(f"PowerPoint export failed: {res.stderr.strip()}")


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
    args = ap.parse_args()

    deck = Path(args.deck).expanduser().resolve()
    outdir = Path(args.outdir).expanduser().resolve() if args.outdir else deck.parent / "render"
    outdir.mkdir(parents=True, exist_ok=True)
    for stale in outdir.glob(f"{deck.stem}-slide-*.png"):
        stale.unlink()
    pdf = outdir / f"{deck.stem}.pdf"
    pdf.unlink(missing_ok=True)

    engines = [args.engine] if args.engine != "auto" else (
        ["powerpoint", "libreoffice"] if sys.platform == "darwin" and POWERPOINT_APP.exists() else ["libreoffice"])
    errors = []
    for engine in engines:
        try:
            if engine == "powerpoint":
                via_powerpoint(deck, pdf)
            else:
                via_libreoffice(deck, pdf, args.pptx_skill)
            break
        except (RuntimeError, subprocess.TimeoutExpired, OSError) as e:
            errors.append(f"{engine}: {e}")
    else:
        sys.exit("could not render:\n  " + "\n  ".join(errors))

    paths = pdf_to_pngs(pdf, outdir, deck.stem, args.dpi)
    sheet = outdir / f"{deck.stem}-overview.jpg"
    contact_sheet(paths, sheet)
    print(f"Rendered with {engine}: {len(paths)} slides")
    for p in paths:
        print(p)
    print(f"Overview: {sheet}")


if __name__ == "__main__":
    main()
