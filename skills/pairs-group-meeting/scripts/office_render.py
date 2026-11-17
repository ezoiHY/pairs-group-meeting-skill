#!/usr/bin/env python3
"""Render PowerPoint decks on macOS through one fixed render folder per project.

PowerPoint for Mac is sandboxed: the first time a script makes it save into a folder, it shows
"Grant File Access" and waits for the user, and it remembers every folder it was granted (with
everything inside). Rendering decks where they lie therefore asks again for each new folder. This
script does all PowerPoint work inside <project>/.office-render, so PowerPoint asks once per project.

    office_render.py --status [--project DIR]
    office_render.py --grant  [--project DIR] [--wait SECONDS]
    office_render.py DECK.pptx [DECK ...] [--project DIR] [--out DIR] [--png] [--dpi 110]

Exit codes: 0 done (--status: granted) · 2 usage error, no PowerPoint, or no Automation permission
            3 render folder not granted yet: run --grant · 4 PowerPoint is waiting on a dialog
            5 PowerPoint could not open or export the file

From Python, export_pdf(deck, pdf, project=...) does one render and raises Stop, whose .code is the
exit code above. Source: https://github.com/ezoiHY/office-render-skill. Other skills bundle copies of
this file; change it there first, then copy it over.
"""

from __future__ import annotations

import argparse
import glob
import itertools
import json
import os
import shlex
import shutil
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

__version__ = "0.1.0"

FOLDER = ".office-render"
MARKER = ".granted"
TMP = "office-render-"   # every temporary file in the render folder starts with this, or with ~$ and this
PP_APP = Path("/Applications/Microsoft PowerPoint.app")
PP_ID = "com.microsoft.Powerpoint"
SCRIPT = Path(__file__).resolve()
BLANK = SCRIPT.parent.parent / "assets" / "blank.pptx"
OPEN_WAIT = 90           # s for a deck to open
DIALOG_AFTER = 5         # s a new dialog-sized PowerPoint window must stay up, with no PDF yet, to count as a dialog
STALE_AFTER = 3600       # s after which temporary files nobody owns any more are deleted
_renders = itertools.count(1)

# The copy is named ~$…: PowerPoint's sandbox lets it read and write ~$* files anywhere (meant for Office's
# lock files), so opening never prompts. Not waiting for the reply keeps a dialog from hanging this call.
OPEN = """on run argv
	tell application "Microsoft PowerPoint"
		ignoring application responses
			open (POSIX file (item 1 of argv))
		end ignoring
	end tell
end run"""

# Variables inside a tell block must not collide with PowerPoint's own terms ("out" is one of its constants).
NAMES = """with timeout of 10 seconds
	tell application "Microsoft PowerPoint"
		set deckNames to {}
		repeat with i from 1 to (count of presentations)
			set end of deckNames to (name of presentation i)
		end repeat
	end tell
end timeout
set AppleScript's text item delimiters to linefeed
return deckNames as text"""

# Saving is what needs the folder grant. Retries cover big decks (embedded video) that refuse the first
# attempts; a timeout (-1712: stuck on a dialog) or a cancel (-128) is not retried.
EXPORT = """on run argv
	set tries to (item 3 of argv) as integer
	tell application "Microsoft PowerPoint"
		set pres to missing value
		repeat with i from 1 to (count of presentations)
			if name of presentation i is (item 2 of argv) then set pres to presentation i
		end repeat
		if pres is missing value then error "the deck is not open"
		set lastErr to ""
		repeat tries times
			try
				with timeout of 1800 seconds
					save pres in (POSIX file (item 1 of argv)) as save as PDF
				end timeout
				return "ok"
			on error errMsg number errNum
				if errNum is in {-1712, -128} then error errMsg number errNum
				set lastErr to errMsg
				delay 1
			end try
		end repeat
		error "save as PDF failed: " & lastErr
	end tell
end run"""

CLOSE = """on run argv
	with timeout of 20 seconds
		tell application "Microsoft PowerPoint"
			repeat with i from (count of presentations) to 1 by -1
				if (name of presentation i) is in argv then close presentation i saving no
			end repeat
		end tell
	end timeout
end run"""

ACTIVATE = 'tell application "Microsoft PowerPoint" to activate'

JS_FRONT = """ObjC.import('AppKit');
function run() {
  const app = $.NSWorkspace.sharedWorkspace.frontmostApplication;
  return app.isNil() ? '' : (ObjC.unwrap(app.bundleIdentifier) || '');
}"""

JS_HIDE = """ObjC.import('AppKit');
function run(argv) {
  const apps = $.NSRunningApplication.runningApplicationsWithBundleIdentifier(argv[0]);
  for (let i = 0; i < apps.count; i++) apps.objectAtIndex(i).hide;
  return apps.count;
}"""

# A window's owner and size need no screen-recording permission (its title would). "Grant File Access" starts
# as a normal window (layer 0) and becomes a modal panel (layer 8) once PowerPoint is active, so both count.
JS_WINDOWS = """ObjC.import('CoreGraphics');
function run() {
  const all = ObjC.deepUnwrap(ObjC.castRefToObject($.CGWindowListCopyWindowInfo(0, 0))) || [];
  return all.filter(w => w.kCGWindowOwnerName === 'Microsoft PowerPoint' && w.kCGWindowLayer < 20)
    .map(w => [w.kCGWindowNumber, Math.round(w.kCGWindowBounds.Width), Math.round(w.kCGWindowBounds.Height)].join(' '))
    .join('\\n');
}"""

# PDFKit draws each page into a white bitmap: no third-party packages needed.
JS_PDF2PNG = """ObjC.import('Cocoa');
ObjC.import('Quartz');
function run(argv) {
  const [pdfPath, outDir, stem, dpi] = argv;
  const scale = parseFloat(dpi) / 72;
  const doc = $.PDFDocument.alloc.initWithURL($.NSURL.fileURLWithPath(pdfPath));
  if (doc.isNil()) throw new Error('cannot read ' + pdfPath);
  const n = doc.pageCount, digits = String(n).length, out = [];
  for (let i = 0; i < n; i++) {
    const page = doc.pageAtIndex(i);
    const box = page.boundsForBox(0);
    const w = Math.round(box.size.width * scale), h = Math.round(box.size.height * scale);
    const rep = $.NSBitmapImageRep.alloc.initWithBitmapDataPlanesPixelsWidePixelsHighBitsPerSampleSamplesPerPixelHasAlphaIsPlanarColorSpaceNameBytesPerRowBitsPerPixel(
      null, w, h, 8, 4, true, false, $.NSDeviceRGBColorSpace, 0, 0);
    const cg = $.NSGraphicsContext.graphicsContextWithBitmapImageRep(rep).CGContext;
    $.CGContextSetRGBFillColor(cg, 1, 1, 1, 1);
    $.CGContextFillRect(cg, $.CGRectMake(0, 0, w, h));
    $.CGContextScaleCTM(cg, scale, scale);
    $.CGContextTranslateCTM(cg, -box.origin.x, -box.origin.y);
    page.drawWithBoxToContext(0, cg);
    const png = rep.representationUsingTypeProperties(4, $({}));
    const p = outDir + '/' + stem + '-slide-' + String(i + 1).padStart(digits, '0') + '.png';
    if (!png.writeToFileAtomically(p, true)) throw new Error('cannot write ' + p);
    out.push(p);
  }
  return out.join('\\n');
}"""

NO_POWERPOINT = ("Microsoft PowerPoint for Mac is not installed (/Applications/Microsoft PowerPoint.app); this "
                 "script needs it. Render another way, e.g. with LibreOffice.")
AUTOMATION_WAIT = ("osascript got no answer within 60 s. macOS is probably asking whether this app may control "
                   "Microsoft PowerPoint (Automation): ask the user to click Allow, then run again.")
AUTOMATION_DENIED = ("macOS does not let this app control PowerPoint. The user can allow it in System Settings > "
                     "Privacy & Security > Automation (Microsoft PowerPoint, under the app that runs the agent).")
DIALOG = ("PowerPoint is waiting on a dialog, so the render stopped. Most likely it is \"Grant File Access\" for "
          "{rd}: PowerPoint lost its access (the folder was deleted or moved, or PowerPoint was reset). It can "
          "also be a sign-in, update or recovery prompt. Its window may be on another desktop. Tell the user, "
          "then run --grant: it brings PowerPoint forward and waits while they answer. Don't force-quit PowerPoint.")
NOT_OPENED = ("PowerPoint did not open the deck within {wait:.0f} s. It is most likely waiting on a dialog "
              "(sign-in, update, recovery, or an unanswered \"Grant File Access\"), maybe on another desktop. "
              "Tell the user, then run --grant to bring PowerPoint forward. Don't force-quit PowerPoint.")
GRANT_WAIT = ("PowerPoint is still waiting on its dialog after {wait:.0f} s. Ask the user to answer it "
              "(Select... > the folder \"{name}\" > Grant Access), then run --grant again.")
GRANT_FAILED = ("The grant did not go through ({err}); the dialog was probably cancelled. Run --grant again and "
                "ask the user to choose Select..., select the folder \"{name}\" itself, then Grant Access.")


class Stop(Exception):
    def __init__(self, code, msg):
        super().__init__(msg)
        self.code = code


def run(cmd, timeout):
    return subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)


def osa(script, *args, timeout=30):
    return run(["osascript", "-e", script, *map(str, args)], timeout)


def jxa(script, *args, timeout=30):
    return run(["osascript", "-l", "JavaScript", "-e", script, *map(str, args)], timeout)


def front_app():
    try:
        return jxa(JS_FRONT, timeout=10).stdout.strip()
    except (OSError, subprocess.TimeoutExpired):
        return ""


def hide_powerpoint():
    try:
        jxa(JS_HIDE, PP_ID, timeout=10)
    except (OSError, subprocess.TimeoutExpired):
        pass


def activate_powerpoint():
    try:
        osa(ACTIVATE, timeout=10)
    except subprocess.TimeoutExpired:
        pass


def powerpoint_running():
    try:
        return run(["pgrep", "-x", "Microsoft PowerPoint"], 10).returncode == 0
    except (OSError, subprocess.TimeoutExpired):
        return False


def dialog_windows():
    """Numbers of PowerPoint windows the size of a dialog. "Grant File Access" is about 456 x 299;
    slide windows are far bigger and toolbars far thinner."""
    try:
        out = jxa(JS_WINDOWS, timeout=10).stdout
    except (OSError, subprocess.TimeoutExpired):
        return set()
    found = set()
    for line in out.splitlines():
        parts = line.split()
        if len(parts) == 3 and all(p.isdigit() for p in parts):
            num, w, h = map(int, parts)
            if 250 <= w <= 800 and 150 <= h <= 650:
                found.add(num)
    return found


def open_names():
    try:
        return osa(NAMES, timeout=15).stdout.splitlines()
    except subprocess.TimeoutExpired:
        return []


class KeepFocus:
    """Hands focus back if PowerPoint took it (when launching, say), but never while it shows a dialog:
    hiding PowerPoint would hide the dialog the user has to answer."""

    def __enter__(self):
        self.front = front_app()
        return self

    def __exit__(self, *exc):
        if self.front and self.front != PP_ID and front_app() == PP_ID and not dialog_windows():
            hide_powerpoint()


def has_git():
    # On a Mac without the developer tools, /usr/bin/git pops up an installer dialog; xcode-select -p doesn't.
    git = shutil.which("git")
    if not git:
        return False
    if sys.platform != "darwin" or Path(git).resolve() != Path("/usr/bin/git"):
        return True
    try:
        return run(["xcode-select", "-p"], 10).returncode == 0
    except (OSError, subprocess.TimeoutExpired):
        return False


def project_root(project=None):
    start = Path(project or os.environ.get("CLAUDE_PROJECT_DIR") or os.getcwd()).expanduser().resolve()
    if not has_git():
        return start
    try:
        r = run(["git", "-C", str(start), "rev-parse", "--path-format=absolute", "--git-common-dir",
                 "--show-toplevel"], 15)
    except (OSError, subprocess.TimeoutExpired):
        return start
    lines = r.stdout.splitlines()
    if r.returncode != 0 or len(lines) < 2:
        return start
    common, top = Path(lines[0]), Path(lines[1])
    # A worktree's common dir is the main checkout's .git, so every worktree shares one render folder.
    return common.parent if common.name == ".git" else top


def render_folder(project=None, render_dir=None):
    """(project root, render folder). A render_dir is used as it is, with no project."""
    if render_dir:
        return None, Path(render_dir).expanduser().resolve()
    root = project_root(project)
    return root, root / FOLDER


def prepare(rd):
    rd.mkdir(parents=True, exist_ok=True)
    ignore = rd / ".gitignore"
    if not ignore.exists():
        ignore.write_text("# PowerPoint render folder (office-render skill), not part of the project\n*\n")


def granted(rd):
    return (rd / MARKER).exists()


def mark_granted(rd):
    try:
        version = run(["defaults", "read", str(PP_APP / "Contents/Info.plist"), "CFBundleShortVersionString"],
                      10).stdout.strip()
    except (OSError, subprocess.TimeoutExpired):
        version = ""
    (rd / MARKER).write_text(json.dumps({"granted": datetime.now().isoformat(timespec="seconds"),
                                         "powerpoint": version}) + "\n")


def grant_command(root, rd):
    where = f"--project {shlex.quote(str(root))}" if root else f"--render-dir {shlex.quote(str(rd))}"
    return f"python3 {shlex.quote(str(SCRIPT))} --grant {where}"


def not_granted(root, rd):
    return Stop(3, f"PowerPoint has not been allowed into this project's render folder yet:\n  {rd}\n"
                   "Tell the user that PowerPoint will ask once, in \"Grant File Access\": click Select..., "
                   f"select the folder \"{rd.name}\" itself, not a file in it (it is hidden: press Cmd-Shift-. "
                   "to show it, or select the project folder above it), then Grant Access. This project will "
                   f"not ask again. Then run:\n  {grant_command(root, rd)}")


def alive(pid):
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        pass
    return True


def owner_pid(name):
    head = name.lstrip("~$")[len(TMP):].split("-", 1)[0]
    return int(head) if head.isdigit() else None


def tidy(rd):
    """Close decks left open by renders that died, then delete the temporary files they left behind.
    While PowerPoint shows a dialog a close would only queue behind it, so that waits for a later run."""
    if powerpoint_running() and not dialog_windows():
        stale = [n for n in open_names()
                 if n.startswith("~$" + TMP) and owner_pid(n) is not None and not alive(owner_pid(n))]
        if stale:
            try:
                osa(CLOSE, *stale, timeout=30)
            except subprocess.TimeoutExpired:
                pass
    now = time.time()
    for p in rd.iterdir():
        if not (p.is_file() and p.name.lstrip("~$").startswith(TMP)):
            continue
        pid = owner_pid(p.name)
        if (pid is not None and not alive(pid)) or now - p.stat().st_mtime > STALE_AFTER:
            p.unlink(missing_ok=True)


def pdf_complete(p):
    try:
        with open(p, "rb") as f:
            if f.read(5) != b"%PDF-":
                return False
            f.seek(max(0, p.stat().st_size - 2048))
            return b"%%EOF" in f.read()
    except OSError:
        return False


def export(src, rd, pdf, tag, wait, grant):
    """Copy src into the render folder, have PowerPoint save it as pdf there, close it, delete the copy.
    In grant mode a dialog is expected: PowerPoint is brought forward and the user gets `wait` seconds."""
    copy = rd / f"~${TMP}{os.getpid()}-{tag}{src.suffix}"
    shutil.copy2(src, copy)
    name, opened, stuck = copy.name, False, False
    blocked_before = dialog_windows()   # a dialog already up would hold back our requests too
    try:
        try:
            r = osa(OPEN, copy, timeout=60)
        except subprocess.TimeoutExpired:
            stuck = True
            raise Stop(4, AUTOMATION_WAIT)
        if r.returncode != 0:
            raise Stop(2, AUTOMATION_DENIED) if "-1743" in r.stderr else Stop(5, f"PowerPoint could not open "
                                                                                  f"{src.name}: {r.stderr.strip()}")
        t0 = time.monotonic()
        while name not in open_names():
            waited = time.monotonic() - t0
            if waited > (wait if grant else OPEN_WAIT) or (blocked_before and waited > 10 and not grant):
                stuck = True
                raise Stop(4, NOT_OPENED.format(wait=waited))
            time.sleep(0.5)
        opened = True
        before = dialog_windows()
        proc = subprocess.Popen(["osascript", "-e", EXPORT, str(pdf), name, "1" if grant else "30"],
                                stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        t0, since, shown = time.monotonic(), None, False
        while proc.poll() is None:
            new = dialog_windows() - before
            since = (since or time.monotonic()) if new and not pdf.exists() else None
            if since and time.monotonic() - since >= DIALOG_AFTER:
                if not grant:
                    proc.kill()
                    stuck = True
                    raise Stop(4, DIALOG.format(rd=rd))
                if not shown:
                    activate_powerpoint()
                    print(f"PowerPoint is asking for access to {rd}; waiting for the user...", file=sys.stderr)
                    shown = True
            if time.monotonic() - t0 > wait:
                proc.kill()
                stuck = True
                raise Stop(4, GRANT_WAIT.format(wait=wait, name=rd.name) if grant else DIALOG.format(rd=rd))
            time.sleep(0.5)
        _, err = proc.communicate()
        if proc.returncode != 0 or not pdf_complete(pdf):
            err = err.strip() or "no PDF was written"
            if grant:
                raise Stop(4, GRANT_FAILED.format(err=err, name=rd.name))
            raise Stop(5, f"PowerPoint could not export {src.name}: {err}")
    finally:
        # A deck stuck behind a dialog stays open; the next run closes it once this process is gone.
        if opened and not stuck:
            try:
                osa(CLOSE, name, timeout=30)
            except subprocess.TimeoutExpired:
                pass
        copy.unlink(missing_ok=True)
        if not stuck:
            (rd / f"~${name}").unlink(missing_ok=True)   # PowerPoint's own lock file for the copy


def do_grant(rd, wait):
    if not BLANK.exists():
        raise Stop(5, f"missing {BLANK}")
    if dialog_windows():
        # PowerPoint already waits on a dialog (from an earlier render, say). Show it and let the user answer
        # it first: anything sent now would queue behind it.
        activate_powerpoint()
        t0 = time.monotonic()
        while dialog_windows():
            if time.monotonic() - t0 > wait:
                raise Stop(4, GRANT_WAIT.format(wait=wait, name=rd.name))
            time.sleep(1)
    pdf = rd / f"{TMP}{os.getpid()}-grant.pdf"
    export(BLANK, rd, pdf, "grant", wait, grant=True)
    pdf.unlink(missing_ok=True)
    mark_granted(rd)
    print(f"PowerPoint may now use {rd}; renders in this project will not ask again.")
    return 0


def export_pdf(deck, pdf, project=None, render_dir=None, wait=600):
    """Render deck to pdf with PowerPoint, through the project's render folder. Raises Stop."""
    deck, pdf = Path(deck).expanduser().resolve(), Path(pdf).expanduser().resolve()
    if sys.platform != "darwin" or not PP_APP.exists():
        raise Stop(2, NO_POWERPOINT)
    if not deck.is_file():
        raise Stop(2, f"no such file: {deck}")
    root, rd = render_folder(project, render_dir)
    prepare(rd)
    if not granted(rd):
        raise not_granted(root, rd)
    with KeepFocus():
        tidy(rd)
        tag = str(next(_renders))
        tmp = rd / f"{TMP}{os.getpid()}-{tag}.pdf"
        try:
            export(deck, rd, tmp, tag, wait, grant=False)
            pdf.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(str(tmp), str(pdf))
        finally:
            tmp.unlink(missing_ok=True)
    return pdf


def render(deck, rd, args):
    deck = deck.expanduser().resolve()
    t0 = time.monotonic()
    out = args.out.expanduser().resolve() if args.out else rd / deck.stem
    pdf = export_pdf(deck, out / f"{deck.stem}.pdf", render_dir=rd, wait=args.wait or 600)
    print(f"Rendered {deck.name} in {time.monotonic() - t0:.1f} s")
    print(f"PDF: {pdf}")
    if args.png:
        for old in out.glob(f"{glob.escape(deck.stem)}-slide-*.png"):
            old.unlink()
        r = jxa(JS_PDF2PNG, pdf, out, deck.stem, args.dpi, timeout=600)
        if r.returncode != 0:
            raise Stop(5, f"could not turn {pdf} into PNGs: {r.stderr.strip()}")
        for line in r.stdout.splitlines():
            print(f"PNG: {line}")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("decks", nargs="*", type=Path, help="decks to render (.pptx or anything PowerPoint opens)")
    ap.add_argument("--project", help="project folder (default: $CLAUDE_PROJECT_DIR, else the current folder); "
                                      "inside a git repo the main checkout is used")
    ap.add_argument("--render-dir", help="use exactly this folder instead of <project>/.office-render")
    ap.add_argument("--status", action="store_true", help="show the render folder and whether it is granted")
    ap.add_argument("--grant", action="store_true", help="one-time setup: the user grants PowerPoint the folder")
    ap.add_argument("--out", type=Path, help="where to write <deck>.pdf and PNGs (default: <render folder>/<deck>/)")
    ap.add_argument("--png", action="store_true", help="also write one PNG per slide")
    ap.add_argument("--dpi", type=int, default=110, help="PNG resolution (default 110)")
    ap.add_argument("--wait", type=float, help="seconds to wait for PowerPoint (default 300 with --grant, else 600)")
    ap.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    args = ap.parse_args()
    if args.status + args.grant + bool(args.decks) != 1:
        ap.error("give exactly one of --status, --grant, or decks to render")
    try:
        sys.exit(dispatch(args))
    except Stop as e:
        print(e, file=sys.stderr)
        sys.exit(e.code)


def dispatch(args):
    if sys.platform != "darwin" or not PP_APP.exists():
        raise Stop(2, NO_POWERPOINT)
    root, rd = render_folder(args.project, args.render_dir)
    if args.status:
        print(f"project:       {root or '(set with --render-dir)'}")
        print(f"render folder: {rd}")
        if granted(rd):
            try:
                when = json.loads((rd / MARKER).read_text()).get("granted", "?")
            except (OSError, ValueError, AttributeError):
                when = "?"
            print(f"granted:       yes, since {when}")
            return 0
        print(f"granted:       no; run once, after telling the user: {grant_command(root, rd)}")
        return 3
    prepare(rd)
    with KeepFocus():
        if args.grant:
            tidy(rd)
            return do_grant(rd, args.wait or 300)
        if not granted(rd):
            raise not_granted(root, rd)
        for deck in args.decks:
            render(deck, rd, args)
        return 0


if __name__ == "__main__":
    main()
