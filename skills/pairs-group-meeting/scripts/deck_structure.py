"""Deck structure for build_deck.py, made with the pptx skill's own scripts.

add_slide.py clones the template's sample slides in spec order and clean.py drops
the unused originals. The finished deck is checked with the pptx skill's
validate.py against the template.

This is the Claude version. The Codex version of the skill has its own
deck_structure.py with the same three functions and no pptx skill.
"""

import os
import re
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path


def add_arguments(ap):
    ap.add_argument("--pptx-skill", help="base directory of the pptx skill (auto-detected under ~/.claude)")


def find_pptx_skill(given):
    candidates = []
    if given:
        candidates.append(Path(given).expanduser())
    if os.environ.get("PPTX_SKILL_DIR"):
        candidates.append(Path(os.environ["PPTX_SKILL_DIR"]).expanduser())
    claude = Path.home() / ".claude"
    for pattern in ("skills/pptx", "skills/*/pptx", "skills/*/*/pptx", "plugins/**/skills/pptx"):
        candidates.extend(sorted(claude.glob(pattern)))
    for c in candidates:
        if (c / "scripts" / "add_slide.py").exists():
            return c.resolve()
    raise SystemExit("pptx skill not found. Load it with the Skill tool and pass its base directory "
                     "as --pptx-skill (or set PPTX_SKILL_DIR).")


def run(cmd):
    res = subprocess.run(cmd, capture_output=True, text=True)
    if res.returncode != 0:
        raise SystemExit(f"{' '.join(map(str, cmd))} failed:\n{res.stdout}{res.stderr}")
    return res.stdout


def build(args, template, sources, originals, out_path):
    """Write a deck made of the template slides named in `sources`, in that order, without `originals`."""
    scripts = find_pptx_skill(args.pptx_skill) / "scripts"
    sys.path.insert(0, str(scripts))
    from office.helpers import rezip, safe_extract

    with tempfile.TemporaryDirectory(prefix="gm_struct_") as tmp:
        unpacked = Path(tmp) / "deck"
        with zipfile.ZipFile(template) as zf:
            safe_extract(zf, unpacked)
        for name in sources:
            run([sys.executable, str(scripts / "add_slide.py"), str(unpacked), name])

        rels = (unpacked / "ppt" / "_rels" / "presentation.xml.rels").read_text(encoding="utf-8")
        dropped = set()
        for rel in re.findall(r"<Relationship\b[^>]*>", rels):
            target = re.search(r'Target="([^"]+)"', rel).group(1)
            if target.rsplit("/", 1)[-1] in originals and "/slides/" in "/" + target:
                dropped.add(re.search(r'Id="([^"]+)"', rel).group(1))
        pres_path = unpacked / "ppt" / "presentation.xml"
        pres = pres_path.read_text(encoding="utf-8")
        pres = re.sub(r'<p:sldId\b[^>]*r:id="([^"]+)"[^>]*/>',
                      lambda m: "" if m.group(1) in dropped else m.group(0), pres)
        pres_path.write_text(pres, encoding="utf-8")

        run([sys.executable, str(scripts / "clean.py"), str(unpacked)])
        rezip(unpacked, out_path)


def validate(args, out_path, template):
    """(ok, report) from the pptx skill's validate.py."""
    scripts = find_pptx_skill(args.pptx_skill) / "scripts"
    res = subprocess.run([sys.executable, str(scripts / "office" / "validate.py"), str(out_path),
                          "--original", str(template)], capture_output=True, text=True)
    return res.returncode == 0, (res.stdout + res.stderr).strip()
