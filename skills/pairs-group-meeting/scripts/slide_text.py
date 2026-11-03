#!/usr/bin/env python3
"""Print the text of every slide in one or more decks, in slide order.

Used to check a built deck's content, and to read the previous meeting's deck
(its Todo slide is this meeting's Done candidate list). Shapes are read top to
bottom, and a line break inside a paragraph (Shift-Enter) starts a new line, as
it does on the slide. Tables print one row per line, charts their categories and
series, and the speaker notes follow the slide's text.

Usage:
    python slide_text.py deck.pptx [more.pptx ...]
"""

import sys

from pptx import Presentation
from pptx.shapes.group import GroupShape


def reading_order(shapes):
    return sorted(shapes, key=lambda sh: (sh.top or 0, sh.left or 0))


def shape_lines(shape):
    if isinstance(shape, GroupShape):
        for child in reading_order(shape.shapes):
            yield from shape_lines(child)
    elif shape.has_text_frame:
        for p in shape.text_frame.paragraphs:
            # python-pptx gives a line break inside a paragraph as \v.
            yield from (line.strip() for line in p.text.split("\v") if line.strip())
    elif getattr(shape, "has_table", False) and shape.has_table:
        for row in shape.table.rows:
            yield " | ".join(" ".join(cell.text.split()) for cell in row.cells)
    elif getattr(shape, "has_chart", False) and shape.has_chart:
        for plot in shape.chart.plots:
            yield "chart: " + ", ".join(map(str, plot.categories))
            yield "series: " + ", ".join(str(series.name) for series in plot.series)


def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    for path in sys.argv[1:]:
        prs = Presentation(path)
        if len(sys.argv) > 2:
            print(f"=== {path}")
        for i, slide in enumerate(prs.slides, 1):
            print(f"--- slide {i}")
            for shape in reading_order(slide.shapes):
                for line in shape_lines(shape):
                    print(line)
            notes = slide.notes_slide.notes_text_frame if slide.has_notes_slide else None
            if notes is not None and notes.text.strip():
                print("notes: " + " ".join(notes.text.split()))


if __name__ == "__main__":
    main()
