# /// script
# requires-python = ">=3.10"
# dependencies = ["python-pptx==1.0.2"]
# ///
"""Write a small brand template for testing template mode.

    uv run make_template.py brand.pptx [--logo-over-title]

It has what real company templates have and the default one lacks: its own
theme colours and fonts, and a logo on the master in the top-right corner.
The content layouts' titles are narrowed to clear the logo, as a designer
would; --logo-over-title leaves them full width, so the lint should catch
every title covering the logo.
"""

import argparse
import os
import sys

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "..", "scripts"))

import ooxml  # noqa: E402
import themes  # noqa: E402

BRAND = {"base": "paper", "colors": {"bg": "FFFFFF", "text": "14213D", "muted": "4A5568",
                                     "accent": "B5179E", "series": "3A0CA3"},
         "fonts": {"major": "Arial", "minor": "Arial"}, "title_bold": True}
LOGO = (Inches(11.9), Inches(0.3), Inches(0.95), Inches(0.5))
LOGO_CLEARANCE = Inches(0.25)
NARROWED = ("Title Only", "Title and Content", "Two Content", "Comparison")


def make_template(path, logo_over_title=False):
    prs = Presentation()
    ooxml.scale_master(prs, themes.SLIDE_W / prs.slide_width, themes.SLIDE_H / prs.slide_height)
    prs.slide_width, prs.slide_height = themes.SLIDE_W, themes.SLIDE_H
    ooxml.apply_theme(prs, themes.resolve_theme(BRAND))
    # python-pptx cannot add shapes to a master, so the logo is drawn on a
    # throwaway slide and its element moved across.
    scratch = prs.slides.add_slide(prs.slide_layouts[6])
    logo = scratch.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, *LOGO)
    logo.name = "Brand logo"
    logo.text_frame.text = "BRAND"
    logo.text_frame.word_wrap = False
    logo.text_frame.paragraphs[0].runs[0].font.size = Pt(12)
    master_tree = prs.slide_master.shapes._spTree
    logo._element.nvSpPr.cNvPr.set("id", str(max(int(e.get("id")) for e in master_tree.iter(qn("p:cNvPr"))) + 1))
    master_tree.append(logo._element)
    ooxml.remove_all_slides(prs)
    if not logo_over_title:
        for lay in prs.slide_layouts:
            if lay.name not in NARROWED:
                continue
            for ph in lay.placeholders:
                if ph._element.ph.get("type") == "title":
                    # All four at once: an inherited placeholder given only a
                    # width gets a zero height.
                    left, top, height = ph.left, ph.top, ph.height
                    ph.left, ph.top, ph.width, ph.height = left, top, LOGO[0] - LOGO_CLEARANCE - left, height
    prs.save(path)
    return path


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("out")
    ap.add_argument("--logo-over-title", action="store_true")
    args = ap.parse_args(argv)
    make_template(args.out, args.logo_over_title)
    print(f"Wrote {args.out}")


if __name__ == "__main__":
    main()
