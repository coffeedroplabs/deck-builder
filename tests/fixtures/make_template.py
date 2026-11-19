# /// script
# requires-python = ">=3.10"
# dependencies = ["python-pptx==1.0.2"]
# ///
"""Write a small brand template for testing template mode.

    uv run make_template.py brand.pptx [--logo-over-title] [--footer-logo] [--high-footer]

It has what real company templates have and the default one lacks: its own
theme colours and fonts, and a logo on the master in the top-right corner.
The content layouts' titles are narrowed to clear the logo, as a designer
would; --logo-over-title leaves them full width, so the lint should catch
every title covering the logo. --footer-logo adds a second mark bottom left,
where the date placeholder sits: clear of the footer and slide-number
placeholders, but inside the default grid's footer row. --high-footer raises
the footer row above the default grid's, as some company templates do.
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
FOOTER_LOGO = (Inches(0.67), Inches(6.9), Inches(1.4), Inches(0.45))
LOGO_CLEARANCE = Inches(0.25)
HIGH_FOOTER_Y = Inches(6.6)
NARROWED = ("Title Only", "Title and Content", "Two Content", "Comparison")


def _add_master_art(prs, name, box, text):
    # python-pptx cannot add shapes to a master, so the art is drawn on a
    # throwaway slide and its element moved across.
    scratch = prs.slides.add_slide(prs.slide_layouts[6])
    art = scratch.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, *box)
    art.name = name
    art.text_frame.text = text
    art.text_frame.word_wrap = False
    art.text_frame.paragraphs[0].runs[0].font.size = Pt(12)
    master_tree = prs.slide_master.shapes._spTree
    art._element.nvSpPr.cNvPr.set("id", str(max(int(e.get("id")) for e in master_tree.iter(qn("p:cNvPr"))) + 1))
    master_tree.append(art._element)
    ooxml.remove_all_slides(prs)


def make_template(path, logo_over_title=False, footer_logo=False, accent=None, high_footer=False):
    prs = Presentation()
    ooxml.scale_master(prs, themes.SLIDE_W / prs.slide_width, themes.SLIDE_H / prs.slide_height)
    prs.slide_width, prs.slide_height = themes.SLIDE_W, themes.SLIDE_H
    brand = {**BRAND, "colors": {**BRAND["colors"], **({"accent": accent} if accent else {})}}
    ooxml.apply_theme(prs, themes.resolve_theme(brand))
    _add_master_art(prs, "Brand logo", LOGO, "BRAND")
    if footer_logo:
        _add_master_art(prs, "Footer mark", FOOTER_LOGO, "brand.example")
    if high_footer:
        # As in the Fernhill eval template, the content placeholders end where
        # the raised footer row starts.
        for part in [prs.slide_master, *prs.slide_layouts]:
            for ph in part.placeholders:
                kind = ph._element.ph.get("type", "obj")
                left, top, width, height = ph.left, ph.top, ph.width, ph.height
                if None in (left, top, width, height):
                    continue
                if kind in ("dt", "ftr", "sldNum"):
                    ph.left, ph.top, ph.width, ph.height = left, HIGH_FOOTER_Y, width, height
                elif kind not in ("title", "ctrTitle") and top + height > HIGH_FOOTER_Y:
                    ph.left, ph.top, ph.width, ph.height = left, top, width, HIGH_FOOTER_Y - top
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
    ap.add_argument("--footer-logo", action="store_true")
    ap.add_argument("--high-footer", action="store_true")
    args = ap.parse_args(argv)
    make_template(args.out, args.logo_over_title, args.footer_logo, high_footer=args.high_footer)
    print(f"Wrote {args.out}")


if __name__ == "__main__":
    main()
