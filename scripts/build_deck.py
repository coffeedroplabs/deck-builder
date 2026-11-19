# /// script
# requires-python = ">=3.10"
# dependencies = ["python-pptx==1.0.2", "pillow==13.0.0"]
# ///
"""Build an editable .pptx from a deck spec.

    uv run build_deck.py spec.json -o deck.pptx
    uv run build_deck.py spec.json --check-only

The spec is checked first, then every text is fitted to its box at the
profile's sizes. Any error (a spec problem, or text that does not fit at the
smallest allowed size) stops the build before a file is written, and the
report says which slide and field to change. Exit 0 built, 1 refused, 2 crashed.
"""

import argparse
import hashlib
import json
import os
import sys
from datetime import datetime, timezone

from pptx import Presentation
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.oxml.ns import qn

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import charts  # noqa: E402
import layout  # noqa: E402
import ooxml  # noqa: E402
import spec_check  # noqa: E402
import themes  # noqa: E402

DEFAULT_TEMPLATE_LAYOUTS = {"cover": "Title Slide", "section": "Section Header"}
CONTENT_TEMPLATE_LAYOUTS = ("Title and Content", "Title Only")
TITLE_ONLY_INDEX = 5  # in python-pptx's default template
BODY_PLACEHOLDERS = {"body", "obj", "subTitle"}
ANCHORS = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}
ALIGN = {"left": "l", "right": "r", "center": "ctr", "l": "l", "r": "r", "ctr": "ctr"}
HEADER_RULE = (1.0, "tx1")
ROW_RULE = (0.75, "bg2")
TABLE_ALT_MAX_ROWS = 12


def sha256_file(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()


def _box(shape):
    return (shape.left, shape.top, shape.width, shape.height)


# -- presentation setup -------------------------------------------------------

def open_deck(spec, spec_dir):
    """(prs, layout chooser, frames, theme fonts, title bold, problems)."""
    problems = []
    if not spec.get("template"):
        prs = Presentation()
        # python-pptx's template is 4:3; scale its master before widening.
        ooxml.scale_master(prs, themes.SLIDE_W / prs.slide_width, themes.SLIDE_H / prs.slide_height)
        prs.slide_width, prs.slide_height = themes.SLIDE_W, themes.SLIDE_H
        theme = themes.resolve_theme(spec.get("theme"))
        ooxml.apply_theme(prs, theme)
        title_only = prs.slide_layouts[TITLE_ONLY_INDEX]
        return prs, (lambda name: title_only), {}, theme["title_bold"], problems

    path = os.path.join(spec_dir, spec["template"])
    prs = ooxml.open_presentation(path)
    ooxml.remove_all_slides(prs)
    expected = spec.get("template_fingerprint")
    if expected and expected != ooxml.fingerprint(prs):
        problems.append(_problem("template-changed", None, "template",
                                 "the template's layouts or theme changed since this spec was written; "
                                 "re-run inspect_template.py and check layout_map"))
    by_name = {lay.name: lay for lay in prs.slide_layouts}
    mapping = spec.get("layout_map") or {}
    content_default = next((by_name[n] for n in CONTENT_TEMPLATE_LAYOUTS if n in by_name), None)
    chosen, frames = {}, {}
    used = {s["layout"] for s in spec["slides"]}
    for name in sorted(used):
        wanted = mapping.get(name) or DEFAULT_TEMPLATE_LAYOUTS.get(name)
        lay = by_name.get(wanted) if wanted else content_default
        if lay is None and wanted and name not in mapping:
            lay = content_default
        if lay is None:
            problems.append(_problem("template-layout-missing", None, f"layout_map.{name}",
                                     f"the template has no layout {wanted or CONTENT_TEMPLATE_LAYOUTS[0]!r}; "
                                     f"it has: {', '.join(by_name)}"))
            continue
        chosen[name] = lay
        frame = {"title": None, "body": None}
        for ph in lay.placeholders:
            kind = ph._element.ph.get("type", "obj")
            part = "title" if kind in ("title", "ctrTitle") else "body" if kind in BODY_PLACEHOLDERS else None
            if part and frame[part] is None:
                frame[part] = _box(ph)
                frame[f"{part}_align"], frame[f"{part}_anchor"] = ooxml.placeholder_text_style(ph, lay.slide_master)
        for kind, part in (("ftr", "footer"), ("sldNum", "number")):
            ph = ooxml.footer_placeholders(lay).get(kind)
            if ph is not None:
                frame[part] = _box(ph)
                frame[f"{part}_align"], frame[f"{part}_anchor"] = ooxml.placeholder_text_style(ph, lay.slide_master)
        frames[name] = frame
    return prs, (lambda name: chosen.get(name)), frames, _template_title_bold(prs), problems


def _template_title_bold(prs):
    lvl1 = prs.slide_master._element.find(f"{qn('p:txStyles')}/{qn('p:titleStyle')}/{qn('a:lvl1pPr')}")
    rpr = lvl1.find(qn("a:defRPr")) if lvl1 is not None else None
    return rpr is not None and rpr.get("b") in ("1", "true")


def _problem(code, slide, field, message, severity="error"):
    return {"severity": severity, "code": code, "slide": slide, "field": field, "message": message}


# -- drawing ------------------------------------------------------------------

def _prepare_frame(tf, anchor):
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    if anchor:
        tf.vertical_anchor = ANCHORS[anchor]


# PowerPoint's own palette shades of a theme colour ("Darker 25%", "Darker 50%";
# "Lighter 40%", "Lighter 60%"), nearest first, so shaded text still follows the theme.
DARKER_SHADES = ((("lumMod", 75000),), (("lumMod", 50000),))
LIGHTER_SHADES = ((("lumMod", 60000), ("lumOff", 40000)), (("lumMod", 40000), ("lumOff", 60000)))


def text_ink(info, bg):
    """Returns ink(scheme, size, bold) -> (scheme, mods) for one slide background.

    A template's brand accent is often too light for text on white (a green
    that passes as a bar fails as a word), so a colour that fails contrast is
    written as its nearest palette shade that passes, and as tx1 if none does.
    """
    light_bg = bool(bg) and themes.contrast_ratio(bg, "000000") >= themes.contrast_ratio(bg, "FFFFFF")
    shades = DARKER_SHADES if light_bg else LIGHTER_SHADES

    def ink(scheme, size, bold):
        base = ooxml.scheme_hex(info, scheme)
        if not (base and bg):
            return scheme, ()
        need = themes.contrast_need(size, bold)
        for mods in ((),) + shades:
            if themes.contrast_ratio(ooxml.apply_mods(base, mods), bg) >= need:
                return scheme, mods
        return "tx1", ()
    return ink


def _write_paragraphs(tf, paras, ink):
    for i, p in enumerate(paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        ooxml.set_paragraph(para, align=ALIGN[p["align"]], indent_pt=p["indent"], hanging_pt=p["hanging"],
                            line_spacing=p["line_spacing"], space_before_pt=p["space_before"],
                            bullet=p["bullet"])
        font_ref = "+mj-lt" if p["font"] == "major" else "+mn-lt"
        for j, segment in enumerate(p["text"].replace("\v", "\n").split("\n")):
            if j:
                para.add_line_break()
            run = para.add_run()
            run.text = segment
            color, mods = ink(p["color"], p["size"], p["bold"])
            ooxml.style_run_props(run._r.get_or_add_rPr(), size_pt=p["size"], bold=p["bold"],
                                  italic=p["italic"], color=color, mods=mods, font_ref=font_ref)


def draw_text(slide, el, ink):
    shape = None
    if el["placeholder"]:
        shape = slide.shapes.title
    if shape is None:
        shape = slide.shapes.add_textbox(*el["box"])
    else:
        shape.left, shape.top, shape.width, shape.height = el["box"]
    shape.name = el["name"]
    tf = shape.text_frame
    tf.clear()
    _prepare_frame(tf, el["anchor"])
    _write_paragraphs(tf, el["paras"], ink)
    return shape


def draw_slide_number(slide, el, number, ink):
    shape = slide.shapes.add_textbox(*el["box"])
    shape.name = el["name"]
    tf = shape.text_frame
    _prepare_frame(tf, el["anchor"])
    para = tf.paragraphs[0]
    ooxml.set_paragraph(para, align=el["align"])
    r_pr = ooxml.add_slide_number(para, number)
    color, mods = ink("tx2", el["size"], False)
    ooxml.style_run_props(r_pr, size_pt=el["size"], color=color, mods=mods, font_ref="+mn-lt")


def draw_table(slide, el, ink):
    cols, rows = el["columns"], el["rows"]
    frame = slide.shapes.add_table(len(rows) + 1, len(cols), *el["box"])
    frame.name = el["name"]
    table = frame.table
    ooxml.set_table_style_none(table)
    for j, w in enumerate(el["widths"]):
        table.columns[j].width = w
    for i, h in enumerate(el["heights"]):
        table.rows[i].height = h
    margin_x, margin_y = int(layout.CELL_MARGIN_X_PT * 12700), int(layout.CELL_MARGIN_Y_PT * 12700)
    for i, values in enumerate([cols] + rows):
        header = i == 0
        highlighted = el["highlight_row"] is not None and i - 1 == el["highlight_row"]
        for j, value in enumerate(values):
            cell = table.cell(i, j)
            cell.margin_left = cell.margin_right = margin_x
            cell.margin_top = cell.margin_bottom = margin_y
            cell.vertical_anchor = MSO_ANCHOR.BOTTOM if header else MSO_ANCHOR.TOP
            tf = cell.text_frame
            tf.word_wrap = True
            para = tf.paragraphs[0]
            ooxml.set_paragraph(para, align=ALIGN[el["align"][j]])
            run = para.add_run()
            run.text = value
            bold = header or highlighted
            color, mods = ink("accent1" if highlighted else ("tx2" if header else "tx1"), el["size"], bold)
            ooxml.style_run_props(run._r.get_or_add_rPr(), size_pt=el["size"], bold=bold,
                                  color=color, mods=mods, font_ref="+mn-lt")
            last = i == len(rows)
            ooxml.set_cell_borders(cell, bottom=HEADER_RULE if header else (None if last else ROW_RULE))
    lines = [", ".join(cols)] + [", ".join(r) for r in rows[:TABLE_ALT_MAX_ROWS]]
    more = len(rows) - TABLE_ALT_MAX_ROWS
    alt = f"Table with {len(rows)} rows. Columns: " + "; ".join(lines) + (f"; and {more} more rows" if more > 0 else "")
    ooxml.set_alt_text(frame, alt + ".")
    return frame


def draw_image(slide, el):
    pic = slide.shapes.add_picture(el["path"], *el["box"])
    pic.name = el["name"]
    ooxml.set_alt_text(pic, el["alt"])
    return pic


def remove_unused_placeholders(slide):
    for ph in list(slide.placeholders):
        if ph._element.ph.get("type", "obj") not in ("title", "ctrTitle"):
            ph._element.getparent().remove(ph._element)


# -- build --------------------------------------------------------------------

def plan_deck(spec, spec_dir, prs, frames, title_bold):
    info = ooxml.theme_info(prs)
    prof = themes.profile(spec["profile"])
    fonts = {"major": info["fonts"]["major"], "minor": info["fonts"]["minor"]}
    meta = spec.get("meta") or {}
    planner = layout.Planner(prof, fonts, title_bold, prs.slide_width, prs.slide_height, spec_dir,
                             frames=frames, slide_numbers=meta.get("slide_numbers", True))
    plans = [planner.plan(n, s) for n, s in enumerate(spec["slides"], 1)]
    return plans, planner.problems, info


def render(prs, spec, plans, choose_layout, info):
    colors = ooxml.role_colors(info)
    for n, (slide_spec, elements) in enumerate(zip(spec["slides"], plans), 1):
        slide = prs.slides.add_slide(choose_layout(slide_spec["layout"]))
        remove_unused_placeholders(slide)
        ink = text_ink(info, ooxml.slide_background_hex(slide, info))
        for el in elements:
            if el["kind"] == "text":
                draw_text(slide, el, ink)
            elif el["kind"] == "chart":
                charts.add_chart(slide, el, colors)
            elif el["kind"] == "table":
                draw_table(slide, el, ink)
            elif el["kind"] == "image":
                draw_image(slide, el)
            elif el["kind"] == "slidenum":
                draw_slide_number(slide, el, n, ink)
        if slide_spec.get("notes"):
            slide.notes_slide.notes_text_frame.text = slide_spec["notes"]
    meta = spec.get("meta") or {}
    cp = prs.core_properties
    cp.title = meta.get("title") or spec["slides"][0]["title"]
    if meta.get("author"):
        cp.author = meta["author"]
        cp.last_modified_by = meta["author"]
    if meta.get("subject"):
        cp.subject = meta["subject"]


def summarize(plans, spec):
    out = []
    for n, (s, els) in enumerate(zip(spec["slides"], plans), 1):
        items = []
        for el in els:
            item = {"name": el["name"], "kind": el["kind"], "box_emu": list(el["box"])}
            if el["kind"] == "text":
                item["sizes"] = sorted({p["size"] for p in el["paras"]})
            elif "size" in el:
                item["sizes"] = [el["size"]]
            items.append(item)
        out.append({"slide": n, "layout": s["layout"], "title": s["title"], "elements": items})
    return out


def build(spec_path, out_path, check_only=False):
    spec_dir = os.path.dirname(os.path.abspath(spec_path))
    with open(spec_path, encoding="utf-8") as f:
        spec = json.load(f)
    report = {"tool": "deck-builder build_deck", "version": themes.VERSION,
              "generated_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
              "spec": os.path.abspath(spec_path), "spec_sha256": sha256_file(spec_path),
              "deck": None, "deck_sha256": None, "built": False, "problems": []}
    problems = spec_check.check_spec(spec, spec_dir)
    if not spec_check.errors(problems):
        prs, choose_layout, frames, title_bold, setup_problems = open_deck(spec, spec_dir)
        problems += setup_problems
        if not spec_check.errors(problems):
            plans, fit_problems, info = plan_deck(spec, spec_dir, prs, frames, title_bold)
            problems += fit_problems
            report["slides"] = summarize(plans, spec)
            report["profile"] = spec["profile"]
            report["template_fingerprint"] = ooxml.fingerprint(prs) if spec.get("template") else None
            if not spec_check.errors(problems) and not check_only:
                render(prs, spec, plans, choose_layout, info)
                os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
                prs.save(out_path)
                problems += ooxml.validate_package(out_path)
                report.update(deck=os.path.abspath(out_path), deck_sha256=sha256_file(out_path),
                              built=not spec_check.errors(problems))
    report["problems"] = problems
    return report


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("spec")
    ap.add_argument("-o", "--out", help="output .pptx (required unless --check-only)")
    ap.add_argument("--check-only", action="store_true", help="check and fit without writing a deck")
    ap.add_argument("--report", help="where to write build_report.json (default: beside the deck or spec)")
    args = ap.parse_args(argv)
    if not args.check_only and not args.out:
        ap.error("-o/--out is required unless --check-only")
    anchor = args.out if args.out and not args.check_only else args.spec
    report_path = args.report or os.path.join(os.path.dirname(os.path.abspath(anchor)), "build_report.json")
    try:
        report = build(args.spec, args.out, args.check_only)
    except Exception as e:  # a crash must still leave a report that says the build failed
        crash = spec_check.crash_message(e)
        report = {"tool": "deck-builder build_deck", "version": themes.VERSION, "built": False,
                  "crash": crash, "problems": []}
        print(crash, file=sys.stderr)
        try:
            with open(report_path, "w", encoding="utf-8") as f:
                json.dump(report, f, indent=2)
        except OSError as write_error:
            print(f"The build report could not be written either: {write_error}", file=sys.stderr)
        return 2
    with open(report_path, "w", encoding="utf-8") as f:
        json.dump(report, f, indent=2, ensure_ascii=False)
    for p in report["problems"]:
        print(spec_check.format_problem(p))
    errs = spec_check.errors(report["problems"])
    warns = len(report["problems"]) - len(errs)
    if errs:
        print(f"REFUSED: {len(errs)} error(s), {warns} warning(s). Nothing was written. Report: {report_path}")
        return 1
    if args.check_only:
        print(f"OK (check only): {len(report.get('slides', []))} slides fit, {warns} warning(s). Report: {report_path}")
    else:
        print(f"BUILT {report['deck']}: {len(report['slides'])} slides, {warns} warning(s). Report: {report_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
