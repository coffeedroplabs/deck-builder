# /// script
# requires-python = ">=3.11"
# dependencies = ["python-pptx==1.0.2"]
# ///
"""Read a company template and suggest how a deck spec should use it.

    uv run inspect_template.py brand.potx [--json template_report.json]

Prints the slide size, theme fonts and colours, every layout with its
placeholders and its fixed art (logos, bands), the problems a build would hit,
and the spec keys to paste: template, a suggested layout_map and the template
fingerprint. The suggestion is a starting point; check each mapping against
the layout list, because layout names and intent vary between templates.

Exit 0 usable, 1 the build cannot use this template as it is, 2 cannot read it.
"""

import argparse
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import build_deck  # noqa: E402
import layout  # noqa: E402
import ooxml  # noqa: E402
import rules  # noqa: E402
import themes  # noqa: E402

TITLE_KINDS = {"title", "ctrTitle"}
# Areas the builder fills itself; a subtitle counts for the frame but is not
# a content area, so it does not make a layout a "Title and Content" one.
BODY_KINDS = {"body", "obj"}
WIDE = 16 / 9


def _inches(emu):
    return round(emu / themes.EMU_PER_INCH, 2)


def _box_in(box):
    x, y, w, h = box
    return [_inches(x), _inches(y), _inches(w), _inches(h)]


def _overlaps(a, b):
    ax, ay, aw, ah = a
    bx, by, bw, bh = b
    return min(ax + aw, bx + bw) > max(ax, bx) and min(ay + ah, by + bh) > max(ay, by)


def build_regions(lay, slide_w, slide_h):
    """Where the builder will draw on this layout: the frames open_deck takes,
    including the footer placeholders, else the default grid's footer boxes."""
    regions = {}
    for ph in lay.placeholders:
        kind = ph._element.ph.get("type", "obj")
        if kind in TITLE_KINDS:
            regions.setdefault("title placeholder", (ph.left, ph.top, ph.width, ph.height))
        elif kind in build_deck.BODY_PLACEHOLDERS:
            regions.setdefault("body placeholder", (ph.left, ph.top, ph.width, ph.height))
    sx, sy = slide_w / themes.SLIDE_W, slide_h / themes.SLIDE_H
    grid = {"ftr": (layout.MARGIN_IN, layout.SOURCE_W_IN),
            "sldNum": (layout.MARGIN_IN + layout.CONTENT_W_IN - layout.NUMBER_W_IN, layout.NUMBER_W_IN)}
    found = ooxml.footer_placeholders(lay)
    for kind, label in (("ftr", "source line"), ("sldNum", "slide number")):
        ph = found.get(kind)
        if ph is not None:
            regions[f"footer placeholder ({label})"] = (ph.left, ph.top, ph.width, ph.height)
        else:
            x, w = grid[kind]
            regions[f"footer row ({label})"] = tuple(
                int(v * themes.EMU_PER_INCH * f) for v, f in
                ((x, sx), (layout.FOOTER_Y_IN, sy), (w, sx), (layout.FOOTER_H_IN, sy)))
    return regions


def describe_layout(index, lay, slide_w, slide_h):
    placeholders = []
    for ph in lay.placeholders:
        kind = ph._element.ph.get("type", "obj")
        placeholders.append({"kind": kind, "idx": ph.placeholder_format.idx, "name": ph.name,
                             "box_in": _box_in((ph.left, ph.top, ph.width, ph.height))})
    kinds = [p["kind"] for p in placeholders]
    regions = build_regions(lay, slide_w, slide_h)
    art = []
    for a in ooxml.template_art(lay, slide_w, slide_h):
        clashes = [] if a["background"] else [r for r, box in regions.items() if _overlaps(a["box"], box)]
        art.append({"name": a["name"], "from": a["from"], "background": a["background"],
                    "box_in": _box_in(a["box"]), "clashes_with": clashes})
    return {"index": index, "name": lay.name, "placeholders": placeholders, "art": art,
            "has_title": any(k in TITLE_KINDS for k in kinds), "centred_title": "ctrTitle" in kinds,
            "bodies": sum(k in BODY_KINDS for k in kinds), "has_subtitle": "subTitle" in kinds}


def _first(layouts, test):
    return next((lay["name"] for lay in layouts if test(lay)), None)


def suggest_layout_map(layouts):
    """Pick a template layout for each spec layout, preferring standard names, then shape."""
    def named(*words):
        return lambda lay: any(w in lay["name"].lower() for w in words)

    cover = (_first(layouts, lambda lay: lay["name"] == build_deck.DEFAULT_TEMPLATE_LAYOUTS["cover"])
             or _first(layouts, lambda lay: lay["centred_title"])
             or _first(layouts, named("title slide", "cover", "title page")))
    section = (_first(layouts, lambda lay: lay["name"] == build_deck.DEFAULT_TEMPLATE_LAYOUTS["section"])
               or _first(layouts, lambda lay: named("section", "divider", "chapter")(lay) and lay["has_title"]))
    # One body area: the builder lays out charts, columns and steps inside it
    # itself, so a two-content layout would only hand it the left half.
    content = (_first(layouts, lambda lay: lay["name"] == "Title and Content")
               or _first(layouts, lambda lay: lay["has_title"] and lay["bodies"] == 1
                         and not lay["centred_title"] and not named("section")(lay)))
    title_only = (_first(layouts, lambda lay: lay["name"] == "Title Only")
                  or _first(layouts, lambda lay: lay["has_title"] and lay["bodies"] == 0
                            and not lay["has_subtitle"] and not lay["centred_title"]))
    body_default = content or title_only
    mapping = {}
    for name in sorted(rules.LAYOUTS):
        if name == "cover":
            choice = cover or body_default
        elif name == "section":
            choice = section or title_only or body_default
        else:
            choice = body_default
        if choice:
            mapping[name] = choice
    return mapping


def inspect(path):
    prs = ooxml.open_presentation(path)
    w, h = prs.slide_width, prs.slide_height
    info = ooxml.theme_info(prs)
    layouts = [describe_layout(i, lay, w, h) for i, lay in enumerate(prs.slide_layouts)]
    mapping = suggest_layout_map(layouts)
    problems = []

    def add(severity, code, message):
        problems.append({"severity": severity, "code": code, "slide": None, "field": "template",
                         "message": message})

    if abs(w / h - WIDE) > 0.01:
        add("warning", "not-widescreen",
            f"slides are {_inches(w)} x {_inches(h)} in, not 16:9; the builder scales its grid to fit, "
            "so check the render closely")
    for key in ("major", "minor"):
        font = info["fonts"][key]
        if font not in themes.METRIC_SAFE_FONTS:
            add("warning", "font-not-metric-safe",
                f"{key} font {font!r} has no metric-compatible fallback; the build's fit is a "
                "prediction, so only the PowerPoint render can sign off overflow, and viewers "
                "without the font will see different line breaks")
    roles = ooxml.role_colors(info)
    if all(roles.get(k) for k in ("text", "bg", "muted", "accent", "series", "third")):
        for msg in themes.theme_contrast_problems({"colors": roles}):
            if msg.startswith("text "):
                add("warning", "theme-contrast", f"{msg}; body text will fail the lint")
            else:
                add("warning", "theme-contrast",
                    f"{msg}; words in that colour are written in its nearest darker or lighter palette "
                    "shade that passes, so they look darker than the brand colour; chart fills keep it as is")
    if not any(lay["has_title"] for lay in layouts):
        add("error", "no-title-layout", "no layout has a title placeholder; the build needs one for every slide")
    if not (mapping.get("bullets")):
        add("error", "no-content-layout", "no layout with a title and one body area, or a title alone, was found; "
                                          "write layout_map by hand from the list above")
    for name in sorted(set(mapping.values())):
        lay = next(lay for lay in layouts if lay["name"] == name)
        for a in lay["art"]:
            if a["clashes_with"]:
                add("warning", "template-art",
                    f"layout {name!r}: {a['name']!r} at {a['box_in']} in sits in the "
                    f"{' and the '.join(a['clashes_with'])}; the lint fails any slide content that covers it, "
                    "so map these slides to another layout or expect to move content")
    mx, my = themes.SAFE_MARGIN * w, themes.SAFE_MARGIN * h
    for name in sorted(set(mapping.values())):
        lay = next(lay for lay in prs.slide_layouts if lay.name == name)
        for region, (x, y, rw, rh) in build_regions(lay, w, h).items():
            # The default grid's footer row is the builder's, not the designer's.
            if region.startswith("footer row"):
                continue
            if x < mx or y < my or x + rw > w - mx or y + rh > h - my:
                add("warning", "template-safe-area",
                    f"layout {name!r}: the {region} at {_box_in((x, y, rw, rh))} in comes within 5% of the "
                    "slide edge, where projectors and video calls may crop; the lint leaves the template's "
                    "own placement alone, so check the render")
    return {"template": os.path.abspath(path), "slide_size_in": [_inches(w), _inches(h)],
            "theme": {"name": info["name"], "fonts": info["fonts"], "colors": info["colors"], "roles": roles},
            "layouts": layouts, "layout_map": mapping, "fingerprint": ooxml.fingerprint(prs),
            "problems": problems}


def _print(report, template_arg):
    w, h = report["slide_size_in"]
    theme = report["theme"]
    print(f"Template: {report['template']}")
    print(f"Slide size: {w} x {h} in")
    print(f"Theme {theme['name']!r}: headings {theme['fonts']['major']}, body {theme['fonts']['minor']}")
    print("Roles the build colours by: " + ", ".join(f"{k} {v}" for k, v in theme["roles"].items() if v))
    print("\nLayouts (boxes are x, y, width, height in inches):")
    for lay in report["layouts"]:
        print(f"  {lay['index']:>2}  {lay['name']}")
        for p in lay["placeholders"]:
            print(f"        {p['kind']:<9} idx {p['idx']:<3} {p['box_in']}")
        for a in lay["art"]:
            tag = "background" if a["background"] else "art"
            clash = f"  <- in the {', '.join(a['clashes_with'])}" if a["clashes_with"] else ""
            print(f"        {tag:<9} {a['from']:<6} {a['name']} {a['box_in']}{clash}")
    for p in report["problems"]:
        print(f"{p['severity'].upper():7} {p['code']:22} {p['message']}")
    keys = {"template": template_arg, "layout_map": report["layout_map"],
            "template_fingerprint": report["fingerprint"]}
    print("\nSpec keys to paste (template is relative to the spec file; check each mapping):")
    print(json.dumps(keys, indent=2, ensure_ascii=False))


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("template", help=".pptx or .potx")
    ap.add_argument("--json", help="also write the full report to this path")
    args = ap.parse_args(argv)
    try:
        report = inspect(args.template)
    except Exception as e:  # any unreadable package, whatever the library raises
        print(f"Cannot read {args.template} as a PowerPoint template: {e}", file=sys.stderr)
        return 2
    _print(report, args.template)
    if args.json:
        with open(args.json, "w", encoding="utf-8") as f:
            json.dump(report, f, indent=2, ensure_ascii=False)
    return 1 if any(p["severity"] == "error" for p in report["problems"]) else 0


if __name__ == "__main__":
    sys.exit(main())
