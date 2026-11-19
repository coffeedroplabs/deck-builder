# /// script
# requires-python = ">=3.10"
# dependencies = ["python-pptx==1.0.2"]
# ///
"""Check a built deck and write gate_result.json, the only pass/fail verdict.

    uv run lint_deck.py deck.pptx --spec spec.json [--measure render/measure.json]

Reads the deck as it is on disk, so a hand edit after the build is checked
too. With a PowerPoint measure.json from render.py it also checks what
PowerPoint actually laid out: overflowing text, titles that wrapped too far,
tables that grew. Without one the gate can still pass, but it records that
nothing was measured, and that has to be said when the deck is delivered.

Exit 0 passed, 1 failed, 2 crashed (gate_result.json then says passed: false).
"""

import argparse
import hashlib
import json
import os
import sys
from datetime import datetime, timezone

from pptx import Presentation
from pptx.oxml.ns import qn

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import check_story  # noqa: E402
import ooxml  # noqa: E402
import rules  # noqa: E402
import spec_check  # noqa: E402
import themes  # noqa: E402

TOOL = "deck-builder lint_deck"
SMALL_TEXT_SHAPES = {"Source", "Slide number"}  # held to the source floor, not the body floor
NOTES_EXEMPT = {"cover", "section"}
TOL_PT = 1.0
TOL_EMU = int(TOL_PT * themes.EMU_PER_PT)
APPROVAL_FILE = check_story.APPROVAL_FILE


def sha256_file(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()


class Lint:
    def __init__(self, deck_path, spec, spec_dir, prof):
        self.deck_path = deck_path
        self.spec = spec
        self.spec_dir = spec_dir
        self.prof = prof
        self.problems = []
        self.prs = Presentation(deck_path)
        self.info = ooxml.theme_info(self.prs)
        self.W, self.H = self.prs.slide_width, self.prs.slide_height

    def add(self, severity, code, slide, field, message):
        self.problems.append({"severity": severity, "code": code, "slide": slide, "field": field,
                              "message": message})

    # -- the deck file --------------------------------------------------------

    def check_package(self):
        for p in ooxml.validate_package(self.deck_path):
            self.add(p["severity"], p["code"], None, p["part"], p["message"])

    def check_matches_spec(self):
        slides = self.spec["slides"]
        if len(self.prs.slides) != len(slides):
            self.add("error", "deck-spec-mismatch", None, "",
                     f"the deck has {len(self.prs.slides)} slides but the spec has {len(slides)}; "
                     "rebuild from the spec, or put the hand edit into the spec")
            return
        for n, (slide, s) in enumerate(zip(self.prs.slides, slides), 1):
            title = _title_shape(slide)
            if title is None:
                self.add("error", "title-missing", n, "Title", "the slide has no title shape")
            elif rules.normalize(title.text_frame.text) != rules.normalize(s.get("title")):
                self.add("error", "deck-spec-mismatch", n, "Title",
                         "the title in the deck differs from the spec; rebuild, or put the edit into the spec")

    def check_slides(self):
        families = set()
        for n, slide in enumerate(self.prs.slides, 1):
            bg = ooxml.slide_background_hex(slide, self.info)
            boxes = []
            footer_boxes = [(ph.left, ph.top, ph.width, ph.height)
                            for ph in ooxml.footer_placeholders(slide.slide_layout).values()]
            for shape in slide.shapes:
                name = shape.name
                box = (shape.left, shape.top, shape.width, shape.height)
                if None not in box:
                    boxes.append((name, box))
                    # The source line and slide number fill the template's
                    # footer placeholders' boxes, which its designer placed.
                    owned = _at_template_position(shape, box) or (
                        name in SMALL_TEXT_SHAPES and any(_same_box(box, f) for f in footer_boxes))
                    self._check_geometry(n, name, box, owned)
                el = shape._element
                if el.tag in (qn("p:pic"), qn("p:graphicFrame")) and not ooxml.alt_text(el):
                    self.add("error", "alt-missing", n, name,
                             "picture, chart or table has no alt text; set alt in the spec")
                for para, cell in _paragraphs(shape):
                    families |= self._check_runs(n, name, para, bg)
                    text = "".join(r.text for r in para.runs)
                    for sev, code, msg in rules.text_problems(text):
                        self.add(sev, code, n, name, msg)
                if getattr(shape, "has_chart", False) and shape.has_chart:
                    families |= self._check_chart_text(n, name, shape.chart)
            self._check_overlap(n, boxes)
            self._check_template_art(n, slide, boxes)
            self._check_notes(n, slide)
        if len(families) > themes.MAX_FONT_FAMILIES:
            self.add("error", "too-many-fonts", None, "",
                     f"{len(families)} font families ({', '.join(sorted(families))}); "
                     f"use at most {themes.MAX_FONT_FAMILIES}")

    def _floor(self, name):
        return self.prof["source_floor"] if name in SMALL_TEXT_SHAPES else self.prof["floor"]

    def _family(self, ref):
        if ref in (None, "", "+mn-lt"):
            return self.info["fonts"]["minor"]
        if ref == "+mj-lt":
            return self.info["fonts"]["major"]
        return ref

    def _check_runs(self, n, name, para, bg):
        families = set()
        floor = self._floor(name)
        for run in para.runs:
            if not run.text.strip():
                continue
            r_pr = run._r.find(qn("a:rPr"))
            sz = r_pr.get("sz") if r_pr is not None else None
            size = int(sz) / 100 if sz else None
            if size is not None and size < floor:
                self.add("error", "font-below-floor", n, name,
                         f"{size:g} pt text; the {self.spec['profile']} floor is {floor} pt")
            latin = r_pr.find(qn("a:latin")) if r_pr is not None else None
            families.add(self._family(latin.get("typeface") if latin is not None else None))
            fill = r_pr.find(qn("a:solidFill")) if r_pr is not None else None
            fg = ooxml.color_hex(fill, self.info) if fill is not None else ooxml.scheme_hex(self.info, "tx1")
            if bg and fg:
                bold = r_pr is not None and r_pr.get("b") in ("1", "true")
                large = size is not None and (size >= 18 or (bold and size >= 14))
                need = themes.CONTRAST["large" if large else "text"]
                ratio = themes.contrast_ratio(fg, bg)
                if ratio < need:
                    self.add("error", "low-contrast", n, name,
                             f"text #{fg} on #{bg} is {ratio:.1f}:1; needs {need}:1")
        return families

    def _check_chart_text(self, n, name, chart):
        families = set()
        root = chart._chartSpace
        for tag in ("a:defRPr", "a:rPr"):
            for r_pr in root.iter(qn(tag)):
                sz = r_pr.get("sz")
                if sz and int(sz) / 100 < self.prof["floor"]:
                    self.add("error", "font-below-floor", n, name,
                             f"{int(sz) / 100:g} pt chart text; the floor is {self.prof['floor']} pt")
                latin = r_pr.find(qn("a:latin"))
                if latin is not None:
                    families.add(self._family(latin.get("typeface")))
        return families

    def _check_geometry(self, n, name, box, template_owned=False):
        x, y, w, h = box
        if x < -TOL_EMU or y < -TOL_EMU or x + w > self.W + TOL_EMU or y + h > self.H + TOL_EMU:
            self.add("error", "off-slide", n, name, "shape extends past the slide edge")
            return
        # The template's designer chose where its placeholders sit; the deck
        # author cannot act on that, and inspect_template.py reports it once.
        if template_owned:
            return
        mx, my = themes.SAFE_MARGIN * self.W, themes.SAFE_MARGIN * self.H
        if x < mx - TOL_EMU or y < my - TOL_EMU or x + w > self.W - mx + TOL_EMU or y + h > self.H - my + TOL_EMU:
            self.add("warning", "outside-safe-area", n, name,
                     "shape is within 5% of the slide edge; projectors and video calls may crop it")

    def _check_overlap(self, n, boxes):
        for i in range(len(boxes)):
            for j in range(i + 1, len(boxes)):
                (a, (ax, ay, aw, ah)), (b, (bx, by, bw, bh)) = boxes[i], boxes[j]
                dx = min(ax + aw, bx + bw) - max(ax, bx)
                dy = min(ay + ah, by + bh) - max(ay, by)
                if dx > TOL_EMU and dy > TOL_EMU:
                    self.add("error", "shapes-overlap", n, f"{a} / {b}", f"{a!r} and {b!r} overlap")

    def _check_template_art(self, n, slide, boxes):
        # A template's logo or badge sits on the layout, not the slide, so the
        # slide-level overlap check never sees it.
        for art in ooxml.template_art(slide.slide_layout, self.W, self.H):
            if art["background"]:
                continue
            ax, ay, aw, ah = art["box"]
            for name, (bx, by, bw, bh) in boxes:
                dx = min(ax + aw, bx + bw) - max(ax, bx)
                dy = min(ay + ah, by + bh) - max(ay, by)
                if dx > TOL_EMU and dy > TOL_EMU:
                    self.add("error", "covers-template-art", n, name,
                             f"{name!r} covers the template's {art['name']!r}; move it, or map this "
                             "layout to one without that art")

    def _check_notes(self, n, slide):
        layout = self.spec["slides"][n - 1].get("layout") if n <= len(self.spec["slides"]) else None
        notes = slide.notes_slide.notes_text_frame.text if slide.has_notes_slide else ""
        if notes.strip():
            for sev, code, msg in rules.text_problems(notes):
                if code != "manual-bullet":
                    self.add(sev, code, n, "notes", msg)
        elif self.prof["notes_required"] and layout not in NOTES_EXEMPT:
            self.add("warning", "notes-missing", n, "notes", "add speaker notes; the talk track is where detail goes")

    # -- what PowerPoint laid out ------------------------------------------------

    def check_measure(self, measure, deck_sha):
        if measure is None:
            self.add("warning", "not-measured", None, "",
                     "no measure.json; run render.py so PowerPoint measures the text")
            return False
        if measure.get("deck_sha256") != deck_sha:
            self.add("error", "measure-stale", None, "",
                     "measure.json is from a different build of this deck; run render.py again")
            return False
        if not measure.get("measured"):
            self.add("warning", "not-measured", None, "",
                     f"rendered with {measure.get('backend')}, which cannot measure text; "
                     "overflow was predicted from font metrics, not checked in PowerPoint")
            return False
        m_slides = measure.get("slides") or []
        if len(m_slides) != len(self.prs.slides):
            self.add("error", "measure-stale", None, "", "measure.json has a different slide count; re-render")
            return False
        for n, (slide, ms) in enumerate(zip(self.prs.slides, m_slides), 1):
            by_id = {s["id"]: s for s in _as_list(ms.get("shapes"))}
            for shape in slide.shapes:
                m = by_id.get(shape.shape_id)
                if m is None:
                    self.add("error", "measure-stale", n, shape.name, "shape missing from measure.json; re-render")
                    continue
                if m.get("has_text"):
                    self._check_measured_text(n, shape.name, m)
                if m.get("has_table") and shape.height is not None:
                    planned = shape.height / themes.EMU_PER_PT
                    if m["height"] > planned + TOL_PT:
                        self.add("error", "table-overflow", n, shape.name,
                                 f"rows wrapped: the table grew to {m['height']:.0f} pt from {planned:.0f} pt; "
                                 "shorten cell text or cut rows")
        return True

    def _check_measured_text(self, n, name, m):
        # Width is not checked: BoundWidth overshoots the box for centred and
        # right-aligned text, and with word wrap on a too-long word breaks onto
        # another line, which the height check below already catches.
        bottom = m["bound_top"] + m["bound_height"]
        if bottom > m["top"] + m["height"] + TOL_PT or m["bound_top"] < m["top"] - TOL_PT:
            self.add("error", "text-overflow", n, name,
                     f"text is {m['bound_height']:.0f} pt tall in a {m['height']:.0f} pt box; "
                     "cut words, split the slide, or move detail to the notes")
        floor = self._floor(name)
        for s in _as_list(m.get("font_sizes")):
            if s not in ("", None) and float(s) < floor:
                self.add("error", "font-below-floor", n, name,
                         f"PowerPoint shows {float(s):g} pt text; the floor is {floor} pt")
        if name == "Title":
            lines = m.get("lines") or 0
            if lines > self.prof["max_title_lines"]:
                self.add("error", "title-lines", n, name,
                         f"title runs to {lines} lines; keep it to {self.prof['max_title_lines']}")
            texts = [t for t in _as_list(m.get("line_texts")) if t and t.strip()]
            if len(texts) >= 2 and rules.word_count(texts[-1]) == 1:
                self.add("warning", "title-widow", n, name,
                         f"the title's last line is one word ({texts[-1].strip()!r}); cut words until it fits "
                         "one line or the lines balance (cutting never trips title-long, padding can)")

    # -- the outline the user approved -------------------------------------------

    def check_approval(self, approval_path):
        state = check_story.approval_state(self.spec["slides"], approval_path)
        if state["status"] == "missing":
            self.add("warning", "story-unapproved", None, "",
                     "no approved outline; run check_story.py and get the user's approval of the titles")
        elif state["status"] == "stale":
            self.add("warning", "story-changed", None, "",
                     "the titles changed since the user approved the outline; show them the new titles")
        return state

    def check_structure(self):
        self.problems += check_story.story_problems(self.spec["slides"])


def _as_list(v):
    # ConvertTo-Json writes a one-element array as a bare value.
    if v is None:
        return []
    return v if isinstance(v, list) else [v]


def _at_template_position(shape, box):
    """Whether a placeholder sits exactly where its layout puts it."""
    base = getattr(shape, "_base_placeholder", None) if shape.is_placeholder else None
    if base is None:
        return False
    base_box = (base.left, base.top, base.width, base.height)
    return None not in base_box and _same_box(box, base_box)


def _same_box(a, b):
    return all(abs(p - q) <= TOL_EMU for p, q in zip(a, b))


def _title_shape(slide):
    for shape in slide.shapes:
        if shape.name == "Title" and shape.has_text_frame:
            return shape
    try:
        return slide.shapes.title
    except (KeyError, AttributeError):
        return None


def _paragraphs(shape):
    if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
        for p in shape.text_frame.paragraphs:
            yield p, None
    if getattr(shape, "has_table", False) and shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                for p in cell.text_frame.paragraphs:
                    yield p, cell


def _dedupe(spec_problems, deck_problems):
    # The spec check and the deck check share text rules; keep one copy.
    seen = {(p["code"], p["slide"]) for p in spec_problems}
    out = list(spec_problems)
    for p in deck_problems:
        key = (p["code"], p["slide"])
        if key in seen and p["code"] not in ("font-below-floor", "text-overflow"):
            continue
        seen.add(key)
        out.append(p)
    return out


def lint(deck_path, spec_path, measure_path=None, approval_path=None):
    spec_dir = os.path.dirname(os.path.abspath(spec_path))
    with open(spec_path, encoding="utf-8") as f:
        spec = json.load(f)
    prof = themes.profile(spec["profile"])
    deck_sha = sha256_file(deck_path)
    result = {"tool": TOOL, "version": themes.VERSION, "deck": os.path.abspath(deck_path),
              "deck_sha256": deck_sha, "spec": os.path.abspath(spec_path), "spec_sha256": sha256_file(spec_path),
              "generated_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
              "passed": False, "measured": False, "measure_backend": None, "approval": None,
              "errors": [], "warnings": [], "counts": {}, "crash": None}
    spec_problems = spec_check.check_spec(spec, spec_dir)
    linter = Lint(deck_path, spec, spec_dir, prof)
    linter.check_package()
    linter.check_matches_spec()
    linter.check_slides()
    linter.check_structure()
    if measure_path is None:
        default = os.path.join(os.path.dirname(os.path.abspath(deck_path)), "render", "measure.json")
        measure_path = default if os.path.exists(default) else None
    measure = None
    if measure_path:
        with open(measure_path, encoding="utf-8") as f:
            measure = json.load(f)
        result["measure_backend"] = measure.get("backend")
    result["measured"] = linter.check_measure(measure, deck_sha)
    result["approval"] = linter.check_approval(approval_path or os.path.join(spec_dir, APPROVAL_FILE))
    problems = _dedupe(spec_problems, linter.problems)
    result["errors"] = [p for p in problems if p["severity"] == "error"]
    result["warnings"] = [p for p in problems if p["severity"] != "error"]
    result["counts"] = {"slides": len(linter.prs.slides), "errors": len(result["errors"]),
                        "warnings": len(result["warnings"])}
    result["passed"] = not result["errors"]
    return result


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("deck")
    ap.add_argument("--spec", required=True)
    ap.add_argument("--measure", help="render.py's measure.json (default: render/measure.json beside the deck)")
    ap.add_argument("--approval", help=f"approved outline (default: {APPROVAL_FILE} beside the spec)")
    ap.add_argument("-o", "--out", help="where to write gate_result.json (default: beside the deck)")
    args = ap.parse_args(argv)
    out = args.out or os.path.join(os.path.dirname(os.path.abspath(args.deck)), "gate_result.json")
    try:
        result = lint(args.deck, args.spec, args.measure, args.approval)
    except Exception as e:  # a crash must leave a failed verdict, never a stale passing one
        crash = spec_check.crash_message(e)
        result = {"tool": TOOL, "version": themes.VERSION, "deck": os.path.abspath(args.deck),
                  "generated_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
                  "passed": False, "errors": [], "warnings": [], "crash": crash}
        print(crash, file=sys.stderr)
        try:
            with open(out, "w", encoding="utf-8") as f:
                json.dump(result, f, indent=2)
            print(f"NOT CHECKED. {out} records passed: false.")
        except OSError as write_error:
            print(f"NOT CHECKED, and the verdict could not be written: {write_error}")
        return 2
    with open(out, "w", encoding="utf-8") as f:
        json.dump(result, f, indent=2, ensure_ascii=False)
    for p in result["errors"] + result["warnings"]:
        print(spec_check.format_problem(p))
    measured = "measured in PowerPoint" if result["measured"] else "NOT measured"
    verdict = "PASSED" if result["passed"] else "FAILED"
    print(f"{verdict}: {result['counts']['errors']} error(s), {result['counts']['warnings']} warning(s), "
          f"text {measured}, outline {result['approval']['status']}. Verdict: {out}")
    return 0 if result["passed"] else 1


if __name__ == "__main__":
    sys.exit(main())
