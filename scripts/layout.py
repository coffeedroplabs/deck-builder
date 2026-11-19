"""Slide geometry and text fitting: turns one spec slide into placed elements.

Pure computation, no python-pptx, so fitting can be tested without writing a
file. Every text element is sized from the profile: start at the role's base
size, step down 2 pt to its smallest, and if it still does not fit, record a
problem naming the field and the room it needed. The builder refuses to write
a deck while any such problem stands.
"""

import os
import re

from PIL import Image

import textfit
import themes

EMU_IN = themes.EMU_PER_INCH
EMU_PT = themes.EMU_PER_PT
REF_W_IN = themes.SLIDE_W / EMU_IN
REF_H_IN = themes.SLIDE_H / EMU_IN
MARGIN_IN = themes.GRID["margin_x"] * REF_W_IN
CONTENT_W_IN = REF_W_IN - 2 * MARGIN_IN
GUTTER_IN = themes.GRID["gutter"] * REF_W_IN
BODY_Y_IN = themes.GRID["body_y"] * REF_H_IN
BODY_H_IN = themes.GRID["body_bottom"] * REF_H_IN - BODY_Y_IN
FOOTER_Y_IN = themes.GRID["footer_y"] * REF_H_IN
FOOTER_H_IN = themes.GRID["footer_h"] * REF_H_IN
SOURCE_W_IN = 10.0
NUMBER_W_IN = 1.0
PROCESS_GAP_IN = 0.3
CHART_SHARE = 0.62  # chart width when bullets sit beside it
IMAGE_SHARE = 0.55
CELL_MARGIN_X_PT = 0.08 * 72
CELL_MARGIN_Y_PT = 0.04 * 72
TITLE_LINE_SPACING = 0.9
NUMERIC_CELL_RE = re.compile(r"^[\s$€£¥+\-−(]*[\d.,]+\s*(%|x|pts?|[kmb]n?)?[)\s]*$", re.IGNORECASE)

# Where each special layout puts its title and body in the default grid, in
# inches on the 13.333 x 7.5 in reference slide: (title_y, title_h, anchor,
# body_y, body_h). Layouts not listed use the standard title band and body.
SPECIAL_FRAMES = {
    "cover": (2.2, 2.0, "b", 4.3, 2.35),
    "section": (2.6, 1.6, "b", 4.3, 1.2),
    "statement": (1.6, 3.0, "m", 4.8, 1.85),
}
TITLE_ROLE = {"cover": "cover_title", "section": "statement", "statement": "statement"}


def emu(inches):
    return int(round(inches * EMU_IN))


def pt(emu_value):
    return emu_value / EMU_PT


def box_w_pt(box):
    return pt(box[2])


class Planner:
    """Plans slides for one deck. `frames` maps a layout to template geometry
    ({"title": box or None, "body": box or None}) in template mode."""

    def __init__(self, prof, fonts, title_bold, slide_w, slide_h, spec_dir=".",
                 frames=None, slide_numbers=True):
        self.prof = prof
        self.fonts = fonts
        self.title_bold = title_bold
        self.sx = slide_w / themes.SLIDE_W
        self.sy = slide_h / themes.SLIDE_H
        self.spec_dir = spec_dir
        self.frames = frames or {}
        self.slide_numbers = slide_numbers
        self.problems = []
        self.n = 0

    # -- geometry -----------------------------------------------------------

    def box(self, x_in, y_in, w_in, h_in):
        return (emu(x_in * self.sx), emu(y_in * self.sy), emu(w_in * self.sx), emu(h_in * self.sy))

    def frame(self, layout):
        """(title box, title anchor, body box) for a layout."""
        default_title = self.box(MARGIN_IN, themes.GRID["title_y"] * REF_H_IN, CONTENT_W_IN,
                                 themes.GRID["title_h"] * REF_H_IN)
        default_body = self.box(MARGIN_IN, BODY_Y_IN, CONTENT_W_IN, BODY_H_IN)
        anchor = "t"
        if layout in SPECIAL_FRAMES:
            ty, th, anchor, by, bh = SPECIAL_FRAMES[layout]
            default_title = self.box(MARGIN_IN, ty, CONTENT_W_IN, th)
            default_body = self.box(MARGIN_IN, by, CONTENT_W_IN, bh)
        tmpl = self.frames.get(layout) or {}
        if not tmpl:
            return default_title, anchor, default_body
        if layout == "statement":
            # Templates have no statement layout: keep the template's title
            # column so the left edge matches the other slides, and centre the
            # message as the default grid does.
            x, _, w, _ = tmpl.get("title") or default_title
            return ((x, default_title[1], w, default_title[3]), anchor,
                    (x, default_body[1], w, default_body[3]))
        # A template's own placeholder decides its anchor; ours would fight it.
        return tmpl.get("title") or default_title, None, tmpl.get("body") or default_body

    def frame_style(self, layout, part):
        """(align, anchor) the template gives a frame's text; None anchor means the default."""
        tmpl = self.frames.get(layout) or {}
        return tmpl.get(f"{part}_align") or "l", tmpl.get(f"{part}_anchor")

    @staticmethod
    def sub_box(body, y_frac, h_frac, x_frac=0.0, w_frac=1.0):
        x, y, w, h = body
        return (x + int(w * x_frac), y + int(h * y_frac), int(w * w_frac), int(h * h_frac))

    def columns(self, body, count, gap_in):
        x, y, w, h = body
        gap = emu(gap_in * self.sx)
        col_w = (w - gap * (count - 1)) // count
        return [(x + i * (col_w + gap), y, col_w, h) for i in range(count)]

    def split(self, body, share):
        """Left area taking `share` of the width, right area after a gutter."""
        x, y, w, h = body
        gutter = emu(GUTTER_IN * self.sx)
        left_w = int(w * share)
        return (x, y, left_w, h), (x + left_w + gutter, y, w - left_w - gutter, h)

    # -- paragraphs and fitting ---------------------------------------------

    def para(self, text, size, *, font="minor", bold=False, italic=False, color="tx1", align="l",
             indent=0.0, hanging=0.0, bullet=None, space_before=0.0, line_spacing=1.0, max_lines=None):
        return {"text": text, "font": font, "family": self.fonts[font], "size": size, "bold": bold,
                "italic": italic, "color": color, "align": align, "indent": indent, "hanging": hanging,
                "bullet": bullet, "space_before": space_before, "line_spacing": line_spacing,
                "max_lines": max_lines}

    def bullet_paras(self, bullets, size, sub_size, color="tx1"):
        out = []
        cap = self.prof["max_bullet_lines"]
        for b in bullets:
            text, subs = (b, []) if isinstance(b, str) else (b["text"], b.get("sub") or [])
            ind = round(0.9 * size, 1)
            out.append(self.para(text, size, color=color, indent=ind, hanging=ind, bullet="•",
                                 space_before=0.5 * size if out else 0.0, max_lines=cap))
            for s in subs:
                hang = round(0.9 * sub_size, 1)
                out.append(self.para(s, sub_size, color=color, indent=ind + hang, hanging=hang, bullet="–",
                                     space_before=0.25 * sub_size, max_lines=cap))
        return out

    def sizes(self, role, sub_role=None):
        base, smallest = self.prof["sizes"][role]
        steps = textfit.size_steps(base, smallest)
        if not sub_role:
            return [(s, None) for s in steps]
        sub_base, sub_smallest = self.prof["sizes"][sub_role]
        return [(s, max(s - (base - sub_base), sub_smallest)) for s in steps]

    @staticmethod
    def _check(paras, box):
        res = textfit.fits(paras, pt(box[2]), pt(box[3]))
        over = [i for i, (p, lines) in enumerate(zip(paras, res["per_paragraph"]))
                if p.get("max_lines") is not None and lines > p["max_lines"]]
        return res, over

    def fit_group(self, field_items, role, sub_role=None, code="text-overflow"):
        """Fit several boxes at one shared size (columns read as a set).

        field_items: list of (field, box, make) where make(size, sub_size)
        returns the paragraph list. Returns a list of paragraph lists, at the
        largest size where every box fits, or at the smallest size with a
        problem recorded.
        """
        last = None
        for size, sub in self.sizes(role, sub_role):
            results = []
            for field, box, make in field_items:
                paras = make(size, sub)
                res, over = self._check(paras, box)
                results.append((field, box, paras, res, over))
            if all(r[3]["fits"] and not r[4] for r in results):
                return [r[2] for r in results]
            last = (size, results)
        size, results = last
        for field, box, paras, res, over in results:
            if over:
                p = paras[over[0]]
                lines = res["per_paragraph"][over[0]]
                kind = "title-lines" if code == "title-lines" else ("source-long" if code == "source-long" else "bullet-lines")
                self.problem(kind, field, f"{p['text'][:60]!r} takes {lines} lines at {size} pt; "
                                          f"at most {p['max_lines']}. Shorten it")
            elif not res["fits"]:
                self.problem(code, field, f"needs {res['needed']} pt of height at {size} pt but the box has "
                                          f"{res['available']} pt ({res['lines']} lines). Cut words, move "
                                          "detail to the notes, or split the slide")
        return [r[2] for r in results]

    def fit_one(self, field, box, make, role, sub_role=None, code="text-overflow"):
        return self.fit_group([(field, box, make)], role, sub_role, code)[0]

    def problem(self, code, field, message, severity="error"):
        self.problems.append({"severity": severity, "code": code, "slide": self.n, "field": field,
                              "message": message})

    @staticmethod
    def text_el(name, box, paras, anchor="t", placeholder=False):
        return {"kind": "text", "name": name, "box": box, "paras": paras, "anchor": anchor,
                "placeholder": placeholder}

    # -- slides -------------------------------------------------------------

    def plan(self, n, slide):
        """Elements for one slide, in reading order."""
        self.n = n
        layout = slide["layout"]
        title_box, anchor, body = self.frame(layout)
        els = [self.title(slide, layout, title_box, anchor)]
        els += getattr(self, "plan_" + layout.replace("-", "_"))(slide, body)
        els += self.footer(slide, layout)
        return els

    def title(self, slide, layout, box, anchor):
        role = TITLE_ROLE.get(layout, "title")
        max_lines = self.prof["max_title_lines"] if role == "title" else None
        code = "title-lines" if max_lines else "text-overflow"

        align, _ = self.frame_style(layout, "title")

        def make(size, _sub):
            return [self.para(slide["title"], size, font="major", bold=self.title_bold, align=align,
                              line_spacing=TITLE_LINE_SPACING, max_lines=max_lines)]
        paras = self.fit_one("title", box, make, role, code=code)
        return self.text_el("Title", box, paras, anchor, placeholder=True)

    def simple(self, name, field, box, text, role, *, color="tx1", font="minor", bold=False,
               italic=False, align="l", anchor="t", max_lines=None):
        def make(size, _sub):
            return [self.para(text, size, font=font, bold=bold, italic=italic, color=color,
                              align=align, max_lines=max_lines)]
        code = "source-long" if field == "source" else "text-overflow"
        return self.text_el(name, box, self.fit_one(field, box, make, role, code=code), anchor)

    def plan_cover(self, slide, body):
        els = []
        align, anchor = self.frame_style("cover", "body")
        if slide.get("subtitle"):
            els.append(self.simple("Subtitle", "subtitle", self.sub_box(body, 0, 0.55), slide["subtitle"],
                                   "body", color="tx2", align=align, anchor=anchor or "t"))
        if slide.get("byline"):
            x, y, w, h = body
            bh = min(emu(0.65 * self.sy), int(h * 0.4))
            els.append(self.simple("Byline", "byline", (x, y + h - bh, w, bh), slide["byline"], "small",
                                   color="tx2", align=align, anchor="b"))
        return els

    def plan_section(self, slide, body):
        if not slide.get("subtitle"):
            return []
        align, anchor = self.frame_style("section", "body")
        return [self.simple("Subtitle", "subtitle", body, slide["subtitle"], "body", color="tx2",
                            align=align, anchor=anchor or "t")]

    def plan_statement(self, slide, body):
        if not slide.get("body"):
            return []
        # The supporting line sits under the message in the title's column,
        # so it follows the title's alignment, not the body placeholder's.
        align, _ = self.frame_style("statement", "title")
        return [self.simple("Body", "body", body, slide["body"], "body2", color="tx2", align=align)]

    def plan_bullets(self, slide, body):
        paras = self.fit_one("bullets", body, lambda s, sub: self.bullet_paras(slide["bullets"], s, sub),
                             "body", "body2")
        return [self.text_el("Body", body, paras)]

    def plan_conclusion(self, slide, body):
        els = []
        bullets_box = body
        if slide.get("ask"):
            bullets_box = self.sub_box(body, 0, 0.66)
            ask_box = self.sub_box(body, 0.71, 0.29)
        paras = self.fit_one("bullets", bullets_box,
                             lambda s, sub: self.bullet_paras(slide["bullets"], s, sub), "body", "body2")
        els.append(self.text_el("Body", bullets_box, paras))
        if slide.get("ask"):
            els.append(self.simple("Ask", "ask", ask_box, slide["ask"], "body", color="accent1", bold=True))
        return els

    def plan_two_column(self, slide, body):
        boxes = self.columns(body, 2, GUTTER_IN)
        items = []
        for side, box in zip(("left", "right"), boxes):
            col = slide[side]

            def make(size, sub, col=col):
                paras = []
                if col.get("heading"):
                    paras.append(self.para(col["heading"], size, bold=True, max_lines=2))
                bullets = self.bullet_paras(col["bullets"], size, sub)
                if paras and bullets:
                    bullets[0]["space_before"] = 0.5 * size
                return paras + bullets
            items.append((side, box, make))
        groups = self.fit_group(items, "body2", "small")
        return [self.text_el(f"{side.title()} column", box, paras)
                for (side, box, _), paras in zip(items, groups)]

    def side_bullets(self, slide, box):
        paras = self.fit_one("bullets", box, lambda s, sub: self.bullet_paras(slide["bullets"], s, sub),
                             "body2", "small")
        return self.text_el("Body", box, paras)

    def plan_chart(self, slide, body):
        chart_box, els = body, []
        if slide.get("bullets"):
            chart_box, text_box = self.split(body, CHART_SHARE)
            els.append(self.side_bullets(slide, text_box))
        size = self.prof["sizes"]["chart"][0]
        self.check_categories(slide["chart"], chart_box, size)
        return [{"kind": "chart", "name": "Chart", "box": chart_box, "chart": slide["chart"], "size": size}] + els

    def check_categories(self, chart, box, size):
        ctype = chart["type"]
        cats = [str(c) for c in chart.get("categories") or []]
        if not cats or ctype in ("pie", "scatter"):
            return
        family = self.fonts["minor"]
        if ctype in ("bar", "stacked-bar"):
            limit = pt(box[2]) * 0.35
            for c in cats:
                if textfit.text_width_pt(c, family, size) > limit:
                    self.problem("category-labels-collide", "chart.categories",
                                 f"category {c!r} is wider than a third of the chart at {size} pt; "
                                 "shorten it", severity="warning")
                    return
            return
        # Categories along the bottom: each gets an equal slot. Past two wrapped
        # lines, or a word wider than its slot, PowerPoint tilts or skips labels.
        slot = pt(box[2]) * 0.9 / len(cats)
        for c in cats:
            words_fit = all(textfit.text_width_pt(w, family, size) <= slot * 0.95 for w in c.split())
            lines = len(textfit.wrap_lines(c, family, size, slot * 0.95))
            if not words_fit or lines > 2:
                self.problem("category-labels-collide", "chart.categories",
                             f"category {c!r} does not fit its {slot:.0f} pt slot at {size} pt; use a bar "
                             "chart for long names, or shorten them")
                return

    def plan_table(self, slide, body):
        table = slide["table"]
        cols = [str(c) for c in table["columns"]]
        rows = [[str(c) for c in row] for row in table["rows"]]
        numeric = [all(NUMERIC_CELL_RE.match(r[j]) for r in rows) for j in range(len(cols))]
        align = table.get("align") or ["right" if numeric[j] else "left" for j in range(len(cols))]
        hr = table.get("highlight_row")
        family = self.fonts["minor"]
        last = None
        for size, _ in self.sizes("table"):
            widths = self.table_widths(cols, rows, size, pt(body[2]), family)
            if widths is None:
                last = (size, None)
                continue
            heights = [self.row_height(cols, widths, size, family, True)]
            heights += [self.row_height(r, widths, size, family, i == hr) for i, r in enumerate(rows)]
            last = (size, sum(heights))
            if sum(heights) <= pt(body[3]) + 0.5:
                w_emu = [int(w * EMU_PT) for w in widths]
                w_emu[-1] += body[2] - sum(w_emu)
                h_emu = [int(round(h * EMU_PT)) for h in heights]
                box = (body[0], body[1], body[2], sum(h_emu))
                return [{"kind": "table", "name": "Table", "box": box, "columns": cols, "rows": rows,
                         "widths": w_emu, "heights": h_emu, "size": size, "align": align,
                         "highlight_row": hr}]
        size, needed = last
        if needed is None:
            self.problem("text-overflow", "table", f"a word is wider than its column at {size} pt; "
                                                   "shorten cells or drop a column")
        else:
            self.problem("text-overflow", "table", f"needs {needed:.0f} pt of height at {size} pt but the "
                                                   f"body has {pt(body[3]):.0f} pt. Cut rows or shorten cells")
        return []

    @staticmethod
    def table_widths(cols, rows, size, total, family):
        pad = 2 * CELL_MARGIN_X_PT + 1
        natural, minimum = [], []
        for j, head in enumerate(cols):
            cells = [(head, True)] + [(r[j], False) for r in rows]
            natural.append(max(textfit.text_width_pt(t, family, size, bold=b) for t, b in cells) + pad)
            minimum.append(max(textfit.text_width_pt(w, family, size, bold=b)
                               for t, b in cells for w in (t.split() or [""])) + pad)
        if sum(natural) <= total:
            return [w + (total - sum(natural)) * w / sum(natural) for w in natural]
        if sum(minimum) > total:
            return None
        spare = [n - m for n, m in zip(natural, minimum)]
        room = total - sum(minimum)
        return [m + room * s / sum(spare) for m, s in zip(minimum, spare)]

    @staticmethod
    def row_height(cells, widths, size, family, bold):
        lines = max(len(textfit.wrap_lines(c, family, size, w - 2 * CELL_MARGIN_X_PT, bold=bold))
                    for c, w in zip(cells, widths))
        return lines * textfit.LINE_HEIGHT * size + 2 * CELL_MARGIN_Y_PT

    def plan_image(self, slide, body):
        area, els = body, []
        if slide.get("bullets"):
            area, text_box = self.split(body, IMAGE_SHARE)
            els.append(self.side_bullets(slide, text_box))
        path = os.path.join(self.spec_dir, slide["image"]["path"])
        with Image.open(path) as im:
            iw, ih = im.size
        x, y, w, h = area
        scale = min(w / iw, h / ih)
        fw, fh = int(iw * scale), int(ih * scale)
        # Beside bullets the picture lines up with the text's top-left; alone it is centred.
        if slide.get("bullets"):
            fitted = (x, y, fw, fh)
        else:
            fitted = (x + (w - fw) // 2, y + (h - fh) // 2, fw, fh)
        return [{"kind": "image", "name": "Picture", "box": fitted, "path": path,
                 "alt": slide["image"]["alt"]}] + els

    def plan_big_number(self, slide, body):
        # Numbers use the minor font: a serif major font like Georgia has
        # old-style figures, which make a big number look uneven.
        els = [self.simple("Number", "number", self.sub_box(body, 0.0215, 0.43), slide["number"],
                           "big_number", color="accent1", font="minor", bold=True, anchor="b", max_lines=1),
               self.simple("Label", "label", self.sub_box(body, 0.473, 0.172), slide["label"], "body",
                           max_lines=2)]
        if slide.get("context"):
            els.append(self.simple("Context", "context", self.sub_box(body, 0.667, 0.333), slide["context"],
                                   "body2", color="tx2"))
        return els

    def plan_quote(self, slide, body):
        text = slide["quote"].strip()
        if not text.startswith(("“", '"', "‘")):
            text = f"“{text}”"
        return [self.simple("Quote", "quote", self.sub_box(body, 0.0215, 0.688), text, "quote", font="major"),
                self.simple("Attribution", "attribution", self.sub_box(body, 0.731, 0.269),
                            slide["attribution"], "small", color="tx2")]

    def plan_process(self, slide, body):
        steps = slide["steps"]
        boxes = self.columns(body, len(steps), PROCESS_GAP_IN)
        items = []
        for i, (step, box) in enumerate(zip(steps, boxes), 1):
            def make(size, sub, i=i, step=step):
                paras = [self.para(str(i), round(size * 1.5), font="minor", bold=True, color="accent1"),
                         self.para(step["label"], size, bold=True, space_before=0.3 * size)]
                if step.get("detail"):
                    paras.append(self.para(step["detail"], sub, space_before=0.4 * sub))
                return paras
            items.append((f"steps[{i - 1}]", box, make))
        groups = self.fit_group(items, "body2", "small")
        return [self.text_el(f"Step {i}", box, paras) for i, ((_, box, _), paras) in enumerate(zip(items, groups), 1)]

    def footer(self, slide, layout):
        els = []
        if slide.get("source"):
            box = self.box(MARGIN_IN, FOOTER_Y_IN, SOURCE_W_IN, FOOTER_H_IN)
            els.append(self.simple("Source", "source", box, slide["source"], "source", color="tx2",
                                   anchor="b", max_lines=1))
        if self.slide_numbers and layout != "cover":
            box = self.box(MARGIN_IN + CONTENT_W_IN - NUMBER_W_IN, FOOTER_Y_IN, NUMBER_W_IN, FOOTER_H_IN)
            size = self.prof["sizes"]["source"][0]
            els.append({"kind": "slidenum", "name": "Slide number", "box": box, "size": size})
        return els
