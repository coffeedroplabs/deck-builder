"""Validate a deck spec before anything is built.

Every problem is a dict {severity, code, slide, field, message}; slide is
1-based or None for deck-level problems. Errors stop the build; warnings are
reported and must be read, but a deck may ship with them.
"""

import os
import traceback

import rules
import themes

TOP_KEYS = {"profile", "theme", "template", "layout_map", "meta", "slides", "template_fingerprint"}
META_KEYS = {"title", "author", "subject", "slide_numbers"}
COMMON_KEYS = {"layout", "title", "notes", "source"}
LAYOUT_KEYS = {
    "cover": {"subtitle", "byline"},
    "section": {"subtitle"},
    "statement": {"body"},
    "bullets": {"bullets"},
    "two-column": {"left", "right"},
    "chart": {"chart", "bullets"},
    "table": {"table"},
    "image": {"image", "bullets"},
    "big-number": {"number", "label", "context"},
    "quote": {"quote", "attribution"},
    "process": {"steps"},
    "conclusion": {"bullets", "ask"},
}
REQUIRED = {
    "bullets": ("bullets",), "two-column": ("left", "right"), "chart": ("chart",),
    "table": ("table",), "image": ("image",), "big-number": ("number", "label"),
    "quote": ("quote", "attribution"), "process": ("steps",), "conclusion": ("bullets",),
}
STRING_FIELDS = {"title", "notes", "source", "subtitle", "byline", "body", "number", "label",
                 "context", "quote", "attribution", "ask"}
SOURCED = {"chart", "table", "big-number"}
SIDE_BULLETS_MAX = 3  # beside a chart or image the bullets get under half the width
CONCLUSION_BULLETS_MAX = 3
CHART_TYPES = {"bar", "column", "stacked-bar", "stacked-column", "line", "pie", "waterfall", "scatter"}
CHART_KEYS = {"type", "categories", "series", "highlight", "number_format", "labels",
              "value_axis_title", "x_axis_title", "alt", "totals"}
NO_LABEL_TOGGLE = {"stacked-bar", "stacked-column", "waterfall", "scatter", "pie"}
TABLE_KEYS = {"columns", "rows", "highlight_row", "align"}
ALIGNS = {"left", "right", "center"}


def _is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _is_text(v):
    return isinstance(v, str) and v.strip() != ""


class _Checker:
    def __init__(self, spec, spec_dir):
        self.spec = spec
        self.spec_dir = spec_dir
        self.problems = []

    def add(self, severity, code, slide, field, message):
        self.problems.append({"severity": severity, "code": code, "slide": slide,
                              "field": field, "message": message})

    def error(self, code, slide, field, message):
        self.add("error", code, slide, field, message)

    def warn(self, code, slide, field, message):
        self.add("warning", code, slide, field, message)

    def unknown(self, obj, allowed, slide, field):
        for key in sorted(set(obj) - allowed):
            where = f"{field}.{key}" if field else key
            self.error("unknown-key", slide, where, f"unknown key {key!r}; allowed: {', '.join(sorted(allowed))}")

    # -- deck level ---------------------------------------------------------

    def run(self):
        spec = self.spec
        if not isinstance(spec, dict):
            self.error("bad-type", None, "", "the spec must be a JSON object")
            return self.problems
        self.unknown(spec, TOP_KEYS, None, "")
        prof = self.check_profile()
        self.check_theme()
        self.check_template()
        self.check_meta()
        slides = spec.get("slides")
        if not isinstance(slides, list) or not slides:
            self.error("missing-field", None, "slides", "slides must be a non-empty list")
            return self.problems
        clean = []
        for n, slide in enumerate(slides, 1):
            before = len([p for p in self.problems if p["severity"] == "error"])
            self.check_slide(n, slide, prof)
            if len([p for p in self.problems if p["severity"] == "error"]) == before:
                clean.append((n, slide))
        if prof:
            for n, slide in clean:
                self.check_slide_text(n, slide, prof)
            self.check_deck_rules(slides, clean)
        return self.problems

    def check_profile(self):
        name = self.spec.get("profile")
        if name is None:
            self.error("missing-field", None, "profile", f"profile is required: {', '.join(themes.PROFILES)}")
            return None
        try:
            return themes.profile(name)
        except (ValueError, TypeError) as e:
            self.error("bad-profile", None, "profile", str(e))
            return None

    def check_theme(self):
        if self.spec.get("template"):
            if "theme" in self.spec:
                self.warn("theme-ignored", None, "theme", "template mode uses the template's own theme; theme is ignored")
            return
        try:
            theme = themes.resolve_theme(self.spec.get("theme"))
        except (ValueError, TypeError) as e:
            self.error("bad-theme", None, "theme", str(e))
            return
        for msg in themes.theme_contrast_problems(theme):
            self.error("theme-contrast", None, "theme", msg)

    def check_template(self):
        template = self.spec.get("template")
        layout_map = self.spec.get("layout_map")
        if template is not None:
            if not _is_text(template):
                self.error("bad-type", None, "template", "template must be a path string")
            else:
                path = os.path.join(self.spec_dir, template)
                if not os.path.isfile(path):
                    self.error("template-missing", None, "template", f"template file not found: {path}")
                elif not template.lower().endswith((".pptx", ".potx")):
                    self.error("bad-type", None, "template", "template must be a .pptx or .potx file")
        if layout_map is not None:
            if not isinstance(layout_map, dict):
                self.error("bad-type", None, "layout_map", "layout_map must map layout names to template layout names")
            else:
                for key, value in layout_map.items():
                    if key not in rules.LAYOUTS:
                        self.error("bad-layout", None, f"layout_map.{key}", f"unknown layout {key!r}")
                    if not _is_text(value):
                        self.error("bad-type", None, f"layout_map.{key}", "template layout name must be a string")
            if template is None:
                self.warn("layout-map-unused", None, "layout_map", "layout_map only applies in template mode")
        fp = self.spec.get("template_fingerprint")
        if fp is not None and not _is_text(fp):
            self.error("bad-type", None, "template_fingerprint", "template_fingerprint must be a string")

    def check_meta(self):
        meta = self.spec.get("meta")
        if meta is None:
            return
        if not isinstance(meta, dict):
            self.error("bad-type", None, "meta", "meta must be an object")
            return
        self.unknown(meta, META_KEYS, None, "meta")
        for key in ("title", "author", "subject"):
            if key in meta and not isinstance(meta[key], str):
                self.error("bad-type", None, f"meta.{key}", f"meta.{key} must be a string")
        if "slide_numbers" in meta and not isinstance(meta["slide_numbers"], bool):
            self.error("bad-type", None, "meta.slide_numbers", "slide_numbers must be true or false")

    # -- slide structure ----------------------------------------------------

    def check_slide(self, n, slide, prof):
        if not isinstance(slide, dict):
            self.error("bad-type", n, "", "each slide must be an object")
            return
        layout = slide.get("layout")
        if layout not in rules.LAYOUTS:
            self.error("bad-layout", n, "layout", f"layout {layout!r} is not one of {', '.join(sorted(rules.LAYOUTS))}")
            return
        self.unknown(slide, COMMON_KEYS | LAYOUT_KEYS[layout], n, "")
        for key in STRING_FIELDS & set(slide):
            if not isinstance(slide[key], str):
                self.error("bad-type", n, key, f"{key} must be a string")
        if not _is_text(slide.get("title")):
            self.error("title-missing", n, "title", "every slide needs a title")
        for key in REQUIRED.get(layout, ()):
            if key not in slide or slide[key] in (None, "", [], {}):
                self.error("missing-field", n, key, f"{layout} slides need {key!r}")
        if layout in SOURCED and not _is_text(slide.get("source")):
            self.error("missing-source", n, "source",
                       f"{layout} slides need a source line saying where the numbers came from")
        if prof is None:
            return
        if layout in ("bullets", "conclusion") and "bullets" in slide:
            cap = CONCLUSION_BULLETS_MAX if layout == "conclusion" else prof["max_bullets"]
            code = "conclusion-bullets" if layout == "conclusion" else "too-many-bullets"
            self.check_bullets(n, "bullets", slide["bullets"], cap, prof, code)
        if layout in ("chart", "image") and "bullets" in slide:
            self.check_bullets(n, "bullets", slide["bullets"], SIDE_BULLETS_MAX, prof, "too-many-bullets")
        if layout == "two-column":
            for side in ("left", "right"):
                if side in slide:
                    self.check_column(n, side, slide[side], prof)
        if layout == "chart" and "chart" in slide:
            self.check_chart(n, slide["chart"], prof)
        if layout == "table" and "table" in slide:
            self.check_table(n, slide["table"], prof)
        if layout == "image" and "image" in slide:
            self.check_image(n, slide["image"])
        if layout == "process" and "steps" in slide:
            self.check_steps(n, slide["steps"])

    def check_bullets(self, n, field, bullets, cap, prof, code):
        if not isinstance(bullets, list) or not bullets:
            self.error("bad-type", n, field, "bullets must be a non-empty list")
            return
        if len(bullets) > cap:
            self.error(code, n, field, f"{len(bullets)} bullets; at most {cap} here. Cut, merge, or split the slide")
        for i, b in enumerate(bullets):
            where = f"{field}[{i}]"
            if isinstance(b, str):
                if not b.strip():
                    self.error("bad-type", n, where, "empty bullet")
                continue
            if not isinstance(b, dict):
                self.error("bad-type", n, where, "a bullet is a string or {\"text\", \"sub\"}")
                continue
            self.unknown(b, {"text", "sub"}, n, where)
            if not _is_text(b.get("text")):
                self.error("missing-field", n, f"{where}.text", "bullet text is required")
            sub = b.get("sub", [])
            if not isinstance(sub, list):
                self.error("bad-type", n, f"{where}.sub", "sub must be a list of strings")
                continue
            if len(sub) > prof["max_bullets"]:
                self.error("too-many-subbullets", n, f"{where}.sub",
                           f"{len(sub)} sub-points; at most {prof['max_bullets']}")
            for j, s in enumerate(sub):
                if isinstance(s, (dict, list)):
                    self.error("too-deep", n, f"{where}.sub[{j}]",
                               f"at most {prof['max_levels']} levels; a third level belongs in the notes")
                elif not _is_text(s):
                    self.error("bad-type", n, f"{where}.sub[{j}]", "sub-points are non-empty strings")

    def check_column(self, n, side, col, prof):
        if not isinstance(col, dict):
            self.error("bad-type", n, side, f"{side} must be {{\"heading\", \"bullets\"}}")
            return
        self.unknown(col, {"heading", "bullets"}, n, side)
        if "heading" in col and not _is_text(col["heading"]):
            self.error("bad-type", n, f"{side}.heading", "heading must be a non-empty string")
        if "bullets" not in col:
            self.error("missing-field", n, f"{side}.bullets", "each column needs bullets")
            return
        self.check_bullets(n, f"{side}.bullets", col["bullets"], prof["max_bullets"], prof, "too-many-bullets")

    def check_steps(self, n, steps):
        if not isinstance(steps, list):
            self.error("bad-type", n, "steps", "steps must be a list")
            return
        if not 3 <= len(steps) <= 5:
            self.error("steps-count", n, "steps", f"{len(steps)} steps; a process slide shows 3 to 5")
        for i, step in enumerate(steps):
            where = f"steps[{i}]"
            if not isinstance(step, dict):
                self.error("bad-type", n, where, "a step is {\"label\", \"detail\"}")
                continue
            self.unknown(step, {"label", "detail"}, n, where)
            if not _is_text(step.get("label")):
                self.error("missing-field", n, f"{where}.label", "each step needs a label")
            if "detail" in step and not isinstance(step["detail"], str):
                self.error("bad-type", n, f"{where}.detail", "detail must be a string")

    def check_image(self, n, image):
        if not isinstance(image, dict):
            self.error("bad-type", n, "image", "image must be {\"path\", \"alt\"}")
            return
        self.unknown(image, {"path", "alt"}, n, "image")
        if not _is_text(image.get("alt")):
            self.error("missing-alt", n, "image.alt", "describe what the picture shows and why it is here")
        path = image.get("path")
        if not _is_text(path):
            self.error("missing-field", n, "image.path", "image path is required")
        elif not os.path.isfile(os.path.join(self.spec_dir, path)):
            self.error("image-missing", n, "image.path", f"image not found: {os.path.join(self.spec_dir, path)}")

    def check_table(self, n, table, prof):
        if not isinstance(table, dict):
            self.error("bad-type", n, "table", "table must be {\"columns\", \"rows\"}")
            return
        self.unknown(table, TABLE_KEYS, n, "table")
        cols, rows = table.get("columns"), table.get("rows")
        if not isinstance(cols, list) or not cols or not all(_is_text(c) for c in cols):
            self.error("missing-field", n, "table.columns", "columns must be a non-empty list of strings")
            return
        if not isinstance(rows, list) or not rows:
            self.error("missing-field", n, "table.rows", "rows must be a non-empty list of rows")
            return
        if len(cols) > prof["max_table_cols"]:
            self.error("too-many-cols", n, "table.columns",
                       f"{len(cols)} columns; at most {prof['max_table_cols']} for this profile")
        if len(rows) > prof["max_table_rows"]:
            self.error("too-many-rows", n, "table.rows",
                       f"{len(rows)} rows; at most {prof['max_table_rows']}. Split it or move detail to an appendix")
        for i, row in enumerate(rows):
            if not isinstance(row, list):
                self.error("bad-type", n, f"table.rows[{i}]", "each row is a list of cells")
            elif len(row) != len(cols):
                self.error("table-row-length", n, f"table.rows[{i}]", f"{len(row)} cells for {len(cols)} columns")
            elif not all(isinstance(c, str) or _is_num(c) for c in row):
                self.error("bad-type", n, f"table.rows[{i}]", "cells are strings or numbers")
        hr = table.get("highlight_row")
        if hr is not None and (not isinstance(hr, int) or isinstance(hr, bool) or not 0 <= hr < len(rows)):
            self.error("bad-type", n, "table.highlight_row", f"highlight_row is a 0-based data row index below {len(rows)}")
        align = table.get("align")
        if align is not None and (not isinstance(align, list) or len(align) != len(cols)
                                  or any(a not in ALIGNS for a in align)):
            self.error("bad-type", n, "table.align", f"align lists one of left/right/center per column ({len(cols)})")

    def check_chart(self, n, chart, prof):
        if not isinstance(chart, dict):
            self.error("bad-type", n, "chart", "chart must be an object")
            return
        self.unknown(chart, CHART_KEYS, n, "chart")
        ctype = chart.get("type")
        if ctype not in CHART_TYPES:
            self.error("bad-type", n, "chart.type", f"chart type must be one of {', '.join(sorted(CHART_TYPES))}")
            return
        series = chart.get("series")
        if not isinstance(series, list) or not series:
            self.error("missing-field", n, "chart.series", "series must be a non-empty list")
            return
        for i, s in enumerate(series):
            if not isinstance(s, dict):
                self.error("bad-type", n, f"chart.series[{i}]", "a series is {\"name\", \"values\"}")
                return
            self.unknown(s, {"name", "values"}, n, f"chart.series[{i}]")
            if not _is_text(s.get("name")):
                self.error("missing-field", n, f"chart.series[{i}].name", "each series needs a name")
            if not isinstance(s.get("values"), list) or not s["values"]:
                self.error("missing-field", n, f"chart.series[{i}].values", "values must be a non-empty list")
                return
        names = [s.get("name") for s in series]
        cap = 1 if ctype in ("pie", "waterfall") else prof["max_series"]
        if len(series) > cap:
            self.error("too-many-series", n, "chart.series", f"{len(series)} series; {ctype} charts take at most {cap}")
        for key in ("number_format", "value_axis_title", "x_axis_title", "alt"):
            if chart.get(key) is not None and not isinstance(chart[key], str):
                self.error("bad-type", n, f"chart.{key}", f"{key} must be a string")
        if "labels" in chart:
            if not isinstance(chart["labels"], bool):
                self.error("bad-type", n, "chart.labels", "labels must be true or false")
            elif ctype in NO_LABEL_TOGGLE:
                self.error("labels-unsupported", n, "chart.labels",
                           f"{ctype} charts set their own labels; remove 'labels'")

        if ctype == "scatter":
            if "categories" in chart:
                self.error("bad-type", n, "chart.categories", "scatter charts take [x, y] pairs, not categories")
            for i, s in enumerate(series):
                if not all(isinstance(p, list) and len(p) == 2 and all(_is_num(v) for v in p) for p in s["values"]):
                    self.error("bad-type", n, f"chart.series[{i}].values", "scatter values are [x, y] number pairs")
            if not _is_text(chart.get("x_axis_title")) or not _is_text(chart.get("value_axis_title")):
                self.error("scatter-axes", n, "chart", "scatter charts need x_axis_title and value_axis_title")
            self.check_highlight(n, chart, names)
            return
        if chart.get("x_axis_title") is not None:
            self.error("bad-type", n, "chart.x_axis_title", "x_axis_title is for scatter charts only")

        cats = chart.get("categories")
        if not isinstance(cats, list) or not cats or not all(isinstance(c, str) or _is_num(c) for c in cats):
            self.error("missing-field", n, "chart.categories", "categories must be a non-empty list")
            return
        cats = [str(c) for c in cats]
        if len(cats) > prof["max_categories"]:
            self.error("too-many-categories", n, "chart.categories",
                       f"{len(cats)} categories; at most {prof['max_categories']}. Group the tail into 'Other'")
        if ctype == "pie" and len(cats) > prof["max_pie_slices"]:
            self.error("too-many-slices", n, "chart.categories",
                       f"{len(cats)} slices; at most {prof['max_pie_slices']}. Use a bar chart")
        ok = True
        for i, s in enumerate(series):
            vals = s["values"]
            if len(vals) != len(cats):
                self.error("series-length", n, f"chart.series[{i}].values",
                           f"{len(vals)} values for {len(cats)} categories")
                ok = False
            if not all(_is_num(v) for v in vals):
                self.error("bad-type", n, f"chart.series[{i}].values", "values must be numbers")
                ok = False
            elif ctype == "pie" and any(v < 0 for v in vals):
                self.error("bad-type", n, f"chart.series[{i}].values", "pie values cannot be negative")
        if ctype == "waterfall":
            if "highlight" in chart:
                self.error("highlight-unsupported", n, "chart.highlight",
                           "waterfalls colour rises and falls; highlight is not used")
            if ok:
                self.check_waterfall(n, chart, cats, series[0]["values"])
            return
        if "totals" in chart:
            self.error("bad-type", n, "chart.totals", "totals is for waterfall charts only")
        self.check_highlight(n, chart, cats if len(series) == 1 else names)

    def check_highlight(self, n, chart, known):
        hl = chart.get("highlight")
        if hl is None:
            return
        wanted = hl if isinstance(hl, list) else [hl]
        known = [str(k) for k in known]
        for h in wanted:
            if str(h) not in known:
                self.error("highlight-unknown", n, "chart.highlight",
                           f"highlight {h!r} is not one of: {', '.join(known)}")

    def check_waterfall(self, n, chart, cats, values):
        totals = chart.get("totals") or []
        if not isinstance(totals, list) or any(str(t) not in cats for t in totals):
            self.error("bad-type", n, "chart.totals", f"totals must list category names from: {', '.join(cats)}")
            return
        totals = {str(t) for t in totals}
        level = 0.0
        for i, (cat, v) in enumerate(zip(cats, values)):
            if cat in totals:
                # A total after the first bar must equal the bars before it, or
                # the chart shows a sum that does not add up.
                if i > 0 and abs(v - level) > 0.005 * max(abs(v), abs(level), 1e-9):
                    self.error("waterfall-mismatch", n, "chart.series[0].values",
                               f"total {cat!r} is {v} but the bars before it sum to {round(level, 6)}")
                level = v
            else:
                level += v
            if level < 0:
                self.error("waterfall-negative", n, "chart.series[0].values",
                           f"the running total goes below zero at {cat!r}; use a bar chart of changes instead")
                return

    # -- text, per slide ----------------------------------------------------

    def check_slide_text(self, n, slide, prof):
        layout = slide["layout"]
        for field, text in rules.slide_text_fields(slide):
            for sev, code, msg in rules.text_problems(text):
                self.add(sev, code, n, field, msg)
        for sev, code, msg in rules.title_problems(slide, prof):
            self.add(sev, code, n, "title", msg)
        notes = slide.get("notes", "")
        for sev, code, msg in rules.text_problems(notes):
            if code != "manual-bullet":  # notes may be written as a numbered talk track
                self.add(sev, code, n, "notes", msg)
        if layout not in rules.CONTENT_LAYOUTS:
            return
        if prof["notes_required"] and not _is_text(notes):
            self.warn("notes-missing", n, "notes", "add speaker notes; the talk track is where detail goes")
        load = rules.slide_word_load(slide)
        if load > prof["max_words"]:
            self.warn("word-load", n, "", f"{load} words on the slide; keep it to {prof['max_words']}. "
                                          "Move detail to the notes or split the slide")
        expected, items = rules.title_count(slide.get("title")), rules.top_level_items(slide)
        if expected is not None and items is not None and expected != items:
            self.warn("count-mismatch", n, "title", f"title promises {expected} but the slide shows {items}")
        if layout not in SOURCED and not _is_text(slide.get("source")):
            texts = [t for f, t in rules.slide_text_fields(slide) if f != "source"]
            if any(rules.has_number_claim(t) for t in texts):
                self.warn("number-unsourced", n, "source", "the slide states a figure; add a source line")

    # -- deck rules ---------------------------------------------------------

    def check_deck_rules(self, slides, clean):
        seen = {}
        for n, slide in clean:
            if slide["layout"] not in rules.CONTENT_LAYOUTS:
                continue
            key = rules.normalize(slide["title"]).lower()
            if key in seen:
                self.warn("duplicate-title", n, "title", f"same title as slide {seen[key]}; each slide makes its own point")
            seen.setdefault(key, n)
        big = [n for n, s in enumerate(slides, 1) if isinstance(s, dict) and s.get("layout") == "big-number"]
        if len(big) > themes.MAX_BIG_NUMBERS:
            self.error("big-number-limit", big[themes.MAX_BIG_NUMBERS], "layout",
                       f"{len(big)} big-number slides; at most {themes.MAX_BIG_NUMBERS} per deck")
        for a, b in zip(big, big[1:]):
            if b == a + 1:
                self.error("big-number-adjacent", b, "layout", f"big-number slides {a} and {b} are adjacent")
        run = 0
        for n, slide in enumerate(slides, 1):
            run = run + 1 if isinstance(slide, dict) and slide.get("layout") == "bullets" else 0
            if run == themes.MAX_SAME_LAYOUT_RUN + 1:
                self.warn("layout-run", n, "layout",
                          f"{run} bullet slides in a row; show one as a chart, process, two-column or statement")


def check_spec(spec, spec_dir="."):
    return _Checker(spec, spec_dir).run()


def errors(problems):
    return [p for p in problems if p["severity"] == "error"]


def crash_message(e):
    """What a script prints, and records, when it stops on an exception."""
    # A wrong path is the usual cause and needs no traceback to fix.
    if isinstance(e, FileNotFoundError):
        return f"Cannot read {e.filename}: {e.strerror}"
    return traceback.format_exc()


def format_problem(p):
    where = f"slide {p['slide']}" if p.get("slide") else "deck"
    field = f" {p['field']}" if p.get("field") else ""
    return f"{p['severity'].upper():7} {p['code']:22} {where}{field}: {p['message']}"
