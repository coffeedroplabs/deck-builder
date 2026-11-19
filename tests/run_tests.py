# /// script
# requires-python = ">=3.10"
# dependencies = ["python-pptx==1.0.2", "pillow==13.0.0", "pytest==9.2.0"]
# ///
"""The deck-builder test suite.

    uv run -q tests/run_tests.py [pytest arguments]

DECK_BUILDER_RENDER=1 also renders the canary in PowerPoint and lints it
measured. That test is opt-in because it starts PowerPoint and takes a minute.
"""

import copy
import json
import os
import shutil
import sys

import pytest
from lxml import etree
from pptx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
FIXTURES = os.path.join(HERE, "fixtures")
sys.path[:0] = [os.path.join(HERE, "..", "scripts"), FIXTURES]

import build_deck  # noqa: E402
import charts  # noqa: E402
import check_story  # noqa: E402
import inspect_template  # noqa: E402
import lint_deck  # noqa: E402
import ooxml  # noqa: E402
import render  # noqa: E402
import spec_check  # noqa: E402
import textfit  # noqa: E402
from make_template import make_template  # noqa: E402

EM_DASH = "—"


def canary():
    with open(os.path.join(FIXTURES, "canary.json"), encoding="utf-8") as f:
        return json.load(f)


def codes(problems):
    return {p["code"] for p in problems}


def slide_index(spec, layout, chart_type=None):
    for i, s in enumerate(spec["slides"]):
        if s["layout"] == layout and (chart_type is None or s.get("chart", {}).get("type") == chart_type):
            return i
    raise LookupError(layout)


@pytest.fixture
def deck_dir(tmp_path):
    shutil.copy(os.path.join(FIXTURES, "warehouse.png"), tmp_path)
    return tmp_path


def write_spec(folder, spec):
    path = folder / "deck.json"
    path.write_text(json.dumps(spec, indent=2), encoding="utf-8")
    return str(path)


def build_and_lint(folder, spec):
    spec_path = write_spec(folder, spec)
    deck = str(folder / "deck.pptx")
    report = build_deck.build(spec_path, deck)
    assert report["built"], spec_check.errors(report["problems"])
    return report, lint_deck.lint(deck, spec_path)


# -- the canary: every layout, one deck ---------------------------------------


def test_canary_spec_is_clean():
    assert spec_check.errors(spec_check.check_spec(canary(), FIXTURES)) == []


def test_canary_check_only_writes_no_deck(deck_dir):
    report = build_deck.build(write_spec(deck_dir, canary()), None, check_only=True)
    assert spec_check.errors(report["problems"]) == []
    assert report["built"] is False and report["deck"] is None
    assert len(report["slides"]) == len(canary()["slides"])
    assert not (deck_dir / "deck.pptx").exists()


def test_canary_builds_and_passes_the_gate(deck_dir):
    report, result = build_and_lint(deck_dir, canary())
    assert report["deck_sha256"] == result["deck_sha256"]
    assert result["passed"], result["errors"]
    assert result["measured"] is False
    assert result["approval"]["status"] == "missing"
    assert codes(result["warnings"]) == {"not-measured", "story-unapproved"}


def test_lint_reports_an_approved_story(deck_dir):
    spec = canary()
    spec_path = write_spec(deck_dir, spec)
    check_story.write_approval(spec_path, spec["slides"], "user, in chat",
                               str(deck_dir / "outline.approved.json"))
    _, result = build_and_lint(deck_dir, spec)
    assert result["approval"]["status"] == "approved"
    assert "story-unapproved" not in codes(result["warnings"])


def test_lint_fails_a_deck_edited_after_the_build(deck_dir):
    spec = canary()
    build_and_lint(deck_dir, spec)
    spec["slides"][1]["title"] = "One in five orders comes back, at twice the cost of sending it"
    result = lint_deck.lint(str(deck_dir / "deck.pptx"), write_spec(deck_dir, spec))
    assert "deck-spec-mismatch" in codes(result["errors"])


def test_lint_crash_overwrites_a_passing_verdict(deck_dir):
    spec_path = write_spec(deck_dir, canary())
    deck = deck_dir / "deck.pptx"
    out = deck_dir / "gate_result.json"
    assert build_deck.build(spec_path, str(deck))["built"]
    assert lint_deck.main([str(deck), "--spec", spec_path, "-o", str(out)]) == 0
    assert json.loads(out.read_text())["passed"] is True
    deck.write_bytes(b"not a zip file")
    assert lint_deck.main([str(deck), "--spec", spec_path, "-o", str(out)]) == 2
    verdict = json.loads(out.read_text())
    assert verdict["passed"] is False and verdict["crash"]


def test_build_crash_leaves_a_failed_report(deck_dir):
    spec_path = deck_dir / "deck.json"
    spec_path.write_text("{ not json", encoding="utf-8")
    assert build_deck.main([str(spec_path), "--check-only"]) == 2
    report = json.loads((deck_dir / "build_report.json").read_text())
    assert report["built"] is False and report["crash"]


def test_a_wrong_path_fails_cleanly(deck_dir, capsys):
    missing = str(deck_dir / "missing.json")
    assert build_deck.main([missing, "--check-only"]) == 2
    assert lint_deck.main([str(deck_dir / "missing.pptx"), "--spec", missing]) == 2
    assert render.main([str(deck_dir / "missing.pptx"), "-o", str(deck_dir / "render")]) == 2
    err = capsys.readouterr().err
    assert "Traceback" not in err
    assert f"Cannot read {missing}" in err and "missing.pptx" in err


# -- spec problems the build refuses ------------------------------------------


def _chart(spec, chart_type):
    return spec["slides"][slide_index(spec, "chart", chart_type)]["chart"]


def _unknown_key(spec):
    spec["slides"][1]["colour"] = "red"


def _waterfall_labels(spec):
    _chart(spec, "waterfall")["labels"] = True


def _waterfall_highlight(spec):
    _chart(spec, "waterfall")["highlight"] = "Trail jackets"


def _waterfall_total_off(spec):
    _chart(spec, "waterfall")["series"][0]["values"][-1] = 2.5


def _waterfall_below_zero(spec):
    chart = _chart(spec, "waterfall")
    chart["totals"] = ["Current cost"]
    chart["series"][0]["values"] = [0.2, -0.28, -0.19, -0.14, 0.1]


def _scatter_without_axes(spec):
    spec["slides"][slide_index(spec, "chart", "bar")]["chart"] = {
        "type": "scatter", "series": [{"name": "Orders", "values": [[1, 2], [3, 4], [5, 7]]}]}


def _pie_two_series(spec):
    spec["slides"][slide_index(spec, "chart", "bar")]["chart"] = {
        "type": "pie", "categories": ["A", "B"],
        "series": [{"name": "One", "values": [1, 2]}, {"name": "Two", "values": [3, 4]}]}


def _big_numbers_adjacent(spec):
    i = slide_index(spec, "big-number")
    twin = copy.deepcopy(spec["slides"][i])
    twin["title"] = "Handling a return costs twice the outbound trip"
    spec["slides"].insert(i + 1, twin)


def _low_contrast_theme(spec):
    spec["theme"] = {"base": "paper", "colors": {"text": "DDDDDD"}}


def _em_dash(spec):
    spec["slides"][slide_index(spec, "bullets")]["bullets"][0] = f"Only the chart changes {EM_DASH} not the products"


def _leftover(spec):
    spec["slides"][slide_index(spec, "bullets")]["bullets"][2] = "TODO: add the test length"


@pytest.mark.parametrize("mutate, code", [
    (_unknown_key, "unknown-key"),
    (_waterfall_labels, "labels-unsupported"),
    (_waterfall_highlight, "highlight-unsupported"),
    (_waterfall_total_off, "waterfall-mismatch"),
    (_waterfall_below_zero, "waterfall-negative"),
    (_scatter_without_axes, "scatter-axes"),
    (_pie_two_series, "too-many-series"),
    (_big_numbers_adjacent, "big-number-adjacent"),
    (_low_contrast_theme, "theme-contrast"),
    (_em_dash, "em-dash"),
    (_leftover, "leftover-placeholder"),
], ids=lambda v: v if isinstance(v, str) else None)
def test_bad_spec_is_refused(deck_dir, mutate, code):
    spec = canary()
    mutate(spec)
    assert code in codes(spec_check.errors(spec_check.check_spec(spec, str(deck_dir))))
    report = build_deck.build(write_spec(deck_dir, spec), str(deck_dir / "deck.pptx"))
    assert report["built"] is False
    assert not (deck_dir / "deck.pptx").exists()


def test_scatter_with_axes_is_accepted(deck_dir):
    spec = canary()
    _scatter_without_axes(spec)
    chart = spec["slides"][slide_index(spec, "chart", "scatter")]["chart"]
    chart.update(x_axis_title="Orders per day", value_axis_title="Returns per day")
    assert spec_check.errors(spec_check.check_spec(spec, str(deck_dir))) == []


# -- chart helpers ------------------------------------------------------------


@pytest.mark.parametrize("value, fmt, shown", [
    (1234.4, "#,##0", "1,234"),
    (0.164, "0%", "16%"),
    (2.8, '"$"0.0"m"', "$2.8m"),
    (-0.28, '"$"0.0"m"', "-$0.3m"),
    (1500000, '#,##0.0,,"m"', "1.5m"),
    (-0.01, "0.0", "0.0"),
    (12345, None, "12,345"),
    (0.5, "General", "0.5"),
])
def test_format_number(value, fmt, shown):
    assert charts.format_number(value, fmt) == shown


def _line_slide(spec, names):
    i = slide_index(spec, "chart", "bar")
    spec["slides"][i]["chart"] = {
        "type": "line", "categories": ["Jan", "Feb", "Mar", "Apr", "May", "Jun"], "number_format": "0%",
        "series": [{"name": n, "values": [0.1, 0.11, 0.19, 0.22, 0.14, 0.09 + k / 100]}
                   for k, n in enumerate(names)]}
    return i


def _plot_layout(deck, i):
    from pptx import Presentation
    chart = next(s.chart for s in Presentation(deck).slides[i].shapes if s.has_chart)
    return chart._chartSpace.chart.plotArea.find(qn("c:layout"))


def test_line_names_get_room_at_the_line_ends(deck_dir):
    spec = canary()
    i = _line_slide(spec, ["Running tights", "Trail jackets", "Socks"])
    report, result = build_and_lint(deck_dir, spec)
    assert "series-name-long" not in codes(report["problems"])
    assert result["passed"], result["errors"]
    layout = _plot_layout(str(deck_dir / "deck.pptx"), i)
    assert layout is not None
    width = float(layout.find(f"{qn('c:manualLayout')}/{qn('c:w')}").get("val"))
    assert 0.75 <= width < 1


def test_a_single_line_keeps_the_full_width(deck_dir):
    spec = canary()
    i = _line_slide(spec, ["Return rate"])
    build_and_lint(deck_dir, spec)
    assert _plot_layout(str(deck_dir / "deck.pptx"), i) is None


def test_a_long_series_name_warns(deck_dir):
    spec = canary()
    _line_slide(spec, ["Trail jackets and every other outerwear line we sell", "Socks"])
    report, result = build_and_lint(deck_dir, spec)
    assert "series-name-long" in codes(report["problems"])
    assert result["passed"], result["errors"]


MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def test_a_year_of_months_fits_a_line_but_not_a_column(deck_dir):
    spec = canary()
    i = _line_slide(spec, ["Return rate"])
    chart = spec["slides"][i]["chart"]
    chart["categories"] = MONTHS
    chart["series"][0]["values"] = [0.1 + k / 100 for k in range(12)]
    report, result = build_and_lint(deck_dir, spec)
    assert result["passed"], result["errors"]
    chart["type"] = "column"
    assert "too-many-categories" in codes(spec_check.errors(spec_check.check_spec(spec, str(deck_dir))))


def test_waterfall_series_floats_changes_on_the_running_total():
    assert charts.waterfall_series(["Start", "Up", "Down", "End"], [10, 5, -3, 12], ["Start", "End"]) == {
        "Base": [None, 10, 12, None],
        "Up": [None, 5, None, None],
        "Down": [None, None, 3, None],
        "Total": [10, None, None, 12],
    }


# -- the story gate -----------------------------------------------------------


def outline(*slides, profile="live"):
    return {"profile": profile, "slides": [{"layout": layout, "title": title} for layout, title in slides]}


def test_canary_story_is_clean():
    assert check_story.check_outline(canary()) == []


@pytest.mark.parametrize("slides, code", [
    ([("bullets", "Overview")], "title-topic"),
    ([("bullets", "Returns doubled after the size chart changed"),
      ("bullets", "Returns doubled after the size chart change")], "title-near-duplicate"),
    ([("bullets", "Returns doubled after the size chart changed"), ("bullets", "Questions?")], "closing-slide"),
    ([("cover", "Returns cost more than shipping"),
      ("bullets", "Returns doubled after March"),
      ("bullets", "Three lines drive most returned units"),
      ("bullets", "A sizing fix saves about $0.6m a year")], "conclusion-missing"),
    ([("cover", "Returns cost more than shipping"),
      ("bullets", "Returns doubled after March"),
      ("conclusion", "Approve the trail jacket test this week"),
      ("bullets", "A sizing fix saves about $0.6m a year")], "conclusion-not-last"),
    ([("bullets", "Returns doubled after March"),
      ("bullets", "Three lines drive most returned units"),
      ("bullets", "Returns doubled after March")], "duplicate-title"),
], ids=lambda v: v if isinstance(v, str) else None)
def test_story_problem(slides, code):
    assert code in codes(check_story.check_outline(outline(*slides)))


def test_appendix_needs_no_conclusion():
    problems = check_story.check_outline(outline(
        ("cover", "Returns cost more than shipping"),
        ("bullets", "Returns doubled after March"),
        ("conclusion", "Approve the trail jacket test this week"),
        ("section", "Appendix"),
        ("bullets", "Return rates by product line, January to June")))
    assert "conclusion-not-last" not in codes(problems)


def test_approval_goes_stale_when_a_title_changes(tmp_path):
    spec = canary()
    path = str(tmp_path / "outline.approved.json")
    assert check_story.approval_state(spec["slides"], path)["status"] == "missing"
    check_story.write_approval(str(tmp_path / "outline.json"), spec["slides"], "user, in chat", path)
    state = check_story.approval_state(spec["slides"], path)
    assert state["status"] == "approved" and state["approved_by"] == "user, in chat"
    spec["slides"][2]["title"] = "Returns cost twice what outbound shipping does"
    assert check_story.approval_state(spec["slides"], path)["status"] == "stale"


def test_approval_ignores_layout_changes(tmp_path):
    spec = canary()
    path = str(tmp_path / "outline.approved.json")
    check_story.write_approval(str(tmp_path / "outline.json"), spec["slides"], "user, in chat", path)
    spec["slides"][1]["layout"] = "bullets"
    assert check_story.approval_state(spec["slides"], path)["status"] == "approved"


def test_approve_requires_who_approved(tmp_path):
    path = tmp_path / "outline.json"
    path.write_text(json.dumps(canary()), encoding="utf-8")
    with pytest.raises(SystemExit):
        check_story.main([str(path), "--approve"])
    assert check_story.main([str(path), "--approve", "--by", "user, in chat"]) == 0
    assert (tmp_path / "outline.approved.json").exists()


# -- template mode ------------------------------------------------------------


@pytest.fixture(scope="module")
def templates(tmp_path_factory):
    folder = tmp_path_factory.mktemp("templates")
    return {"brand": make_template(str(folder / "brand.pptx")),
            "brand-bad": make_template(str(folder / "brand-bad.pptx"), logo_over_title=True),
            "brand-footer": make_template(str(folder / "brand-footer.pptx"), footer_logo=True)}


def test_inspector_accepts_the_brand_template(templates):
    report = inspect_template.inspect(templates["brand"])
    assert spec_check.errors(report["problems"]) == []
    assert report["layout_map"]["cover"] == "Title Slide"
    assert report["layout_map"]["bullets"] == "Title and Content"
    assert "template-art" not in codes(report["problems"])
    assert "template-safe-area" in codes(report["problems"])
    assert inspect_template.main([templates["brand"]]) == 0


def test_inspector_warns_about_a_logo_in_the_title(templates):
    assert "template-art" in codes(inspect_template.inspect(templates["brand-bad"])["problems"])


def test_inspector_rejects_an_unreadable_file(tmp_path):
    path = tmp_path / "broken.pptx"
    path.write_bytes(b"not a zip file")
    assert inspect_template.main([str(path)]) == 2


def template_spec(folder, template_path):
    shutil.copy(template_path, folder / "brand.pptx")
    spec = canary()
    del spec["theme"]
    spec["template"] = "brand.pptx"
    spec["template_fingerprint"] = inspect_template.inspect(str(folder / "brand.pptx"))["fingerprint"]
    return spec


def test_template_canary_passes_the_gate(deck_dir, templates):
    _, result = build_and_lint(deck_dir, template_spec(deck_dir, templates["brand"]))
    assert result["passed"], result["errors"]


def test_titles_covering_the_logo_fail_the_gate(deck_dir, templates):
    _, result = build_and_lint(deck_dir, template_spec(deck_dir, templates["brand-bad"]))
    assert "covers-template-art" in codes(result["errors"])


def test_the_footer_fills_the_templates_footer_placeholders(deck_dir, templates):
    _, result = build_and_lint(deck_dir, template_spec(deck_dir, templates["brand-footer"]))
    assert result["passed"], result["errors"]
    prs = ooxml.open_presentation(str(deck_dir / "deck.pptx"))
    checked = set()
    for slide in prs.slides:
        wanted = {name: ooxml.footer_placeholders(slide.slide_layout)[kind]
                  for name, kind in (("Source", "ftr"), ("Slide number", "sldNum"))}
        for shape in slide.shapes:
            if shape.name in wanted:
                ph = wanted[shape.name]
                assert (shape.left, shape.top, shape.width, shape.height) == (ph.left, ph.top, ph.width, ph.height)
                checked.add(shape.name)
    assert checked == {"Source", "Slide number"}
    assert "template-art" not in codes(inspect_template.inspect(templates["brand-footer"])["problems"])


def test_a_changed_template_is_refused(deck_dir, templates):
    spec = template_spec(deck_dir, templates["brand"])
    shutil.copy(templates["brand-bad"], deck_dir / "brand.pptx")
    report = build_deck.build(write_spec(deck_dir, spec), str(deck_dir / "deck.pptx"))
    assert report["built"] is False
    assert "template-changed" in codes(report["problems"])


def test_a_missing_mapped_layout_is_refused(deck_dir, templates):
    spec = template_spec(deck_dir, templates["brand"])
    spec["layout_map"] = {"chart": "No Such Layout"}
    report = build_deck.build(write_spec(deck_dir, spec), str(deck_dir / "deck.pptx"))
    assert "template-layout-missing" in codes(report["problems"])


def _title_placeholder(layout):
    return next(ph for ph in layout.placeholders if ph._element.ph.get("type") in ("title", "ctrTitle"))


def _lvl1(ph_element):
    lst = ph_element.find(f"{qn('p:txBody')}/{qn('a:lstStyle')}")
    lvl1 = lst.find(qn("a:lvl1pPr"))
    return lvl1 if lvl1 is not None else etree.SubElement(lst, qn("a:lvl1pPr"))


def _layout(path, name):
    prs = ooxml.open_presentation(path)
    return prs, next(lay for lay in prs.slide_layouts if lay.name == name)


def test_placeholder_style_comes_from_the_layout_first(templates):
    prs, layout = _layout(templates["brand"], "Title Only")
    ph = _title_placeholder(layout)
    _lvl1(ph._element).set("algn", "r")
    ph._element.find(f"{qn('p:txBody')}/{qn('a:bodyPr')}").set("anchor", "b")
    assert ooxml.placeholder_text_style(ph, prs.slide_master) == ("r", "b")


def test_placeholder_style_falls_back_to_the_master_text_style(templates):
    prs, layout = _layout(templates["brand"], "Title Only")
    ph = _title_placeholder(layout)
    for el in (ph._element, ph._base_placeholder._element):
        body = el.find(qn("p:txBody"))
        if body is None:
            continue
        for node in body.iter(qn("a:lvl1pPr"), qn("a:bodyPr")):
            node.attrib.pop("algn", None)
            node.attrib.pop("anchor", None)
    style = prs.slide_master._element.find(f"{qn('p:txStyles')}/{qn('p:titleStyle')}/{qn('a:lvl1pPr')}")
    style.set("algn", "ctr")
    assert ooxml.placeholder_text_style(ph, prs.slide_master) == ("ctr", "t")


# -- text fitting -------------------------------------------------------------


def test_size_steps():
    assert textfit.size_steps(24, 18) == [24, 22, 20, 18]
    assert textfit.size_steps(20, 19) == [20, 19]
    assert textfit.size_steps(18, 18) == [18]


def test_wrapped_lines_each_fit_the_width():
    text = "Fixing sizing on the top three product lines would remove most of the cost of returns"
    lines = textfit.wrap_lines(text, "Arial", 24, 300)
    assert len(lines) > 1
    assert " ".join(lines) == text
    assert all(textfit.text_width_pt(line, "Arial", 24) <= 300.01 for line in lines)


def test_fit_size_steps_down_then_gives_up():
    text = "Returns now cost more than outbound shipping across every product line"

    def paragraphs(size):
        return [{"text": text, "family": "Arial", "size": size}]

    size, result = textfit.fit_size(paragraphs, [24, 22, 20, 18], 900, 40, max_lines=1)
    assert size is not None and result["fits"] and result["lines"] == 1
    size, result = textfit.fit_size(paragraphs, [24, 22, 20, 18], 200, 40, max_lines=1)
    assert size is None and not result["fits"]


# -- measured in PowerPoint (opt-in) ------------------------------------------


@pytest.mark.skipif(os.environ.get("DECK_BUILDER_RENDER") != "1",
                    reason="set DECK_BUILDER_RENDER=1 to render in PowerPoint")
def test_canary_passes_measured_in_powerpoint(deck_dir):
    if not render.powerpoint_available():
        pytest.skip("PowerPoint is not reachable from here")
    build_and_lint(deck_dir, canary())
    deck = str(deck_dir / "deck.pptx")
    measure = render.render(deck, str(deck_dir / "render"), backend="powerpoint")
    assert measure["measured"] is True
    result = lint_deck.lint(deck, str(deck_dir / "deck.json"))
    assert result["measured"] is True
    assert result["passed"], result["errors"]
    assert "not-measured" not in codes(result["warnings"])


if __name__ == "__main__":
    sys.exit(pytest.main([__file__, "-p", "no:cacheprovider", *sys.argv[1:]]))
