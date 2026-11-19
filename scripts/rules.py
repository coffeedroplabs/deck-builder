"""Text rules shared by the story check, the spec check and the deck lint.

The same functions run at outline time, at build time and on the finished
.pptx, so a slide cannot pass one stage and fail the next on a different
definition of a word, a title or a leftover placeholder.
"""

import hashlib
import re
import unicodedata

CONTENT_LAYOUTS = {
    "statement", "bullets", "two-column", "chart", "table", "image",
    "big-number", "quote", "process", "conclusion",
}
LAYOUTS = CONTENT_LAYOUTS | {"cover", "section"}

# Template text, unfinished drafting and assistant chatter. Any hit is an
# error: these never belong on a delivered slide.
LEFTOVER_RE = re.compile(
    r"click to (add|edit)|lorem ipsum|\bTODO\b|\bTBD\b|\bTBC\b|\bFIXME\b|\bXX+\b|\?\?\?|"
    r"\[(insert|add|your|placeholder|source|citation)[^\]]*\]|\{\{|\}\}|<placeholder>|"
    r"\byour (title|text|name|company|logo) here\b|\bplaceholder\b|"
    r"\bas an ai\b|\bi hope this helps\b|\bcertainly[,!]|^here (is|are) (a|an|the|your)\b",
    re.IGNORECASE | re.MULTILINE,
)

# Words that mark text as machine-written or as filler. A warning, because
# some have honest uses ("a robust estimate"), but each one gets looked at.
SLOP_RE = re.compile(
    r"\b(delve[sd]?|delving|leverag(e|es|ed|ing)|unlock(s|ed|ing)?|seamless(ly)?|"
    r"game[- ]chang(er|ing)|cutting[- ]edge|fast[- ]paced|elevat(e|es|ed|ing)|"
    r"empower(s|ed|ing)?|synerg(y|ies)|harness(es|ed|ing)?|revolutioni[sz](e|es|ed|ing)|"
    r"transformative|paradigm|holistic|best[- ]in[- ]class|world[- ]class|"
    r"next[- ]gen(eration)?|unparalleled|tapestry|ever[- ]evolving|"
    r"in today's|it's not just|not only .{1,40} but also|robust)\b",
    re.IGNORECASE,
)

EM_DASH = "—"
EMOJI_RE = re.compile("[\U0001F300-\U0001FAFF\U00002600-\U000027BF\U0001F000-\U0001F2FF️]")
MARKDOWN_RE = re.compile(r"\*\*|__|^#{1,6} |`|\[[^\]]+\]\([^)]+\)", re.MULTILINE)
MANUAL_BULLET_RE = re.compile(r"^\s*([-*•–▪●]|\d+[.)])\s+")

# Whole titles that name a topic instead of making a point.
TOPIC_TITLES = {
    "introduction", "intro", "overview", "background", "agenda", "summary",
    "executive summary", "conclusion", "conclusions", "next steps", "questions",
    "q&a", "q & a", "thank you", "thanks", "results", "findings", "key findings",
    "recommendations", "context", "the problem", "problem", "solution",
    "the solution", "market", "market opportunity", "team", "our team",
    "traction", "financials", "roadmap", "timeline", "approach", "methodology",
    "key takeaways", "takeaways", "discussion", "update", "status",
}
CLOSING_TITLES = {"thank you", "thanks", "questions", "questions?", "any questions?", "q&a", "q & a"}

NUMBER_WORDS = {"two": 2, "three": 3, "four": 4, "five": 5, "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10}
COUNT_RE = re.compile(
    r"\b(two|three|four|five|six|seven|eight|nine|ten|[2-9]|10)\s+(?:[\w-]+\s+){0,2}?"
    r"(reasons?|ways?|steps?|things?|risks?|options?|priorities|lessons?|takeaways?|"
    r"pillars?|drivers?|changes?|levers?|principles?|questions?|ideas?|goals?|actions?|"
    r"moves?|bets?|phases?|stages?|areas?|factors?|problems?|issues?|benefits?|"
    r"features?|trends?|findings?|recommendations?|decisions?|asks?)\b",
    re.IGNORECASE,
)
DIGIT_RE = re.compile(r"\d")
WORD_RE = re.compile(r"[^\W_][\w'’%.,$€£/-]*", re.UNICODE)


def normalize(text):
    return " ".join(unicodedata.normalize("NFC", text or "").split())


def word_count(text):
    return len(WORD_RE.findall(text or ""))


def _bullet_texts(bullets):
    for b in bullets or []:
        if isinstance(b, str):
            yield "bullet", b
        elif isinstance(b, dict):
            yield "bullet", b.get("text", "")
            for s in b.get("sub") or []:
                yield "sub-bullet", s


def slide_text_fields(slide):
    """(field, text) for every piece of text the slide shows, in reading order.

    Notes are excluded: they are checked on their own, because the talk track
    is where detail is meant to go.
    """
    out = [("title", slide.get("title", ""))]
    for key in ("subtitle", "byline", "body", "number", "label", "context", "quote", "attribution"):
        if slide.get(key):
            out.append((key, slide[key]))
    out.extend(_bullet_texts(slide.get("bullets")))
    for side in ("left", "right"):
        col = slide.get(side) or {}
        if col.get("heading"):
            out.append((f"{side}.heading", col["heading"]))
        out.extend((f"{side}.{f}", t) for f, t in _bullet_texts(col.get("bullets")))
    for i, step in enumerate(slide.get("steps") or []):
        out.append((f"steps[{i}].label", step.get("label", "")))
        if step.get("detail"):
            out.append((f"steps[{i}].detail", step["detail"]))
    if slide.get("ask"):
        out.append(("ask", slide["ask"]))
    table = slide.get("table") or {}
    for c in table.get("columns") or []:
        out.append(("table.columns", str(c)))
    for row in table.get("rows") or []:
        for cell in row:
            out.append(("table.cell", str(cell)))
    chart = slide.get("chart") or {}
    for c in chart.get("categories") or []:
        out.append(("chart.categories", str(c)))
    for s in chart.get("series") or []:
        out.append(("chart.series", str(s.get("name", ""))))
    if slide.get("source"):
        out.append(("source", slide["source"]))
    return [(f, t) for f, t in out if t]


def slide_word_load(slide):
    """Words on the slide face, not counting the title, source line or notes.

    This is the number the word limits in themes.PROFILES apply to, at outline
    time and on the built deck alike.
    """
    return sum(word_count(t) for f, t in slide_text_fields(slide) if f not in ("title", "source"))


def top_level_items(slide):
    """How many items a title's count ("three reasons") should match, or None."""
    layout = slide.get("layout")
    if layout in ("bullets", "conclusion"):
        return len(slide.get("bullets") or [])
    if layout == "process":
        return len(slide.get("steps") or [])
    if layout == "two-column":
        return 2
    if layout == "table":
        return len((slide.get("table") or {}).get("rows") or [])
    if layout == "chart":
        return len((slide.get("chart") or {}).get("categories") or [])
    return None


def title_count(title):
    m = COUNT_RE.search(title or "")
    if not m:
        return None
    word = m.group(1).lower()
    return NUMBER_WORDS.get(word, int(word) if word.isdigit() else None)


def text_problems(text):
    """(severity, code, message) for one piece of slide text."""
    problems = []
    if not text:
        return problems
    m = LEFTOVER_RE.search(text)
    if m:
        problems.append(("error", "leftover-placeholder", f"leftover or placeholder text {m.group(0)!r}"))
    if EM_DASH in text:
        problems.append(("error", "em-dash", "em dash; use a plain dash, a comma or two sentences"))
    if EMOJI_RE.search(text):
        problems.append(("error", "emoji", "emoji in slide text"))
    if MARKDOWN_RE.search(text):
        problems.append(("error", "markdown", "Markdown syntax; slide text is plain text"))
    if MANUAL_BULLET_RE.match(text):
        problems.append(("error", "manual-bullet", "typed bullet or number at the start; the layout adds bullets"))
    for m in SLOP_RE.finditer(text):
        problems.append(("warning", "slop-word", f"filler word {m.group(0)!r}; say the specific thing"))
    return problems


def title_problems(slide, prof):
    """(severity, code, message) for a slide title against the profile."""
    title = normalize(slide.get("title"))
    layout = slide.get("layout")
    if not title:
        return [("error", "title-missing", "every slide needs a title")]
    if layout not in CONTENT_LAYOUTS:
        return []
    problems = []
    words = word_count(title)
    if words > prof["max_title_words"]:
        problems.append(("warning", "title-long", f"title has {words} words; keep it to {prof['max_title_words']}"))
    lowered = title.lower().rstrip(".:!")
    if lowered in TOPIC_TITLES or (words <= 3 and not DIGIT_RE.search(title)):
        problems.append(("warning", "title-topic",
                         "title names a topic; write the point the slide proves as a sentence"))
    return problems


def titles_hash(slides):
    """Hash of the ordered titles, binding an approved outline to the built deck.

    Layouts are left out on purpose: changing how a point is shown does not
    change the story the user approved; changing, adding or reordering a
    title does.
    """
    joined = "\n".join(normalize(s.get("title")) for s in slides)
    return hashlib.sha256(joined.encode("utf-8")).hexdigest()
