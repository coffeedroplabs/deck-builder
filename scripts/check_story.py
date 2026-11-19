# /// script
# requires-python = ">=3.11"
# dependencies = []
# ///
"""Check a deck's story from its titles alone, and record the user's approval.

    uv run check_story.py outline.json
    uv run check_story.py outline.json --approve --by "user, in chat"

The outline is the spec's skeleton, {"profile": "live", "slides": [{"layout":
"cover", "title": "..."}, ...]}; a full spec works too. The script prints the
ghost deck (the titles in order and nothing else) and the story problems.
Show the user the ghost deck and get their yes before writing the spec: a
wrong argument found here costs a line, found after the build it costs every
slide. --approve writes outline.approved.json beside the outline; lint_deck.py
warns when the built deck's titles drift from it.

Exit 0 no errors (approval written, with --approve), 1 errors, 2 bad input.
"""

import argparse
import json
import os
import sys
from datetime import datetime, timezone

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import rules  # noqa: E402
import themes  # noqa: E402

TOOL = "deck-builder check_story"
APPROVAL_FILE = "outline.approved.json"

# Two neighbouring titles sharing this share of their content words usually
# make one point twice, the duplicate-slide defect reviewers flag in AI decks.
NEAR_DUPLICATE = 0.6
MIN_SHARED_WORDS = 3
# Shorter decks are usually a slide or two for someone else's talk, which
# need no conclusion of their own.
MIN_DECK_FOR_CONCLUSION = 4

STOPWORDS = {
    "a", "an", "the", "and", "or", "but", "of", "to", "in", "on", "at", "by", "for", "with",
    "from", "as", "is", "are", "was", "were", "be", "been", "it", "its", "this", "that",
    "these", "those", "we", "our", "us", "you", "your", "they", "their", "than", "then",
    "so", "if", "into", "over", "per", "can", "will", "would", "should", "do", "does",
    "did", "has", "have", "had", "not", "no", "more", "most", "less", "all", "each",
}


def _problem(severity, code, slide, message, field="title"):
    return {"severity": severity, "code": code, "slide": slide, "field": field, "message": message}


def content_words(title):
    words = set()
    for w in rules.WORD_RE.findall(rules.normalize(title).lower()):
        w = w.strip("'’.,%$€£/-")
        if not w or w in STOPWORDS:
            continue
        # Crude plural folding is enough to match "return" with "returns".
        if len(w) > 3 and w.endswith("s") and not w.endswith("ss"):
            w = w[:-1]
        words.add(w)
    return words


def similarity(a, b):
    wa, wb = content_words(a), content_words(b)
    shared = wa & wb
    if len(shared) < MIN_SHARED_WORDS:
        return 0.0
    return len(shared) / len(wa | wb)


def _is_appendix(slide):
    return (slide.get("layout") == "section"
            and rules.normalize(slide.get("title")).lower().startswith("appendix"))


def _is_closing(title):
    closing = {t.rstrip("?!.") for t in rules.CLOSING_TITLES}
    return rules.normalize(title).lower().rstrip("?!.:") in closing


def story_problems(slides):
    """Deck-structure problems, shared with lint_deck.py so a story cannot pass here and fail there."""
    problems = []
    for n, (a, b) in enumerate(zip(slides, slides[1:]), 2):
        if a.get("layout") in rules.CONTENT_LAYOUTS and b.get("layout") in rules.CONTENT_LAYOUTS:
            score = similarity(a.get("title"), b.get("title"))
            if score >= NEAR_DUPLICATE:
                problems.append(_problem("warning", "title-near-duplicate", n,
                                         f"makes nearly the same point as slide {n - 1}; merge them, "
                                         "or say what this slide adds"))
    for n, s in enumerate(slides, 1):
        if _is_closing(s.get("title")):
            problems.append(_problem("warning", "closing-slide", n,
                                     "a courtesy slide; end on the conclusion and its ask, "
                                     "and leave it up during questions"))
    main = []
    for n, s in enumerate(slides, 1):
        if _is_appendix(s):
            break
        main.append((n, s))
    if len(main) >= MIN_DECK_FOR_CONCLUSION:
        conclusions = [n for n, s in main if s.get("layout") == "conclusion"]
        if not conclusions:
            problems.append(_problem("warning", "conclusion-missing", None,
                                     "no conclusion slide; end the main deck on the takeaways and the ask",
                                     field="layout"))
        elif main[-1][0] != conclusions[-1]:
            problems.append(_problem("warning", "conclusion-not-last", conclusions[-1],
                                     "the conclusion should be the last slide before the appendix; "
                                     "slides after it read as an afterthought", field="layout"))
    return problems


def check_outline(outline):
    if not isinstance(outline, dict) or not isinstance(outline.get("slides"), list) or not outline["slides"]:
        return [_problem("error", "bad-type", None, "the outline needs a non-empty \"slides\" list", field="")]
    try:
        prof = themes.profile(outline.get("profile"))
    except (ValueError, TypeError) as e:
        return [_problem("error", "bad-profile", None, str(e), field="profile")]
    problems = []
    seen = {}
    slides = outline["slides"]
    for n, s in enumerate(slides, 1):
        if not isinstance(s, dict):
            problems.append(_problem("error", "bad-type", n, "each slide is an object", field=""))
            continue
        if s.get("layout") not in rules.LAYOUTS:
            problems.append(_problem("error", "bad-layout", n,
                                     f"layout {s.get('layout')!r} is not one of {', '.join(sorted(rules.LAYOUTS))}",
                                     field="layout"))
        for sev, code, msg in rules.title_problems(s, prof) + rules.text_problems(rules.normalize(s.get("title"))):
            problems.append(_problem(sev, code, n, msg))
        key = rules.normalize(s.get("title")).lower()
        if key and s.get("layout") in rules.CONTENT_LAYOUTS:
            if key in seen:
                problems.append(_problem("warning", "duplicate-title", n,
                                         f"same title as slide {seen[key]}; each slide makes its own point"))
            seen.setdefault(key, n)
    if all(isinstance(s, dict) for s in slides):
        problems += story_problems(slides)
    return problems


def ghost_deck(slides):
    lines = []
    for n, s in enumerate(slides, 1):
        layout = s.get("layout") if isinstance(s, dict) else "?"
        title = rules.normalize(s.get("title")) if isinstance(s, dict) else ""
        lines.append(f"{n:>3}  {str(layout):<10}  {title or '(no title)'}")
    return "\n".join(lines)


def approval_state(slides, approval_path):
    """Whether the titles still match what the user approved: missing, approved or stale."""
    if not os.path.exists(approval_path):
        return {"status": "missing", "path": approval_path}
    with open(approval_path, encoding="utf-8") as f:
        record = json.load(f)
    status = "approved" if record.get("titles_hash") == rules.titles_hash(slides) else "stale"
    return {"status": status, "path": approval_path, "approved_by": record.get("approved_by"),
            "approved_at": record.get("approved_at")}


def write_approval(outline_path, slides, by, out):
    record = {"tool": TOOL, "version": themes.VERSION, "outline": os.path.abspath(outline_path),
              "titles_hash": rules.titles_hash(slides), "approved_by": by,
              "approved_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
              "titles": [rules.normalize(s.get("title")) for s in slides]}
    with open(out, "w", encoding="utf-8") as f:
        json.dump(record, f, indent=2, ensure_ascii=False)
    return record


def _format(p):
    where = f"slide {p['slide']}" if p.get("slide") else "deck"
    return f"{p['severity'].upper():7} {p['code']:22} {where}: {p['message']}"


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("outline", help="outline.json, or a full deck spec")
    ap.add_argument("--approve", action="store_true", help="record that the user approved these titles")
    ap.add_argument("--by", help="who approved, and how: \"user, in chat\"")
    ap.add_argument("-o", "--out", help=f"approval file (default: {APPROVAL_FILE} beside the outline)")
    args = ap.parse_args(argv)
    if args.approve != bool(args.by):
        ap.error("--approve and --by go together: approval records who gave it")
    try:
        with open(args.outline, encoding="utf-8") as f:
            outline = json.load(f)
    except (OSError, json.JSONDecodeError) as e:
        print(f"Cannot read {args.outline}: {e}", file=sys.stderr)
        return 2
    problems = check_outline(outline)
    errors = [p for p in problems if p["severity"] == "error"]
    slides = outline.get("slides") if isinstance(outline, dict) else None
    if isinstance(slides, list):
        print("Ghost deck: read only the titles. They should carry the whole argument.\n")
        print(ghost_deck(slides))
        print()
    for p in problems:
        print(_format(p))
    if problems:
        print()
    out = args.out or os.path.join(os.path.dirname(os.path.abspath(args.outline)), APPROVAL_FILE)
    if errors:
        print(f"{len(errors)} error(s). Fix them before showing the user the outline.")
        return 1
    warnings = len(problems)
    if args.approve:
        write_approval(args.outline, slides, args.by, out)
        print(f"Approved by {args.by}. Recorded in {out}.")
        return 0
    state = approval_state(slides, out)
    if state["status"] == "approved":
        tail = f"Titles match the approval by {state['approved_by']} ({state['approved_at']})."
    elif state["status"] == "stale":
        tail = (f"Titles changed since {state['approved_by']} approved them. Show the user this ghost "
                "deck again, then re-run with --approve.")
    else:
        tail = ("Show the user this ghost deck and ask them to approve or change the titles. "
                f"Then run: check_story.py {args.outline} --approve --by \"user, in chat\"")
    print(f"{len(slides)} slides, {warnings} warning(s). {tail}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
