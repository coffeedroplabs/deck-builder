---
name: deck-builder
description: Build new, editable PowerPoint decks (.pptx) that pass measured checks instead of an eyeball - a story you can follow from the titles alone, text that really fits (measured in PowerPoint), font floors, contrast, alt text, sourced numbers - in a default theme or the user's company template. Use this whenever the user wants a presentation, slide deck, slides, pitch deck, board or investor deck, read-ahead, keynote, training deck or .pptx made, wants notes, a report, an outline or data turned into slides, or wants content put into their company's PowerPoint template, even if they never say "deck". For reading, extracting from or lightly editing a .pptx someone else made, use the pptx skill instead.
---

# Deck builder

Decks fail in four places: the story (titles that name a topic instead of
making a point, slides that repeat each other), the geometry (text that
overflows or overlaps, which you cannot see in the XML), the template (company
templates go off the rails), and the file (repair prompts). This skill closes
each one with a script you run, not a judgement you make about your own work.

You write a JSON spec. A deterministic builder lays it out from locked layouts,
PowerPoint measures the result, and a lint writes the verdict. You never place
coordinates, write XML or edit the .pptx by hand: hand-placed geometry is where
overlap and overflow come from, and an edited .pptx no longer matches the spec
the gates checked.

**The verdict is `gate_result.json`, and only that.** Do not tell the user a
deck passed, fits or is ready unless the latest lint run wrote `"passed": true`
for the current deck file. Your own look at the pictures is a review, not a
gate.

## Running the scripts

Scripts live in this skill's `scripts/` folder. Run each with
`uv run -q <skill>/scripts/NAME.py ...`; their dependencies are pinned inline,
so nothing is installed globally. Every script prints what to do next, writes
a JSON report, and exits 0 for success, 1 for problems you fix in the spec,
2 for a crash or unreadable input.

| Script | Does |
|---|---|
| `check_story.py outline.json` | Story gate on the titles alone; `--approve --by "user, in chat"` records approval |
| `inspect_template.py brand.potx` | Reads a company template, suggests `layout_map` (template mode only) |
| `build_deck.py deck.json -o deck.pptx` | Checks the spec, fits every text, builds; refuses rather than shrink or overflow |
| `render.py deck.pptx -o render/` | PowerPoint renders PNGs and measures every text box; LibreOffice fallback renders only |
| `lint_deck.py deck.pptx --spec deck.json` | The gate: geometry, measured overflow, floors, contrast, alt text, story; writes `gate_result.json` |

Keep one folder per deck (the user's choice, else `./<deck-name>/` in the
working directory) holding `brief.md`, `outline.json`,
`outline.approved.json`, `deck.json`, `deck.pptx`, `render/` and
`gate_result.json`. Renders run one at a time (the script holds a lock), so
never start two.

## The workflow

Go through the phases in order. Each names the reference to read at that
point; do not load them all up front.

### 1. Brief

Find out, from the conversation and the user's files first and by asking only
for what is missing:

- **Profile**: `live` (presented in a room or on a call), `pitch` (short,
  persuasive, presented or sent) or `read-ahead` (read on a screen without a
  presenter). It sets every size and limit, so it comes first.
- **Audience and the ask**: who decides what, and what you want from them.
- **Material**: the data, documents and sources. Every number on a slide comes
  from the user's material or a source you can cite. Never invent a figure,
  round one into a nicer one, or fill a gap with a plausible estimate; ask, or
  leave the claim out.
- **Template**: a company .pptx or .potx, or none (then a preset theme). Brand
  colours come from the user or their assets, never from your guess.
- **Length**: a rough slide count or talk length.

Write the answers to `brief.md`.

### 2. Outline, then the user's yes

Write `outline.json`: `{"profile": "...", "slides": [{"layout": "...",
"title": "..."}]}`, titles only. Read `references/design-rules.md` (deck
structure and titles) first. Every content title is a full-sentence claim with
a verb, at most 15 words: "Returns cost more than outbound shipping", not
"Returns overview". Read in order, the titles alone should carry the whole
argument.

Run `check_story.py outline.json`. Fix its errors, consider its warnings, then
show the user the ghost deck it prints and ask them to approve or change the
titles. When they say yes, run it again with `--approve --by "user, in chat"`.
A wrong argument found here costs a line; found after the build it costs every
slide. If the user said up front to skip approval, record that instead
(`--by "user waived review, in chat"`).

### 3. Spec

Write `deck.json` from the approved outline. `references/spec.md` has every
field and layout; `references/design-rules.md` says which layout and chart to
choose for a message. With a template, read `references/template-mode.md` and
run `inspect_template.py` before writing the spec.

Titles in the spec should match the approved outline. If writing the evidence
shows a title is wrong, change it, but show the user the changed titles before
delivery; the lint warns `story-changed` until then.

### 4. Build

Run `build_deck.py deck.json -o deck.pptx`. A refusal writes nothing and names
the slide, field and code. Fix the spec with the fix ladder below and rebuild;
`--check-only` checks and fits without writing a file.

### 5. Render and lint

Run `render.py deck.pptx -o render/`, then
`lint_deck.py deck.pptx --spec deck.json`. Rebuild, re-render and re-lint
after every spec change: the lint refuses a measurement taken from an older
deck (`measure-stale`). Errors fail the gate. Warnings do not, but read each
one and either fix it or have a reason to keep it, which you will report.

### 6. Review

Once the gate passes, read `references/review.md` and run the independent
review: a blind reader that sees only the slide pictures, then three critics,
never more than three agents at a time. The review catches what scripts
cannot: a chart that does not show what its title claims, a number that does
not match its source, a deck that reads as generic.

### 7. Fix ladder

When a slide fails a gate or a review finding, fix the spec, in this order:

1. **Cut words.** Drop qualifiers, merge overlapping bullets, shorten labels.
2. **Move detail out.** Into the speaker notes (`live`, `pitch`) or an
   appendix slide (`read-ahead`).
3. **Split the slide** into two, each with its own claim. That changes the
   titles, so show the user the new ghost deck.
4. **Change the layout.** Bullets that compare become `two-column`; numbers in
   a sentence become a `chart`; a sequence becomes `process`.

Never shrink text below the profile's sizes, switch to a denser profile to fit
more words, or edit the .pptx. The builder already steps sizes down as far as
the profile allows; past that, the slide holds too much.

Then go back to step 4. After three rounds of review fixes, stop and deliver
with what is left listed.

### 8. Deliver

Read `references/compatibility.md`, then tell the user:

- the path of the .pptx, and of `render/contact.png` for a quick look;
- the gate result: errors (none, or you would not be delivering), the warnings
  you kept and why, and whether text was measured in PowerPoint or only
  predicted (LibreOffice render);
- what the review changed and what it raised that you did not change;
- that the final visual check belongs in PowerPoint on their machine, and
  anything from `compatibility.md` that applies (Keynote, Google Slides, fonts).

## Later changes

Change the spec and rebuild, then re-run render and lint; never patch the
.pptx. Object names are stable across rebuilds, so "make slide 4's chart a
bar" is a spec edit. If the user has edited the .pptx themselves, the lint
reports `deck-spec-mismatch`; ask before rebuilding, because a rebuild from
the spec discards their edits.

## Rules that are never traded away

- No invented numbers. Every figure is in the user's material or carries a
  source line.
- No text in images, no image-only slides, no emoji, no Markdown on slides, no
  em dash (the lint fails it; use a plain hyphen or reword).
- No "Thank you" or "Questions?" closer; end on the conclusion and the ask.
- No grids of big-number tiles: they invite made-up round figures.
- The gate's verdict is not yours to override. If a rule is wrong for this
  deck, say so to the user; do not work around the script.
