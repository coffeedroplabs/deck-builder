# Deck spec reference

A deck is one JSON file. `build_deck.py` turns it into a .pptx; you never place
coordinates or write XML yourself, because hand-placed geometry is where overlap
and overflow come from. If a slide does not fit a layout, change the content or
the layout, not the renderer.

Paths in the spec (`template`, `image.path`) are relative to the spec file.
Unknown keys are an error, not a silent drop, so a typo cannot lose content.
All text is plain text: no Markdown, no HTML, no emoji.

## Top level

```json
{
  "profile": "live",
  "theme": "paper",
  "template": null,
  "layout_map": {},
  "meta": {"title": "Q3 churn review", "author": "Ops team", "subject": "", "slide_numbers": true},
  "slides": [ ... ]
}
```

| Key | Required | Meaning |
|---|---|---|
| `profile` | yes | `live`, `pitch` or `read-ahead`. Sets every size, floor and limit (see `scripts/themes.py`). |
| `theme` | no | A preset name (`paper`, `consulting`, `ink`) or `{"base": "paper", "colors": {...}, "fonts": {"major": "...", "minor": "..."}, "title_bold": true}`. Ignored in template mode. Brand colours must come from the user or their assets, never invented; the build refuses colours that fail contrast. |
| `template` | no | Path to a .pptx or .potx. Switches to template mode (see `template-mode.md`). |
| `layout_map` | template mode | Maps this spec's layout names to the template's layout names. |
| `meta` | no | Document properties. `slide_numbers` (default true) adds a slide number to every slide but the cover. |
| `slides` | yes | One object per slide, in order. |

## Fields every slide may have

| Key | Meaning |
|---|---|
| `layout` | Required. One of the layouts below. |
| `title` | Required on every layout. A full-sentence claim (a verb and a point), at most 15 words, except on `cover` and `section`. The title is the slide's message; the body is its evidence. |
| `notes` | Speaker notes. Expected on every content slide for `live` and `pitch` (the gate warns), because the talk track is where detail goes instead of the slide. |
| `source` | Source line in the footer. Required on `chart`, `table` and `big-number`, and on any slide whose text contains a number. Say where the number came from: "Source: Stripe export, Jul to Sep 2026". |

## Layouts

Bullets are strings, or `{"text": "...", "sub": ["...", "..."]}` for one level
of sub-points. Lists are capped by profile (live: 4 bullets, 4 sub-points each,
2 lines per bullet).

### `cover`
`title`, optional `subtitle`, optional `byline` (presenter and date).

### `section`
`title`, optional `subtitle`. Use between parts of a long deck, and once as
`"title": "Appendix"` before backup slides; the story check then expects the
conclusion right before the appendix.

### `statement`
`title` (the statement, set large and centred vertically), optional `body`
(one or two supporting sentences). For the one idea the audience must leave
with. Not more than a few per deck.

### `bullets`
`title`, `bullets`. The workhorse, so the gate warns when three slides in a row
share it; vary with `chart`, `two-column`, `process` or `statement`.

### `two-column`
`title`, `left` and `right`, each `{"heading": "...", "bullets": [...]}`
(heading optional). For comparisons: before and after, option A and B, us and
them. Both columns share one font size.

### `chart`
`title`, `chart`, `source`, optional `bullets` (at most 3, placed right of the
chart, which then takes 62% of the width).

```json
"chart": {
  "type": "column",
  "categories": ["Jul", "Aug", "Sep"],
  "series": [{"name": "Churn rate", "values": [0.041, 0.038, 0.029]}],
  "highlight": "Sep",
  "number_format": "0.0%",
  "labels": true,
  "value_axis_title": null,
  "alt": null
}
```

- `type`: `bar` (horizontal; use for long category names or rankings),
  `column`, `stacked-bar`, `stacked-column`, `line`, `pie` (at most 5 slices;
  prefer a bar), `waterfall`, `scatter`.
- `highlight`: a category name or list of names (single series) or a series
  name (several series). Highlighted marks take the accent; the rest take the
  muted series colour. Without it, a single series is all accent and several
  series take accent, muted, third in order. Highlight the thing the title
  claims.
- `labels`: data labels on the marks (default true for single-series bar,
  column and waterfall, false otherwise). With labels on, the value axis is
  hidden.
- `number_format`: Excel format code for labels and axis, such as `0%`,
  `0.0%`, `#,##0`, `$#,##0`, `$#,##0.0,,"M"`.
- `waterfall`: one series of changes; list the bars that are totals (drawn from
  zero) in `"totals": ["2025", "2026"]`. Rises take the accent, falls the third
  colour, totals the muted colour.
- `scatter`: each series' `values` is a list of `[x, y]` pairs; `categories` is
  omitted; `x_axis_title` and `value_axis_title` are required.
- `alt`: alt text. When omitted the builder writes one from the data ("Column
  chart: Churn rate by month. Jul 4.1%, Aug 3.8%, Sep 2.9%."), which is usually
  better than a hand-written summary because it carries the numbers.
- Multi-series line charts are labelled at the line ends instead of a legend.
  Bar and column charts with several series keep a legend at the top.
- The build refuses category labels that would collide at the profile's chart
  size: use `bar` instead of `column` for long names, or shorten them.

### `table`
`title`, `table`, `source`.

```json
"table": {
  "columns": ["Plan", "Price", "Churn"],
  "rows": [["Monthly", "$29", "4.1%"], ["Annual", "$290", "1.2%"]],
  "highlight_row": 1,
  "align": ["left", "right", "right"]
}
```

`align` is optional; columns whose cells all look numeric are right-aligned.
`highlight_row` (0-based, data rows only) sets that row in bold with the accent
colour. The profile caps rows and columns; past that, split the table or move
it to the appendix of a read-ahead.

### `image`
`title`, `image: {"path": "...", "alt": "..."}`, optional `bullets` (at most 3,
placed right of the image, which then takes 55% of the width). `alt` is
required: describe what the picture shows and why it is on the slide. Images
are fitted inside their area without cropping. Never use an image that has text
baked into it; put the words on the slide.

### `big-number`
`title`, `number` (the figure as it should read: `"42%"`, `"$1.2M"`, `"3.4x"`),
`label` (what it measures, one line), optional `context` (one comparison
line), `source`. At most two per deck, never adjacent, never in a row of tiles:
a grid of big numbers invites invented figures.

### `quote`
`title` (the claim the quote supports), `quote`, `attribution` (who said it and
where: "Head of Ops, logistics customer, interview 12 Aug").

### `process`
`title`, `steps`: 3 to 5 of `{"label": "...", "detail": "..."}` (detail
optional). For a real sequence only; numbering a list that is not a sequence is
a template habit.

### `conclusion`
`title`, `bullets` (at most 3 takeaways), optional `ask` (the decision or next
step you want, set apart in the accent colour). The last content slide; never
end on "Thank you" or "Questions".

## Example

```json
{
  "profile": "live",
  "theme": "paper",
  "meta": {"title": "Annual plans review", "author": "Growth team"},
  "slides": [
    {"layout": "cover", "title": "Annual plans cut churn by two thirds", "byline": "Growth team, 4 March 2027"},
    {"layout": "chart", "title": "Churn fell from 4.1% to 2.9% in the quarter annual plans launched",
     "chart": {"type": "column", "categories": ["Jul", "Aug", "Sep"],
               "series": [{"name": "Monthly churn", "values": [0.041, 0.038, 0.029]}],
               "highlight": "Sep", "number_format": "0.0%"},
     "source": "Source: Stripe export, Jul to Sep 2026",
     "notes": "Annual plans launched 1 September. The drop is all in September."},
    {"layout": "conclusion", "title": "Make annual the default plan from April",
     "bullets": ["Annual subscribers churn at a third of the monthly rate", "No change to monthly pricing"],
     "ask": "Approve annual-first checkout for the April release",
     "notes": "Ask for the decision, then stop talking."}
  ]
}
```
