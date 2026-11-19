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
| `template_fingerprint` | template mode | The hash `inspect_template.py` prints. The build refuses a template that changed since (`template-changed`). |
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
  claims. On a single-series line, the highlighted category gets a marker and,
  with labels off, its value.
- Bars, columns and waterfalls always start at zero, so a bar's length is its
  value. A small change on a large base (margin from 38% to 42%) then looks
  flat: show it as a `line`, or chart the change itself.
- `labels`: data labels on the marks, for `bar`, `column` and `line` only
  (default true for single-series bar and column, false otherwise). With
  labels on, the value axis is hidden. Stacked charts, waterfalls, pies and
  scatters set their own labels, and the build refuses `labels` on them
  (`labels-unsupported`): pies label each slice with its name and value,
  waterfalls label every bar, stacked and scatter charts keep the axis.
- `number_format`: Excel format code for labels and axis, such as `0%`,
  `0.0%`, `#,##0`, `$#,##0`, `$#,##0.0,,"M"`. Write the currency the way the
  slide text does: if the titles say A$, use `"A$"#,##0.0,,"m"`, or the chart
  shows "$5.90m" beside a title that says "A$5.90m".
- `waterfall`: one series of changes; list the bars that are totals (drawn from
  zero) in `"totals": ["2025", "2026"]`, by category name. Rises take the
  accent, falls the third colour, totals the muted colour, so `highlight` is
  refused (`highlight-unsupported`). The build checks the arithmetic: a total
  after the first bar must equal the bars before it, within 0.5%
  (`waterfall-mismatch`), and the running total must not go below zero
  (`waterfall-negative`; show the changes as a bar chart instead). `totals` on
  any other chart type is an error.
- `scatter`: each series' `values` is a list of `[x, y]` number pairs;
  `categories` is omitted; `x_axis_title` and `value_axis_title` are required
  (`scatter-axes`). `x_axis_title` on any other chart type is an error.
  `highlight` names a series; several series get a legend.
- Series per chart: one for `pie` and `waterfall`; otherwise the profile's
  cap (3 for `live` and `pitch`, 4 for `read-ahead`).
- Categories per chart: the profile's cap (8 for `live`, 10 for `pitch`, 14
  for `read-ahead`, `too-many-categories`), except `line`, which is limited
  only by labels that collide. Never drop points to get under the cap: group a
  ranking's tail into "Other", or show a time series as a `line` or in coarser
  periods (quarters instead of months).
- A target, plan or threshold: add it as a second `line` series with the same
  value in every category, named for what it is ("SLA 4 h", "Plan"), and
  `highlight` the actual series so the target stays muted. The name at the
  line end labels it, so it needs no legend.
- `alt`: alt text. When omitted the builder writes one from the data ("Column
  chart: Churn rate by month. Jul 4.1%, Aug 3.8%, Sep 2.9%."), which is usually
  better than a hand-written summary because it carries the numbers.
- Multi-series line charts are labelled at the line ends instead of a legend.
  The build narrows the plot so the longest name fits right of its line;
  without that room PowerPoint pushes the names back across the lines. A name
  that would take more than a quarter of the chart's width warns
  `series-name-long`: shorten it, because every line loses that width. Bar and
  column charts with several series keep a legend at the top.
- The build refuses category labels that would collide at the profile's chart
  size: use `bar` instead of `column` for long names, or shorten them.
- It also refuses value labels wider than the room their marks leave
  (`value-labels-collide`). Columns in one cluster touch, so two series of
  six months with labels like "A$640k" collide. Set `"labels": false` so the
  axis carries the scale (usual for actual against plan), use a shorter
  `number_format`, or show fewer series or categories.

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
line), `source`. The figure is set in the body font, as chart and table
figures are, so the same number looks the same everywhere in the deck. At most two per deck, never adjacent, never in a row of tiles:
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
step you want, set in bold accent one line under the last bullet), and
`source` when a takeaway quotes a number, as on any slide. The last content
slide; never end on "Thank you" or "Questions".

## Outline

Before the spec, `check_story.py` checks an outline: the same top level with
only `profile` and, per slide, `layout` and `title`.

```json
{"profile": "live",
 "slides": [{"layout": "cover", "title": "Annual plans cut churn by two thirds"},
            {"layout": "chart", "title": "Churn fell from 4.1% to 2.9% in the quarter annual plans launched"},
            {"layout": "conclusion", "title": "Make annual the default plan from April"}]}
```

`check_story.py outline.json --approve --by "user, in chat"` records the user's
approval in `outline.approved.json` beside it, bound to a hash of the ordered
titles. The lint reads that file from beside the spec (or `--approval`) and
warns `story-unapproved` when it is missing and `story-changed` when the
deck's titles no longer match it. Layouts are not part of the hash: showing a
point differently does not change the approved story; changing, adding or
reordering a title does.

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
