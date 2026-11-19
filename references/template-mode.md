# Template mode

Company templates are where generated decks most often go wrong: content
lands on the logo, titles take the wrong font, layouts get picked by guess.
Template mode keeps the template in charge of everything it defines (theme,
fonts, colours, title and body positions, alignment, logos) and keeps the
builder in charge only of what goes inside those areas.

## 1. Inspect the template first

```
uv run -q <skill>/scripts/inspect_template.py path/to/brand.potx --json template_report.json
```

It accepts .pptx and .potx. It prints:

- the slide size, and a `not-widescreen` warning if it is not 16:9 (the
  builder scales its grid to any size, but check the render closely);
- the theme fonts and the colours the build will use for each role (text,
  background, muted, accent, series), with `font-not-metric-safe` and
  `theme-contrast` warnings where they apply;
- every layout with its placeholders and its fixed art (logos, bands,
  pictures on the layout or the master), each art piece marked with the area
  it sits in;
- problems: `no-title-layout` and `no-content-layout` (errors: the build
  cannot use the template as it is), `template-art` (a logo inside a title,
  body or footer area), `template-safe-area` (a placeholder within 5% of the
  slide edge);
- the spec keys to paste: `template`, a suggested `layout_map` and
  `template_fingerprint`.

Exit 0 means usable, 1 means the build cannot use it without a hand-written
`layout_map`, 2 means the file is not a readable template.

Show the user anything surprising before writing the spec: a template with
no content layout, a logo inside the title area, a 4:3 slide size.

## 2. Map the layouts

The spec's layouts (`cover`, `bullets`, `chart` and so on) are mapped to the
template's layouts by name:

```json
"template": "brand.potx",
"layout_map": {
  "cover": "Title Slide",
  "section": "Section Header",
  "bullets": "Title and Content",
  "chart": "Title Only"
},
"template_fingerprint": "3f9a..."
```

- `template` is relative to the spec file.
- Without a `layout_map` entry, `cover` uses "Title Slide", `section` uses
  "Section Header", and every other layout uses "Title and Content", else
  "Title Only". When a default name is missing from the template, the build
  falls back to the content layout; when a name you wrote in `layout_map` is
  missing, it refuses (`template-layout-missing`) and lists the template's
  layouts.
- Map content layouts to a layout with one title and at most one body area.
  The builder lays out charts, columns, tables and steps inside the body area
  itself, so a two-content layout would only give it the left half. The
  inspector's suggestion already follows this; check it against the layout
  list, because layout names and intent vary between templates.
- `theme` is ignored in template mode (`theme-ignored` warning), and
  `layout_map` without a template does nothing (`layout-map-unused`).

## 3. What the builder takes from the template

- **Theme**: fonts, colours and the colour map. Text is coloured by theme
  role, so a brand palette flows through without hex values in the spec.
- **Title position**: every slide keeps the layout's title placeholder,
  filled with the title, at the template's position, alignment and anchor.
- **Body area**: the layout's body placeholder marks where content goes. The
  builder removes the placeholder itself and draws its own text boxes,
  charts and tables inside that area, so no "Click to add text" prompt is
  left behind. Cover and section subtitles follow the placeholder's
  alignment and anchor; content on the other layouts is set top-left inside
  the area, as in the default grid. A layout with no body placeholder uses
  the default grid's body area, scaled to the slide.
- **Insets are zero.** The builder fits text to the exact box, so the
  template's inner margins are set to zero on the boxes it fills. Text may
  sit closer to a placeholder's edge than in the template's own sample
  slides; that is expected.
- **Statement slides** keep the template's title column (its left edge and
  width) so they line up with the other slides, but are centred vertically as
  in the default grid, because templates have no statement layout.
- **Footer row**: the source line and slide number go in the default grid's
  footer row, scaled to the slide. The template's own footer, date and
  slide-number placeholders are not used.
- **Font sizes stay the profile's.** The template decides the typeface; the
  profile decides the sizes and floors, so a template with 14 pt body text
  still gets profile-sized text in a `live` deck.

## 4. Template art

Logos, badges and bands drawn on a layout or on the master (unless the
layout hides master shapes) appear behind every slide, but they are not
placeholders, so the builder's own geometry cannot see them. The inspector
lists them and the lint checks every slide against them:

- Art covering more than 25% of the slide counts as a background (a colour
  band, a photo panel) that content is meant to sit on, and is not checked.
- Smaller art is meant to stay clear. Slide content that covers it fails the
  gate with `covers-template-art`. The fix is to map those slides to another
  layout, usually one whose title area the template's designer already
  narrowed to clear the logo, or to shorten the content so it stays inside
  its area.
- `template-art` from the inspector is the early warning: art already inside
  an area the builder fills.

## 5. The safe area

The lint warns `outside-safe-area` for content within 5% of the slide edge,
where projectors and video calls may crop. Placeholders the template puts
there are the designer's choice, which the deck author cannot change, so the
lint skips placeholders sitting exactly where their layout puts them; the
inspector reports each such layout once as `template-safe-area`. Look at the
render for those layouts and tell the user if anything looks cramped.

## 6. When the template changes

The spec records the template's fingerprint: a hash of its slide size,
layouts (names and placeholder positions), theme colours, colour map and
fonts. If the template file changes after the spec was written, the build
refuses with `template-changed`. Re-run the inspector, check the
`layout_map` against the new layout list, and paste the new fingerprint.
Never delete the fingerprint to get past the refusal: it exists so a moved
placeholder or a new logo cannot silently break a deck that passed before.
