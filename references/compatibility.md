# Compatibility

What the deck has been checked against, what it has not, and what to tell the
user at delivery.

## Renderers

| Renderer | Status | What it gives |
|---|---|---|
| PowerPoint (Windows, or WSL calling Windows) | **Measured.** The renderer the skill was built and tested against. | Pictures and `measure.json`: every shape's position and its text's measured bounds and line count. The lint's overflow check is a measurement. |
| LibreOffice (`soffice` headless) | **Fallback, untested.** Used only when PowerPoint is not reachable. | Pictures only. The lint reports `not-measured`, and overflow is the builder's prediction from font metrics, not a measurement. |
| Keynote | **Untested.** | Nothing; no authoritative source on its .pptx import was found. |
| Google Slides | **Untested.** | Nothing. |

`render.py` picks PowerPoint when it can (`--backend auto`). It opens the deck
read-only and without a window, closes only the presentation it opened, and
quits PowerPoint only if PowerPoint was not already running, so the user's
open decks are never touched. One render runs at a time: a second waits for
the first, because PowerPoint and LibreOffice renders are memory-heavy.

LibreOffice's known gaps (documented in bug reports from 2013 to 2020, so
directional rather than current): it ignores placeholder autofit, draws chart
axis intervals and pie sizes differently, and lays text out differently even
with metric-compatible substitute fonts. A deck rendered only in LibreOffice
has not been measured; say so at delivery, and say the final check needs
PowerPoint.

## Fonts

The build does not embed fonts. PowerPoint embeds only TrueType-outline fonts
that are licensed for it, and embedding makes files large and sometimes
read-only, so the policy is fonts present everywhere instead.

- The metric-safe fonts are Arial, Calibri, Cambria, Georgia, Times New Roman
  and Courier New. They exist on Windows and Mac, and Linux has
  metric-compatible substitutes (Liberation, Carlito, Caladea, Gelasio), so
  line breaks survive. The presets use only these.
- Any other font (a template's brand font, often) triggers
  `font-not-metric-safe`. The PowerPoint render still measures it correctly
  on this machine, but a viewer without the font sees a substitute with
  different widths, and text that fit here may wrap or overflow there. Tell
  the user, and suggest they check the deck on a machine like their
  audience's.
- Google Slides substitutes any font outside Google Fonts. A deck headed for
  Google Slides should use Arial or another font Google Slides has.

## What the build leaves out, and why

The builder writes only native, plain objects: text boxes, the title
placeholder, native charts and tables, and pictures. It avoids the features
that cause repair prompts or render differently between applications:

- no shape styles (`p:style`), connectors, shadows, glow or 3D effects;
- no SmartArt, animation or slide transitions;
- no embedded fonts, OLE objects or macros;
- one slide master, the template's own or the default;
- slide IDs and relationship IDs from python-pptx's own numbering, never
  copied from another file.

Charts are native PowerPoint charts with their data embedded, so the user can
edit the numbers in PowerPoint (right-click, Edit Data). Tables are native
tables. Nothing is a picture of text.

## At delivery

Tell the user, in a sentence or two, only what applies:

- that the gate result was measured in PowerPoint, or that it was rendered
  with LibreOffice only and the text fit is predicted;
- any `font-not-metric-safe` font, and what a viewer without it will see;
- that Keynote and Google Slides were not tested, if they mentioned either;
- that the final look belongs in PowerPoint on their machine, because a
  picture of a slide is not the slide.
