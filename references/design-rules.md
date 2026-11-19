# Design rules

What to put on slides, and why the gates enforce what they do. Read the deck
structure and title sections before writing the outline, and the layout and
chart sections before writing the spec.

## Contents

1. Where the numbers come from
2. Deck structure by profile
3. Titles are claims
4. Numbers you can stand behind
5. Choosing a layout
6. Choosing a chart
7. Colour and type
8. What makes a deck look machine-made

## 1. Where the numbers come from

Every threshold lives in `scripts/themes.py`. Not all of them are equally
sound, and the user should hear which ones are conventions when they push back
on one.

| Rule | Value | Evidence |
|---|---|---|
| Text floor, presented decks | 18 pt (`live`), 16 pt (`pitch`) | **Strong.** Kosslyn et al. 2012 (peer reviewed) puts text above 20 pt; the UK ONS, Cambridge and Microsoft all name 18 pt as the minimum. `pitch` sits 2 pt lower because its decks are often sent as well as shown; that step is a judgement, not a finding. |
| Body target | 24 pt (`live`), 20 pt (`pitch`) | **Moderate.** A compromise between cited floors (Oslo asks for 24 pt; projection guidance recommends about 24 pt on a 7.5 in slide), not a measured optimum. |
| Bullets per list, lines per bullet | 4 and 2 (presented) | **Strong.** Kosslyn et al. 2012. |
| Words per slide | 50 (`live`) | **Moderate.** ONS guidance. 60 and 120 for `pitch` and `read-ahead` scale from it and are conventions. |
| Contrast | 4.5:1 text, 3:1 large text and chart marks | **Strong.** WCAG 2.2, 1.4.3 and 1.4.11. Large is 18 pt, or 14 pt bold. |
| Direct labels over legends | line charts labelled at the ends | **Moderate.** One 2021 experiment (Ajani et al.) found direct labels likely to help; other declutter advice was contested. |
| One accent colour, grey context | `highlight` | **Weak.** Practitioner convention (Knaflic and others), widely repeated, not tested. |
| Title at most 15 words, 2 lines | all profiles | **Weak.** Consulting vendors and blogs. A warning, not an error. |
| Read-ahead sizes | 14 pt target, 10 pt floor, 8 pt source lines | **Weak.** Vendor and blog conventions only. The user may want them larger; never smaller. |
| Safe margin | 5% of each edge | **Weak.** The researchers' synthesis of ONS (2.5%), brand guides (6%) and Office (6.9%). A warning only. |
| Strict contrast for projectors | 7:1 | **Weak.** A recommendation from the observation that projectors wash out colour. Not enforced; mention it for a deck shown in a bright room. |
| Assertion-evidence titles | sentence titles with visual evidence | **Moderate.** One 2013 study (110 students) found better comprehension and recall; from the method's own group, no independent replication. Consulting practice (Minto's titles test) converges on the same rule. |

Popular rules with no evidence behind them, which the skill does not use:
Kawasaki's 10/20/30 and the 6x6 rule.

## 2. Deck structure by profile

**live**: a talk. One idea per slide, the detail in the speaker notes, which
the gate expects on every content slide. Open with the answer, not the agenda:
the cover title or the first content slide states the conclusion. Use a
`statement` slide for the one idea the room must leave with. Ten to twenty
slides for a half-hour talk is typical; fewer slides with more talk beats more
slides with more text.

**pitch**: short and persuasive, often sent ahead and then presented. Large
type, one idea per slide, graphs captioned with their conclusion. The order
that persuades is problem, why now, the answer, evidence it works, the ask.
Ten to fifteen slides. A pitch title still has to be a claim: "Teams lose four
hours a week to manual handoffs", not "The problem".

**read-ahead**: read alone on a screen, so slides carry the detail a presenter
would have said. Still one message per slide, with the message in the title.
Every data slide has a source line. Put backup material after a `section`
slide titled "Appendix"; the story check then expects the conclusion right
before it. An executive summary as the second slide is expected here: a
`bullets` slide whose title is the overall answer and whose bullets are the
next slides' titles, condensed.

Every profile:

- End the main deck on a `conclusion` slide with the takeaways and the ask,
  and leave it up during questions. No "Thank you" or "Questions?" slide.
  Because it stays up, it has to carry the decision on its own: the ask names
  what is decided, what it costs and by when ("Approve the 5-store pilot
  (A$18,000) by 15 Dec"), each takeaway keeps its number, and the slide takes
  a `source` for them like any other.
- Use a `section` slide only in decks long enough to need signposts (about
  fifteen slides or more); in a short deck it is a slide that says nothing.
- An agenda slide is rarely worth it in a talk under twenty minutes. If the
  user wants one, its title still makes a point ("Three changes will get us to
  break-even by June").

## 3. Titles are claims

The title is the slide's message; the body is its evidence. Read in order,
the titles alone should tell the whole argument: this is the ghost deck
`check_story.py` prints, and the test an executive applies when they skim.

A good title:

- is a full sentence with a verb, stating what the slide proves;
- is specific: it names the number, the direction, the comparison;
- fits in 15 words and two lines;
- can be disagreed with. "Churn is important" cannot; "Churn costs more than
  acquisition this year" can.

| Topic title (the lint warns `title-topic`) | Claim title |
|---|---|
| Market overview | Mid-size logistics firms are the only segment still growing |
| Results | Annual plans cut churn from 4.1% to 2.9% in one quarter |
| Next steps | Approve the April pilot so results are in before budget season |
| Our team | The founders built the routing engine two of the top three carriers use |

The checks behind this:

- `title-topic`: the title is a known topic label ("Overview", "Results") or
  three words or fewer without a figure.
- `title-long`, `title-lines`, `title-widow`: too many words, more than two
  lines, or a last line of one word, measured at the title's real size. Fix a
  widow by cutting words, not adding them. `ask-widow` is the same warning for
  the conclusion's bold ask.
- `count-mismatch`: a title that promises "three reasons" must have three
  bullets, steps or rows.
- `duplicate-title` and `title-near-duplicate`: two slides making the same
  point. Merge them or say what the second adds.
- Cover and section titles are exempt from the claim rule; a cover title that
  states the conclusion is still better than one that names the subject.

## 4. Numbers you can stand behind

Invented figures are the fastest way to lose a room, and models invent them
easily: a round "40% faster" fills a gap that the material did not.

- Every number comes from the user's material or a source you can name.
  Never estimate, extrapolate, round to a nicer figure, or carry a number over
  from your own knowledge without saying where it is from and checking it.
- Charts, tables and big numbers need a `source`, and so does any slide whose
  text contains a quantity (money, a percentage, a multiplier, a decimal or a
  figure in thousands). The gate warns `number-unsourced`. It does not catch
  plain counts and durations ("5 stores for 12 weeks", "38 stores"), so give
  those a source too when they come from the material and not from the
  proposal itself.
- Write the source as a reader would check it: "Source: Stripe export, Jul to
  Sep 2026", not "Source: internal data".
- When the material has no number for a claim, the claim goes or becomes
  qualitative. Tell the user what you left out and why.
- A chart's title must match what the chart shows. If the title says churn
  fell by a third, the chart shows that fall, with the falling bar
  highlighted.
- At most two `big-number` slides per deck, never next to each other. A wall
  of big-number tiles is the pattern that most often carries made-up figures.

## 5. Choosing a layout

Pick the layout from what the slide has to prove, not from what the text looks
like.

| The slide shows | Layout |
|---|---|
| The one idea the audience must remember | `statement` |
| A few parallel points of equal weight | `bullets` (at most 4 in a talk) |
| Two things side by side: before and after, option A and B, us and them | `two-column` |
| A trend, a ranking, a share, a change | `chart` |
| Exact values someone will look up, or several attributes per item | `table` |
| A product, a place, a diagram that exists as a picture | `image` |
| One figure that carries the slide on its own | `big-number` (sparingly) |
| A customer's or expert's own words | `quote` |
| A real sequence: steps that happen in order | `process` (3 to 5 steps) |
| The takeaways and the ask | `conclusion` (last) |

The gate warns when three `bullets` slides come in a row (`layout-run`); when
it does, one of them is usually a comparison (`two-column`), a sequence
(`process`), numbers in disguise (`chart`) or the main idea (`statement`).

## 6. Choosing a chart

| The point | Chart |
|---|---|
| Compare a few items, short names | `column` |
| Compare items with long names, or rank them | `bar` (sorted, largest first) |
| Change over time, many points or several series | `line` |
| How a total breaks into parts, two or three parts | `pie` only if the point is "most of it is one thing"; otherwise `bar` |
| Parts of a total compared across groups | `stacked-bar` or `stacked-column`, at most three parts |
| How a start value became an end value | `waterfall` |
| Whether two measures move together | `scatter` |

- Highlight what the title claims (`highlight`): the accent goes on that
  bar, category or series and everything else goes muted. A chart with no
  highlight asks the audience to find the point themselves.
- Prefer one series. Several series in a talk are hard to read from the back;
  split them across slides or move the extra series to the notes.
- Keep categories short: the build refuses labels that collide, and the fix
  is shorter names or a `bar` chart, not smaller type. The same goes for
  value labels on clustered columns: with two or more series, let the axis
  carry the scale.
- Pick the `number_format` the audience reads naturally: `0%` for shares,
  `$#,##0.0,,"M"` for millions, in the material's currency (`"A$"`, `"£"`), so
  the chart and its title show the same figure the same way.
- Show every point the material has. A month left off a trend to get under a
  category cap is a gap the audience will notice, or worse, will not.
- A target or plan is a muted second line (see `spec.md`), so the gap between
  actual and target is the thing the eye lands on.

## 7. Colour and type

- The presets (`paper`, `consulting`, `ink`) use theme colours and two
  metric-safe font families, and pass contrast. Use one unless the user gives
  brand colours or a template.
- Custom colours come only from the user or their brand assets, never from a
  guess at "their brand". The build refuses a theme whose text colours fail
  4.5:1 (`theme-contrast`).
- Accent is for emphasis: the highlighted mark, the ask. When everything is
  accent, nothing is.
- `ink` is a dark theme. Dark slides look good on a screen and wash out on
  many projectors; prefer a light theme for a room you have not seen.
- Two font families at most (`too-many-fonts`). Fonts outside the metric-safe
  list (`font-not-metric-safe`) change line breaks on machines without them;
  see `compatibility.md`.

## 8. What makes a deck look machine-made

Several independent sources (design vendors, an anti-slop design skill,
consulting trainers) list the same tells. Each source has a stake, but the
agreement makes the list usable.

The lint catches some of them:

- filler words: delve, leverage, unlock, seamless, game-changing,
  cutting-edge, empower, synergy, holistic, transformative, robust and others
  (`slop-word`, a warning: some have honest uses, but each gets a second look);
- em dashes, emoji, Markdown, typed bullets (errors);
- leftover template text and assistant chatter ("Click to add title", "TBD",
  "Here is your deck") (`leftover-placeholder`, an error);
- topic titles, near-duplicate slides, courtesy closers, runs of bullets.

The review catches the rest, so watch for them when writing the spec:

- every slide the same shape: three bullets, three boxes, three icons;
- text that could appear in anyone's deck ("Customers want better
  experiences"): name the customer, the number, the case;
- a chart that does not prove its title, or a title that restates the chart
  without saying what it means;
- symmetry for its own sake: two columns where one side is padding;
- decoration standing in for content: icons, stock images, accent bars.
