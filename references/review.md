# Independent review

The gates measure what can be measured. They cannot tell whether a chart shows
what its title claims, whether a number matches its source, or whether the
deck reads as if anyone could have written it. Neither can you, reliably,
about a deck you just wrote: you see what you meant, not what is on the slide.
So the review is done by fresh agents that did not write the deck.

Start it only once `gate_result.json` says `"passed": true`. Reviewing a deck
that still fails the gate wastes the reviewers on problems a script already
found.

## The order

1. **Blind reader**, alone. It sees only the slide pictures: no brief, no
   spec, no notes, no idea what the deck is meant to say. It writes down what
   it thinks the argument and the ask are. You then compare that with the
   brief. Where they differ, the slides do not say what you think they say.
   The blind reader goes first and alone because once it has seen the brief
   it can no longer be blind.
2. **Three critics** in parallel, each with one lens (below), each told
   nothing about what the others look for. Three agents is the ceiling at any
   time: wider fan-outs have hit the account's usage limit and returned
   nothing.

Spawn each with the Agent tool (`general-purpose`), giving it the agent file
from `agents/` to read first and the paths it needs. Do not tell a reviewer
what you think is wrong or what you already fixed: a reviewer handed a
hypothesis tends to confirm it.

| Agent | Reads first | Gets | Never gets |
|---|---|---|---|
| Blind reader | `agents/blind-reader.md` | `render/slide-NN.png` | brief, spec, notes, material |
| Critic: argument and evidence | `agents/critic.md`, lens `argument` | slide PNGs, `deck.json`, `brief.md`, the user's material and sources | your opinion |
| Critic: visual design | `agents/critic.md`, lens `visual` | slide PNGs, `render/contact.png`, `references/design-rules.md` section 8, `gate_result.json` | brief, material |
| Critic: audience and profile | `agents/critic.md`, lens `audience` | slide PNGs, `brief.md`, `deck.json` (for the notes), the profile's row in design-rules section 2 | material |

Point the agents at absolute paths and ask each to write its findings to a
file in the deck folder (`review/blind-reader.json`,
`review/critic-argument.json` and so on) as well as return them, so the
findings survive a lost session.

## Findings

Every reviewer returns findings in one shape, so you can triage them
together:

```json
{
  "lens": "argument",
  "findings": [
    {
      "slide": 4,
      "field": "chart",
      "severity": "must-fix",
      "finding": "The title says churn fell by a third; the chart shows 4.1% to 2.9%, a fall of 29%.",
      "fix": "Change the title to 'Churn fell from 4.1% to 2.9%' or check the source figure."
    }
  ]
}
```

- **must-fix**: wrong, misleading or unreadable. A number that does not
  match its source, a chart that does not show its title's claim, a slide the
  blind reader read the opposite way, text a viewer cannot read in the
  picture.
- **should-fix**: weakens the deck. A vague title, a generic slide, a chart
  type that hides the point, a missing step in the argument.
- **nit**: taste. Fix if cheap.

`fix` is in spec terms (a field, a layout, a chart type, a title), never a
coordinate or a font size, because the spec is the only thing you can change.

## Triage

- **The gate wins on anything it measures.** If a critic says text is too
  small but the lint passed it above the profile's floor, that is a nit about
  the profile, not a must-fix. If a critic says text overflows a box and the
  lint measured it as fitting, look at the picture yourself: if it visibly
  overflows, the gate has a bug. Fix the slide, and tell the user the gate
  missed it, with the slide and shape.
- **Must-fix**: fix every one, through the fix ladder in `SKILL.md`.
- **Should-fix**: fix unless it conflicts with what the user asked for; if it
  does, keep the user's choice and mention the finding at delivery.
- **Nit**: fix the cheap ones.
- **Numbers**: a finding that a number is wrong or unsourced is settled
  against the user's material, not against the critic's memory. If the
  material does not settle it, ask the user.
- **Titles**: a fix that changes a title changes the approved story. Collect
  title changes and show the user the new ghost deck once, not one at a time.

## Rounds

After fixing, rebuild, re-render and re-lint until the gate passes again.
Then re-run only what the fixes touch:

- the blind reader again if any title changed, or if it misread the deck
  last time (a new agent: the old one has seen the deck);
- each critic whose lens had a must-fix, with a note of which slides changed
  so it checks them and nothing else is re-litigated.

Stop after three rounds even if findings remain. More rounds trade the user's
time for diminishing changes, and reviewers start contradicting each other.
Report what is left at delivery: the finding, its severity and why you left
it.
