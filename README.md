# EconoCausal

Project brief for
[snigdhasudheer/Infotact_project1](https://github.com/snigdhasudheer/Infotact_project1)
(main branch, reviewed 27 Oct 2026, commit `b01a204`).

## Deliverable

**`EconoCausal_Project_Brief.pptx`** — 5 slides, the primary deck:

| # | Slide | Purpose |
|---|-------|---------|
| 1 | Title + snapshot | Source commit, headline counts, contents |
| 2 | Today: built vs outstanding | Delivered list beside open list, with KPIs |
| 3 | Predicted further work | Four predicted weeks, confidence per task |
| 4 | Dashboard: how it shows results | Six panels, their data source, state now → after |
| 5 | What the demo can claim | Claim / cannot claim, next 7 days, one decision |

`EconoCausal_Project_Status_Review.pptx` (15 slides) is the longer review deck
from the earlier pass — kept for detail, not needed to present.

## Rebuilding

```bash
pip install python-pptx
python scripts/build_brief.py     # writes the 5-slide brief
python scripts/build_deck.py      # writes the 15-slide review deck
python scripts/render_preview.py EconoCausal_Project_Brief.pptx preview_brief
```

`scripts/deck_style.py` holds the shared palette and layout helpers; the two
builder scripts are the single source of truth for content.

## Scope note

Facts were read from the repository (README, `docs/`, `main.py`,
`frontend/src`, test suite, git history). Slides 3–4 are labelled predictions.
No causal effect, model score or budget saving is claimed anywhere.
