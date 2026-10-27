# EconoCausal

Project status review deck for
[snigdhasudheer/Infotact_project1](https://github.com/snigdhasudheer/Infotact_project1)
(main branch, reviewed 27 Oct 2026).

## Deliverable

`EconoCausal_Project_Status_Review.pptx` — 15 slides covering:

1. Where the project stands today (commits, PRs, re-scored requirements)
2. What has already been delivered on main (data, API, dashboard, tests)
3. What work is still outstanding (causal core + platform fixes)
4. How to close the gap (phased plan, risks, decisions)

## Rebuilding the deck

```bash
pip install python-pptx
python scripts/build_deck.py          # writes the .pptx
python scripts/render_preview.py      # optional PNG previews + overflow check
```

`scripts/build_deck.py` is the single source of truth for the content and
layout; edit it rather than the generated file.

## Scope note

All figures were read from the repository (README, `docs/`, `main.py`,
`frontend/src`, test suite, git history). No causal effect, model score or
budget saving is claimed anywhere in the deck.
