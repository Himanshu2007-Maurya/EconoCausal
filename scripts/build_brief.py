"""Build the 5-slide EconoCausal brief.

Focus: predicted remaining work + how the dashboard is used to show results.
Source review: snigdhasudheer/Infotact_project1 @ main (b01a204), 27 Oct 2026.

Run:  python scripts/build_brief.py
Out:  EconoCausal_Project_Brief.pptx
"""

from pptx import Presentation

from deck_style import (
    AMBER, AMBER_L, AMBER_TINT, BODY, BORDER, CARD, CW, FAINT, GREY_TINT,
    HEAD, INK, M, MUTED, PAPER, ROSE, ROSE_TINT, SH, SLATE, SW, TEAL, TEAL_L,
    TEAL_TINT, WHITE, Deck,
)

prs = Presentation()
prs.slide_width = prs.slide_width = prs.slide_width
from pptx.util import Inches  # noqa: E402

prs.slide_width = Inches(SW)
prs.slide_height = Inches(SH)
d = Deck(prs)

NOTES = {}


# ==================================================== 1. TITLE + SNAPSHOT
s = d.slide(INK)
d.rect(s, 0, 0, 0.10, SH, TEAL)
for i in range(7):
    d.rect(s, 8.55 + i * 0.62, 1.24, 0.012, 1.86,
           __import__("pptx").dml.color.RGBColor(0x16, 0x1F, 0x2E))
for i in range(4):
    d.rect(s, 8.55, 1.24 + i * 0.62, 4.10, 0.012,
           __import__("pptx").dml.color.RGBColor(0x16, 0x1F, 0x2E))

d.tb(s, 0.95, 1.24, 6.6, 0.24, "Project brief  ·  main branch",
     size=10.5, bold=True, color=TEAL_L, caps=True, tracking=1.8)
d.tb(s, 0.95, 1.62, 7.0, 0.80, "EconoCausal", size=50, bold=True,
     color=WHITE, font=HEAD)
d.tb(s, 0.95, 2.50, 7.0, 0.66,
     "What is built today, what comes next,\nand how the dashboard shows it",
     size=18, color=__import__("pptx").dml.color.RGBColor(0xC7, 0xD2, 0xE0),
     spacing=1.24)
d.rect(s, 0.95, 3.48, 1.5, 0.03, TEAL)
d.rich(s, 0.95, 3.76, 6.9, 0.80,
       [("Five slides. ", True, WHITE, 12),
        ("Two of them are forward-looking: the predicted work left on main, "
         "and the dashboard flow used to show the result. No causal number is "
         "claimed anywhere in this deck.", False,
         __import__("pptx").dml.color.RGBColor(0xB6, 0xC3, 0xD4), 12)],
       spacing=1.34)

meta = [("Repository", "snigdhasudheer/Infotact_project1"),
        ("Reviewed", "27 Oct 2026  ·  main @ b01a204")]
y = 4.92
for k, v in meta:
    d.tb(s, 0.95, y, 1.15, 0.22, k, size=9, color=FAINT, caps=True,
         tracking=1.0, bold=True)
    d.tb(s, 2.12, y - 0.015, 5.4, 0.24, v, size=10.5, color=WHITE)
    y += 0.34

# brief contents, right side of the hero
d.tb(s, 8.55, 3.62, 4.2, 0.24, "In this brief", size=9, bold=True,
     color=TEAL_L, caps=True, tracking=1.2)
toc = ["01  Today: built vs outstanding",
       "02  Predicted further work",
       "03  Dashboard: how it shows results",
       "04  What the demo can claim"]
yy = 3.94
for item in toc:
    d.tb(s, 8.55, yy, 4.2, 0.24, item, size=10.5,
         color=__import__("pptx").dml.color.RGBColor(0xC7, 0xD2, 0xE0))
    yy += 0.32

d.rect(s, 0.95, 5.86, CW - 0.17, 0.014,
       __import__("pptx").dml.color.RGBColor(0x2A, 0x36, 0x4A))
stats = [("41", "commits"), ("18", "merged PRs"), ("16", "unit tests"),
         ("5", "API routes"), ("2/12", "requirements done")]
x = 0.95
for n, lab in stats:
    d.rich(s, x, 6.08, 2.3, 0.34,
           [(n, True, TEAL_L, 16), ("  " + lab, False, FAINT, 9.5)])
    x += 2.30
d.tb(s, 0.95, 6.58, 9.0, 0.24,
     "Evidence read from the repository, not from its roadmap.",
     size=8.5, color=FAINT, italic=True)

NOTES[1] = (
    "Five slides, two of them forward-looking. Everything factual here was "
    "counted from the main branch - 41 commits, 18 merged PRs, 16 tests, five "
    "API routes, and 2 of 12 requirements genuinely complete. The second half "
    "of the brief is a prediction, and it is labelled as one.")

# ==================================================== 2. TODAY
s = d.slide()
y = d.header(s, "01  ·  Today", "Built vs outstanding, on one screen",
             "Two of twelve requirements are complete, four are partial. "
             "Nothing downstream of the export estimates an effect.")

kpis = [("41", "commits on main", "GitHub history"),
        ("18", "merged PRs · 0 open", "all merged into main"),
        ("16", "unit tests · no CI", "data utilities only"),
        ("5", "API routes", "upload · export · report")]
cw = (CW - 3 * 0.20) / 4
for i, (n, lab, sub) in enumerate(kpis):
    cx = M + i * (cw + 0.20)
    d.card(s, cx, y, cw, 1.00)
    d.tb(s, cx + 0.28, y + 0.16, cw - 0.5, 0.34, n, size=21, bold=True,
         color=TEAL, font=HEAD)
    d.tb(s, cx + 0.28, y + 0.54, cw - 0.5, 0.22, lab, size=9.6, bold=True,
         color=INK)
    d.tb(s, cx + 0.28, y + 0.76, cw - 0.5, 0.22, sub, size=8.2, color=FAINT)

y2 = y + 1.18
cw2 = (CW - 0.26) / 2

# delivered
d.card(s, M, y2, cw2, 3.52)
d.pill(s, M + cw2 - 1.34, y2 + 0.24, "ON MAIN", WHITE, TEAL, w=1.02)
d.label(s, M + 0.32, y2 + 0.24, 3.2, "Delivered", color=TEAL)
done = [("Dataset decision", "Retail vs Hillstrom compared and documented"),
        ("Cleaning & validation", "8 required columns, 5 tests, 400 on bad input"),
        ("Product-month export", "6-column CSV, 2-dp rounding, 7 tests"),
        ("FastAPI + React shell", "upload → validate → export → report, wired"),
        ("Methodology notes", "8 identification assumptions, limits stated")]
yy = y2 + 0.68
for name, desc in done:
    d.dot(s, M + 0.32, yy + 0.08, 0.11, TEAL)
    d.tb(s, M + 0.58, yy, 2.5, 0.24, name, size=10.4, bold=True, color=INK)
    d.tb(s, M + 0.58, yy + 0.23, cw2 - 0.95, 0.24, desc, size=8.9,
         color=MUTED)
    yy += 0.56

# outstanding
ox = M + cw2 + 0.26
d.card(s, ox, y2, cw2, 3.52)
d.pill(s, ox + cw2 - 1.46, y2 + 0.24, "OUTSTANDING", ROSE, ROSE_TINT,
       w=1.14, size=8)
d.label(s, ox + 0.32, y2 + 0.24, 3.2, "Still open", color=ROSE)
open_items = [("Treatment, outcome, covariates", "blocks everything below"),
              ("Causal DAG", "no graph, no defensible estimator"),
              ("DML estimate with intervals", "EconML not installed"),
              ("DoWhy refutation", "no placebo or sensitivity checks"),
              ("Budget optimisation", "needs a validated estimate"),
              ("Dashboard placeholders", "4 of 6 panels still empty")]
yy = y2 + 0.68
for name, desc in open_items:
    d.dot(s, ox + 0.32, yy + 0.08, 0.11, ROSE)
    d.tb(s, ox + 0.58, yy, 2.7, 0.24, name, size=10.4, bold=True, color=INK)
    d.tb(s, ox + 0.58, yy + 0.23, cw2 - 0.95, 0.24, desc, size=8.9,
         color=MUTED)
    yy += 0.48

d.rich(s, M, y2 + 3.74, CW, 0.26,
       [("Read it this way:  ", True, INK, 10),
        ("the data and delivery layer is credible and tested; the four "
         "requirements that give the project its name are untouched.", False,
         MUTED, 10)])
d.footer(s)

NOTES[2] = (
    "Keep this to one breath. Left column is what a reviewer can click "
    "today; right column is what is missing. The important nuance is at the "
    "bottom: the first item on the right - the treatment, outcome and "
    "covariate definitions - belongs to Section A, which is why the causal "
    "track has not started.")

# ==================================================== 3. PREDICTED WORK
s = d.slide()
y = d.header(s, "02  ·  Predicted", "What comes next, week by week",
             "Four predicted weeks built from the effort each task actually "
             "needs. Confidence reflects how much is outside Section B's "
             "control.")

weeks = [
    ("Week 1", "Unblock", TEAL, [
        ("C1", "Section A definitions signed off", "high"),
        ("P3", "Duplicate id + port mismatch fixed", "high"),
        ("P5–P6", "Untrack .pyc, add CI workflow", "high")],
     "Predicted: DAG committed, CSV loads in the dashboard, CI green"),
    ("Week 2", "Estimate", AMBER, [
        ("C2", "Analysis table build script + graph", "medium"),
        ("C3", "LinearDML, cross-fit, cluster-robust CI", "medium"),
        ("P2a", "Causal Analysis card fed by real output", "high")],
     "Predicted: first ATE with a confidence interval on screen"),
    ("Week 3", "Validate", AMBER, [
        ("C4", "Placebo, random cause, subset refutations", "medium"),
        ("P2b", "Overlap + heterogeneity charts", "high"),
        ("P1", "Product-month CSV into EDA section", "high")],
     "Predicted: estimate survives refutation, or gets re-scoped"),
    ("Week 4", "Decide", ROSE, [
        ("C5", "SciPy allocation, labelled simulated", "low"),
        ("P4", "Fetch script or fixtures for reproducibility", "high"),
        ("—", "Final report + demo run-through", "high")],
     "Predicted: one quantified recommendation plus its caveat"),
]
cw = (CW - 3 * 0.18) / 4
CONF = {"high": (TEAL, TEAL_TINT), "medium": (AMBER, AMBER_TINT),
        "low": (ROSE, ROSE_TINT)}
for i, (wk, name, col, items, pred) in enumerate(weeks):
    cx = M + i * (cw + 0.18)
    d.card(s, cx, y, cw, 3.14)
    d.rect(s, cx, y, cw, 0.05, col)
    d.tb(s, cx + 0.24, y + 0.24, 1.6, 0.22, wk, size=9, bold=True, color=col,
         caps=True, tracking=1.1)
    d.tb(s, cx + 0.24, y + 0.48, cw - 0.46, 0.30, name, size=15, bold=True,
         color=INK, font=HEAD)
    yy = y + 0.88
    for code, what, conf in items:
        fg, _bg = CONF[conf]
        d.tb(s, cx + 0.24, yy, 0.62, 0.22, code, size=8.4, bold=True,
             color=fg, caps=True, tracking=0.6)
        d.tb(s, cx + 0.24, yy + 0.21, cw - 0.46, 0.48, what, size=8.9,
             color=SLATE, spacing=1.26)
        yy += 0.60
    d.rect(s, cx + 0.24, y + 2.74, cw - 0.48, 0.010, BORDER)
    d.tb(s, cx + 0.24, y + 2.82, cw - 0.46, 0.26, pred, size=8.2, bold=True,
         color=col, spacing=1.2)

y3 = y + 3.34
d.card(s, M, y3, CW, 0.90, fill=GREY_TINT, line=None)
d.rich(s, M + 0.32, y3 + 0.20, CW - 0.64, 0.54,
       [("The prediction that matters:  ", True, INK, 10.3),
        ("by week 4 the dashboard's empty ATE card holds a number with a "
         "confidence interval, or the project honestly ships a Hillstrom "
         "e-mail-uplift result instead. Both are defensible; a filled card "
         "with no identification argument is not.", False, SLATE, 10.3)],
       spacing=1.32)
d.rich(s, M, y3 + 1.08, 5.4, 0.24,
       [("Confidence:  ", True, FAINT, 8.6),
        ("high  ", True, TEAL, 8.6), ("·  ", False, FAINT, 8.6),
        ("medium  ", True, AMBER, 8.6), ("·  ", False, FAINT, 8.6),
        ("low", True, ROSE, 8.6)])
d.tb(s, M + 5.4, y3 + 1.08, CW - 5.4, 0.24,
     "Effort: C1 2–3 days · C2 1 wk · C3 1 wk · C4 3–4 days · C5 4–5 days",
     size=8.6, color=FAINT, italic=True, align=__import__("pptx").enum.text.PP_ALIGN.RIGHT)
d.footer(s)

NOTES[3] = (
    "This is the forward-looking slide. Four weeks, not four months: the work "
    "left is bounded, and the only real risk sits in week 1. Point at the "
    "confidence pills - everything Section B controls is high, everything "
    "depending on Section A or on whether price variation exists is medium or "
    "low. The bottom strip is the commitment: a number with an interval, or "
    "an honest re-scope.")

# ==================================================== 4. DASHBOARD
s = d.slide()
y = d.header(s, "03  ·  Dashboard", "How the dashboard is used to show results",
             "Six panels, one story: upload → describe → estimate → recommend "
             "→ allocate. Each panel names its own data source.")

cols = [2.05, 3.55, 2.45, 3.72]
heads = ["Panel", "What it shows", "Where the data comes from",
         "State now  →  after the predicted work"]
x = M
d.rect(s, M, y, CW, 0.40, INK, None, radius=0.12)
for i, h in enumerate(heads):
    d.tb(s, x + 0.12, y + 0.09, cols[i] - 0.2, 0.24, h, size=8.6, bold=True,
         color=WHITE, caps=True, tracking=1.0)
    x += cols[i]

rows = [
    ("Dataset & EDA", "Validation summary, 5-row preview, quality tiles",
     "/api/upload → pandas", "live, one file at a time  →  adds missing-value "
     "and month tiles", TEAL),
    ("Product-month charts", "Price vs quantity by month, per product",
     "product_month_summary.csv", "component exists, no data  →  real Plotly "
     "trend per product", AMBER),
    ("Causal Analysis", "ATE, 95% CI, sample size, overlap diagnostics",
     "DML output (JSON)", "four “Pending” cards  →  estimate + interval + "
     "diagnostics", ROSE),
    ("Recommendations", "Products and segments ranked by estimated effect",
     "ATE + heterogeneity", "placeholder  →  ranked table with effect sizes",
     ROSE),
    ("Budget Optimization", "Allocation vs baseline, explicitly simulated",
     "SciPy scenario", "placeholder  →  labelled simulation, never a promise",
     ROSE),
]
ry = y + 0.48
for i, (panel, shows, src, state, col) in enumerate(rows):
    h = 0.52
    if i % 2 == 0:
        d.rect(s, M, ry - 0.04, CW, h, WHITE, BORDER, radius=0.05)
    else:
        d.rect(s, M, ry - 0.04, CW, h, GREY_TINT, BORDER, radius=0.05)
    d.rect(s, M + 0.02, ry - 0.04, 0.05, h, col)
    d.tb(s, M + 0.24, ry + 0.08, cols[0] - 0.3, 0.24, panel, size=10.2,
         bold=True, color=INK)
    d.tb(s, M + cols[0], ry + 0.06, cols[1] - 0.2, 0.50, shows, size=8.7,
         color=SLATE, spacing=1.20)
    d.tb(s, M + cols[0] + cols[1], ry + 0.08, cols[2] - 0.2, 0.44, src,
         size=8.4, color=MUTED, font="Consolas", spacing=1.20)
    d.tb(s, M + cols[0] + cols[1] + cols[2], ry + 0.06, cols[3] - 0.2, 0.50,
         state, size=8.7, color=col, spacing=1.20)
    ry += h + 0.05

# demo flow
y4 = ry + 0.14
d.label(s, M, y4, 4.0, "The five-minute demo flow")
flow = [("1", "Upload sample_upload.xlsx", "validation + summary appears"),
        ("2", "Export the product-month CSV", "the hand-off table downloads"),
        ("3", "Open Dataset & EDA", "one real trend chart, real numbers"),
        ("4", "Open Causal Analysis", "ATE, interval, overlap — or say why not"),
        ("5", "Open Budget Optimization", "scenario, labelled simulated")]
fx = M
fw = (CW - 4 * 0.14) / 5
for i, (n, title, sub) in enumerate(flow):
    cx = fx + i * (fw + 0.14)
    d.card(s, cx, y4 + 0.26, fw, 0.74, fill=TEAL_TINT, line=None)
    d.tb(s, cx + 0.18, y4 + 0.33, 0.30, 0.22, n, size=11, bold=True,
         color=TEAL, font=HEAD)
    d.tb(s, cx + 0.46, y4 + 0.32, fw - 0.60, 0.42, title, size=8.8, bold=True,
         color=INK, spacing=1.2)
    d.tb(s, cx + 0.18, y4 + 0.70, fw - 0.32, 0.24, sub, size=7.8, color=MUTED)

d.rich(s, M, y4 + 1.08, CW, 0.28,
       [("Never show a placeholder:  ", True, INK, 10),
        ("hide an unfinished panel or label it “in progress”. A reviewer reads "
         "“Pending” as unfinished scaffolding, which costs more than an empty "
         "tab.", False, MUTED, 10)])
d.footer(s)

NOTES[4] = (
    "This is the showing strategy. Each row says what a panel will display "
    "and which file or endpoint feeds it - that is the credibility test. The "
    "flow at the bottom is the demo order: upload, export, describe, "
    "estimate, allocate. Five minutes, no placeholder cards on screen. If "
    "step 4 is not ready, say so out loud rather than showing a blank card.")

# ==================================================== 5. CLAIMS
s = d.slide(INK)
d.rect(s, 0, 0, 0.10, SH, TEAL)
d.tb(s, M + 0.05, 0.52, CW, 0.24, "04  ·  What the demo can claim",
     size=10.5, bold=True, color=TEAL_L, caps=True, tracking=1.6)
d.tb(s, M + 0.05, 0.82, 10.0, 0.56, "A defensible pitch, and its limits",
     size=26, bold=True, color=WHITE, font=HEAD)
d.rect(s, M + 0.05, 1.56, 6.0, 0.03,
       __import__("pptx").dml.color.RGBColor(0x2A, 0x36, 0x4A))

cwc = (CW - 0.40) / 2
d.card(s, M + 0.05, 1.86, cwc, 2.20,
       fill=__import__("pptx").dml.color.RGBColor(0x1A, 0x25, 0x3A), line=None)
d.label(s, M + 0.35, 2.08, 4.0, "Can claim", color=TEAL_L)
can = ["A validated, tested pipeline on real retail data",
       "Descriptive price–demand evidence, clearly labelled as association",
       "A DML estimate with a confidence interval and refutation checks, "
       "once C3–C4 land",
       "A simulated allocation showing how budget would move"]
yy = 2.40
for c in can:
    d.dot(s, M + 0.35, yy + 0.07, 0.09, TEAL_L)
    d.tb(s, M + 0.58, yy, cwc - 0.85, 0.46, c, size=9.8,
         color=__import__("pptx").dml.color.RGBColor(0xC7, 0xD2, 0xE0),
         spacing=1.24)
    yy += 0.44

d.card(s, M + cwc + 0.35, 1.86, cwc, 2.20,
       fill=__import__("pptx").dml.color.RGBColor(0x1A, 0x25, 0x3A), line=None)
d.label(s, M + cwc + 0.65, 2.08, 4.0, "Cannot claim", color=AMBER_L)
cant = ["A causal discount effect from historical retail prices",
        "ROI or savings produced by the optimiser",
        "That an e-mail uplift equals a price elasticity"]
yy = 2.40
for c in cant:
    d.dot(s, M + cwc + 0.65, yy + 0.07, 0.09, AMBER_L)
    d.tb(s, M + cwc + 0.88, yy, cwc - 0.85, 0.46, c, size=9.8,
         color=__import__("pptx").dml.color.RGBColor(0xC7, 0xD2, 0xE0),
         spacing=1.24)
    yy += 0.44

# next steps + decision
y5 = 4.42
d.tb(s, M + 0.05, y5, 6.4, 0.22, "Next seven days", size=9, bold=True,
     color=TEAL_L, caps=True, tracking=1.2)
nxt = [("1", "Get Section A definitions in writing", "project lead"),
       ("2", "Ship the unblocked fixes: P1, P3, P5, P6", "Section B"),
       ("3", "Build the analysis table and commit the DAG", "Section B")]
yy = y5 + 0.30
for n, txt, who in nxt:
    d.tb(s, M + 0.05, yy, 0.30, 0.24, n, size=10, bold=True, color=TEAL_L)
    d.tb(s, M + 0.42, yy, 4.5, 0.24, txt, size=10,
         color=__import__("pptx").dml.color.RGBColor(0xC7, 0xD2, 0xE0))
    d.tb(s, M + 5.05, yy, 1.55, 0.24, who, size=8, color=FAINT,
         align=__import__("pptx").enum.text.PP_ALIGN.RIGHT, caps=True,
         tracking=0.6)
    yy += 0.36

d.card(s, 7.55, y5 - 0.16, 4.78, 1.86,
       fill=__import__("pptx").dml.color.RGBColor(0x1A, 0x25, 0x3A), line=None)
d.label(s, 7.85, y5 + 0.02, 4.0, "One decision, today", color=AMBER_L)
d.tb(s, 7.85, y5 + 0.32, 4.24, 0.86,
     "Do we estimate the e-mail uplift Hillstrom can support, or hold for a "
     "pricing design we do not have yet?", size=10.5, bold=True, color=WHITE,
     spacing=1.28)
d.tb(s, 7.85, y5 + 1.22, 4.24, 0.24,
     "It decides whether week 4 ships a number or a caveat.", size=8.6,
     color=FAINT)

d.tb(s, M + 0.05, 6.72, CW, 0.26,
     "Sources: snigdhasudheer/Infotact_project1 @ main · README, docs/day2, "
     "docs/day5, requirements-checklist, main.py, frontend/src, test suite",
     size=7.6, color=__import__("pptx").dml.color.RGBColor(0x64, 0x74, 0x8B))
d.footer(s, dark=True)

NOTES[5] = (
    "Close on the boundary, not a summary. Left: what this project can stand "
    "behind today and after week 4. Right: the three claims that would break "
    "under questioning - the discount effect, the optimiser's savings, and "
    "relabelling an e-mail uplift as price elasticity. Then ask for the one "
    "decision that determines whether week 4 produces a number or a caveat.")

for i, sl in enumerate(prs.slides, start=1):
    if i in NOTES:
        sl.notes_slide.notes_text_frame.text = NOTES[i]

prs.save("EconoCausal_Project_Brief.pptx")
print("saved:", len(prs.slides._sldIdLst), "slides")
