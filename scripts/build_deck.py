"""
Build the EconoCausal project status deck.

Source of truth: github.com/snigdhasudheer/Infotact_project1 @ main (b01a204),
reviewed 27 Oct 2026.

Run:  python scripts/build_deck.py
Out:  EconoCausal_Project_Status_Review.pptx
"""

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

# ---------------------------------------------------------------- palette
INK = RGBColor(0x0B, 0x12, 0x20)
INK_2 = RGBColor(0x1E, 0x29, 0x3B)
SLATE = RGBColor(0x33, 0x41, 0x55)
MUTED = RGBColor(0x64, 0x74, 0x8B)
FAINT = RGBColor(0x94, 0xA3, 0xB8)
TEAL = RGBColor(0x0D, 0x94, 0x88)
TEAL_L = RGBColor(0x5E, 0xEA, 0xD4)
AMBER = RGBColor(0xB4, 0x53, 0x09)
AMBER_L = RGBColor(0xFB, 0xBF, 0x24)
ROSE = RGBColor(0xBE, 0x12, 0x3C)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
PAPER = RGBColor(0xF7, 0xF8, 0xFA)
CARD = RGBColor(0xFF, 0xFF, 0xFF)
BORDER = RGBColor(0xE2, 0xE8, 0xF0)
TEAL_TINT = RGBColor(0xEC, 0xFD, 0xF5)
AMBER_TINT = RGBColor(0xFF, 0xFB, 0xEB)
ROSE_TINT = RGBColor(0xFF, 0xF1, 0xF2)
GREY_TINT = RGBColor(0xF1, 0xF5, 0xF9)

HEAD = "Segoe UI"
BODY = "Segoe UI"

SW, SH = 13.333, 7.5
M = 0.78  # page margin
CW = SW - 2 * M  # content width

NOTES = {1: "Open by naming the source: the main branch of snigdhasudheer/Infotact_project1, reviewed today (commit b01a204). Say up front that every number in this deck was counted from the repository, because the repository's own roadmap and requirements checklist are out of date.", 2: 'Four questions. Set the expectation that the last two are the useful ones: what is missing, and what closes it.', 3: 'One-line framing: this is a causal-inference study delivered as documentation plus a dashboard. Note the two workstreams - Section A owns the definitions, Section B owns the methodology and delivery - because that split explains why the causal track is idle.', 4: "Hard evidence only. 41 commits, 18 merged PRs, none open, 16 unit tests, 5 API routes. Then the re-score: 2 of 12 requirements complete and 4 partial, which is 33 percent. The committed checklist still says 0/12 and 'backend pending' - main has overtaken it.", 5: "Six milestones, each traceable to a PR range. Keep this brief; it is the 'credit where due' slide.", 6: 'Online Retail: 541,909 rows down to 524,878 after cleaning, with the four quality counts on screen. Hillstrom: 64,000 rows, zero missing, but 6,562 duplicate profiles flagged and not investigated. Neither dataset can be joined to the other, so the merger was rejected deliberately.', 7: 'Walk the four routes left to right, then flag that the frontend is a real pipeline: upload, validate, export, report. The Plotly EDA component exists but is rendered without data, and Recommendations and Budget are still placeholders.', 8: '16 tests is real coverage of the data utilities - column checks, empty files, bad months, negative values, duplicates, exit codes. The gaps are structural: no API tests, no frontend tests, no CI, and requirements.txt has none of the causal libraries.', 9: "The pivotal slide. Three reasons the headline result is missing: no randomised price in Online Retail, Hillstrom randomises e-mail not price, and Section A has not signed off treatment, outcome, or covariates. The dashboard showing 'Not yet estimated' is correct behaviour, not a defect.", 10: 'The re-scored checklist, requirement by requirement. Land the bottom line: the four requirements that give the project its name - DAG, DML, refutation, optimisation - are all untouched.', 11: 'Causal track: five tasks in strict order, with effort and the gate for each. C1 is the blocker and it belongs to Section A. C3 and C4 are where this becomes a causal study.', 12: 'Platform track: six fixes, none blocked, start today. P1 and P3 are the two that most affect what a reviewer sees, because wiring the CSV and removing the duplicate section id are small but visible.', 13: 'Four phases with gates. The point of the sequence is that nobody builds an optimiser on top of an unvalidated estimate. If Section A slips, run Phase 0 and Phase 2 in parallel and keep the causal slides empty.', 14: 'Four risks, each with a mitigation someone owns. The second is the technical one: without valid price variation, any DML elasticity is a confounded artefact.', 15: 'Close on the decision, not a summary. Ask the reviewer: Hillstrom e-mail uplift that the data can actually identify, or wait for a pricing design we do not have? The answer decides whether the next milestone ships a number or a caveat.'}

prs = Presentation()
prs.slide_width = Inches(SW)
prs.slide_height = Inches(SH)
BLANK = prs.slide_layouts[6]

_page = {"n": 0}


# ---------------------------------------------------------------- helpers
def slide(bg=PAPER):
    s = prs.slides.add_slide(BLANK)
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = bg
    _page["n"] += 1
    return s


def tb(s, l, t, w, h, text, size=13, bold=False, color=SLATE, font=BODY,
       align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=1.0, italic=False,
       caps=False, space_after=0, tracking=None):
    box = s.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    p.line_spacing = spacing
    p.space_after = Pt(space_after)
    r = p.add_run()
    r.text = text.upper() if caps else text
    f = r.font
    f.name = font
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.color.rgb = color
    if tracking is not None:
        r.font._rPr.set("spc", str(int(tracking * 100)))
    return box


def rich(s, l, t, w, h, chunks, size=13, color=SLATE, align=PP_ALIGN.LEFT,
         spacing=1.0, anchor=MSO_ANCHOR.TOP):
    """chunks: list of (text, bold, color_or_None, size_or_None)."""
    box = s.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    p.line_spacing = spacing
    for text, bold, col, sz in chunks:
        r = p.add_run()
        r.text = text
        r.font.name = BODY
        r.font.size = Pt(sz or size)
        r.font.bold = bold
        r.font.color.rgb = col or color
    return box


def card(s, l, t, w, h, fill=CARD, line=BORDER, radius=0.05, shadow=False):
    sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(l), Inches(t),
                            Inches(w), Inches(h))
    sh.adjustments[0] = radius
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(0.9)
    sh.shadow.inherit = shadow
    sh.text_frame.word_wrap = True
    return sh


def rect(s, l, t, w, h, fill, line=None, radius=None):
    shape = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    sh = s.shapes.add_shape(shape, Inches(l), Inches(t), Inches(w), Inches(h))
    if radius:
        sh.adjustments[0] = radius
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(0.9)
    sh.shadow.inherit = False
    return sh


def dot(s, l, t, d, color):
    sh = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(l), Inches(t), Inches(d),
                            Inches(d))
    sh.fill.solid()
    sh.fill.fore_color.rgb = color
    sh.line.fill.background()
    sh.shadow.inherit = False
    return sh


def pill(s, l, t, text, fg, bg, w=None, size=8.5):
    w = w or (0.16 + 0.075 * len(text))
    sh = rect(s, l, t, w, 0.26, bg, None, radius=0.5)
    tf = sh.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = 0
    tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = text
    r.font.name = BODY
    r.font.size = Pt(size)
    r.font.bold = True
    r.font.color.rgb = fg
    return sh


def header(s, kicker, title, sub=None):
    rect(s, M, 0.445, 0.055, 0.155, TEAL)
    tb(s, M + 0.22, 0.44, CW, 0.22, kicker, size=10, bold=True, color=TEAL,
       caps=True, tracking=1.4)
    tb(s, M, 0.70, CW, 0.52, title, size=27, bold=True, color=INK, font=HEAD,
       spacing=0.95)
    y = 1.34
    if sub:
        tb(s, M, y, CW - 1.2, 0.30, sub, size=12, color=MUTED, spacing=1.15)
        y += 0.34
    rect(s, M, y + 0.10, CW, 0.012, BORDER)
    return y + 0.32


def footer(s, dark=False):
    c = RGBColor(0x64, 0x74, 0x8B) if not dark else RGBColor(0x7C, 0x8C, 0xA6)
    tb(s, M, SH - 0.48, 8.0, 0.24,
       "EconoCausal  ·  Project status review  ·  main branch",
       size=8.5, color=c)
    tb(s, SW - M - 1.2, SH - 0.48, 1.2, 0.24, f"{_page['n']:02d}",
       size=8.5, color=c, align=PP_ALIGN.RIGHT, bold=True)


def section_label(s, l, t, w, text, color=INK):
    tb(s, l, t, w, 0.22, text, size=9.5, bold=True, color=color, caps=True,
       tracking=1.1)


# ================================================================ 1. TITLE
s = slide(INK)
rect(s, 0, 0, 0.10, SH, TEAL)
# faint grid motif (kept clear of the right edge)
GRID = RGBColor(0x16, 0x1F, 0x2E)
for i in range(8):
    rect(s, 7.62 + i * 0.60, 1.18, 0.012, 4.62, GRID)
for i in range(8):
    rect(s, 7.62, 1.18 + i * 0.58, 4.32, 0.012, GRID)

tb(s, 0.95, 1.28, 6.4, 0.24, "Project status review  ·  main branch",
   size=10.5, bold=True, color=TEAL_L, caps=True, tracking=1.8)
tb(s, 0.95, 1.70, 6.9, 0.80, "EconoCausal", size=54, bold=True, color=WHITE,
   font=HEAD)
tb(s, 0.95, 2.62, 6.6, 0.72,
   "Causal dynamic pricing via\nDouble Machine Learning",
   size=19, color=RGBColor(0xC7, 0xD2, 0xE0), spacing=1.22)
rect(s, 0.95, 3.72, 1.5, 0.03, TEAL)

rich(s, 0.95, 4.02, 6.6, 0.9,
     [("Where the project actually stands: ", True, WHITE, 12.5),
      ("what is built on main, what is still open, and what it takes to close "
       "the gap — evidence taken from the repository, not from the roadmap.",
       False, RGBColor(0xB6, 0xC3, 0xD4), 12.5)], spacing=1.32)

meta = [("Repository", "snigdhasudheer/Infotact_project1"),
        ("Reviewed", "27 October 2026  ·  main @ b01a204"),
        ("Program", "Infotact Project 1  ·  Section B delivery")]
y = 5.30
for k, v in meta:
    tb(s, 0.95, y, 1.15, 0.22, k, size=9, color=FAINT, caps=True,
       tracking=1.0, bold=True)
    tb(s, 2.12, y - 0.015, 5.2, 0.24, v, size=10.5, color=WHITE)
    y += 0.36

# stat strip
stats = [("41", "commits"), ("18", "merged PRs"), ("0", "open PRs"),
         ("16", "unit tests"), ("5", "API routes")]
x = 0.95
for i, (n, lab) in enumerate(stats):
    rich(s, x, 6.52, 2.2, 0.34, [(n, True, TEAL_L, 17),
                                 ("  " + lab, False, FAINT, 9.5)])
    x += 2.35

# ================================================================ 2. AGENDA
s = slide()
y = header(s, "Contents", "What this deck answers",
           "Four questions, in the order a reviewer would ask them.")
items = [
    ("01", "Where the project stands",
     "Commits, pull requests and a measured read of the 12-point requirements "
     "checklist."),
    ("02", "What has been delivered",
     "Data validation, product-month export, FastAPI endpoints, React "
     "dashboard and the test suite."),
    ("03", "What is still missing",
     "Treatment definition, causal DAG, DML estimates, refutation, budget "
     "optimizer and the final report."),
    ("04", "How to close the gap",
     "A sequenced four-phase plan with gates, owners, risks and the one open "
     "decision."),
]
cw, gap = 5.86, 0.22
for i, (num, title, body) in enumerate(items):
    cx = M + (i % 2) * (cw + gap)
    cy = y + (i // 2) * 2.10
    card(s, cx, cy, cw, 1.92)
    tb(s, cx + 0.38, cy + 0.30, 1.0, 0.40, num, size=24, bold=True,
       color=RGBColor(0xD5, 0xDF, 0xE8), font=HEAD)
    tb(s, cx + 1.15, cy + 0.34, cw - 1.5, 0.32, title, size=15, bold=True,
       color=INK, font=HEAD)
    tb(s, cx + 1.15, cy + 0.80, cw - 1.55, 0.98, body, size=11.5, color=MUTED,
       spacing=1.35)

tb(s, M, 6.62, CW, 0.30,
   "Scope: only the main branch of snigdhasudheer/Infotact_project1, as "
   "reviewed on 27 October 2026. No projected results appear in this deck.",
   size=9.5, color=FAINT, italic=True)
footer(s)

# ================================================================ 3. GLANCE
s = slide()
y = header(s, "Orientation", "The project at a glance",
           "A causal-inference study of pricing and marketing, delivered as a "
           "documented analysis plus a dashboard.")
card(s, M, y, 7.05, 1.52)
section_label(s, M + 0.38, y + 0.26, 6.0, "Objective")
tb(s, M + 0.38, y + 0.56, 6.3, 0.86,
   "Estimate how a pricing or marketing intervention changes customer "
   "behaviour, then spend a limited promotional budget using those estimates "
   "rather than intuition. Online Retail and Hillstrom were both assessed as "
   "candidate datasets.",
   size=11.5, color=SLATE, spacing=1.38)

# two workstreams
y2 = y + 1.72
ws = [("Section A", "Problem definition",
       "Treatment, outcome, covariates, causal DAG, identification assumptions",
       "PENDING", ROSE, ROSE_TINT),
      ("Section B", "Methodology & delivery",
       "Documentation, validation utilities, API, dashboard, reporting",
       "IN PROGRESS", AMBER, AMBER_TINT)]
for i, (tag, name, desc, status, fg, bg) in enumerate(ws):
    cx = M + i * 3.60
    card(s, cx, y2, 3.40, 1.56, fill=bg, line=None)
    tb(s, cx + 0.30, y2 + 0.24, 2.2, 0.22, tag, size=9, bold=True, color=fg,
       caps=True, tracking=1.1)
    tb(s, cx + 0.30, y2 + 0.50, 2.85, 0.28, name, size=13, bold=True,
       color=INK, font=HEAD)
    tb(s, cx + 0.30, y2 + 0.82, 2.85, 0.66, desc, size=10, color=SLATE,
       spacing=1.28)

# stack panel
px = M + 7.40
card(s, px, y, 4.28, 3.58)
tb(s, px + 0.34, y + 0.28, 3.6, 0.24, "Stack", size=10, bold=True, color=INK,
   caps=True, tracking=1.1)

live = ["Python · pandas · openpyxl", "FastAPI · Uvicorn",
        "React 19 · Vite", "Plotly.js · react-plotly"]
planned = ["EconML (Double ML)", "DoWhy (refutation)", "SciPy (optimizer)",
           "scikit-learn (nuisance models)"]
tb(s, px + 0.34, y + 0.58, 3.6, 0.2, "In use on main", size=9, bold=True,
   color=TEAL, caps=True, tracking=0.8)
yy = y + 0.84
for item in live:
    dot(s, px + 0.34, yy + 0.075, 0.10, TEAL)
    tb(s, px + 0.58, yy, 3.4, 0.24, item, size=10.5, color=SLATE)
    yy += 0.28
tb(s, px + 0.34, yy + 0.18, 3.6, 0.2, "Planned, not installed",
   size=9, bold=True, color=ROSE, caps=True, tracking=0.8)
yy += 0.44
for item in planned:
    dot(s, px + 0.34, yy + 0.075, 0.10, RGBColor(0xCB, 0xD5, 0xE1))
    tb(s, px + 0.58, yy, 3.4, 0.24, item, size=10.5, color=MUTED)
    yy += 0.28

# timeline
y3 = y + 3.50
tb(s, M, y3, 3.0, 0.2, "Delivery so far", size=9.5, bold=True, color=INK,
   caps=True, tracking=1.0)
days = [("Day 1", "Dataset overview"), ("Day 2", "Decision + schema"),
        ("Day 3", "Hillstrom QA"), ("Day 4", "Dashboard shell"),
        ("Day 5", "Methodology"), ("Day 6", "Export + validator")]
rect(s, M, y3 + 0.56, CW - 0.4, 0.014, RGBColor(0xD8, 0xE1, 0xEA))
for i, (d, lab) in enumerate(days):
    x = M + 0.06 + i * 2.0
    dot(s, x - 0.055, y3 + 0.505, 0.13, TEAL)
    tb(s, x - 0.05, y3 + 0.72, 1.85, 0.22, d, size=10, bold=True, color=INK)
    tb(s, x - 0.05, y3 + 0.94, 1.85, 0.22, lab, size=9.5, color=MUTED)
footer(s)

# ================================================================ 4. KPI
s = slide()
y = header(s, "Evidence", "Where main stands today",
           "Counted directly from the repository, not from the roadmap in the "
           "codebase.")
kpis = [("41", "commits on main", "18 merged pull requests · 0 open"),
        ("18", "merged PRs", "every PR since Day 1 merged cleanly"),
        ("16", "unit tests", "Python utilities only · no CI workflow"),
        ("5", "API routes", "health, upload, export, report, root")]
cw = (CW - 3 * 0.22) / 4
for i, (big, lab, sub) in enumerate(kpis):
    cx = M + i * (cw + 0.22)
    card(s, cx, y, cw, 1.52)
    tb(s, cx + 0.30, y + 0.24, cw - 0.5, 0.55, big, size=30, bold=True,
       color=TEAL, font=HEAD)
    tb(s, cx + 0.30, y + 0.80, cw - 0.5, 0.24, lab, size=11, bold=True,
       color=INK)
    tb(s, cx + 0.30, y + 1.06, cw - 0.5, 0.34, sub, size=8.8, color=MUTED,
       spacing=1.2)

# score
y4 = y + 1.86
card(s, M, y4, CW, 2.30)
section_label(s, M + 0.38, y4 + 0.26, 5.0, "Requirements checklist, re-scored")
rich(s, M + 0.38, y4 + 0.56, 4.2, 0.62,
     [("33", True, INK, 34), (" %", True, MUTED, 18)], spacing=0.95)
tb(s, M + 0.40, y4 + 1.22, 3.4, 0.5,
   "2 of 12 requirements complete,\n4 partially delivered", size=10,
   color=SLATE, spacing=1.3)

# segmented bar
seg_y = y4 + 0.72
seg_x = M + 4.05
bw, bg_ = 0.50, 0.075
status = ["done", "done", "partial", "todo", "todo", "todo", "todo",
          "todo", "partial", "partial", "partial", "todo"]
for i, st in enumerate(status):
    x = seg_x + i * (bw + bg_)
    if st == "done":
        rect(s, x, seg_y, bw, 0.30, TEAL, None, radius=0.35)
    elif st == "partial":
        rect(s, x, seg_y, bw, 0.30, AMBER_TINT, None, radius=0.35)
        rect(s, x, seg_y, bw * 0.5, 0.30, AMBER_L, None, radius=0.35)
    else:
        rect(s, x, seg_y, bw, 0.30, GREY_TINT, None, radius=0.35)
lx = seg_x
for lab, col in [("Complete", TEAL), ("Partial", AMBER_L),
                 ("Not started", RGBColor(0xD3, 0xDB, 0xE4))]:
    dot(s, lx, seg_y + 0.55, 0.12, col)
    tb(s, lx + 0.20, seg_y + 0.50, 1.8, 0.22, lab, size=9.5, color=MUTED)
    lx += 1.75
tb(s, seg_x, y4 + 1.60, 8.0, 0.5,
   "Ticks are this review's assessment; docs/requirements-checklist.md still "
   "shows 0/12 and calls the backend “pending”, which main has already "
   "overtaken.", size=9.5, color=FAINT, italic=True, spacing=1.3)

# strip facts
y5 = y4 + 2.52
facts = [("3", "contributors"), ("5", "notebooks"), ("3", "docs"),
         ("5.9", "k lines JSX"), ("0", "CI workflows")]
x = M
for n, lab in facts:
    rich(s, x, y5, 2.2, 0.28, [(n, True, INK, 13), ("   " + lab, False, MUTED, 9.5)])
    x += 2.42
footer(s)

# ================================================================ 5. MILESTONES
s = slide()
y = header(s, "Delivered", "Six milestones already on main",
           "Each card is backed by merged pull requests and files in the "
           "repository.")
cards = [
    ("Day 1–2", "Dataset decision & schema", 
     "Online Retail and Hillstrom compared, schemas documented, and the "
     "causal limits of each written down.", "#1 – #4", TEAL),
    ("Day 2–3", "Cleaning & upload validation",
     "Eight required columns enforced; duplicates, missing values and "
     "non-positive values handled. 5 tests.", "#3, #9", TEAL),
    ("Day 3–4", "Hillstrom quality & treatment",
     "Missing values, duplicate profiles and the three marketing segments "
     "profiled before any modelling.", "#5, #6", TEAL),
    ("Day 4", "Dashboard shell & navigation",
     "React + Plotly layout with a working sidebar, GitHub commit feed and "
     "merged-PR panel.", "#7, #11", TEAL),
    ("Day 5", "Methodology & assumptions",
     "Eight identification assumptions documented, plus the exact limits of "
     "observational price data.", "#10", TEAL),
    ("Day 6", "Export, validator & report",
     "Product-month CSV export, six-column CSV validator with 7 tests, and a "
     "downloadable dataset report.", "#12 – #18", TEAL),
]
cw, gap = 3.86, 0.20
for i, (day, title, body, prref, col) in enumerate(cards):
    cx = M + (i % 3) * (cw + gap)
    cy = y + (i // 3) * 2.30
    card(s, cx, cy, cw, 2.06)
    rect(s, cx, cy, 0.045, 2.06, col)
    tb(s, cx + 0.30, cy + 0.24, 2.4, 0.22, day, size=8.8, bold=True, color=TEAL,
       caps=True, tracking=1.1)
    tb(s, cx + 0.30, cy + 0.50, cw - 0.55, 0.52, title, size=13, bold=True,
       color=INK, font=HEAD, spacing=1.02)
    tb(s, cx + 0.30, cy + 1.04, cw - 0.58, 0.72, body, size=9.8, color=MUTED,
       spacing=1.3)
    tb(s, cx + 0.30, cy + 1.76, cw - 0.55, 0.2, prref, size=8.5, color=FAINT,
       bold=True)
tb(s, M, 6.62, CW, 0.24,
   "All 18 pull requests were merged into main; nothing is sitting open or "
   "unreviewed.", size=9.5, color=FAINT, italic=True)
footer(s)

# ================================================================ 6. DATA
s = slide()
y = header(s, "Delivered · data layer", "The data foundation is solid",
           "Two datasets inspected and documented — with the caveats that "
           "decide what can be estimated later.")
# online retail
cw = 6.05
card(s, M, y, cw, 3.05)
pill(s, M + cw - 1.32, y + 0.26, "PRIMARY", WHITE, TEAL, w=0.98)
tb(s, M + 0.34, y + 0.26, 3.6, 0.3, "Online Retail  ·  UCI", size=13.5,
   bold=True, color=INK, font=HEAD)
tb(s, M + 0.34, y + 0.60, cw - 0.7, 0.24,
   "541,909 rows × 8 columns  →  524,878 rows × 9 columns after cleaning",
   size=9.5, color=MUTED)
rows = [("Exact duplicate rows removed", "5,268"),
        ("Missing Description", "1,454"),
        ("Missing CustomerID (kept)", "135,080"),
        ("UnitPrice ≤ 0 / Quantity ≤ 0", "2,517 / 10,624")]
yy = y + 0.98
for i, (k, v) in enumerate(rows):
    if i % 2 == 0:
        rect(s, M + 0.22, yy - 0.06, cw - 0.44, 0.42, GREY_TINT, None,
             radius=0.25)
    tb(s, M + 0.36, yy, cw - 1.9, 0.24, k, size=10, color=SLATE)
    tb(s, M + cw - 1.55, yy, 1.3, 0.24, v, size=10, bold=True, color=INK,
       align=PP_ALIGN.RIGHT)
    yy += 0.42
tb(s, M + 0.34, y + 2.70, cw - 0.7, 0.26,
   "Revenue = Quantity × UnitPrice added as the analysis feature.",
   size=9, color=TEAL, bold=True)

# hillstrom
hx = M + cw + 0.26
card(s, hx, y, CW - cw - 0.26, 3.05)
pill(s, hx + CW - cw - 1.62, y + 0.26, "SECONDARY", MUTED, GREY_TINT, w=1.28)
tb(s, hx + 0.34, y + 0.26, 3.4, 0.3, "Hillstrom  ·  e-mail test", size=13.5,
   bold=True, color=INK, font=HEAD)
tb(s, hx + 0.34, y + 0.60, 4.4, 0.24, "64,000 rows × 12 columns  ·  0 missing "
   "values", size=9.5, color=MUTED)
hrows = [("Randomised arms", "Mens / Womens / No e-mail"),
         ("Outcomes", "visit, conversion, spend"),
         ("Duplicate rows flagged", "6,562  (not removed)"),
         ("Treatment type", "e-mail, not price")]
yy = y + 0.98
for i, (k, v) in enumerate(hrows):
    if i % 2 == 0:
        rect(s, hx + 0.22, yy - 0.06, CW - cw - 0.70, 0.42, GREY_TINT, None,
             radius=0.25)
    tb(s, hx + 0.36, yy, 1.85, 0.36, k, size=10, color=SLATE, spacing=1.15)
    tb(s, hx + 2.15, yy, 2.35, 0.36, v, size=9.8, bold=True, color=INK,
       spacing=1.15)
    yy += 0.46
tb(s, hx + 0.34, y + 2.70, 4.35, 0.26,
   "Only randomised arm in either dataset.", size=9, color=AMBER, bold=True)

# bottom notes
y6 = y + 3.28
notes = [("Why two datasets", "No shared column or join key exists, so they "
          "stay separate; merging them was explicitly rejected in Day 2."),
         ("Open item", "Hillstrom's 6,562 duplicate profiles were flagged but "
          "not investigated — removing them now would change the estimand.")]
for i, (h, b) in enumerate(notes):
    cx = M + i * (cw + 0.26)
    w = cw if i == 0 else CW - cw - 0.26
    rich(s, cx, y6, w, 0.66, [(h + "  —  ", True, INK, 10), (b, False, MUTED, 10)],
         spacing=1.32)

# the data contract the export actually produces
y12 = y6 + 0.82
card(s, M, y12, CW, 0.86, fill=GREY_TINT, line=None)
tb(s, M + 0.34, y12 + 0.18, 4.4, 0.22, "Hand-off schema  ·  product_month_summary.csv",
   size=9, bold=True, color=INK, caps=True, tracking=1.0)
tb(s, M + 0.34, y12 + 0.48, CW - 0.7, 0.26,
   "StockCode  ·  Month  ·  Average_Price  ·  Total_Quantity  ·  "
   "Total_Revenue  ·  Transaction_Count", size=10.5, bold=True, color=TEAL,
   font="Consolas")
tb(s, M + 7.15, y12 + 0.18, 4.4, 0.66,
   "Six columns, validated before the dashboard consumes them. This is the "
   "table the causal stage must build on.", size=8.8, color=MUTED,
   spacing=1.28, align=PP_ALIGN.RIGHT)
footer(s)

# ================================================================ 7. BUILD
s = slide()
y = header(s, "Delivered · application layer",
           "What actually runs end to end",
           "A working upload → validate → export → report pipeline, wrapped in "
           "a React dashboard.")
# API panel
cw = 6.55
card(s, M, y, cw, 3.60)
tb(s, M + 0.34, y + 0.24, 3.0, 0.24, "FastAPI  ·  main.py", size=10, bold=True,
   color=INK, caps=True, tracking=1.0)
routes = [("GET", "/api/health", "Liveness probe for the dashboard"),
          ("POST", "/api/upload", ".xlsx validation, 5-row preview, summary"),
          ("POST", "/api/export-product-month",
           "Product-month CSV, 6 columns, 2-dp rounding"),
          ("POST", "/api/dataset-report", "Plain-text data-quality report"),
          ("GET", "/", "Service banner")]
yy = y + 0.62
for i, (m, path, what) in enumerate(routes):
    if i % 2 == 0:
        rect(s, M + 0.22, yy - 0.05, cw - 0.44, 0.56, GREY_TINT, None,
             radius=0.2)
    pill(s, M + 0.34, yy + 0.05, m, WHITE, TEAL if m == "GET" else INK,
         w=0.52, size=7.5)
    tb(s, M + 0.96, yy + 0.02, 2.85, 0.24, path, size=10.5, bold=True,
       color=INK, font="Consolas")
    tb(s, M + 0.96, yy + 0.27, 4.7, 0.24, what, size=9, color=MUTED)
    yy += 0.52
tb(s, M + 0.34, y + 3.22, cw - 0.7, 0.24,
   "Type check on every route; bad input returns 400, never a silent 200.",
   size=8.6, color=MUTED, italic=True)

# frontend panel
fx = M + cw + 0.24
card(s, fx, y, CW - cw - 0.24, 3.60)
tb(s, fx + 0.34, y + 0.26, 3.4, 0.24, "React dashboard  ·  Vite", size=10,
   bold=True, color=INK, caps=True, tracking=1.0)
feats = [("Sidebar navigation", "6 sections wired to smooth scroll", TEAL),
         ("Upload panel", "Drag-drop .xlsx, preview table, summary tiles", TEAL),
         ("Export + report", "Downloads CSV and .txt from the API", TEAL),
         ("GitHub activity", "Live commits and merged PRs from main", TEAL),
         ("Plotly EDA charts", "Component built, waiting on data", AMBER),
         ("Recommendations, budget", "Placeholders only", ROSE)]
yy = y + 0.68
for name, desc, col in feats:
    dot(s, fx + 0.34, yy + 0.09, 0.10, col)
    tb(s, fx + 0.58, yy, 2.35, 0.24, name, size=10, bold=True, color=INK)
    tb(s, fx + 0.58, yy + 0.22, 3.95, 0.24, desc, size=9, color=MUTED)
    yy += 0.47

# footer note
y7 = y + 3.90
rich(s, M, y7, CW, 0.44,
     [("Read this as a pipeline, not a model.  ", True, INK, 10.5),
      ("Upload, validation, export and reporting are production-shaped and "
       "tested. Nothing downstream of the export estimates a causal effect "
       "yet.", False, MUTED, 10.5)], spacing=1.3)
footer(s)

# ================================================================ 8. TESTS
s = slide()
y = header(s, "Delivered · quality", "16 tests, no CI, no API coverage",
           "Real automated coverage of the data utilities — with a gap where "
           "the web layer sits.")
# test split
cw = 6.05
card(s, M, y, cw, 2.60)
tb(s, M + 0.34, y + 0.26, 4.0, 0.24, "Test suite  ·  python -m unittest",
   size=10, bold=True, color=INK, caps=True, tracking=1.0)
groups = [("test_validate_upload.py", 5),
          ("test_export_product_month.py", 4),
          ("test_validate_product_month_csv.py", 7)]
yy = y + 0.80
for name, n in groups:
    tb(s, M + 0.34, yy, 3.3, 0.24, name, size=9.8, color=SLATE, font="Consolas")
    bar_w = 2.6 * (n / 7)
    rect(s, M + 3.62, yy + 0.04, 2.6, 0.18, GREY_TINT, None, radius=0.4)
    rect(s, M + 3.62, yy + 0.04, bar_w, 0.18, TEAL, None, radius=0.4)
    tb(s, M + 6.10, yy, 0.8, 0.24, str(n), size=10, bold=True, color=TEAL,
       align=PP_ALIGN.RIGHT)
    yy += 0.54
rich(s, M + 0.34, y + 2.22, 5.4, 0.42,
     [("16", True, INK, 14), ("  tests pass on synthetic fixtures — no real "
       "dataset is shipped", False, MUTED, 9.5)], spacing=1.25)

# what's covered
hx = M + cw + 0.26
card(s, hx, y, CW - cw - 0.26, 2.60)
tb(s, hx + 0.34, y + 0.26, 4.4, 0.24, "Covered", size=10, bold=True, color=INK,
   caps=True, tracking=1.0)
cov = ["Required vs unexpected columns rejected",
       "Empty files, blank codes, bad YYYY-MM values",
       "Negative prices, quantities and revenue",
       "Non-integer transaction counts",
       "Duplicate StockCode + Month keys",
       "Non-zero exit code on failure"]
yy = y + 0.62
for c in cov:
    tb(s, hx + 0.34, yy, 4.6, 0.26, "·  " + c, size=9.8, color=SLATE)
    yy += 0.31

# what's missing
y8 = y + 2.98
missing = [("No API tests", "The five FastAPI routes have no request-level "
            "tests at all."),
           ("No frontend tests", "oxlint is configured, but no component or "
            "interaction test exists."),
           ("No CI", "No .github/workflows — the 16 tests only run when a "
            "developer remembers."),
           ("No pinned deps", "requirements.txt omits EconML, DoWhy, SciPy and "
            "scikit-learn.")]
cw2 = (CW - 3 * 0.20) / 4
for i, (h, b) in enumerate(missing):
    cx = M + i * (cw2 + 0.20)
    card(s, cx, y8, cw2, 1.82, fill=ROSE_TINT, line=None)
    tb(s, cx + 0.26, y8 + 0.22, cw2 - 0.5, 0.24, h, size=10.5, bold=True,
       color=ROSE)
    tb(s, cx + 0.26, y8 + 0.52, cw2 - 0.5, 0.96, b, size=9.3, color=SLATE,
       spacing=1.3)
footer(s)

# ================================================================ 9. GAP
s = slide(INK)
rect(s, 0, 0, 0.10, SH, AMBER_L)
tb(s, M + 0.05, 0.52, CW, 0.24, "The honest gap", size=10.5, bold=True,
   color=AMBER_L, caps=True, tracking=1.6)
tb(s, M + 0.05, 0.82, 9.6, 0.60,
   "Why the centrepiece result is still missing",
   size=27, bold=True, color=WHITE, font=HEAD)
rect(s, M + 0.05, 1.60, 6.0, 0.03, RGBColor(0x2A, 0x36, 0x4A))

gaps = [("01", "Online Retail has no randomised price",
         "Prices are historical and set by the retailer. Without a treatment "
         "assignment, a price–quantity relationship cannot be read as a causal "
         "effect."),
        ("02", "Hillstrom randomises e-mail, not price",
         "It can support a defensible e-mail uplift estimate. That is a "
         "different estimand from discount elasticity and must not be "
         "relabelled as one."),
        ("03", "No treatment, outcome or covariate set is signed off",
         "Section A still owns those definitions. Until they arrive there is "
         "no DAG, no valid estimator and no defensible budget optimisation.")]
yy = 2.00
for num, head, body in gaps:
    tb(s, M + 0.05, yy, 0.7, 0.34, num, size=17, bold=True, color=AMBER_L,
       font=HEAD)
    tb(s, M + 0.78, yy, 6.35, 0.28, head, size=13, bold=True, color=WHITE,
       font=HEAD)
    tb(s, M + 0.78, yy + 0.34, 6.5, 0.60, body, size=10.3,
       color=RGBColor(0xB6, 0xC3, 0xD4), spacing=1.32)
    yy += 1.32

card(s, 8.05, 2.00, 4.50, 3.72, fill=RGBColor(0x1A, 0x25, 0x3A), line=None)
tb(s, 8.40, 2.28, 3.8, 0.24, "Consequence", size=9, bold=True, color=TEAL_L,
   caps=True, tracking=1.2)
tb(s, 8.40, 2.60, 3.85, 1.10,
   "The dashboard is deliberately right to show “Not yet estimated”.",
   size=14, bold=True, color=WHITE, font=HEAD, spacing=1.15)
tb(s, 8.40, 3.74, 3.85, 1.30,
   "Filling those cards with numbers before the identification argument "
   "exists would be the single most damaging thing this project could do — it "
   "is the first thing a reviewer will test.",
   size=10, color=RGBColor(0xB6, 0xC3, 0xD4), spacing=1.34)
rect(s, 8.40, 5.10, 3.70, 0.012, RGBColor(0x33, 0x41, 0x55))
tb(s, 8.40, 5.24, 3.8, 0.3, "Unblocked by: Section A definitions + a "
   "documented identification strategy", size=9, color=FAINT, spacing=1.3)

tb(s, M + 0.05, 6.28, 9.2, 0.3,
   "Everything on the next three slides depends on resolving 01 – 03.",
   size=10.5, color=AMBER_L, bold=True)
footer(s, dark=True)

# ================================================================ 10. SCORECARD
s = slide()
y = header(s, "Remaining work · scorecard",
           "Twelve requirements, re-scored against main",
           "The checklist flags itself as out of date; this is what the code "
           "supports today.")
items = [("Select and document the dataset", "done"),
         ("Validate and clean the data", "done"),
         ("Exploratory data analysis", "partial"),
         ("Define treatment, outcome, covariates", "todo"),
         ("Create a causal DAG", "todo"),
         ("Implement DML with EconML", "todo"),
         ("Validate assumptions with DoWhy", "todo"),
         ("Optimise budget with SciPy", "todo"),
         ("FastAPI backend endpoints", "partial"),
         ("React dashboard with Plotly", "partial"),
         ("Tests and reproducibility", "partial"),
         ("Final report and project demo", "todo")]
meta = {"done": ("Complete", TEAL, TEAL_TINT),
        "partial": ("Partial", AMBER, AMBER_TINT),
        "todo": ("Not started", ROSE, ROSE_TINT)}
col_w = (CW - 0.60) / 2
for i, (label, st) in enumerate(items):
    col = i // 6
    cx = M + col * (col_w + 0.60)
    cy = y + (i % 6) * 0.72
    txt, fg, bg = meta[st]
    card(s, cx, cy, col_w, 0.60, fill=bg, line=None)
    dot(s, cx + 0.24, cy + 0.23, 0.14, fg)
    rich(s, cx + 0.52, cy + 0.19, col_w - 2.0, 0.26,
         [(f"{i + 1:02d}", True, FAINT, 9), ("   " + label, False, INK, 10.8)])
    tb(s, cx + col_w - 1.38, cy + 0.19, 1.16, 0.24, txt, size=8.8, bold=True,
       color=fg, align=PP_ALIGN.RIGHT)

y9 = y + 6 * 0.72 + 0.22
rich(s, M, y9, CW, 0.3,
     [("Bottom line:  ", True, INK, 10.5),
      ("the project has built a credible data and delivery layer. The four "
       "requirements that give it its name — DAG, DML, refutation, "
       "optimisation — are all still untouched.", False, MUTED, 10.5)])
footer(s)

# ================================================================ 11. TRACK 1
s = slide()
y = header(s, "Remaining work · track 1", "Causal core — five tasks, strictly ordered",
           "Each task is a gate: the next one cannot start honestly until the "
           "previous one is signed off.")
efforts = {"C1": "2–3 days", "C2": "1 week",
           "C3": "1 week", "C4": "3–4 days", "C5": "4–5 days"}
tasks = [("C1", "Freeze Section A definitions",
          "Treatment, outcome, covariates, estimand and unit of analysis, "
          "written into docs/.", "Blocker for C2 – C5",
          "Signed definitions doc", ROSE),
         ("C2", "Build the analysis table and DAG",
          "Product-month or customer-level panel, plus a committed DAG "
          "(dagitty / DoWhy graph).", "Needs C1",
          "Reproducible build script + DAG", AMBER),
         ("C3", "Estimate with EconML DML",
          "LinearDML or CausalForestDML, cross-fitting, cluster-robust "
          "intervals, overlap checks.", "Needs C1, C2",
          "ATE + CI + overlap report", AMBER),
         ("C4", "Refute with DoWhy",
          "Placebo treatment, random common cause, data subset and "
          "sensitivity analysis.", "Needs C3",
          "Refutation table in the report", AMBER),
         ("C5", "Allocate budget with SciPy",
          "Optimise spend from estimated responses; label every output as a "
          "simulation.", "Needs C3, C4",
          "Scenario + stated assumptions", ROSE)]
cw = 2.34
for i, (code, title, body, dep, done, col) in enumerate(tasks):
    cx = M + i * (cw + 0.16)
    card(s, cx, y, cw, 3.40)
    rect(s, cx, y, cw, 0.05, col)
    tb(s, cx + 0.24, y + 0.26, 1.0, 0.24, code, size=10, bold=True, color=col,
       caps=True, tracking=1.0)
    tb(s, cx + 0.24, y + 0.56, cw - 0.46, 0.62, title, size=12, bold=True,
       color=INK, font=HEAD, spacing=1.05)
    tb(s, cx + 0.24, y + 1.24, cw - 0.48, 1.10, body, size=9.2, color=MUTED,
       spacing=1.3)
    tb(s, cx + 0.24, y + 1.94, cw - 0.46, 0.22, "EFFORT", size=7.5, bold=True,
       color=FAINT, caps=True, tracking=0.9)
    tb(s, cx + 0.24, y + 2.12, cw - 0.46, 0.22, efforts[code], size=9.5,
       bold=True, color=INK)
    rect(s, cx + 0.24, y + 2.40, cw - 0.48, 0.010, BORDER)
    tb(s, cx + 0.24, y + 2.50, cw - 0.46, 0.22, "DEPENDS ON", size=7.5, bold=True,
       color=FAINT, caps=True, tracking=0.9)
    tb(s, cx + 0.24, y + 2.68, cw - 0.46, 0.22, dep, size=8.8, color=SLATE)
    tb(s, cx + 0.24, y + 2.96, cw - 0.46, 0.36, "Done when: " + done,
       size=8.2, color=TEAL, spacing=1.2)

y10 = y + 3.60
card(s, M, y10, CW, 1.02, fill=GREY_TINT, line=None)
rich(s, M + 0.34, y10 + 0.22, CW - 0.7, 0.60,
     [("Sizing:  ", True, INK, 10.3),
      ("C1 is days, not weeks — but it belongs to Section A and is the reason "
       "this track has not started. C3 and C4 are the two tasks that turn this "
       "from a data dashboard into a causal study; budget roughly one week "
       "each with a clean analysis table.", False, SLATE, 10.3)],
     spacing=1.32)
footer(s)

# ================================================================ 12. TRACK 2
s = slide()
y = header(s, "Remaining work · track 2",
           "Product and platform — six fixes, none blocked",
           "This track can start today while Section A finishes its "
           "definitions.")
rows = [("P1", "Wire product-month CSV end to end",
         "Export and validator exist, but the dashboard never loads "
         "product_month_summary.csv.", "Unblocked"),
        ("P2", "Replace dashboard placeholders",
         "Four panels still say “will appear here”. Needs real EDA charts and "
         "the ATE cards.", "Unblocked"),
        ("P3", "Fix wiring defects",
         "Duplicate section id, report endpoint on port 8001 while the rest "
         "use 8000, EDACharts rendered without props.", "Unblocked"),
        ("P4", "Restore reproducibility",
         "data/ is git-ignored, so notebooks cannot run from a fresh clone. "
         "Add a fetch script or sample fixtures.", "Unblocked"),
        ("P5", "Repository hygiene",
         "Two .pyc files are tracked, and README.md holds Day-6 notes instead "
         "of a project overview.", "Unblocked"),
        ("P6", "Add CI",
         "GitHub Actions running the 16 tests plus oxlint on every pull "
         "request.", "Unblocked")]
cw = 6.02
for i, (code, title, body, status) in enumerate(rows):
    col = i % 2
    row = i // 2
    cx = M + col * (cw + 0.28)
    cy = y + row * 1.64
    card(s, cx, cy, cw, 1.42)
    tb(s, cx + 0.26, cy + 0.20, 0.55, 0.24, code, size=10, bold=True,
       color=TEAL, caps=True, tracking=0.9)
    tb(s, cx + 0.78, cy + 0.18, cw - 2.1, 0.26, title, size=11.5, bold=True,
       color=INK, font=HEAD)
    tb(s, cx + 0.78, cy + 0.56, cw - 1.05, 0.72, body, size=9.3, color=MUTED,
       spacing=1.30)
    pill(s, cx + cw - 1.20, cy + 0.19, status, TEAL, TEAL_TINT, w=0.94,
         size=7.4)
footer(s)

# ================================================================ 13. ROADMAP
s = slide()
y = header(s, "Plan", "Four phases to a defensible result",
           "Sequenced so that visible progress never outruns the identification "
           "argument.")
phases = [("Phase 0", "Unblock", "Week 1",
           ["C1 definitions signed off", "P3 defects fixed", "P4–P5 hygiene"],
           "Gate: DAG + data contract agreed", TEAL),
          ("Phase 1", "Estimate", "Weeks 1–2",
           ["C2 analysis table + DAG", "C3 DML with intervals", "C4 refutation"],
           "Gate: ATE with CI survives refutation", AMBER),
          ("Phase 2", "Surface", "Weeks 2–3",
           ["P1 CSV into dashboard", "P2 real charts + ATE cards", "P6 CI green"],
           "Gate: no placeholder panels left", AMBER),
          ("Phase 3", "Decide", "Weeks 3–4",
           ["C5 SciPy allocation", "Recommendations", "Report + demo"],
           "Gate: simulated allocation, labelled", ROSE)]
cw = 2.94
# timeline rule
rect(s, M, y + 0.34, CW - 0.5, 0.014, RGBColor(0xD8, 0xE1, 0xEA))
for i, (ph, name, when, items_, gate, col) in enumerate(phases):
    cx = M + i * (cw + 0.20)
    dot(s, cx + 0.16, y + 0.28, 0.14, col)
    card(s, cx, y + 0.70, cw, 3.24)
    tb(s, cx + 0.26, y + 0.92, 2.4, 0.22, ph, size=9, bold=True, color=col,
       caps=True, tracking=1.1)
    tb(s, cx + 0.26, y + 1.18, 2.4, 0.32, name, size=16, bold=True, color=INK,
       font=HEAD)
    tb(s, cx + 0.26, y + 1.54, 2.4, 0.24, when, size=9.5, color=MUTED)
    yy = y + 1.88
    for it in items_:
        tb(s, cx + 0.26, yy, cw - 0.48, 0.42, "·  " + it, size=9.2, color=SLATE,
           spacing=1.2)
        yy += 0.40 if len(it) < 26 else 0.56
    rect(s, cx + 0.26, y + 3.34, cw - 0.52, 0.010, BORDER)
    tb(s, cx + 0.26, y + 3.44, cw - 0.50, 0.44, gate, size=8.4, bold=True,
       color=col, spacing=1.22)
y11 = y + 4.16
rich(s, M, y11, CW, 0.5,
     [("If Section A slips:  ", True, INK, 10.3),
      ("run Phase 0 and Phase 2 in parallel and hold the causal slides empty. "
       "A dashboard with honest blanks is recoverable; a dashboard with "
       "invented elasticities is not.", False, MUTED, 10.3)], spacing=1.3)
footer(s)

# ================================================================ 14. RISKS
s = slide()
y = header(s, "Risks", "Four things that could sink this",
           "Each has an owner-facing mitigation, not just a warning.")
# table header
tw = CW
cols = [0.35, 3.35, 3.90, 4.15]
heads = ["", "Risk", "Why it matters", "Mitigation"]
x = M
rect(s, M, y, tw, 0.42, INK, None, radius=0.12)
for i, (h, w) in enumerate(zip(heads, cols)):
    if h:
        tb(s, x + 0.12, y + 0.10, w - 0.2, 0.24, h, size=9, bold=True,
           color=WHITE, caps=True, tracking=1.1)
    x += w
risks = [("Section A definitions never land",
          "The entire causal track stays at zero and the deck repeats itself "
          "next week.",
          "Set a hard date; escalate to the project lead; ship P1–P6 meanwhile."),
         ("No valid price variation in the data",
          "A DML price elasticity would be an artefact of confounding, not a "
          "causal number.",
          "Re-scope to Hillstrom e-mail uplift, or state a simulated-price "
          "design explicitly."),
         ("Placeholders still visible at demo",
          "Reviewer sees “Pending” cards and reads the project as unfinished "
          "scaffolding.",
          "Hide unfinished sections behind a “in progress” tab until real "
          "output exists."),
         ("Nothing is reproducible from the clone",
          "data/ is ignored, .pyc files are tracked, no CI: results cannot be "
          "re-run.",
          "Fetch script or fixtures, untrack build artefacts, add Actions "
          "workflow.")]
ry = y + 0.52
for i, (risk, why, mit) in enumerate(risks):
    h = 1.02
    if i % 2 == 0:
        rect(s, M, ry - 0.04, tw, h, WHITE, BORDER, radius=0.06)
    else:
        rect(s, M, ry - 0.04, tw, h, GREY_TINT, BORDER, radius=0.06)
    rect(s, M + 0.14, ry + 0.22, 0.10, 0.10, ROSE)
    tb(s, M + 0.35, ry + 0.14, 3.10, 0.66, risk, size=10.8, bold=True,
       color=INK, spacing=1.2)
    tb(s, M + 3.70, ry + 0.14, 3.70, 0.74, why, size=9.5, color=MUTED,
       spacing=1.28)
    tb(s, M + 7.60, ry + 0.14, 4.10, 0.74, mit, size=9.5, color=SLATE,
       spacing=1.28)
    ry += h + 0.08
footer(s)

# ================================================================ 15. CLOSE
s = slide(INK)
rect(s, 0, 0, 0.10, SH, TEAL)
tb(s, M + 0.05, 0.60, CW, 0.24, "Next seven days", size=10.5, bold=True,
   color=TEAL_L, caps=True, tracking=1.6)
tb(s, M + 0.05, 0.92, 9.0, 0.55, "What needs to happen", size=30, bold=True,
   color=WHITE, font=HEAD)

actions = [("1", "Get the Section A definitions in writing",
            "Treatment, outcome, covariates, estimand. This unlocks everything "
            "else.", "Project lead"),
           ("2", "Choose the causal target",
            "Hillstrom e-mail uplift (identifiable now) or retail price "
            "elasticity (needs a design).", "Section A + B"),
           ("3", "Ship the unblocked platform fixes",
            "P1, P3, P5, P6 — CSV into the dashboard, defects, hygiene, CI.",
            "Section B"),
           ("4", "Stand up the analysis table and DAG",
            "Reproducible build script plus a committed graph.", "Section B"),
           ("5", "Then, and only then, train DML",
            "One estimator, cluster-robust CIs, overlap diagnostics, refutation "
            "table.", "Section B")]
yy = 1.74
for i, (num, head, body, who) in enumerate(actions):
    tb(s, M + 0.05, yy - 0.02, 0.42, 0.32, num, size=13, bold=True,
       color=TEAL_L, font=HEAD)
    tb(s, M + 0.60, yy - 0.02, 4.9, 0.30, head, size=12.5, bold=True,
       color=WHITE, font=HEAD)
    tb(s, M + 5.62, yy - 0.01, 4.55, 0.42, body, size=10, color=RGBColor(
        0xB6, 0xC3, 0xD4), spacing=1.25)
    tb(s, 11.05, yy - 0.02, 1.60, 0.28, who, size=8.5, color=FAINT,
       align=PP_ALIGN.RIGHT, caps=True, tracking=0.6)
    yy += 0.74

rect(s, M + 0.05, 5.68, 1.6, 0.03, TEAL)
rich(s, M + 0.05, 5.88, 9.9, 0.70,
     [("Open decision for the reviewer:  ", True, TEAL_L, 11),
      ("do we estimate the e-mail uplift that Hillstrom can actually support, "
       "or hold for a pricing design we do not yet have? The answer determines "
       "whether the next milestone produces a number or a caveat.",
       False, RGBColor(0xC7, 0xD2, 0xE0), 11)], spacing=1.35)
tb(s, M + 0.05, 6.86, CW, 0.26,
   "Sources: snigdhasudheer/Infotact_project1 @ main · README, docs/day2, "
   "docs/day5, docs/requirements-checklist, main.py, frontend/src, test suite",
   size=7.8, color=RGBColor(0x64, 0x74, 0x8B))

# speaker notes
for i, sl in enumerate(prs.slides, start=1):
    if i in NOTES:
        sl.notes_slide.notes_text_frame.text = NOTES[i]

prs.save("EconoCausal_Project_Status_Review.pptx")
print("saved:", len(prs.slides.__iter__.__self__._sldIdLst), "slides")
