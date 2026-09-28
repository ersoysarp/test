#!/usr/bin/env python3
"""Build an editable widescreen PowerPoint of the Blue Currency deck."""

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import nsmap, qn
from pptx.util import Emu, Inches, Pt
from lxml import etree

# 16:9 widescreen — 13.333" x 7.5", matching the 1280x720 HTML stage
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

INK = RGBColor(0x26, 0x30, 0x3A)
INK2 = RGBColor(0x3A, 0x45, 0x4F)
GREY = RGBColor(0x6B, 0x75, 0x80)
GREY_L = RGBColor(0x9A, 0xA3, 0xAC)
RULE = RGBColor(0xE3, 0xE8, 0xE9)
SOFT = RGBColor(0xF7, 0xF9, 0xF9)
TEAL = RGBColor(0x0B, 0x7F, 0x77)
TEAL_D = RGBColor(0x0A, 0x5F, 0x59)
TEAL_SOFT = RGBColor(0xE4, 0xF1, 0xEF)
AMBER = RGBColor(0xB9, 0x7D, 0x1E)
AMBER_SOFT = RGBColor(0xFA, 0xF1, 0xDF)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREEN = RGBColor(0x2E, 0x9E, 0x5B)
RED = RGBColor(0xC0, 0x45, 0x3B)
RED_SOFT = RGBColor(0xF5, 0xE4, 0xE1)
GREEN_SOFT = RGBColor(0xE9, 0xF4, 0xED)
AMBER_BOX = RGBColor(0xFB, 0xF3, 0xE3)
RED_BOX = RGBColor(0xF9, 0xEA, 0xE7)
MOD_BG = RGBColor(0xED, 0xF0, 0xF2)
DARK = RGBColor(0x0E, 0x2A, 0x31)
CRIT = RGBColor(0xC0, 0x45, 0x3B)
HIGH = RGBColor(0xD9, 0x98, 0x2F)
MED = RGBColor(0x98, 0xA2, 0xAB)
TRACK = RGBColor(0xED, 0xF0, 0xF2)


def rgb(color):
    return color


def set_fill(shape, color):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color


def set_line(shape, color, pt=1.0):
    shape.line.color.rgb = color
    shape.line.width = Pt(pt)


def no_line(shape):
    shape.line.fill.background()


def set_run(run, text, size=14, bold=False, color=INK, font="Calibri"):
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font


def add_text(slide, l, t, w, h, text, size=14, bold=False, color=INK,
             align=PP_ALIGN.LEFT, font="Calibri", anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    try:
        tf._txBody.bodyPr.set("anchor", {
            MSO_ANCHOR.TOP: "t",
            MSO_ANCHOR.MIDDLE: "ctr",
            MSO_ANCHOR.BOTTOM: "b",
        }.get(anchor, "t"))
    except Exception:
        pass
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    set_run(run, text, size, bold, color, font)
    return box


def add_rich(slide, l, t, w, h, parts, align=PP_ALIGN.LEFT, anchor="t", size=14):
    """parts = (text, bold, color) or (text, size, bold, color)."""
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    try:
        tf._txBody.bodyPr.set("anchor", anchor)
    except Exception:
        pass
    p = tf.paragraphs[0]
    p.alignment = align
    for part in parts:
        if len(part) == 3:
            text, bold, color = part
            sz = size
        else:
            text, sz, bold, color = part
        set_run(p.add_run(), text, sz, bold, color)
    return box


def add_rect(slide, l, t, w, h, fill, line=None, line_pt=1.0, radius=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE,
                                   l, t, w, h)
    set_fill(shape, fill)
    if line is None:
        no_line(shape)
    else:
        set_line(shape, line, line_pt)
    if radius is not None:
        try:
            shape.adjustments[0] = radius
        except Exception:
            pass
    return shape


def add_oval(slide, l, t, w, h, fill, color=WHITE, text="", size=12, bold=True):
    sh = slide.shapes.add_shape(MSO_SHAPE.OVAL, l, t, w, h)
    set_fill(sh, fill)
    no_line(sh)
    if text:
        tf = sh.text_frame
        tf.word_wrap = False
        try:
            tf._txBody.bodyPr.set("anchor", "ctr")
        except Exception:
            pass
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        set_run(p.add_run(), text, size, bold, color)
    return sh


def add_shape_text(shape, text, size=14, bold=False, color=INK, align=PP_ALIGN.LEFT, anchor="ctr"):
    tf = shape.text_frame
    tf.word_wrap = True
    try:
        tf._txBody.bodyPr.set("anchor", anchor)
    except Exception:
        pass
    # inset so text doesn't hug the edge
    tf.margin_left = Inches(0.08)
    tf.margin_right = Inches(0.08)
    tf.margin_top = Inches(0.04)
    tf.margin_bottom = Inches(0.04)
    p = tf.paragraphs[0]
    p.alignment = align
    set_run(p.add_run(), text, size, bold, color)
    return tf


def add_notes(slide, text):
    notes = slide.notes_slide
    notes.notes_text_frame.text = text


def kicker_title(slide, kicker, title_parts):
    add_text(slide, Inches(0.42), Inches(0.22), Inches(12.4), Inches(0.28),
             kicker, size=14, bold=True, color=TEAL)
    box = slide.shapes.add_textbox(Inches(0.42), Inches(0.46), Inches(12.5), Inches(0.48))
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    for text, bold, color in title_parts:
        set_run(p.add_run(), text, 25, bold, color)


def build_slide1(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, WHITE)
    kicker_title(
        slide,
        "BLUE CURRENCY  |  AGENTIC AI FOR REAL ESTATE DEVELOPMENT AND CONSTRUCTION",
        [
            ("Blue Currency is the ", True, INK),
            ("first agentic AI platform", True, TEAL_D),
            (" that runs the entire construction lifecycle, end to end", True, INK),
        ],
    )

    # left diagram panel
    add_rect(slide, Inches(0.42), Inches(1.08), Inches(7.15), Inches(6.12),
             RGBColor(0xF3, 0xFA, 0xF9), RULE, 0.75, radius=0.08)

    orch = add_rect(slide, Inches(0.62), Inches(1.22), Inches(6.75), Inches(0.58),
                    TEAL_D, radius=0.12)
    add_shape_text(orch, "1     End to End Agentic Orchestration for Construction Lifecycle Management",
                   15, True, WHITE, PP_ALIGN.CENTER)

    mods = [
        (2, "AI Generated Unified BoQ"),
        (3, "Dynamic Bottom Up CAPEX Estimation"),
        (4, "Smart Tender Bid Evaluation"),
        (5, "Agentic Cost Realization Tracking"),
    ]
    mw = Inches(1.58)
    gap = Inches(0.12)
    x0 = Inches(0.62)
    y = Inches(1.96)
    for i, (n, label) in enumerate(mods):
        x = x0 + i * (mw + gap)
        box = add_rect(slide, x, y, mw, Inches(1.18), WHITE, RULE, 0.75, radius=0.1)
        add_oval(slide, x + Inches(0.58), y + Inches(0.10), Inches(0.38), Inches(0.38),
                 TEAL, WHITE, str(n), 13)
        add_text(slide, x + Inches(0.06), y + Inches(0.52), mw - Inches(0.12), Inches(0.60),
                 label, 14, True, INK, PP_ALIGN.CENTER)

    val = add_rect(slide, Inches(0.62), Inches(3.30), Inches(6.75), Inches(0.50),
                   WHITE, AMBER, 1.25, radius=0.1)
    add_oval(slide, Inches(2.95), Inches(3.38), Inches(0.34), Inches(0.34), AMBER, WHITE, "6", 12)
    add_text(slide, Inches(3.35), Inches(3.38), Inches(2.9), Inches(0.34),
             "AI-Based Value Optimizer", 15, True, AMBER, PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)

    core = add_rect(slide, Inches(0.62), Inches(3.96), Inches(6.75), Inches(0.72),
                    DARK, radius=0.1)
    add_shape_text(core, "AGENTIC AI CORE\nthe reasoning engine that powers end to end orchestration",
                   14, True, WHITE, PP_ALIGN.CENTER)

    add_text(slide, Inches(0.62), Inches(4.80), Inches(6.75), Inches(0.26),
             "BUILT ON EXISTING MCKINSEY ASSETS", 12, True, GREY, PP_ALIGN.CENTER)
    add_rect(slide, Inches(1.4), Inches(5.08), Inches(5.2), Inches(0.015), TEAL)

    assets = [
        "SourceAI", "ContractAI", "CleanSheetAI",
        "Procurement Control Tower", "Construction Control Tower", "Other assets",
    ]
    aw = Inches(1.05)
    for i, name in enumerate(assets):
        x = Inches(0.62) + i * (aw + Inches(0.08))
        pill = add_rect(slide, x, Inches(5.28), aw, Inches(0.68), WHITE, TEAL_D, 1.0, radius=0.5)
        add_shape_text(pill, name, 13, True, TEAL_D, PP_ALIGN.CENTER)

    # right column
    add_text(slide, Inches(7.78), Inches(1.08), Inches(5.12), Inches(0.28),
             "WHAT EACH MODULE DOES", 12, True, INK)
    add_rect(slide, Inches(7.78), Inches(1.36), Inches(5.12), Inches(0.02), INK)

    items = [
        (TEAL_D, "1", "End to End Agentic Orchestration",
         "Single platform view to manage projects from project brief to end of construction"),
        (TEAL, "2", "AI Generated Unified BoQ",
         "AI-based unified BoQ catalog creator and data cleaner at spec level, using the client's dirty historical data"),
        (TEAL, "3", "Dynamic Bottom Up CAPEX",
         "Item-level cost breakdown structure creator, with live commodity and macro-economic price index integration for forecasting"),
        (TEAL, "4", "Smart Tender Bid Evaluation",
         "Automatically collects bids, flags outliers and generates an AI-recommended price range to score each bid"),
        (TEAL, "5", "Agentic Cost Realization Tracking",
         "Agentic contract and IPC collection to track cost realization and project progress in real time"),
        (AMBER, "6", "AI-Based Value Optimizer",
         "Agentic workflow to predict risks and procurement value optimization opportunities by learning from historical data"),
    ]
    y = Inches(1.46)
    for fill, num, name, sub in items:
        add_oval(slide, Inches(7.78), y + Inches(0.04), Inches(0.32), Inches(0.32),
                 fill, WHITE, num, 11)
        add_text(slide, Inches(8.18), y, Inches(4.72), Inches(0.24), name, 14, True, INK)
        add_text(slide, Inches(8.18), y + Inches(0.22), Inches(4.72), Inches(0.28), sub, 14, False, GREY)
        add_rect(slide, Inches(8.18), y + Inches(0.50), Inches(4.72), Inches(0.01), RULE)
        y += Inches(0.56)

    add_text(slide, Inches(7.78), Inches(4.90), Inches(5.12), Inches(0.26),
             "KEY VALUE PROPOSITIONS", 12, True, INK)
    add_rect(slide, Inches(7.78), Inches(5.16), Inches(5.12), Inches(0.02), INK)

    vps = [
        "More complete than AI point tools, more AI-native than legacy suites",
        "Easy to use: covering all the fundamentals, without suite complexities",
        "Plug and play: limited data prep or integration need",
        "Modular, and works with other McKinsey assets",
    ]
    vw = Inches(2.48)
    vh = Inches(0.82)
    for i, vp in enumerate(vps):
        col = i % 2
        row = i // 2
        x = Inches(7.78) + col * (vw + Inches(0.14))
        yy = Inches(5.30) + row * (vh + Inches(0.10))
        card = add_rect(slide, x, yy, vw, vh, TEAL_SOFT, radius=0.08)
        add_rect(slide, x, yy, Inches(0.07), vh, TEAL)
        add_text(slide, x + Inches(0.14), yy + Inches(0.08), vw - Inches(0.22), vh - Inches(0.12),
                 vp, 14, True, TEAL_D)

    add_notes(slide, "Page 1. Introduction. Blue Currency is the first agentic AI platform that runs the entire construction lifecycle, end to end. Modules: orchestration, unified BoQ, dynamic CAPEX, tender evaluation, procurement and progress tracking, plus optional Value Optimizer. Built on existing McKinsey assets.")


def build_slide2(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, WHITE)
    kicker_title(
        slide,
        "BLUE CURRENCY  |  THE PROBLEM WE SOLVE",
        [
            ("Even the market leaders still run these workflows manually, at ", True, INK),
            ("severe and compounding cost", True, TEAL_D),
        ],
    )

    headers = [
        (Inches(0.42), Inches(2.15), INK, "Module"),
        (Inches(2.65), Inches(7.20), INK2,
         "How it is done today\nleaders and developing players, with the severity of each"),
        (Inches(9.95), Inches(2.96), TEAL, "What Blue Currency does"),
    ]
    hy = Inches(1.32)
    for x, w, fill, text in headers:
        h = add_rect(slide, x, hy, w, Inches(0.58), fill, radius=0.04)
        add_shape_text(h, text, 14, True, WHITE, PP_ALIGN.LEFT, "ctr")

    LEVELS = {"HIGH": CRIT, "MEDIUM": HIGH, "LOW": MED}

    rows = [
        ("1", TEAL_D, "End to End Agentic Orchestration",
         "Core ERP with ad hoc point tools and limited integration", "MEDIUM",
         "No end-to-end tracking, only e-mails, Excels and random SharePoints", "HIGH",
         "One governed agentic layer across every stage, smarter with each project"),
        ("2", TEAL, "AI Generated Unified BoQ",
         "Unified BoQ exists, but not always at spec level and manually maintained", "HIGH",
         "Only standardization in WBS, but near-zero at BoQ level", "HIGH",
         "Using the client's own historical data, an AI gatekeeper maintains the structure"),
        ("3", TEAL, "Dynamic Bottom Up CAPEX",
         "Dependent on expensive cost consultants or top-down estimation models", "LOW",
         "Last cost plus inflation, with no accuracy tracking", "MEDIUM",
         "Item-level bottom-up cost breakdown with live index data integration"),
        ("4", TEAL, "Smart Tender Bid Evaluation",
         "E-tender solutions are adopted, but with limited AI use in comparison", "LOW",
         "Manual consolidation and comparison in random Excels", "MEDIUM",
         "Comparison in seconds, with an AI-based price recommendation"),
        ("5", TEAL, "Agentic Cost Realization Tracking",
         "PM and ERP solutions used, mostly not integrated", "MEDIUM",
         "Manual tracking in Excels and e-mails", "MEDIUM",
         "Predicts future delays from historical realizations"),
        ("6", AMBER, "AI-Based Value Optimizer",
         "Expert-led value engineering with limited AI use", "HIGH",
         "Most value opportunities are not recognized", "HIGH",
         "Not just savings: identifies opportunities and predicts future risks and issues"),
    ]

    y = Inches(1.98)
    rh = Inches(0.86)
    for num, fill, name, ltxt, llvl, dtxt, dlvl, sol in rows:
        add_rect(slide, Inches(0.42), y + rh - Inches(0.01), Inches(12.49), Inches(0.01), RULE)
        add_oval(slide, Inches(0.50), y + Inches(0.26), Inches(0.32), Inches(0.32), fill, WHITE, num, 11)
        add_text(slide, Inches(0.90), y + Inches(0.18), Inches(1.68), Inches(0.56), name, 13, True, INK)

        for idx, (lab, txt, lvl, lab_col) in enumerate([
            ("LEADERS", ltxt, llvl, TEAL_D),
            ("DEVELOPING", dtxt, dlvl, GREY_L),
        ]):
            ry = y + Inches(0.06) + idx * Inches(0.40)
            add_text(slide, Inches(2.68), ry, Inches(1.00), Inches(0.30), lab, 11, True, lab_col)
            add_text(slide, Inches(3.70), ry, Inches(4.55), Inches(0.36), txt, 12, False, INK2)
            lvlbox = add_rect(slide, Inches(8.55), ry + Inches(0.02), Inches(1.15), Inches(0.28),
                              LEVELS[lvl], radius=0.04)
            add_shape_text(lvlbox, lvl, 11, True, WHITE, PP_ALIGN.CENTER)

        solb = add_rect(slide, Inches(9.95), y + Inches(0.06), Inches(2.96), Inches(0.72),
                        TEAL_SOFT, radius=0.04)
        add_shape_text(solb, sol, 12, True, TEAL_D, PP_ALIGN.LEFT, "ctr")
        y += rh

    add_notes(slide, "Page 2. The problem. Severity of the pain is rated High / Medium / Low separately for market leaders and developing players. The 'what it costs them' column was removed; severity carries that message.")


def chip(slide, x, y, w, h, label, fill=WHITE, line=RULE, color=INK2):
    sh = add_rect(slide, x, y, w, h, fill, line, 0.75, radius=0.5)
    add_shape_text(sh, label, 12, True, color, PP_ALIGN.CENTER)
    return sh


def build_slide3(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, WHITE)
    kicker_title(
        slide,
        "BLUE CURRENCY  |  COMPETITIVE LANDSCAPE",
        [
            ("The market falls into ", True, INK),
            ("four archetypes", True, TEAL_D),
            (", and we sit deliberately between them, not at the far corner", True, INK),
        ],
    )

    # y-axis
    yax = slide.shapes.add_textbox(Inches(0.08), Inches(1.8), Inches(0.36), Inches(4.4))
    tf = yax.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    set_run(run, "AI-NATIVE REASONING  →", 11, True, GREY)
    # rotate the axis label
    spPr = yax._element.spPr
    xfrm = spPr.find(qn("a:xfrm"))
    if xfrm is None:
        xfrm = etree.SubElement(spPr, qn("a:xfrm"))
    xfrm.set("rot", str(int(-5400000)))  # -90 degrees in 1/60000 deg

    # four quadrants
    q_w, q_h = Inches(5.95), Inches(2.28)
    qx1, qx2 = Inches(0.52), Inches(6.64)
    qy1, qy2 = Inches(1.10), Inches(4.28)

    # AI specialists
    add_rect(slide, qx1, qy1, q_w, q_h, SOFT, RULE, 0.75, radius=0.08)
    add_rich(slide, qx1 + Inches(0.14), qy1 + Inches(0.08), Inches(5.6), Inches(0.30),
             [("AI specialists  ", True, INK), ("AI-native, one step", True, GREY)])
    add_text(slide, qx1 + Inches(0.14), qy1 + Inches(0.36), Inches(5.6), Inches(0.28),
             "Built on models from day one, excellent at the one step they own.", 14, False, INK2)
    ai = ["Kreo", "Buildots", "Costify", "Zebel", "ConWize",
          "Ediphi", "Doxel", "OpenSpace", "ALICE", "nPlan"]
    cw = Inches(1.10)
    for i, name in enumerate(ai):
        chip(slide, qx1 + Inches(0.14) + (i % 5) * (cw + Inches(0.06)),
             qy1 + Inches(0.70) + (i // 5) * Inches(0.36),
             cw, Inches(0.30), name, WHITE, RGBColor(0xBB, 0xD9, 0xD5), TEAL_D)
    add_text(slide, qx1 + Inches(0.14), qy1 + Inches(1.50), Inches(5.6), Inches(0.24),
             "Strong in  takeoff, cost classification, tendering, site progress", 14, False, GREY)
    add_rich(slide, qx1 + Inches(0.14), qy1 + Inches(1.76), Inches(5.6), Inches(0.40),
             [("Each solves a slice, and ", True, INK),
              ("nothing carries across a handover.", True, RED)])

    # empty / agentic platforms
    add_rect(slide, qx2, qy1, q_w, q_h, RGBColor(0xEE, 0xF8, 0xF6),
             RGBColor(0x9C, 0xCB, 0xC5), 1.5, radius=0.08)
    add_rich(slide, qx2 + Inches(0.14), qy1 + Inches(0.08), Inches(5.6), Inches(0.30),
             [("Agentic platforms  ", True, TEAL_D), ("AI-native, across the lifecycle", True, GREY)])
    add_text(slide, qx2 + Inches(0.14), qy1 + Inches(0.40), Inches(5.6), Inches(0.36),
             "One platform that reasons across every stage, from BoQ through to site.", 14, False, INK2)
    empty = add_rect(slide, qx2 + Inches(1.35), qy1 + Inches(0.90), Inches(3.25), Inches(0.46),
                     RGBColor(0xDC, 0xEE, 0xEA), radius=0.08)
    add_shape_text(empty, "NO ONE IS HERE YET", 14, True, TEAL_D, PP_ALIGN.CENTER)
    add_rich(slide, qx2 + Inches(0.14), qy1 + Inches(1.60), Inches(5.6), Inches(0.50),
             [("Where the market is heading, ", True, TEAL_D),
              ("and what Blue Currency is built for.", True, TEAL_D)])

    # point tools
    add_rect(slide, qx1, qy2, q_w, q_h, SOFT, RULE, 0.75, radius=0.08)
    add_rich(slide, qx1 + Inches(0.14), qy2 + Inches(0.08), Inches(5.6), Inches(0.30),
             [("Point tools  ", True, INK), ("Rules-based, one step", True, GREY)])
    add_text(slide, qx1 + Inches(0.14), qy2 + Inches(0.36), Inches(5.6), Inches(0.28),
             "Built around a single task and driven entirely by the operator.", 14, False, INK2)
    pts = ["Bluebeam Revu", "PlanSwift", "RIB CostX", "RIB Candy",
           "Sage Estimating", "Causeway", "Excel templates"]
    cw2 = Inches(1.55)
    for i, name in enumerate(pts):
        chip(slide, qx1 + Inches(0.14) + (i % 4) * (cw2 + Inches(0.06)),
             qy2 + Inches(0.70) + (i // 4) * Inches(0.36),
             cw2, Inches(0.30), name)
    add_text(slide, qx1 + Inches(0.14), qy2 + Inches(1.50), Inches(5.6), Inches(0.24),
             "Strong in  measurement, pricing, drawing markup", 14, False, GREY)
    add_rich(slide, qx1 + Inches(0.14), qy2 + Inches(1.76), Inches(5.6), Inches(0.40),
             [("Where most of the market still sits, with ", True, INK),
              ("no automation and no memory.", True, RED)])

    # enterprise
    add_rect(slide, qx2, qy2, q_w, q_h, SOFT, RULE, 0.75, radius=0.08)
    add_rich(slide, qx2 + Inches(0.14), qy2 + Inches(0.08), Inches(5.6), Inches(0.30),
             [("Enterprise suites  ", True, INK), ("Rules-based, broad", True, GREY)])
    add_text(slide, qx2 + Inches(0.14), qy2 + Inches(0.36), Inches(5.6), Inches(0.28),
             "Systems of record spanning stages through modules and configuration.", 14, False, INK2)
    ents = ["Oracle", "Procore", "RIB iTWO", "Cleopatra", "SAP Ariba",
            "Autodesk ACC", "InEight", "Trimble", "Jaggaer"]
    cw3 = Inches(1.22)
    for i, name in enumerate(ents):
        chip(slide, qx2 + Inches(0.14) + (i % 5) * (cw3 + Inches(0.06)),
             qy2 + Inches(0.70) + (i // 5) * Inches(0.36),
             cw3, Inches(0.30), name)
    add_text(slide, qx2 + Inches(0.14), qy2 + Inches(1.50), Inches(5.6), Inches(0.24),
             "Strong in  cost control, scheduling, procurement, documents", 14, False, GREY)
    add_rich(slide, qx2 + Inches(0.14), qy2 + Inches(1.76), Inches(5.6), Inches(0.40),
             [("Breadth without intelligence: ", True, INK),
              ("more features, and harder to use.", True, RED)])

    # Blue Currency centre band
    band = add_rect(slide, Inches(2.15), Inches(3.30), Inches(8.95), Inches(0.88),
                    DARK, WHITE, 1.5, radius=0.5)
    add_shape_text(band, "", 14)
    add_text(slide, Inches(2.32), Inches(3.40), Inches(2.35), Inches(0.68),
             "Blue Currency", 16, True, WHITE, PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(slide, Inches(4.70), Inches(3.48), Inches(0.015), Inches(0.52), WHITE)
    add_rich(slide, Inches(4.86), Inches(3.42), Inches(6.05), Inches(0.68),
             [("Deliberately in the middle.  ", True, WHITE),
              ("Broader than the AI specialists, more AI-native than the enterprise suites.", False, RGBColor(0xDC, 0xED, 0xEA))],
             anchor="ctr")

    add_text(slide, Inches(0.52), Inches(6.68), Inches(12.1), Inches(0.28),
             "BREADTH OF LIFECYCLE COVERAGE  →", 12, True, GREY, PP_ALIGN.CENTER)

    add_notes(slide, "Page 3. Competitive landscape as four archetypes. Axes are AI-native reasoning (up) and breadth of lifecycle coverage (right). Blue Currency sits in the middle of the matrix, not in the top-right corner. Top-right (agentic plus end to end) is empty: no one is there yet. AI specialists: Kreo, Buildots, Costify, Zebel, ConWize, Ediphi, Doxel, OpenSpace, ALICE, nPlan. Point tools: Bluebeam, PlanSwift, RIB CostX, RIB Candy, Sage Estimating, Causeway, Excel. Enterprise suites: Oracle, Procore, RIB iTWO, Cleopatra, SAP Ariba, Autodesk ACC, InEight, Trimble, Jaggaer.")


def build_slide4(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, WHITE)
    kicker_title(
        slide,
        "BLUE CURRENCY  |  COMPETITIVE DEEP DIVE",
        [
            ("Module by module, legacy suites are feature-heavy and emerging AI tools are ", True, INK),
            ("smart but narrow", True, TEAL_D),
        ],
    )

    add_text(slide, Inches(0.42), Inches(1.00), Inches(6.4), Inches(0.24),
             "Badge = how much of the module that group already covers", 13, True, INK)
    for i, (col, lab) in enumerate([
        (RGBColor(0xCF, 0xE8, 0xD8), "Covers most"),
        (RGBColor(0xF5, 0xE3, 0xBF), "Partial"),
        (RGBColor(0xF0, 0xD4, 0xCF), "Barely"),
    ]):
        x = Inches(7.0) + i * Inches(1.55)
        add_rect(slide, x, Inches(1.04), Inches(0.26), Inches(0.14), col, RULE, 0.75)
        add_text(slide, x + Inches(0.32), Inches(0.98), Inches(1.18), Inches(0.24), lab, 12, False, INK2)

    headers = [
        (Inches(0.42), Inches(2.95), INK, "Module and opportunity area"),
        (Inches(3.45), Inches(4.70), INK2, "Legacy players"),
        (Inches(8.23), Inches(4.68), TEAL_D, "Emerging AI players"),
    ]
    hy = Inches(1.28)
    for x, w, fill, text in headers:
        h = add_rect(slide, x, hy, w, Inches(0.36), fill, radius=0.04)
        add_shape_text(h, text, 14, True, WHITE, PP_ALIGN.LEFT, "ctr")

    rows = [
        ("1", TEAL_D, "End to End Agentic Orchestration",
         ("WHITE SPACE", "high opportunity", INK, WHITE, RGBColor(0xAE, 0xB6, 0xBC)),
         (RED_BOX, RED, RGBColor(0xF0, 0xD4, 0xCF), RGBColor(0x9B, 0x33, 0x28), "BARELY",
          "Not agentic either: no automation and no learning across stages.",
          "RIB iTWO   ·   Oracle   ·   InEight"),
         (RED_BOX, RED, RGBColor(0xF0, 0xD4, 0xCF), RGBColor(0x9B, 0x33, 0x28), "BARELY",
          "Each tool owns a single stage; nobody governs the whole lifecycle.",
          "Kreo   ·   Buildots   ·   ConWize")),
        ("2", TEAL, "AI Generated Unified BoQ",
         ("OPEN GAP", "medium opportunity", RGBColor(0xDC, 0xE2, 0xE6), INK, GREY),
         (AMBER_BOX, HIGH, RGBColor(0xF5, 0xE3, 0xBF), RGBColor(0x8A, 0x5B, 0x14), "PARTIAL",
          "Provides structure, but no BoQ catalog creation from scratch out of historical data.",
          "Cleopatra   ·   RIB CostX   ·   Trimble"),
         (AMBER_BOX, HIGH, RGBColor(0xF5, 0xE3, 0xBF), RGBColor(0x8A, 0x5B, 0x14), "PARTIAL",
          "Document to data translation, but limited BoQ hierarchy focus.",
          "Costify   ·   Zebel   ·   ConWize")),
        ("3", TEAL, "Dynamic Bottom Up CAPEX",
         ("CROWDED", "low opportunity", RGBColor(0xF4, 0xF6, 0xF7), GREY, GREY_L),
         (GREEN_SOFT, GREEN, RGBColor(0xCF, 0xE8, 0xD8), RGBColor(0x1E, 0x6B, 0x3D), "COVERS MOST",
          "Mature estimating engines; this step is already well served.",
          "RIB CostX   ·   Cleopatra   ·   InEight"),
         (GREEN_SOFT, GREEN, RGBColor(0xCF, 0xE8, 0xD8), RGBColor(0x1E, 0x6B, 0x3D), "COVERS MOST",
          "Genuinely strong AI estimating; also well served.",
          "Zebel   ·   Kreo   ·   Ediphi")),
        ("4", TEAL, "Smart Tender Bid Evaluation",
         ("OPEN GAP", "medium opportunity", RGBColor(0xDC, 0xE2, 0xE6), INK, GREY),
         (AMBER_BOX, HIGH, RGBColor(0xF5, 0xE3, 0xBF), RGBColor(0x8A, 0x5B, 0x14), "PARTIAL",
          "Standard templates make comparison workable; the gap is AI recommendation and spec analysis.",
          "Cleopatra   ·   SAP Ariba   ·   Jaggaer"),
         (GREEN_SOFT, GREEN, RGBColor(0xCF, 0xE8, 0xD8), RGBColor(0x1E, 0x6B, 0x3D), "COVERS MOST",
          "They cover bid comparison well, so no real gap here.",
          "ConWize   ·   Zebel   ·   Ediphi")),
        ("5", TEAL, "Agentic Cost Realization Tracking",
         ("OPEN GAP", "medium opportunity", RGBColor(0xDC, 0xE2, 0xE6), INK, GREY),
         (AMBER_BOX, HIGH, RGBColor(0xF5, 0xE3, 0xBF), RGBColor(0x8A, 0x5B, 0x14), "PARTIAL",
          "Systems of record are entrenched, but recording is fully manual.",
          "Procore   ·   Oracle Primavera   ·   Autodesk ACC"),
         (AMBER_BOX, HIGH, RGBColor(0xF5, 0xE3, 0xBF), RGBColor(0x8A, 0x5B, 0x14), "PARTIAL",
          "Physical progress capture only, with no contract or IPC cost tracking.",
          "Buildots   ·   Doxel   ·   OpenSpace")),
        ("6", AMBER, "AI-Based Value Optimizer",
         ("WHITE SPACE", "high opportunity", INK, WHITE, RGBColor(0xAE, 0xB6, 0xBC)),
         (RED_BOX, RED, RGBColor(0xF0, 0xD4, 0xCF), RGBColor(0x9B, 0x33, 0x28), "BARELY",
          "Static what-if analysis at gates; no continuous savings loop anywhere.",
          "Cleopatra   ·   RIB iTWO   ·   Oracle"),
         (AMBER_BOX, HIGH, RGBColor(0xF5, 0xE3, 0xBF), RGBColor(0x8A, 0x5B, 0x14), "PARTIAL",
          "Optimization exists for schedule; cost value is left untouched.",
          "ALICE Technologies   ·   nPlan")),
    ]

    y = Inches(1.70)
    rh = Inches(0.92)
    for num, fill, name, opp, left, right in rows:
        add_rect(slide, Inches(0.42), y + rh - Inches(0.07), Inches(2.95), Inches(0.01), RULE)
        add_oval(slide, Inches(0.50), y + Inches(0.10), Inches(0.28), Inches(0.28), fill, WHITE, num, 11)
        add_text(slide, Inches(0.84), y + Inches(0.08), Inches(2.45), Inches(0.36), name, 14, True, INK)
        ol, os, ofill, ocol, oscol = opp
        badge = add_rect(slide, Inches(0.50), y + Inches(0.48), Inches(2.76), Inches(0.30), ofill, radius=0.04)
        tf = badge.text_frame
        tf.word_wrap = False
        try:
            tf._txBody.bodyPr.set("anchor", "ctr")
        except Exception:
            pass
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        tf.margin_left = Inches(0.08)
        set_run(p.add_run(), ol + "  ", 11, True, ocol)
        set_run(p.add_run(), os, 11, False, oscol)

        def camp(x, w, data):
            bg, border, bbg, bfg, blab, comment, vendors = data
            add_rect(slide, x, y + rh - Inches(0.07), w, Inches(0.01), RULE)
            # badge + comment
            add_rect(slide, x + Inches(0.14), y + Inches(0.08), Inches(1.18), Inches(0.24), bbg, radius=0.04)
            add_text(slide, x + Inches(0.14), y + Inches(0.08), Inches(1.18), Inches(0.24),
                     blab, 10, True, bfg, PP_ALIGN.CENTER)
            add_text(slide, x + Inches(1.38), y + Inches(0.06), w - Inches(1.50), Inches(0.46),
                     comment, 13, True, INK)
            add_text(slide, x + Inches(0.14), y + Inches(0.54), w - Inches(0.28), Inches(0.26),
                     vendors, 13, True, INK2)

        camp(Inches(3.45), Inches(4.70), left)
        camp(Inches(8.23), Inches(4.68), right)
        y += rh

    notes = """Page 4. Competitive deep dive. Box colour is the group's coverage of that module, not each vendor.

VENDOR EVIDENCE (from the HTML hover notes)

End to End Agentic Orchestration
• RIB iTWO: Connects estimating and controls inside one suite. Integration between modules, not an agentic learning model.
• Oracle: Spans schedule, cost and capital planning. Broad, but configured by experts rather than reasoned by agents.
• InEight: Estimating through to field execution in one platform, driven by workflow and rules rather than agents.
• Kreo: Takeoff and estimating only. Nothing carries forward into tendering, procurement or site control.
• Buildots: Site progress only. No link back to the BoQ, the budget or the procurement pipeline.
• ConWize: Tendering only. The model stops once the contract is awarded.

AI Generated Unified BoQ
• Cleopatra: Configurable cost-breakdown structures. Mapping items into it stays expert-led.
• RIB CostX: Measured takeoff and BoQs from drawings. No automatic coding of any incoming format.
• Trimble: Maintains cost libraries. Mapping third-party formats remains manual.
• Costify: Turns BoQs into a classified private database. Data, not a live BoQ engine.
• Zebel: Standardizes historical spreadsheets into a cost database, not one unified BoQ.
• ConWize: Imports and manages BoQs for tendering, limited automatic standardization.

Dynamic Bottom Up CAPEX
• RIB CostX: Bottom-up estimating from measured quantities, using maintained rates rather than live market drivers.
• Cleopatra: Estimating and benchmarking driven by rules and cost libraries.
• InEight: Cost estimating and project controls. Capable engine, still planner-driven.
• Zebel: Conceptual estimates from a firm's own history. Preconstruction focus.
• Kreo: AI takeoff from drawings through to quantities and estimates. One step, done well.
• Ediphi: Preconstruction estimating and early bid leveling, not live bottom-up CAPEX.

Smart Tender Bid Evaluation
• Cleopatra: Standardized bid comparison once data is already structured. Setup is expert-heavy.
• SAP Ariba: Sourcing and bid workflow. Generic, not construction-BoQ native.
• Jaggaer: Generic sourcing and scoring. Not built to normalize construction BoQs line by line.
• ConWize: Line-by-line BoQ bid comparison that flags anomalies. Strong at tendering only.
• Zebel: Bid leveling against historical unit prices. A preconstruction worksheet.
• Ediphi: Early-stage bid leveling, not full tender evaluation through to award.

Procurement and Progress Tracking
• Procore: Project and field management. Manual site reporting still feeds it.
• Oracle Primavera: Schedule and earned-value controls. Strong on plan, weak as a quantity gatekeeper.
• Autodesk ACC: Construction cloud for models, issues and field data. Coordination, not agentic cost control.
• Buildots: AI compares site captures against BIM and schedule. Best-in-class for that step alone.
• Doxel: Reality capture versus plan. Progress evidence, not procurement control.
• OpenSpace: 360 capture and visual progress documentation. No agentic check on quantities.

Value Optimizer
• Cleopatra: Static what-if analysis on an estimate. Not a continuous savings loop.
• RIB iTWO: Scenario comparison inside the estimate. No always-on value-engineering loop.
• Oracle: Capital planning and portfolio what-ifs, run by analysts at decision points.
• ALICE Technologies: Schedule optioneering. Time, not a cost-value loop.
• nPlan: AI schedule-risk forecasting from past programmes. Time risk, not continuous cost value.
"""
    add_notes(slide, notes)


def main():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    build_slide1(prs)
    build_slide2(prs)
    build_slide3(prs)
    build_slide4(prs)
    out = "/workspace/BlueCurrency.pptx"
    prs.save(out)
    print("wrote", out, "slides:", len(prs.slides))


if __name__ == "__main__":
    main()
