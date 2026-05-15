#!/usr/bin/env python3
"""
LeadGurus Elevate — fully-editable PowerPoint deck.
Every text, shape, table and colour is a native PowerPoint object so the
user can click any element and edit it directly in PowerPoint.

Slide size: A4 landscape (297mm × 210mm) so the deck matches the PDF.
Fonts: Georgia (serif, cross-platform) and Calibri (sans, PowerPoint default).
"""
from pptx import Presentation
from pptx.util import Mm, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

OUT = "/home/user/sally-chat/leadgurus-elevate-pitch.pptx"

# ─────────────────────────── palette ───────────────────────────
BG          = RGBColor(0x0A, 0x18, 0x28)
BG_PANEL    = RGBColor(0x14, 0x2A, 0x42)
PAPER       = RGBColor(0xFA, 0xF7, 0xF2)
BRONZE      = RGBColor(0xC0, 0x9A, 0x78)
BRONZE_SOFT = RGBColor(0xD4, 0xB8, 0x9C)
BRONZE_DEEP = RGBColor(0x8B, 0x6F, 0x50)
TEXT        = RGBColor(0xF2, 0xEE, 0xE7)
TEXT_SOFT   = RGBColor(0xB8, 0xC0, 0xCC)
TEXT_MUTE   = RGBColor(0x85, 0x93, 0xA4)
INK         = RGBColor(0x0A, 0x18, 0x28)
INK_SOFT    = RGBColor(0x4A, 0x58, 0x68)
LINE        = RGBColor(0x2A, 0x3A, 0x52)   # subtle line on dark
LINE_LIGHT  = RGBColor(0xE3, 0xDD, 0xD2)   # subtle line on paper
RED_SOFT    = RGBColor(0xE5, 0xA2, 0xA2)

SERIF = "Georgia"
SANS  = "Calibri"

# ─────────────────────────── presentation ───────────────────────────
prs = Presentation()
prs.slide_width  = Mm(297)
prs.slide_height = Mm(210)
BLANK = prs.slide_layouts[6]


# ═══════════════════════ helpers ═══════════════════════
def rect(slide, x, y, w, h, fill=None, line=None, line_w=0.5, shape=MSO_SHAPE.RECTANGLE):
    s = slide.shapes.add_shape(shape, Mm(x), Mm(y), Mm(w), Mm(h))
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(line_w)
    s.shadow.inherit = False
    return s

def hline(slide, x, y, w, color=LINE, weight=0.5):
    from pptx.enum.shapes import MSO_CONNECTOR
    ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Mm(x), Mm(y), Mm(x + w), Mm(y))
    ln.line.color.rgb = color
    ln.line.width = Pt(weight)
    return ln

def textbox(slide, x, y, w, h, paragraphs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
            line_spacing=1.15, space_before=0, space_after=0):
    """paragraphs accepted forms:
       - a single run tuple: ("text", {opts}) → one paragraph, one run
       - list of run tuples: [(t,o), (t,o)] → one paragraph, multiple runs
       - list of paragraphs: [[(t,o), (t,o)], [(t,o)]] → multi-paragraph
       - paragraph as tuple-of-runs: [((t,o), (t,o))] → one paragraph, multiple runs
    """
    tb = slide.shapes.add_textbox(Mm(x), Mm(y), Mm(w), Mm(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = 0
    tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor

    def is_run(x):
        return (isinstance(x, tuple) and len(x) == 2 and
                isinstance(x[0], str) and isinstance(x[1], dict))

    # normalize to [ [run, run, ...], [run, ...], ... ]
    if is_run(paragraphs):
        paragraphs = [[paragraphs]]
    elif isinstance(paragraphs, list) and paragraphs and is_run(paragraphs[0]):
        # all (text, opts) tuples → one paragraph
        paragraphs = [list(paragraphs)]
    else:
        # ensure each paragraph is a list of runs (convert tuples → lists)
        paragraphs = [list(p) if not isinstance(p, list) else p for p in paragraphs]

    for i, paragraph_runs in enumerate(paragraphs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        if space_before: p.space_before = Pt(space_before)
        if space_after:  p.space_after  = Pt(space_after)
        for r in list(p.runs):
            r.text = ""
        for item in paragraph_runs:
            if isinstance(item, str):
                text, opts = item, {}
            else:
                text, opts = item
            run = p.add_run()
            run.text = text
            run.font.name = opts.get("font", SANS)
            run.font.size = Pt(opts.get("size", 11))
            run.font.bold = opts.get("bold", False)
            run.font.italic = opts.get("italic", False)
            if "color" in opts:
                run.font.color.rgb = opts["color"]
            if "tracking" in opts:
                rPr = run._r.get_or_add_rPr()
                rPr.set("spc", str(int(opts["tracking"] * 100)))
    return tb

def bg(slide, color=BG):
    rect(slide, 0, 0, 297, 210, fill=color)

def header(slide, slide_no, title, on_light=False):
    """Brand top-left, slide number top-right."""
    fg = INK if on_light else TEXT
    fg_mute = INK_SOFT if on_light else TEXT_MUTE
    bz = BRONZE_DEEP if on_light else BRONZE
    # circle mark
    mark = slide.shapes.add_shape(MSO_SHAPE.OVAL, Mm(18), Mm(13.5), Mm(6), Mm(6))
    mark.fill.solid()
    mark.fill.fore_color.rgb = bz
    mark.line.fill.background()
    mark.shadow.inherit = False
    # "L" inside mark
    textbox(slide, 18, 13.6, 6, 6,
            [("L", {"font": SERIF, "size": 11, "bold": True, "color": PAPER})],
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # Lead Gurus
    textbox(slide, 25.5, 14.5, 60, 5,
            [("Lead", {"size": 11, "bold": True, "color": fg}),
             ("Gurus", {"size": 11, "bold": True, "color": bz})],
            anchor=MSO_ANCHOR.MIDDLE)
    # slide number (right)
    textbox(slide, 220, 14.5, 59, 5,
            [(slide_no.upper(), {"size": 9, "color": fg_mute, "tracking": 2.5})],
            align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

def footer(slide, right_text, on_light=False):
    line_color = LINE_LIGHT if on_light else LINE
    fg_mute = INK_SOFT if on_light else TEXT_MUTE
    hline(slide, 18, 195, 261, color=line_color, weight=0.5)
    textbox(slide, 18, 197, 130, 5,
            [("LEADGURUS ELEVATE", {"size": 8, "color": fg_mute, "tracking": 1.8})],
            anchor=MSO_ANCHOR.TOP)
    textbox(slide, 149, 197, 130, 5,
            [(right_text.upper(), {"size": 8, "color": fg_mute, "tracking": 1.8})],
            align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.TOP)

def eyebrow(slide, x, y, text, on_light=False):
    bz = BRONZE_DEEP if on_light else BRONZE
    # short rule
    hline(slide, x, y + 2.2, 6, color=bz, weight=1)
    textbox(slide, x + 7, y, 200, 5,
            [(text.upper(), {"size": 9, "color": bz, "bold": True, "tracking": 3.5})],
            anchor=MSO_ANCHOR.MIDDLE)

def h2(slide, x, y, w, h, text_runs, on_light=False):
    """text_runs is a list of (text, italic_bool) tuples."""
    fg = INK if on_light else TEXT
    bz = BRONZE_DEEP if on_light else BRONZE
    paragraph_runs = []
    for txt, ital in text_runs:
        opts = {"font": SERIF, "size": 28, "color": (bz if ital else fg), "italic": ital}
        paragraph_runs.append((txt, opts))
    textbox(slide, x, y, w, h, paragraph_runs, line_spacing=1.05)

def lede(slide, x, y, w, h, text, max_size=13, on_light=False):
    fg = INK_SOFT if on_light else TEXT_SOFT
    textbox(slide, x, y, w, h,
            [(text, {"size": max_size, "color": fg})],
            line_spacing=1.4)

def page(prs, slide_no_label, footer_right, on_light=False, bg_color=None):
    s = prs.slides.add_slide(BLANK)
    bg(s, color=(bg_color or (PAPER if on_light else BG)))
    header(s, slide_no_label, "", on_light=on_light)
    footer(s, footer_right, on_light=on_light)
    return s

def panel(slide, x, y, w, h, fill=BG_PANEL, border=LINE):
    return rect(slide, x, y, w, h, fill=fill, line=border, line_w=0.5)

def check_line(slide, x, y, w, text_runs, on_light=False, size=10.5):
    """Render '✓ text' bullet. text_runs may be a string, a list of run-tuples,
       or a list mixing strings and run-tuples."""
    bz = BRONZE_DEEP if on_light else BRONZE
    textbox(slide, x, y, 4, 5,
            ("✓", {"size": size - 1, "bold": True, "color": bz}))
    fg = INK if on_light else TEXT
    if isinstance(text_runs, str):
        text_runs = [text_runs]
    runs = []
    for item in text_runs:
        if isinstance(item, str):
            runs.append((item, {"size": size, "color": fg}))
        else:
            t, opts = item
            o = {"size": size, "color": fg, **opts}
            runs.append((t, o))
    textbox(slide, x + 4.5, y, w - 4.5, 5, runs, line_spacing=1.25)


# ═══════════════════════ SLIDE 1 · TITLE ═══════════════════════
def slide_title():
    s = page(prs, "01 · Cover", "For Small & Medium South African Brokerages")
    # eyebrow
    textbox(s, 22, 50, 260, 6,
            [("PITCH DECK · 2026", {"size": 10, "color": BRONZE, "bold": True, "tracking": 5})])
    # H1
    textbox(s, 22, 60, 260, 60,
            [[("Elevate your", {"font": SERIF, "size": 64, "color": TEXT})],
             [("brokerage.", {"font": SERIF, "size": 64, "color": BRONZE, "italic": True})]],
            line_spacing=1.0)
    # sub
    textbox(s, 22, 138, 240, 16,
            [("Four packages. One growth path. The complete back office "
              "for the brokerage that's too big for a side hustle and too "
              "small for an enterprise rollout.",
              {"size": 13, "color": TEXT_SOFT})],
            line_spacing=1.45)
    # meta row
    textbox(s, 22, 170, 240, 6,
            [(("SILVER", {"size": 9, "bold": True, "color": BRONZE, "tracking": 2.5}),
              (" · R6 999    ", {"size": 9, "color": TEXT_MUTE, "tracking": 2}),
              ("GOLD ★", {"size": 9, "bold": True, "color": BRONZE, "tracking": 2.5}),
              (" · R9 999    ", {"size": 9, "color": TEXT_MUTE, "tracking": 2}),
              ("DIAMOND", {"size": 9, "bold": True, "color": BRONZE, "tracking": 2.5}),
              (" · R13 999    ", {"size": 9, "color": TEXT_MUTE, "tracking": 2}),
              ("PLATINUM", {"size": 9, "bold": True, "color": BRONZE, "tracking": 2.5}),
              (" · R19 999", {"size": 9, "color": TEXT_MUTE, "tracking": 2}))])


# ═══════════════════════ SLIDE 2 · AUDIENCE ═══════════════════════
def slide_audience():
    s = page(prs, "02 · Audience", "Audience · 02")
    eyebrow(s, 18, 32, "Who it's for")
    h2(s, 18, 40, 260, 22,
       [("Built for the ", False), ("2–20 agent", True), (" brokerage.", False)])
    lede(s, 18, 70, 240, 16,
         "If you're scaling from a one-broker show into a real operation — "
         "doing IT, marketing, compliance and recruitment yourself — Elevate "
         "replaces six suppliers with one subscription.")

    # Left list (numbered)
    items = [
        ("01", "Independent brokers and ", "2–20 agent", " brokerages"),
        ("02", "FSPs that need to ", "look and operate professionally — fast", ""),
        ("03", "Owners juggling everything: dialler, data, hiring, the website", "", ""),
        ("04", "Anyone tired of personal Gmail and a Facebook page they call a website", "", ""),
    ]
    y = 100
    for n, pre, em, post in items:
        textbox(s, 18, y, 12, 7,
                [(n, {"font": SERIF, "size": 11, "color": BRONZE, "italic": True})])
        runs = [(pre, {"size": 11.5, "color": TEXT})]
        if em:   runs.append((em, {"size": 11.5, "color": BRONZE_SOFT, "bold": True}))
        if post: runs.append((post, {"size": 11.5, "color": TEXT}))
        textbox(s, 30, y, 130, 8, runs, line_spacing=1.35)
        hline(s, 18, y + 10, 142, color=LINE, weight=0.4)
        y += 14

    # Right stat panel
    panel(s, 175, 95, 102, 75)
    stats = [
        ("2–20", "Agents per brokerage in our sweet spot"),
        ("14",   "Days from sign-up to operational"),
        ("1",    "Invoice instead of six"),
    ]
    sy = 102
    for num, lbl in stats:
        textbox(s, 182, sy, 30, 14,
                [(num, {"font": SERIF, "size": 30, "color": BRONZE})],
                line_spacing=0.95)
        textbox(s, 215, sy + 3, 58, 14,
                [(lbl, {"size": 10, "color": TEXT_SOFT})],
                line_spacing=1.3)
        sy += 22


# ═══════════════════════ SLIDE 3 · PROPOSITION ═══════════════════════
def slide_proposition():
    s = page(prs, "03 · Proposition", "Proposition · 03")
    eyebrow(s, 18, 32, "The proposition")
    h2(s, 18, 40, 260, 22,
       [("One subscription. ", False), ("The whole back office.", True)])
    lede(s, 18, 70, 240, 12,
         "Every Elevate package bundles six things that normally come from "
         "six different vendors — under one contract, with one support line.")

    pillars = [
        ("i.",   "Leads",         "Compliant consumer and director records, delivered monthly into your CRM."),
        ("ii.",  "Airtime",       "Per-minute business rates from R0.28c — no contract lock-in on your numbers."),
        ("iii.", "Telephony",     "VoIP landlines plus a business cell number — call recording included."),
        ("iv.",  "Web presence",  "Site, branded email, Microsoft 365 licences and organic social."),
        ("v.",   "Brand assets",  "Digital business card and a polished company profile PDF."),
        ("vi.",  "Support",       "Named SLA, response time in writing, real humans on the line."),
    ]
    cw, ch = 84, 44
    gap = 4
    start_x, start_y = 18, 100
    for idx, (n, title, body) in enumerate(pillars):
        col = idx % 3
        row = idx // 3
        x = start_x + col * (cw + gap)
        y = start_y + row * (ch + gap)
        panel(s, x, y, cw, ch)
        textbox(s, x + 6, y + 4, 20, 5,
                [(n, {"font": SERIF, "size": 11, "color": BRONZE, "italic": True})])
        textbox(s, x + 6, y + 10, cw - 12, 7,
                [(title, {"size": 13, "bold": True, "color": TEXT})])
        textbox(s, x + 6, y + 19, cw - 12, 20,
                [(body, {"size": 10, "color": TEXT_SOFT})],
                line_spacing=1.35)


# ═══════════════════════ SLIDE 4 · TIER OVERVIEW ═══════════════════════
def slide_tier_overview():
    s = page(prs, "04 · The Packages", "The Packages · 04")
    eyebrow(s, 18, 32, "The packages")
    h2(s, 18, 40, 260, 22,
       [("Four tiers. ", False), ("One per growth stage.", True)])
    lede(s, 18, 70, 240, 10,
         "Start where you are. Upgrade when the team grows. No long-term lock-in.")

    tiers = [
        ("Silver",   "R6 999",  "Solo · Just opened",
         ["500 records / month", "R0.35c / min airtime", "1–3 page template site",
          "1 landline · 1 cell", "5 mailboxes · MS 365 ×1", "4 social posts / month",
          "Same-day support"]),
        ("Gold",     "R9 999",  "2–5 agents · Growing",
         ["1 000 records / month", "R0.32c / min airtime", "3–5 page template site",
          "1 landline · 1 cell", "10 mailboxes · MS 365 ×2", "8 social posts / month",
          "4-hour support"]),
        ("Diamond",  "R13 999", "5–10 agents · Established",
         ["1 000 records / month", "R0.30c / min airtime", "5–6 page custom site",
          "3 landlines · 1 cell", "10 mailboxes · MS 365 ×5", "10 social posts / month",
          "2-hour support"]),
        ("Platinum", "R19 999", "10+ agents · B2B",
         ["500 director records / mo", "R0.28c / min airtime", "5–7 page / e-commerce",
          "3 landlines · 1 cell", "10–20 mailboxes · MS 365 ×10", "10 social posts / month",
          "2-hour VIP support"]),
    ]
    cw = 62
    gap = 3
    start_x, start_y = 18, 90
    ch = 96
    for idx, (name, price, tag, feats) in enumerate(tiers):
        featured = (name == "Gold")
        x = start_x + idx * (cw + gap)
        y = start_y
        if featured:
            panel(s, x, y, cw, ch, fill=PAPER, border=BRONZE)
        else:
            panel(s, x, y, cw, ch)

        # ribbon for featured
        if featured:
            rb = rect(s, x + 13, y - 2.5, cw - 26, 5, fill=BRONZE)
            textbox(s, x + 13, y - 2.5, cw - 26, 5,
                    [("★ RECOMMENDED", {"size": 7, "bold": True, "color": PAPER, "tracking": 2})],
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        name_color = INK if featured else TEXT
        price_color = INK if featured else TEXT
        tag_color = INK_SOFT if featured else TEXT_MUTE

        textbox(s, x + 5, y + (8 if featured else 5), cw - 10, 9,
                [(name, {"font": SERIF, "size": 20, "color": name_color})])
        textbox(s, x + 5, y + 19, cw - 10, 9,
                [((price, {"font": SANS, "size": 18, "bold": True, "color": price_color}),
                  (" /month", {"size": 8, "color": tag_color}))])
        textbox(s, x + 5, y + 30, cw - 10, 5,
                [(tag.upper(), {"size": 8, "color": tag_color, "tracking": 1.5})])

        # divider
        div_color = LINE_LIGHT if featured else LINE
        hline(s, x + 5, y + 37, cw - 10, color=div_color)

        # features
        fy = y + 40
        for f in feats:
            # bullet dash
            hline(s, x + 5, fy + 1.5, 2.5, color=(BRONZE_DEEP if featured else BRONZE), weight=1)
            textbox(s, x + 9, fy, cw - 14, 6,
                    [(f, {"size": 9, "color": (INK if featured else TEXT)})],
                    line_spacing=1.25)
            fy += 6.8


# ═══════════════════════ SLIDES 5–8 · TIER DETAIL ═══════════════════════
def slide_tier_detail(name, price, tagline, who, features, slide_no, on_light=False, featured=False):
    s = page(prs, slide_no, f"{name} · {slide_no.split(' · ')[0]}", on_light=on_light)
    eyebrow(s, 18, 32, slide_no.split(' · ')[1] + (" · ★ Recommended" if featured else ""), on_light=on_light)

    # Left: name + price + tagline + who
    if featured:
        # ribbon
        rb = rect(s, 18, 45, 38, 5, fill=BRONZE_DEEP if on_light else BRONZE)
        textbox(s, 18, 45, 38, 5,
                [("★ RECOMMENDED", {"size": 7.5, "bold": True, "color": PAPER, "tracking": 2})],
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        ny = 54
    else:
        ny = 45

    textbox(s, 18, ny, 100, 14,
            [(name + ".", {"font": SERIF, "size": 36, "color": (INK if on_light else TEXT)})])

    # price block
    bz = BRONZE_DEEP if on_light else BRONZE
    textbox(s, 18, ny + 18, 70, 22,
            [(price, {"font": SERIF, "size": 48, "color": bz})],
            line_spacing=0.95)
    textbox(s, 18, ny + 41, 70, 5,
            [("PER MONTH", {"size": 9, "color": (INK_SOFT if on_light else TEXT_SOFT), "tracking": 2.5})])

    # tagline
    textbox(s, 18, ny + 52, 95, 16,
            [(f'"{tagline}"', {"font": SERIF, "size": 14,
                               "color": (BRONZE_DEEP if on_light else BRONZE_SOFT),
                               "italic": True})],
            line_spacing=1.3)

    # who block
    hline(s, 18, 165, 95, color=(LINE_LIGHT if on_light else LINE), weight=0.4)
    textbox(s, 18, 168, 95, 22,
            [(("Best for: ", {"size": 10, "bold": True, "color": bz}),
              (who, {"size": 10, "color": (INK_SOFT if on_light else TEXT_SOFT)}))],
            line_spacing=1.45)

    # Right: features panel
    panel_fg = RGBColor(0xFF, 0xFD, 0xF9) if on_light else BG_PANEL
    panel_border = LINE_LIGHT if on_light else LINE
    panel(s, 130, 45, 148, 140, fill=panel_fg, border=panel_border)

    # 12 features in 2 cols × 6 rows
    col_w = 70
    row_h = 10.5
    fx0 = 135
    fy0 = 52
    for idx, feat in enumerate(features):
        col = idx % 2
        row = idx // 2
        x = fx0 + col * col_w
        y = fy0 + row * row_h
        check_line(s, x, y, col_w - 4, feat, on_light=on_light, size=10)
        if row < 5:
            hline(s, x, y + row_h - 2.2, col_w - 4,
                  color=(LINE_LIGHT if on_light else LINE), weight=0.3)


# ═══════════════════════ SLIDE 9 · 60-SECOND TEST ═══════════════════════
def slide_pick():
    s = page(prs, "09 · Decision Guide", "Decision Guide · 09")
    eyebrow(s, 18, 32, "How to pick")
    h2(s, 18, 40, 260, 22,
       [("The ", False), ("60-second test.", True)])
    lede(s, 18, 70, 240, 10, "Three questions. One tier.")

    picks = [
        ("i.",   "How many agents are on the phones today?",
         [("1 agent", "Silver"), ("2–5 agents", "Gold"),
          ("5–10 agents", "Diamond"), ("10+ agents", "Platinum")]),
        ("ii.",  "Who are you calling?",
         [("Consumers", "Silver–Diamond"), ("Directors / B2B", "Platinum"),
          ("Mixed book", "Diamond")]),
        ("iii.", "Do you need to sell online — not just be online?",
         [("No", "Any tier"), ("Yes, e-commerce", "Platinum"),
          ("Maybe later", "Diamond")]),
    ]
    cw = 84
    gap = 4
    start_x, start_y = 18, 88
    ch = 95
    for idx, (n, q, opts) in enumerate(picks):
        x = start_x + idx * (cw + gap)
        panel(s, x, start_y, cw, ch)
        textbox(s, x + 6, start_y + 6, 30, 12,
                [(n, {"font": SERIF, "size": 30, "color": BRONZE, "italic": True})],
                line_spacing=0.9)
        textbox(s, x + 6, start_y + 22, cw - 12, 18,
                [(q, {"size": 11.5, "bold": True, "color": TEXT})],
                line_spacing=1.35)
        oy = start_y + 48
        for left_t, right_t in opts:
            hline(s, x + 6, oy - 0.5, cw - 12, color=LINE, weight=0.4)
            textbox(s, x + 6, oy, (cw - 12) * 0.55, 5,
                    [(left_t, {"size": 9.5, "color": TEXT_SOFT})],
                    anchor=MSO_ANCHOR.MIDDLE)
            textbox(s, x + 6 + (cw - 12) * 0.45, oy, (cw - 12) * 0.55, 5,
                    [(right_t.upper(), {"size": 9, "bold": True, "color": BRONZE, "tracking": 1.5})],
                    align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
            oy += 8


# ═══════════════════════ SLIDE 10 · COMPARISON TABLE ═══════════════════════
def slide_comparison():
    s = page(prs, "10 · Comparison", "Comparison · 10")
    eyebrow(s, 18, 32, "Side by side")
    h2(s, 18, 40, 260, 22,
       [("What changes ", False), ("between tiers.", True)])

    # Build a native PowerPoint table
    rows = [
        ("",                     "Silver",         "Gold ★",        "Diamond",       "Platinum"),
        ("Price / month",        "R6 999",         "R9 999",        "R13 999",       "R19 999"),
        ("Records / month",      "500",            "1 000",         "1 000",         "500 Director"),
        ("Airtime / min",        "R0.35c",         "R0.32c",        "R0.30c",        "R0.28c"),
        ("Website",              "1–3 pg template","3–5 pg template","5–6 pg custom","5–7 pg / e-commerce"),
        ("Email mailboxes",      "5",              "10",            "10",            "10–20"),
        ("Microsoft 365 licences","1",             "2",             "5",             "10"),
        ("VoIP landlines",       "1",              "1",             "3",             "3"),
        ("Business cell",        "1",              "1",             "1",             "1"),
        ("Social posts / month", "4",              "8",             "10",            "10"),
        ("Support response",     "Same day",       "4 hour",        "2 hour",        "2 hour VIP"),
    ]
    n_rows = len(rows)
    n_cols = 5
    tbl_x, tbl_y, tbl_w, tbl_h = 18, 80, 261, 105
    table_shape = s.shapes.add_table(n_rows, n_cols, Mm(tbl_x), Mm(tbl_y), Mm(tbl_w), Mm(tbl_h))
    table = table_shape.table

    # Column widths
    col_widths_mm = [55, 50, 52, 52, 52]
    for i, w in enumerate(col_widths_mm):
        table.columns[i].width = Mm(w)
    # Row heights
    for r in table.rows:
        r.height = Mm(105 / n_rows)

    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            cell = table.cell(ri, ci)
            # background
            cell.fill.solid()
            if ri == 0:
                if ci == 2:
                    cell.fill.fore_color.rgb = BRONZE
                else:
                    cell.fill.fore_color.rgb = BG_PANEL
            else:
                if ci == 2:
                    # subtle bronze tint
                    cell.fill.fore_color.rgb = RGBColor(0x1F, 0x2E, 0x3F)
                else:
                    cell.fill.fore_color.rgb = BG_PANEL
            cell.margin_left = Mm(3)
            cell.margin_right = Mm(3)
            cell.margin_top = Mm(1.4)
            cell.margin_bottom = Mm(1.4)

            tf = cell.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            for r in list(p.runs):
                r.text = ""
            run = p.add_run()
            run.text = val
            run.font.name = SANS

            if ri == 0:
                run.font.size = Pt(9)
                run.font.bold = True
                run.font.color.rgb = PAPER if ci == 2 else BRONZE
                # uppercase tracking
                rPr = run._r.get_or_add_rPr()
                rPr.set("spc", "150")
                run.text = val.upper()
            else:
                run.font.size = Pt(10.5)
                if ci == 0:
                    run.font.color.rgb = TEXT_SOFT
                elif ci == 2:
                    run.font.bold = True
                    run.font.color.rgb = BRONZE_SOFT
                else:
                    run.font.color.rgb = TEXT

    # Style table borders by writing the XML (must come BEFORE solidFill per OOXML schema)
    tbl_el = table._tbl
    for cell_el in tbl_el.iter(qn('a:tc')):
        tcPr = cell_el.find(qn('a:tcPr'))
        if tcPr is None:
            tcPr = etree.SubElement(cell_el, qn('a:tcPr'))
        # remove any existing borders
        for side in ('lnL', 'lnR', 'lnT', 'lnB'):
            for el in tcPr.findall(qn(f'a:{side}')):
                tcPr.remove(el)
        # find where solidFill is (borders must precede fill in schema order)
        fill_el = tcPr.find(qn('a:solidFill'))
        insert_at = list(tcPr).index(fill_el) if fill_el is not None else len(tcPr)
        for i, side in enumerate(('lnL', 'lnR', 'lnT', 'lnB')):
            ln = etree.Element(qn(f'a:{side}'))
            ln.set('w', '6350')   # 0.5pt
            fill = etree.SubElement(ln, qn('a:solidFill'))
            clr = etree.SubElement(fill, qn('a:srgbClr'))
            clr.set('val', '20304A')
            tcPr.insert(insert_at + i, ln)


# ═══════════════════════ SLIDE 11 · ADD-ONS ═══════════════════════
def slide_addons():
    s = page(prs, "11 · Operational Add-Ons", "Operational Add-Ons · 11")
    eyebrow(s, 18, 32, "Beyond the package")
    h2(s, 18, 40, 260, 22,
       [("When you outgrow the box — ", False), ("plug in.", True)])
    lede(s, 18, 70, 240, 10,
         "Elevate is the subscription. These are the accelerators you bolt on as the brokerage scales.")

    addons = [
        ("Duplicate Checker",        "Per import — clean lists before dialling",       "R750"),
        ("Number Rotation & Anti-Spam","Save numbers from spam flagging",              "R1 500"),
        ("Scripts & Training",       "Compliance-ready call scripts",                  "R650"),
        ("Vici Dial CRM Setup",      "Campaigns · dispositions · users",               "Quoted"),
        ("Admin / IT Support",       "Monthly retainer · same-day response",           "R3 000 / m"),
        ("Hourly Consulting",        "Agent & admin coaching on tap",                  "R650 / hr"),
        ("Half-Day Training",        "Up to 5 admins per session",                     "R8 000"),
        ("Full-Day Training",        "Up to 5 admins · deep dive",                     "R12 000"),
        ("Recruitment Support",      "Per role · sourcing & screening",                "Quoted"),
    ]
    cw = 84
    ch = 28
    gap_x = 4
    gap_y = 4
    start_x, start_y = 18, 88
    for idx, (name, sub, price) in enumerate(addons):
        col = idx % 3
        row = idx // 3
        x = start_x + col * (cw + gap_x)
        y = start_y + row * (ch + gap_y)
        panel(s, x, y, cw, ch)
        textbox(s, x + 5, y + 6, cw - 30, 6,
                [(name, {"size": 10.5, "bold": True, "color": TEXT})])
        textbox(s, x + 5, y + 13, cw - 30, 11,
                [(sub, {"size": 9, "color": TEXT_MUTE})],
                line_spacing=1.3)
        textbox(s, x + cw - 28, y + 9, 24, 8,
                [(price, {"font": SERIF, "size": 13, "bold": True, "color": BRONZE})],
                align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)


# ═══════════════════════ SLIDE 12 · DIY VS ELEVATE ═══════════════════════
def slide_maths():
    s = page(prs, "12 · The Maths", "The Maths · 12")
    eyebrow(s, 18, 32, "The maths")
    h2(s, 18, 40, 260, 22,
       [("DIY costs more — ", False), ("and gives you six suppliers.", True)])

    rows = [
        ("Domain + 5-page website (freelancer)",        "R1 500 – R3 500", "normal"),
        ("Email hosting + Microsoft 365 (5 seats)",     "R1 200",           "normal"),
        ("VoIP landlines (3)",                          "R900",             "normal"),
        ("Part-time social media manager",              "R3 500",           "normal"),
        ("Lead data (1 000 records)",                   "R3 000+",          "normal"),
        ("Airtime allocation",                          "R1 500 – R2 500", "normal"),
        ("DIY total",                                   "R11 600 – R14 600","diy"),
        ("Diamond — everything, one invoice",           "R13 999",          "elv"),
    ]
    header_row = ("Line item (build it yourself, 5-agent brokerage)", "Realistic monthly cost")

    n_rows = len(rows) + 1
    tbl_x, tbl_y, tbl_w, tbl_h = 18, 78, 261, 90
    table_shape = s.shapes.add_table(n_rows, 2, Mm(tbl_x), Mm(tbl_y), Mm(tbl_w), Mm(tbl_h))
    table = table_shape.table
    table.columns[0].width = Mm(180)
    table.columns[1].width = Mm(81)
    for r in table.rows:
        r.height = Mm(tbl_h / n_rows)

    # header
    for ci, val in enumerate(header_row):
        cell = table.cell(0, ci)
        cell.fill.solid()
        cell.fill.fore_color.rgb = BG_PANEL
        cell.margin_left = Mm(5); cell.margin_right = Mm(5)
        cell.margin_top = Mm(2); cell.margin_bottom = Mm(2)
        cell.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = cell.text_frame.paragraphs[0]
        if ci == 1: p.alignment = PP_ALIGN.RIGHT
        for r in list(p.runs): r.text = ""
        run = p.add_run()
        run.text = val.upper()
        run.font.name = SANS
        run.font.size = Pt(9)
        run.font.bold = True
        run.font.color.rgb = BRONZE
        rPr = run._r.get_or_add_rPr()
        rPr.set("spc", "180")

    for ri, (label, val, kind) in enumerate(rows, start=1):
        for ci, txt in enumerate((label, val)):
            cell = table.cell(ri, ci)
            cell.fill.solid()
            if kind == "diy":
                cell.fill.fore_color.rgb = RGBColor(0x2A, 0x1F, 0x22)
            elif kind == "elv":
                cell.fill.fore_color.rgb = RGBColor(0x24, 0x2C, 0x36)
            else:
                cell.fill.fore_color.rgb = BG_PANEL
            cell.margin_left = Mm(5); cell.margin_right = Mm(5)
            cell.margin_top = Mm(2); cell.margin_bottom = Mm(2)
            cell.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            if ci == 1: p.alignment = PP_ALIGN.RIGHT
            for r in list(p.runs): r.text = ""
            run = p.add_run()
            run.text = txt
            run.font.name = SANS
            run.font.size = Pt(12 if kind == "elv" else 11)
            if kind == "diy":
                run.font.bold = True
                run.font.color.rgb = RED_SOFT
            elif kind == "elv":
                run.font.bold = True
                run.font.color.rgb = BRONZE_SOFT
            else:
                run.font.color.rgb = TEXT

    # subtle table borders (schema-correct: borders before fill)
    tbl_el = table._tbl
    for cell_el in tbl_el.iter(qn('a:tc')):
        tcPr = cell_el.find(qn('a:tcPr'))
        if tcPr is None:
            tcPr = etree.SubElement(cell_el, qn('a:tcPr'))
        for side in ('lnL', 'lnR', 'lnT', 'lnB'):
            for el in tcPr.findall(qn(f'a:{side}')):
                tcPr.remove(el)
        fill_el = tcPr.find(qn('a:solidFill'))
        insert_at = list(tcPr).index(fill_el) if fill_el is not None else len(tcPr)
        for i, side in enumerate(('lnL', 'lnR', 'lnT', 'lnB')):
            ln = etree.Element(qn(f'a:{side}'))
            ln.set('w', '6350')
            fill = etree.SubElement(ln, qn('a:solidFill'))
            clr = etree.SubElement(fill, qn('a:srgbClr'))
            clr.set('val', '20304A')
            tcPr.insert(insert_at + i, ln)

    # caption
    textbox(s, 18, 175, 261, 14,
            [(("…and DIY gives you ", {"size": 10.5, "color": TEXT_SOFT}),
              ("six suppliers, six invoices, six support queues", {"size": 10.5, "bold": True, "color": TEXT}),
              (". Elevate gives you one — at the same price as the middle of the DIY range.",
               {"size": 10.5, "color": TEXT_SOFT}))],
            line_spacing=1.4)


# ═══════════════════════ SLIDE 13 · POPIA ═══════════════════════
def slide_popia():
    s = page(prs, "13 · Compliance", "Compliance · 13")
    eyebrow(s, 18, 32, "Compliance")
    h2(s, 18, 40, 260, 22,
       [("POPIA & compliance ", False), ("built in.", True)])
    lede(s, 18, 70, 240, 12,
         "Every Elevate package ships with the compliance scaffolding small "
         "FSPs typically forget to set up. You stop being one DPA complaint "
         "away from a problem.")

    cards = [
        ("i.",   "Compliant lead data",
         "Records sourced under POPIA. Opt-out workflow handled by us — your agents never touch a non-compliant list."),
        ("ii.",  "Call recording trail",
         "Business numbers and VoIP infrastructure give you a defensible record of every consent on every call."),
        ("iii.", "Branded email identity",
         "No more Gmail. A proper @yourbrokerage.co.za address is the foundation of a legitimate Information Officer footprint."),
        ("iv.",  "Consistent disclosure",
         "Digital business card and company profile PDF mean every client sees the same regulator-ready disclosure."),
    ]
    cw, ch = 128, 38
    gap = 4
    start_x, start_y = 18, 100
    for idx, (n, title, body) in enumerate(cards):
        col = idx % 2
        row = idx // 2
        x = start_x + col * (cw + gap)
        y = start_y + row * (ch + gap)
        panel(s, x, y, cw, ch)
        # icon circle
        ic = slide_circle(s, x + 6, y + 6, 12)
        textbox(s, x + 6, y + 6, 12, 12,
                [(n, {"font": SERIF, "size": 13, "color": BRONZE, "italic": True})],
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # text
        textbox(s, x + 22, y + 6, cw - 28, 7,
                [(title, {"size": 12, "bold": True, "color": BRONZE_SOFT})])
        textbox(s, x + 22, y + 14, cw - 28, 22,
                [(body, {"size": 9.5, "color": TEXT_SOFT})],
                line_spacing=1.4)


def slide_circle(slide, x, y, d):
    c = slide.shapes.add_shape(MSO_SHAPE.OVAL, Mm(x), Mm(y), Mm(d), Mm(d))
    c.fill.solid()
    c.fill.fore_color.rgb = RGBColor(0x1F, 0x2E, 0x3F)
    c.line.color.rgb = BRONZE
    c.line.width = Pt(0.5)
    c.shadow.inherit = False
    return c


# ═══════════════════════ SLIDE 14 · ROADMAP ═══════════════════════
def slide_roadmap():
    s = page(prs, "14 · Growth Path", "Growth Path · 14")
    eyebrow(s, 18, 32, "The growth path")
    h2(s, 18, 40, 260, 22,
       [("Silver to Platinum — ", False), ("at your pace.", True)])
    lede(s, 18, 70, 240, 12,
         "A realistic upgrade path for a small brokerage. You're not locked "
         "in — move up or down as the team flexes.")

    steps = [
        ("Month 1–3",   "Silver",   "Get listed. Get found. Look legitimate from day one."),
        ("Month 4–9",   "Gold",     "Hire 2–5 agents. Double the leads. Step up output."),
        ("Month 10–18", "Diamond",  "Customised brand. Multi-line call operations."),
        ("Month 18+",   "Platinum", "Move upmarket — director sales, online checkout."),
    ]
    cw = 60
    gap = 5
    start_x = 18 + (261 - (cw * 4 + gap * 3)) / 2
    line_y = 118
    # horizontal connecting line
    hline(s, start_x + cw / 2, line_y, cw * 3 + gap * 3, color=BRONZE, weight=1.2)

    for idx, (when, tier, desc) in enumerate(steps):
        x = start_x + idx * (cw + gap)
        # dot
        dot = slide.shapes.add_shape if False else None  # placeholder
        d = s.shapes.add_shape(MSO_SHAPE.OVAL, Mm(x + cw / 2 - 2), Mm(line_y - 2), Mm(4), Mm(4))
        d.fill.solid(); d.fill.fore_color.rgb = BRONZE
        d.line.color.rgb = BG; d.line.width = Pt(2)
        d.shadow.inherit = False

        textbox(s, x, line_y + 8, cw, 6,
                [(when.upper(), {"size": 9, "bold": True, "color": BRONZE, "tracking": 2})],
                align=PP_ALIGN.CENTER)
        textbox(s, x, line_y + 16, cw, 12,
                [(tier, {"font": SERIF, "size": 24, "color": TEXT})],
                align=PP_ALIGN.CENTER, line_spacing=1)
        textbox(s, x, line_y + 32, cw, 24,
                [(desc, {"size": 10, "color": TEXT_SOFT})],
                align=PP_ALIGN.CENTER, line_spacing=1.4)


# ═══════════════════════ SLIDE 15 · CTA ═══════════════════════
def slide_cta():
    s = page(prs, "15 · Next Steps", "Next Steps · 15")
    eyebrow(s, 18, 32, "Next steps")
    h2(s, 18, 40, 260, 22,
       [("Pick a tier. ", False), ("Be live in 14 days.", True)])

    # gradient-feel CTA card (solid bronze, no gradient in pptx easily)
    rect(s, 18, 80, 261, 90, fill=BRONZE)
    # subtle inner darker shadow
    rect(s, 18, 80, 261, 90, fill=None, line=BRONZE_DEEP, line_w=1)

    textbox(s, 32, 92, 230, 16,
            [(("Book your 15-minute ", {"font": SERIF, "size": 28, "color": PAPER}),
              ("discovery call.", {"font": SERIF, "size": 28, "color": INK, "italic": True}))],
            line_spacing=1.1)
    textbox(s, 32, 116, 230, 18,
            [("On the call we'll recommend a tier, scope your onboarding, "
              "and have your brokerage operational within 7–14 working days.",
              {"size": 12, "color": PAPER})],
            line_spacing=1.45)

    # divider
    hline(s, 32, 142, 232, color=RGBColor(0xE9, 0xD8, 0xC4), weight=0.5)

    contacts = [("PHONE", "+27 [number]"), ("EMAIL", "info@leadgurus.co.za"), ("WEB", "leadgurus.co.za")]
    cx = 32
    for label, val in contacts:
        textbox(s, cx, 148, 60, 5,
                [(label, {"size": 8, "bold": True, "color": PAPER, "tracking": 2.5})])
        textbox(s, cx, 154, 80, 8,
                [(val, {"size": 12, "bold": True, "color": PAPER})])
        cx += 75

    # tagline
    textbox(s, 18, 178, 261, 6,
            [(("SILVER · ", {"size": 9.5, "color": TEXT_MUTE, "tracking": 2.5}),
              ("GOLD ★", {"size": 9.5, "bold": True, "color": BRONZE, "tracking": 2.5}),
              (" · DIAMOND · PLATINUM    —    CHOOSE YOUR TIER.",
               {"size": 9.5, "color": TEXT_MUTE, "tracking": 2.5}))],
            align=PP_ALIGN.CENTER)


# ═══════════════════════ build ═══════════════════════
slide_title()
slide_audience()
slide_proposition()
slide_tier_overview()
slide_tier_detail(
    "Silver", "R6 999", "Look like a real business by Monday.",
    "the newly-licensed broker who needs a credible web presence, a business "
    "number, and lead flow — without a three-month procurement project.",
    ["500 records / month", "R0.35c / min airtime",
     "1–3 page template website", "Email hosting · 5 mailboxes",
     "Microsoft 365 · 1 licence", "Digital business card",
     "Company profile (PDF)", "1 VoIP landline number",
     "1 business cell number", "Organic social media setup",
     "4 social posts / month", "Same-day support response"],
    "05 · Silver", on_light=False)

slide_tier_detail(
    "Gold", "R9 999", "Double the leads, double the output, half the headache.",
    "2–5 agent brokerages that need volume without going custom yet. The "
    "sweet spot of price vs. capability for most SMB FSPs.",
    [[("1 000 records / month", {"bold": True})],
     "R0.32c / min airtime",
     "3–5 page template website", "Email hosting · 10 mailboxes",
     "Microsoft 365 · 2 licences", "Digital business card",
     "Company profile (PDF)", "1 VoIP landline number",
     "1 business cell number", "Organic social media setup",
     [("8 social posts / month", {"bold": True})],
     "4-hour support response"],
    "06 · Gold", on_light=True, featured=True)

slide_tier_detail(
    "Diamond", "R13 999", "Your own brand identity and the telephony muscle to back it.",
    "5–10 agent brokerages that need their own brand identity and the "
    "telephony muscle to run multiple campaigns simultaneously.",
    ["1 000 records / month", "R0.30c / min airtime",
     [("5–6 page customised site", {"bold": True})],
     "Email hosting · 10 mailboxes",
     [("Microsoft 365 · 5 licences", {"bold": True})],
     "Digital business card",
     "Company profile (PDF)",
     [("3 VoIP landline numbers", {"bold": True})],
     "1 business cell number", "Organic social media setup",
     "10 social posts / month",
     [("2-hour support response", {"bold": True})]],
    "07 · Diamond")

slide_tier_detail(
    "Platinum", "R19 999", "Director-level data and a site that sells — built for B2B.",
    "brokerages targeting business owners and directors, or running policy "
    "and product sales online with checkout.",
    [[("500 director-level records / mo", {"bold": True})],
     "R0.28c / min airtime (lowest)",
     [("5–7 page custom / e-commerce", {"bold": True})],
     "Email hosting · 10–20 mailboxes",
     "Microsoft 365 · 10 licences",
     "Digital business card",
     "Company profile (PDF)",
     "3 VoIP landline numbers",
     "1 business cell number",
     "Organic social media setup",
     "10 social posts / month",
     [("2-hour VIP support — jump the queue", {"bold": True})]],
    "08 · Platinum")

slide_pick()
slide_comparison()
slide_addons()
slide_maths()
slide_popia()
slide_roadmap()
slide_cta()

prs.save(OUT)
import os
print(f"✓ {OUT}  ·  {os.path.getsize(OUT)//1024} KB  ·  {len(prs.slides)} slides")
