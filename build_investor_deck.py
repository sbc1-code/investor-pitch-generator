#!/usr/bin/env python3
"""
Investor Pitch Deck Generator
Generates a professional investor pitch deck for industrial real estate funds
using python-pptx. Configurable branding, data-driven slides, institutional formatting.

Usage:
    python3 build_investor_deck.py

Requires:
    pip install python-pptx
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import os

# === BRAND CONSTANTS (customize these for your fund) ===
ACCENT = RGBColor(0x8A, 0x9A, 0x7B)       # #8A9A7B - Primary accent
SECONDARY = RGBColor(0xA8, 0xB8, 0x9A)    # Secondary accent
DARK_SLATE = RGBColor(0x34, 0x3D, 0x46)   # #343D46 - Primary text
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
WARM_GRAY = RGBColor(0xDC, 0xD5, 0xCF)    # #DCD5CF - Backgrounds
LIGHT_BG = RGBColor(0xF5, 0xF3, 0xF0)     # Light warm background
FONT = "Inter"

# Slide dimensions - 16:9
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

# Asset paths (place your logo files in ./assets/)
ASSETS = "./assets"
LOGO_DARK = os.path.join(ASSETS, "logo_dark.png")
LOGO_WHITE = os.path.join(ASSETS, "logo_white.png")
ICON_ACCENT = os.path.join(ASSETS, "icon_accent.png")
ICON_CIRCLE = os.path.join(ASSETS, "icon_circle.png")

OUTPUT = "./output/Investor_Pitch_Deck.pptx"


def add_accent_bar(slide, top=0):
    """Add the signature thin accent bar at top of slide"""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), top, SLIDE_W, Inches(0.15)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = ACCENT
    shape.line.fill.background()


def add_logo_watermark(slide):
    """Add small icon watermark to bottom-right"""
    if os.path.exists(ICON_ACCENT):
        slide.shapes.add_picture(
            ICON_ACCENT,
            SLIDE_W - Inches(1.2), SLIDE_H - Inches(0.7),
            Inches(0.8), Inches(0.5)
        )


def add_section_label(slide, text, left, top):
    """Add small caps section label in accent color (e.g., 'THE OPPORTUNITY')"""
    txBox = slide.shapes.add_textbox(left, top, Inches(4), Inches(0.4))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text.upper()
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = ACCENT
    p.font.name = FONT


def add_title(slide, text, left, top, width=Inches(10), size=Pt(36), color=DARK_SLATE):
    """Add main slide title"""
    txBox = slide.shapes.add_textbox(left, top, width, Inches(1))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = size
    p.font.bold = True
    p.font.color.rgb = color
    p.font.name = FONT
    return txBox


def add_body(slide, text, left, top, width=Inches(10), height=Inches(3), size=Pt(14), color=DARK_SLATE):
    """Add body text"""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = size
    p.font.color.rgb = color
    p.font.name = FONT
    p.line_spacing = Pt(22)
    return tf


def add_bullet_list(slide, items, left, top, width=Inches(10), height=Inches(4), size=Pt(14), color=DARK_SLATE):
    """Add a bulleted list"""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = item
        p.font.size = size
        p.font.color.rgb = color
        p.font.name = FONT
        p.line_spacing = Pt(24)
        p.space_after = Pt(6)
        # Bullet character
        p.text = "\u2022  " + item
    return tf


def add_stat_box(slide, number, label, left, top, width=Inches(2.5), num_color=ACCENT):
    """Add a big stat with label below"""
    # Number
    txBox = slide.shapes.add_textbox(left, top, width, Inches(0.8))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = number
    p.font.size = Pt(48)
    p.font.bold = True
    p.font.color.rgb = num_color
    p.font.name = FONT
    p.alignment = PP_ALIGN.LEFT

    # Label
    txBox2 = slide.shapes.add_textbox(left, top + Inches(0.75), width, Inches(0.5))
    tf2 = txBox2.text_frame
    p2 = tf2.paragraphs[0]
    p2.text = label
    p2.font.size = Pt(12)
    p2.font.color.rgb = DARK_SLATE
    p2.font.name = FONT
    p2.alignment = PP_ALIGN.LEFT


def add_kv_row(tf, key, value, key_color=ACCENT, val_color=DARK_SLATE):
    """Add a key-value row to existing text frame"""
    p = tf.add_paragraph()
    run_k = p.add_run()
    run_k.text = key + ": "
    run_k.font.size = Pt(13)
    run_k.font.bold = True
    run_k.font.color.rgb = key_color
    run_k.font.name = FONT

    run_v = p.add_run()
    run_v.text = value
    run_v.font.size = Pt(13)
    run_v.font.color.rgb = val_color
    run_v.font.name = FONT
    p.line_spacing = Pt(22)
    p.space_after = Pt(4)


# === BUILD PRESENTATION ===
prs = Presentation()
prs.slide_width = SLIDE_W
prs.slide_height = SLIDE_H

# Use blank layout
blank_layout = prs.slide_layouts[6]  # Blank


# ============================================================
# SLIDE 1: COVER
# ============================================================
slide = prs.slides.add_slide(blank_layout)

# Dark background
bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
bg.fill.solid()
bg.fill.fore_color.rgb = DARK_SLATE
bg.line.fill.background()

# Accent bar at top
shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, Inches(0.12))
shape.fill.solid()
shape.fill.fore_color.rgb = ACCENT
shape.line.fill.background()

# Logo
if os.path.exists(LOGO_WHITE):
    slide.shapes.add_picture(LOGO_WHITE, Inches(0.8), Inches(0.6), Inches(3.5))

# Fund name
txBox = slide.shapes.add_textbox(Inches(0.8), Inches(2.2), Inches(11), Inches(1.5))
tf = txBox.text_frame
p = tf.paragraphs[0]
run = p.add_run()
run.text = "ACME FUND"
run.font.size = Pt(72)
run.font.bold = True
run.font.color.rgb = WHITE
run.font.name = FONT

p2 = tf.add_paragraph()
run2 = p2.add_run()
run2.text = "Industrial Real Estate Investment Fund"
run2.font.size = Pt(28)
run2.font.color.rgb = ACCENT
run2.font.name = FONT

# Tagline
txBox = slide.shapes.add_textbox(Inches(0.8), Inches(4.2), Inches(10), Inches(1))
tf = txBox.text_frame
p = tf.paragraphs[0]
p.text = "Building a diversified portfolio of USD-leased industrial assets\nacross high-growth nearshoring corridors"
p.font.size = Pt(16)
p.font.color.rgb = WARM_GRAY
p.font.name = FONT
p.line_spacing = Pt(26)

# Stats bar at bottom
stats = [
    ("$1.5B", "Target GAV"),
    ("14-17%", "Net IRR Target"),
    ("7-8%", "Cash Yield"),
    ("60+", "Years Experience"),
]
for i, (num, label) in enumerate(stats):
    x = Inches(0.8) + Inches(3) * i
    y = Inches(5.8)

    # Number
    txBox = slide.shapes.add_textbox(x, y, Inches(2.5), Inches(0.7))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = num
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = ACCENT
    p.font.name = FONT

    # Label
    txBox2 = slide.shapes.add_textbox(x, y + Inches(0.6), Inches(2.5), Inches(0.4))
    tf2 = txBox2.text_frame
    p2 = tf2.paragraphs[0]
    p2.text = label
    p2.font.size = Pt(11)
    p2.font.color.rgb = WARM_GRAY
    p2.font.name = FONT

# Confidential footer
txBox = slide.shapes.add_textbox(Inches(0.8), SLIDE_H - Inches(0.5), Inches(8), Inches(0.3))
tf = txBox.text_frame
p = tf.paragraphs[0]
p.text = "CONFIDENTIAL  |  For Qualified Investors Only"
p.font.size = Pt(9)
p.font.color.rgb = RGBColor(0x6B, 0x72, 0x78)
p.font.name = FONT


# ============================================================
# SLIDE 2: THE OPPORTUNITY
# ============================================================
slide = prs.slides.add_slide(blank_layout)
add_accent_bar(slide)
add_logo_watermark(slide)

add_section_label(slide, "THE OPPORTUNITY", Inches(0.8), Inches(0.5))
add_title(slide, "Nearshoring Is Reshaping North American\nSupply Chains, Creating a\nGenerational Investment Opportunity", Inches(0.8), Inches(0.9), width=Inches(11))

items = [
    "Nearshoring super-cycle: Global supply chains are relocating closer to end demand. Target markets deliver U.S. adjacency without trans-Pacific fragility.",
    "Tariff-volatility mitigation: USMCA helps tenants retain duty-free U.S. consumer access while reducing tariff and export-control shocks.",
    "Labor + logistics advantage: Competitive skilled labor, proximity to U.S. interstates and rail crossings, shorter transit cycles than Asia.",
    "Income & valuation spread: Target market cap rates at ~200 bps spread to U.S. Sunbelt industrial; USD lease structures compress perceived FX risk.",
    "Institutional demand gap: Limited institutional-quality supply in high-demand corridors creates pricing power for well-positioned developers.",
]

add_bullet_list(slide, items, Inches(0.8), Inches(2.8), width=Inches(11.5), size=Pt(14))


# ============================================================
# SLIDE 3: ABOUT THE COMPANY
# ============================================================
slide = prs.slides.add_slide(blank_layout)
add_accent_bar(slide)
add_logo_watermark(slide)

add_section_label(slide, "ABOUT US", Inches(0.8), Inches(0.5))
add_title(slide, "60+ Years of Industrial Real Estate\nin the Target Market", Inches(0.8), Inches(0.9), width=Inches(10))

# Left column - about text
tf = add_body(slide, "", Inches(0.8), Inches(2.4), width=Inches(5.5), height=Inches(4))
p = tf.paragraphs[0]
p.text = "Acme Industrial is a family-owned industrial real estate developer rooted in the target market since the early days of cross-border manufacturing."
p.font.size = Pt(15)
p.font.color.rgb = DARK_SLATE
p.font.name = FONT
p.line_spacing = Pt(24)

items_about = [
    "Full-service: development, build-to-suit, sale-leaseback, spec buildings",
    "1.5M+ sq ft developed across multiple industrial parks",
    "Deep municipal and entitlement access in target market",
    "Established relationships with multinational tenants",
    "In-house general contractor capabilities",
    "Proprietary land bank and ROFR/JV pipeline",
]
for item in items_about:
    p = tf.add_paragraph()
    p.text = "\u2022  " + item
    p.font.size = Pt(13)
    p.font.color.rgb = DARK_SLATE
    p.font.name = FONT
    p.line_spacing = Pt(22)
    p.space_after = Pt(4)

# Right column - key stats
stats_right = [
    ("60+", "Years Manufacturing\nExperience"),
    ("1.5M+", "Sq Ft Developed"),
    ("30+", "Years as Developer"),
    ("90M+", "Sq Ft Industrial Space\nin Target Market"),
]
for i, (num, label) in enumerate(stats_right):
    x = Inches(7.5)
    y = Inches(2.4) + Inches(1.2) * i
    add_stat_box(slide, num, label, x, y, width=Inches(4.5))

# Tagline
txBox = slide.shapes.add_textbox(Inches(0.8), Inches(6.5), Inches(8), Inches(0.5))
tf = txBox.text_frame
p = tf.paragraphs[0]
p.text = "\"Large enough to serve, small enough to care.\""
p.font.size = Pt(18)
p.font.italic = True
p.font.color.rgb = ACCENT
p.font.name = FONT


# ============================================================
# SLIDE 4: MARKET OVERVIEW - THREE CITIES
# ============================================================
slide = prs.slides.add_slide(blank_layout)
add_accent_bar(slide)
add_logo_watermark(slide)

add_section_label(slide, "MARKET OVERVIEW", Inches(0.8), Inches(0.5))
add_title(slide, "Three Strategic Markets Across\nthe Target Region", Inches(0.8), Inches(0.9), width=Inches(11))

cities = [
    {
        "name": "Market Alpha",
        "color": ACCENT,
        "strengths": [
            "Borders two U.S. states",
            "Major electronics manufacturing cluster",
            "4 commercial bridge crossings",
            "60+ years of manufacturing history",
            "261,000+ manufacturing workers",
            "370+ registered companies",
        ]
    },
    {
        "name": "Market Beta",
        "color": SECONDARY,
        "strengths": [
            "Regional industrial and financial capital",
            "Deep automotive and heavy manufacturing base",
            "Largest talent pool of engineers in the region",
            "International airport with direct U.S. flights",
            "Premium institutional-grade industrial parks",
            "Strong infrastructure and logistics networks",
        ]
    },
    {
        "name": "Market Gamma",
        "color": ACCENT,
        "strengths": [
            "Direct border access to Southern California",
            "Medical device manufacturing hub",
            "Aerospace and defense cluster",
            "Gateway to Pacific trade routes",
            "Strong bilingual workforce",
            "Growing institutional investor interest",
        ]
    },
]

for i, city in enumerate(cities):
    x = Inches(0.8) + Inches(4) * i
    y_start = Inches(2.6)

    # City name
    txBox = slide.shapes.add_textbox(x, y_start, Inches(3.5), Inches(0.5))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = city["name"]
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = city["color"]
    p.font.name = FONT

    # Divider line
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y_start + Inches(0.55), Inches(3.5), Inches(0.03))
    line.fill.solid()
    line.fill.fore_color.rgb = city["color"]
    line.line.fill.background()

    # Strengths
    txBox2 = slide.shapes.add_textbox(x, y_start + Inches(0.7), Inches(3.5), Inches(4))
    tf2 = txBox2.text_frame
    tf2.word_wrap = True
    for j, s in enumerate(city["strengths"]):
        if j == 0:
            p = tf2.paragraphs[0]
        else:
            p = tf2.add_paragraph()
        p.text = "\u2022  " + s
        p.font.size = Pt(12)
        p.font.color.rgb = DARK_SLATE
        p.font.name = FONT
        p.line_spacing = Pt(20)
        p.space_after = Pt(4)


# ============================================================
# SLIDE 5: DEMAND DRIVERS
# ============================================================
slide = prs.slides.add_slide(blank_layout)
add_accent_bar(slide)
add_logo_watermark(slide)

add_section_label(slide, "DEMAND DRIVERS", Inches(0.8), Inches(0.5))
add_title(slide, "Why Companies Are Moving to\nthe Target Region Now", Inches(0.8), Inches(0.9), width=Inches(11))

drivers = [
    ("Tariff Pressure", "China tariffs at 25%+ with escalation risk. USMCA provides duty-free U.S. access for locally assembled products."),
    ("Supply Chain Risk", "COVID exposed single-source fragility. Boards are mandating nearshore alternatives for supply chain resilience."),
    ("Customer Mandates", "U.S. buyers increasingly requiring nearshore or domestic production from their suppliers."),
    ("Speed to Market", "2-3 day ground transit vs. 4-6 week ocean freight from Asia. Same-day customs clearance at the border."),
    ("Labor Advantage", "20-30% of China coastal wages. Deep skilled manufacturing workforce with 60+ years of experience."),
    ("USMCA Compliance", "Products assembled locally qualify for duty-free U.S. import under rules of origin provisions."),
]

for i, (title, desc) in enumerate(drivers):
    col = i % 3
    row = i // 3
    x = Inches(0.8) + Inches(4) * col
    y = Inches(2.6) + Inches(2.2) * row

    # Accent box
    accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, Inches(0.08), Inches(0.6))
    accent.fill.solid()
    accent.fill.fore_color.rgb = ACCENT
    accent.line.fill.background()

    # Title
    txBox = slide.shapes.add_textbox(x + Inches(0.2), y, Inches(3.5), Inches(0.4))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = DARK_SLATE
    p.font.name = FONT

    # Description
    txBox2 = slide.shapes.add_textbox(x + Inches(0.2), y + Inches(0.45), Inches(3.5), Inches(1.5))
    tf2 = txBox2.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.text = desc
    p2.font.size = Pt(11)
    p2.font.color.rgb = DARK_SLATE
    p2.font.name = FONT
    p2.line_spacing = Pt(18)


# ============================================================
# SLIDE 6: TENANT DEMAND (Electronics Cluster)
# ============================================================
slide = prs.slides.add_slide(blank_layout)
add_accent_bar(slide)
add_logo_watermark(slide)

add_section_label(slide, "TENANT DEMAND", Inches(0.8), Inches(0.5))
add_title(slide, "Asia's Electronics Giants Are Already\nBuilding in the Target Region", Inches(0.8), Inches(0.9), width=Inches(11))

companies = [
    ("Foxconn", "iPhone, AI servers", "$20B+ regional investment", "Operations since 2005"),
    ("Pegatron", "Apple supplier", "Regional expansion", "Operations since 2014"),
    ("Flex Ltd", "EMS leader", "Long-term presence", "Operations since 2003"),
    ("Inventec", "Server & notebook ODM", "US-proximate production", "Operations since 2008"),
    ("Quanta", "World's largest notebook ODM", "Expanding footprint", "Operations since 2010"),
    ("Wistron", "Consumer electronics", "Supply chain diversification", "Operations since 2012"),
    ("Wiwynn", "Cloud server hardware", "Data center supply chain", "Operations expanding"),
]

for i, (name, product, detail, year) in enumerate(companies):
    col = i % 2
    row = i // 2
    x = Inches(0.8) + Inches(6) * col
    y = Inches(2.6) + Inches(1.1) * row

    # Company name in accent color
    txBox = slide.shapes.add_textbox(x, y, Inches(2.5), Inches(0.4))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = name
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = ACCENT
    p.font.name = FONT

    # Details
    txBox2 = slide.shapes.add_textbox(x + Inches(2.6), y, Inches(3), Inches(0.9))
    tf2 = txBox2.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.text = f"{product} | {detail}"
    p2.font.size = Pt(12)
    p2.font.color.rgb = DARK_SLATE
    p2.font.name = FONT

    p3 = tf2.add_paragraph()
    p3.text = year
    p3.font.size = Pt(10)
    p3.font.italic = True
    p3.font.color.rgb = RGBColor(0x6B, 0x72, 0x78)
    p3.font.name = FONT

# Bottom note
txBox = slide.shapes.add_textbox(Inches(0.8), Inches(6.5), Inches(11), Inches(0.5))
tf = txBox.text_frame
p = tf.paragraphs[0]
p.text = "These OEMs and their supply chain partners need industrial space. Their suppliers represent an additional pipeline of tenant demand for the fund."
p.font.size = Pt(13)
p.font.italic = True
p.font.color.rgb = DARK_SLATE
p.font.name = FONT


# ============================================================
# SLIDE 7: INVESTMENT STRATEGY
# ============================================================
slide = prs.slides.add_slide(blank_layout)
add_accent_bar(slide)
add_logo_watermark(slide)

add_section_label(slide, "INVESTMENT STRATEGY", Inches(0.8), Inches(0.5))
add_title(slide, "Core-Plus Industrial Portfolio\nWith Development Upside", Inches(0.8), Inches(0.9), width=Inches(11))

# Left column: Stabilized Acquisitions
txBox = slide.shapes.add_textbox(Inches(0.8), Inches(2.5), Inches(5.5), Inches(0.5))
tf = txBox.text_frame
p = tf.paragraphs[0]
p.text = "Stabilized Acquisitions (~65%)"
p.font.size = Pt(20)
p.font.bold = True
p.font.color.rgb = ACCENT
p.font.name = FONT

acq_items = [
    "Class-A industrial, >100,000 SF",
    "NNN leases, USD-denominated",
    "WALT ~5 years, annual CPI escalators",
    "Multinational credit tenants",
    "Entry cap range: ~8.0% across target markets",
]
add_bullet_list(slide, acq_items, Inches(0.8), Inches(3.2), width=Inches(5.5), height=Inches(3), size=Pt(13))

# Right column: Build-to-Core
txBox = slide.shapes.add_textbox(Inches(7), Inches(2.5), Inches(5.5), Inches(0.5))
tf = txBox.text_frame
p = tf.paragraphs[0]
p.text = "Build-to-Core Development (\u226435%)"
p.font.size = Pt(20)
p.font.bold = True
p.font.color.rgb = SECONDARY
p.font.name = FONT

dev_items = [
    "\u226512% yield-on-cost target",
    "~65% LTC fixed-rate financing",
    "6-12 months reserves",
    "First building spec; \u226550% leased before next spec",
    "Pre-leased and build-to-suit preferred",
    "Corporate guaranteed leases (where available)",
]
add_bullet_list(slide, dev_items, Inches(7), Inches(3.2), width=Inches(5.5), height=Inches(3), size=Pt(13))

# Vertical divider
line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.5), Inches(2.5), Inches(0.03), Inches(4))
line.fill.solid()
line.fill.fore_color.rgb = WARM_GRAY
line.line.fill.background()

# Portfolio target
txBox = slide.shapes.add_textbox(Inches(0.8), Inches(6.5), Inches(11), Inches(0.5))
tf = txBox.text_frame
p = tf.paragraphs[0]
r = p.add_run()
r.text = "Target: "
r.font.size = Pt(14)
r.font.bold = True
r.font.color.rgb = ACCENT
r.font.name = FONT
r2 = p.add_run()
r2.text = "$1.5B gross asset value across ~60 assets  |  ~$9.3M equity per deal  |  MSA cap 40%  |  6 assets per year"
r2.font.size = Pt(14)
r2.font.color.rgb = DARK_SLATE
r2.font.name = FONT


# ============================================================
# SLIDE 8: RETURN ENGINE
# ============================================================
slide = prs.slides.add_slide(blank_layout)
add_accent_bar(slide)
add_logo_watermark(slide)

add_section_label(slide, "RETURNS", Inches(0.8), Inches(0.5))
add_title(slide, "Five Levers Driving Returns", Inches(0.8), Inches(0.9), width=Inches(10))

levers = [
    ("1", "Going-In Yield Premium", "Core-Plus income from day one. Target market cap rates offer ~200 bps spread vs. U.S. Sunbelt industrial."),
    ("2", "Contractual NOI Growth", "NNN lease structures with U.S. CPI escalators provide built-in annual rent growth."),
    ("3", "Credit Quality & Lease Term", "~5-year WALT with multinational corporate-guaranteed tenants reduces income volatility."),
    ("4", "Execution Alpha", "Proprietary access via ROFR/JVs, land bank, in-house development and GC capabilities."),
    ("5", "Exit Cap Discipline", "Underwrite exit at minimum 100 bps inside entry cap. Rolling sales at IRR/MOIC thresholds."),
]

for i, (num, title, desc) in enumerate(levers):
    y = Inches(2.4) + Inches(0.95) * i

    # Number circle
    circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.8), y, Inches(0.5), Inches(0.5))
    circle.fill.solid()
    circle.fill.fore_color.rgb = ACCENT
    circle.line.fill.background()
    tf_c = circle.text_frame
    tf_c.word_wrap = False
    p_c = tf_c.paragraphs[0]
    p_c.text = num
    p_c.font.size = Pt(18)
    p_c.font.bold = True
    p_c.font.color.rgb = WHITE
    p_c.font.name = FONT
    p_c.alignment = PP_ALIGN.CENTER
    tf_c.paragraphs[0].alignment = PP_ALIGN.CENTER

    # Title
    txBox = slide.shapes.add_textbox(Inches(1.5), y, Inches(3.5), Inches(0.4))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = DARK_SLATE
    p.font.name = FONT

    # Description
    txBox2 = slide.shapes.add_textbox(Inches(5), y, Inches(7.5), Inches(0.8))
    tf2 = txBox2.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.text = desc
    p2.font.size = Pt(12)
    p2.font.color.rgb = DARK_SLATE
    p2.font.name = FONT
    p2.line_spacing = Pt(18)


# ============================================================
# SLIDE 9: FUND STRUCTURE & TERMS
# ============================================================
slide = prs.slides.add_slide(blank_layout)
add_accent_bar(slide)
add_logo_watermark(slide)

add_section_label(slide, "FUND STRUCTURE", Inches(0.8), Inches(0.5))
add_title(slide, "Acme Fund | Key Terms", Inches(0.8), Inches(0.9), width=Inches(10))

# Two columns of terms
left_terms = [
    ("Vehicle", "Delaware LP master with optional Cayman feeder/blocker"),
    ("Fund Life", "Closed-end, up to 15 years"),
    ("Target GAV", "$1.5B across ~60 assets"),
    ("Target Leverage", "~65% LTV at acquisition; 75% hard max"),
    ("Debt Profile", "Fixed-rate bias, 7-10 year tenor, DSCR \u22651.25x"),
    ("Currency", "All USD (leases, debt, distributions)"),
    ("GP Commitment", "5-10% of fund"),
]

right_terms = [
    ("Target Net IRR", "14-17%"),
    ("Cash Yield", "7-8% quarterly"),
    ("Preferred Return", "8% cumulative, non-compounding"),
    ("Promote", "15-20% over 8% IRR\n20-25% over 12% IRR\n30% over 17% IRR"),
    ("AM Fee", "2.0% during IP, then 1.5%"),
    ("Acquisition Fee", "1%"),
    ("Development Fee", "5%"),
]

# Left column
txBox = slide.shapes.add_textbox(Inches(0.8), Inches(2.4), Inches(5.5), Inches(5))
tf = txBox.text_frame
tf.word_wrap = True
for i, (key, val) in enumerate(left_terms):
    if i == 0:
        p = tf.paragraphs[0]
    else:
        p = tf.add_paragraph()

    r_k = p.add_run()
    r_k.text = key
    r_k.font.size = Pt(13)
    r_k.font.bold = True
    r_k.font.color.rgb = ACCENT
    r_k.font.name = FONT

    p2 = tf.add_paragraph()
    p2.text = val
    p2.font.size = Pt(12)
    p2.font.color.rgb = DARK_SLATE
    p2.font.name = FONT
    p2.space_after = Pt(8)
    p2.line_spacing = Pt(18)

# Right column
txBox = slide.shapes.add_textbox(Inches(7), Inches(2.4), Inches(5.5), Inches(5))
tf = txBox.text_frame
tf.word_wrap = True
for i, (key, val) in enumerate(right_terms):
    if i == 0:
        p = tf.paragraphs[0]
    else:
        p = tf.add_paragraph()

    r_k = p.add_run()
    r_k.text = key
    r_k.font.size = Pt(13)
    r_k.font.bold = True
    r_k.font.color.rgb = ACCENT
    r_k.font.name = FONT

    p2 = tf.add_paragraph()
    p2.text = val
    p2.font.size = Pt(12)
    p2.font.color.rgb = DARK_SLATE
    p2.font.name = FONT
    p2.space_after = Pt(8)
    p2.line_spacing = Pt(18)

# Vertical divider
line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.5), Inches(2.4), Inches(0.03), Inches(4.5))
line.fill.solid()
line.fill.fore_color.rgb = WARM_GRAY
line.line.fill.background()


# ============================================================
# SLIDE 10: RISK MANAGEMENT
# ============================================================
slide = prs.slides.add_slide(blank_layout)
add_accent_bar(slide)
add_logo_watermark(slide)

add_section_label(slide, "RISK MANAGEMENT", Inches(0.8), Inches(0.5))
add_title(slide, "Disciplined Risk Framework", Inches(0.8), Inches(0.9), width=Inches(10))

risks = [
    ("Market / Policy", "USMCA + USD leases provide structural hedge; MSA concentration cap at 40%"),
    ("Tenant / Credit", "\u226580% credit or parent-guaranteed tenants; watch-list protocol at 24-30 months to maturity"),
    ("Lease / Income", "NNN + CPI escalators; staggered lease maturities across portfolio"),
    ("Construction", "\u226435% development allocation; \u226512% YoC; 6-12 month reserves; GMP contracts"),
    ("Debt / Refi", "Fixed-rate bias, DSCR \u22651.25x, LTV \u226475%; no subscription lines"),
    ("Legal / Compliance", "FCPA and sanctions diligence; environmental assessment; title insurance on all acquisitions"),
]

for i, (risk, mitigant) in enumerate(risks):
    col = i % 2
    row = i // 2
    x = Inches(0.8) + Inches(6) * col
    y = Inches(2.4) + Inches(1.5) * row

    # Risk category - accent
    txBox = slide.shapes.add_textbox(x, y, Inches(5.5), Inches(0.4))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = risk
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = ACCENT
    p.font.name = FONT

    # Mitigant
    txBox2 = slide.shapes.add_textbox(x, y + Inches(0.4), Inches(5.3), Inches(1))
    tf2 = txBox2.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.text = mitigant
    p2.font.size = Pt(12)
    p2.font.color.rgb = DARK_SLATE
    p2.font.name = FONT
    p2.line_spacing = Pt(18)


# ============================================================
# SLIDE 11: COMPETITIVE ADVANTAGE
# ============================================================
slide = prs.slides.add_slide(blank_layout)
add_accent_bar(slide)
add_logo_watermark(slide)

add_section_label(slide, "COMPETITIVE ADVANTAGE", Inches(0.8), Inches(0.5))
add_title(slide, "Why Acme Fund", Inches(0.8), Inches(0.9), width=Inches(10))

advantages = [
    ("Proprietary Access", "Land bank, ROFR agreements, joint venture relationships, and repeat tenant pipelines that institutional funds cannot replicate."),
    ("Local Execution", "Established entitlement and municipal access. In-house general contractor reduces timelines and costs."),
    ("Data & AI Engine", "Broker pipeline scraping, AI-assisted lead scoring, and site-selection models built on 1,000+ qualified leads and decades of market intelligence."),
    ("Cross-Border Capability", "Bilingual team, U.S.-Mexico capital markets expertise, and dual-border advantage for faster closings and operational flexibility."),
    ("Aligned Incentives", "GP commits 5-10% alongside LPs. Transparent affiliate policies. Independent IC member. Institutional reporting package."),
]

for i, (title, desc) in enumerate(advantages):
    y = Inches(2.2) + Inches(1.0) * i

    # Accent bar
    accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), y + Inches(0.05), Inches(0.08), Inches(0.5))
    accent.fill.solid()
    accent.fill.fore_color.rgb = ACCENT
    accent.line.fill.background()

    # Title
    txBox = slide.shapes.add_textbox(Inches(1.1), y, Inches(3), Inches(0.4))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = DARK_SLATE
    p.font.name = FONT

    # Description
    txBox2 = slide.shapes.add_textbox(Inches(4.5), y, Inches(8), Inches(0.8))
    tf2 = txBox2.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.text = desc
    p2.font.size = Pt(12)
    p2.font.color.rgb = DARK_SLATE
    p2.font.name = FONT
    p2.line_spacing = Pt(18)


# ============================================================
# SLIDE 12: THE ASK
# ============================================================
slide = prs.slides.add_slide(blank_layout)

# Dark background
bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
bg.fill.solid()
bg.fill.fore_color.rgb = DARK_SLATE
bg.line.fill.background()

# Accent bar
shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, Inches(0.12))
shape.fill.solid()
shape.fill.fore_color.rgb = ACCENT
shape.line.fill.background()

add_section_label(slide, "THE ASK", Inches(0.8), Inches(0.5))
# Override color for dark bg
for shape in slide.shapes:
    if hasattr(shape, "text_frame"):
        for p in shape.text_frame.paragraphs:
            if p.text == "THE ASK":
                p.font.color.rgb = ACCENT

txBox = slide.shapes.add_textbox(Inches(0.8), Inches(1.2), Inches(11), Inches(1))
tf = txBox.text_frame
p = tf.paragraphs[0]
p.text = "Invest in the\nNearshoring Super-Cycle"
p.font.size = Pt(42)
p.font.bold = True
p.font.color.rgb = WHITE
p.font.name = FONT
p.line_spacing = Pt(52)

# Key terms summary
terms_summary = [
    ("Target Net IRR", "14-17% (net to LPs, USD)"),
    ("Cash Yield", "7-8% quarterly distributions"),
    ("Fund Life", "Closed-end, up to 15 years"),
    ("Preferred Return", "8% cumulative"),
    ("GP Commitment", "5-10%, fully aligned"),
]

txBox = slide.shapes.add_textbox(Inches(0.8), Inches(3.5), Inches(5.5), Inches(3.5))
tf = txBox.text_frame
tf.word_wrap = True
for i, (key, val) in enumerate(terms_summary):
    if i == 0:
        p = tf.paragraphs[0]
    else:
        p = tf.add_paragraph()

    r1 = p.add_run()
    r1.text = key + "  "
    r1.font.size = Pt(14)
    r1.font.bold = True
    r1.font.color.rgb = ACCENT
    r1.font.name = FONT

    r2 = p.add_run()
    r2.text = val
    r2.font.size = Pt(14)
    r2.font.color.rgb = WHITE
    r2.font.name = FONT
    p.line_spacing = Pt(28)

# Governance highlights
txBox = slide.shapes.add_textbox(Inches(7), Inches(3.5), Inches(5.5), Inches(3))
tf = txBox.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "Governance & Reporting"
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = ACCENT
p.font.name = FONT

gov_items = [
    "GP-controlled IC with 1 independent member",
    "LP observer rights; LPAC for major deviations",
    "Quarterly institutional reporting + annual GAAP audit",
    "Quarterly NAV",
    "Disposition-ready data rooms maintained",
]
for item in gov_items:
    p = tf.add_paragraph()
    p.text = "\u2022  " + item
    p.font.size = Pt(12)
    p.font.color.rgb = WARM_GRAY
    p.font.name = FONT
    p.line_spacing = Pt(20)

# Confidential
txBox = slide.shapes.add_textbox(Inches(0.8), SLIDE_H - Inches(0.5), Inches(8), Inches(0.3))
tf = txBox.text_frame
p = tf.paragraphs[0]
p.text = "CONFIDENTIAL  |  For Qualified Investors Only"
p.font.size = Pt(9)
p.font.color.rgb = RGBColor(0x6B, 0x72, 0x78)
p.font.name = FONT


# ============================================================
# SLIDE 13: CONTACT
# ============================================================
slide = prs.slides.add_slide(blank_layout)

# Dark background
bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
bg.fill.solid()
bg.fill.fore_color.rgb = DARK_SLATE
bg.line.fill.background()

# Accent bar
shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, Inches(0.12))
shape.fill.solid()
shape.fill.fore_color.rgb = ACCENT
shape.line.fill.background()

# Logo
if os.path.exists(LOGO_WHITE):
    slide.shapes.add_picture(LOGO_WHITE, Inches(4.5), Inches(1.5), Inches(4.5))

# Contact info
txBox = slide.shapes.add_textbox(Inches(3), Inches(3.2), Inches(7), Inches(3))
tf = txBox.text_frame
p = tf.paragraphs[0]
p.text = "John Smith"
p.font.size = Pt(32)
p.font.bold = True
p.font.color.rgb = WHITE
p.font.name = FONT
p.alignment = PP_ALIGN.CENTER

p2 = tf.add_paragraph()
p2.text = "CEO, Acme Industrial Real Estate"
p2.font.size = Pt(16)
p2.font.color.rgb = ACCENT
p2.font.name = FONT
p2.alignment = PP_ALIGN.CENTER
p2.space_after = Pt(20)

p3 = tf.add_paragraph()
p3.text = "ceo@example.com"
p3.font.size = Pt(14)
p3.font.color.rgb = WARM_GRAY
p3.font.name = FONT
p3.alignment = PP_ALIGN.CENTER

p4 = tf.add_paragraph()
p4.text = "acme-industrial.com"
p4.font.size = Pt(14)
p4.font.color.rgb = WARM_GRAY
p4.font.name = FONT
p4.alignment = PP_ALIGN.CENTER

p5 = tf.add_paragraph()
p5.text = "Target Market, Country"
p5.font.size = Pt(14)
p5.font.color.rgb = WARM_GRAY
p5.font.name = FONT
p5.alignment = PP_ALIGN.CENTER

# Tagline
txBox = slide.shapes.add_textbox(Inches(2), Inches(5.8), Inches(9), Inches(0.5))
tf = txBox.text_frame
p = tf.paragraphs[0]
p.text = "\"Large enough to serve, small enough to care.\""
p.font.size = Pt(18)
p.font.italic = True
p.font.color.rgb = ACCENT
p.font.name = FONT
p.alignment = PP_ALIGN.CENTER

# Confidential
txBox = slide.shapes.add_textbox(Inches(0.8), SLIDE_H - Inches(0.5), Inches(8), Inches(0.3))
tf = txBox.text_frame
p = tf.paragraphs[0]
p.text = "CONFIDENTIAL  |  For Qualified Investors Only"
p.font.size = Pt(9)
p.font.color.rgb = RGBColor(0x6B, 0x72, 0x78)
p.font.name = FONT


# === SAVE ===
os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)
prs.save(OUTPUT)
print(f"Deck saved to: {OUTPUT}")
print(f"Total slides: {len(prs.slides)}")
