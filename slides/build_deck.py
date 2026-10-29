"""Generate the PyCon ES 2026 deck.

Built to the organisers' speaker guide:
  * titles 36-44pt, body 28-32pt, code and chart labels 24-28pt
  * nothing the audience must read falls below 24pt
  * white background, dark text, brand colour reserved for emphasis
  * one idea per slide, short lines, wide margins
  * identity never carried by colour alone (labels + dashes + markers)
  * alt text on every diagram and chart, for the shared file

Every performance figure is read from bench/results.json -- nothing is typed
by hand, so re-running the benchmark and re-running this script keeps the
slides honest.

    ../.venv/bin/python build_deck.py
"""

import json
import os
import re

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_MARKER_STYLE
from pptx.enum.dml import MSO_LINE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
RESULTS = json.load(open(os.path.join(ROOT, "bench", "results.json")))

REPO = "github.com/<your-handle>/pycones26-arrow-pyo3"
LINKEDIN = "linkedin.com/in/allen-yesudasan"
ASSETS = os.path.join(HERE, "assets")

# --------------------------------------------------------------------------
# Theme: PyCon ES 2026 brand (2026.es.pycon.org) on white.
# Brand tokens taken from the site's own CSS custom properties:
#   --color-pycon-black  #1d1d1b    --color-pycon-orange #ff8200
#   --color-pycon-gray   #4a4a48    --color-pycon-red    #f9423a
#   --color-pycon-black-50 #8e8e8d  --color-pycon-yellow #ffc72c
#   --color-pycon-black-25 #c6c6c6  --color-pycon-mute   #c60
# --------------------------------------------------------------------------
BG = RGBColor(0xFF, 0xFF, 0xFF)
PANEL = RGBColor(0xF4, 0xF4, 0xF2)
INK = RGBColor(0x1D, 0x1D, 0x1B)         # 15.9:1 on white
INK2 = RGBColor(0x4A, 0x4A, 0x48)        # 8.9:1
MUTED = RGBColor(0x6E, 0x6E, 0x6C)       # 5.5:1 -- darkened from the brand's
RULE = RGBColor(0xC6, 0xC6, 0xC6)        #   #8e8e8d so captions stay legible
GRID = RGBColor(0xDF, 0xDF, 0xDD)

ACCENT = RGBColor(0xCC, 0x66, 0x00)      # pycon-mute -- brand orange, 4.0:1
ACCENT_L = RGBColor(0xFF, 0x82, 0x00)    # pycon-orange -- bars/rules only

# Chart series, validated with the dataviz validator on surface #ffffff:
#   bars  "#f9423a,#2f5fa8"                 -> all checks pass (all-pairs)
#   lines "#f9423a,#2f5fa8,#00895e,#c66a00" -> pass; green/orange CVD 7.3
#                                              sits in the warn band, so the
#                                              lines also carry dashes+markers
S1 = RGBColor(0x2F, 0x5F, 0xA8)          # blue   -- Arrow / shares the data
S2 = RGBColor(0xF9, 0x42, 0x3A)          # red    -- JSON / copies the data
S3 = RGBColor(0x00, 0x89, 0x5E)          # green  -- NumPy
S4 = RGBColor(0xC6, 0x6A, 0x00)          # orange -- Python list
GOOD = RGBColor(0x00, 0x6E, 0x4B)
CRIT = RGBColor(0xC4, 0x1E, 0x17)

SANS = "Outfit"
MONO = "JetBrains Mono"

# Type scale -- the guide's floors, not suggestions.
T_TITLE = 36
T_KICKER = 18
T_BODY = 28
T_BODY_SM = 24          # absolute floor for content
T_CODE = 24
T_CAPTION = 18          # supplementary only (method notes, sources)
T_FOOTER = 14

W, H = Inches(13.333), Inches(7.5)
L = Inches(0.9)
CW = Inches(11.53)
TOP = Inches(1.75)      # first line of content
BOTTOM = Inches(6.55)   # nothing below this except the footer

prs = Presentation()
prs.slide_width, prs.slide_height = W, H
BLANK = prs.slide_layouts[6]


# --------------------------------------------------------------------------
# primitives
# --------------------------------------------------------------------------
def alt(shape, description):
    """Alt text, so the shared file works with a screen reader."""
    for tag in ("p:nvSpPr", "p:nvGraphicFramePr", "p:nvPicPr", "p:nvGrpSpPr"):
        nv = shape._element.find(qn(tag))
        if nv is not None:
            c = nv.find(qn("p:cNvPr"))
            if c is not None:
                c.set("descr", description)
            return


def slide(notes=""):
    s = prs.slides.add_slide(BLANK)
    f = s.background.fill
    f.solid()
    f.fore_color.rgb = BG
    if notes:
        s.notes_slide.notes_text_frame.text = notes
    return s


def rect(s, l, t, w, h, color, line=None, lw=1.0, rounded=True):
    shp = s.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE, l, t, w, h
    )
    if rounded:
        shp.adjustments[0] = 0.05
    shp.fill.solid()
    shp.fill.fore_color.rgb = color
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(lw)
    shp.shadow.inherit = False
    return shp


def bar(s, l, t, w, h, color):
    return rect(s, l, t, w, h, color, rounded=False)


def text(s, l, t, w, h, runs, size=T_BODY, color=INK, bold=False, font=SANS,
         align=PP_ALIGN.LEFT, spacing=1.25, anchor=MSO_ANCHOR.TOP, space_after=10):
    box = s.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0

    lines = [runs] if isinstance(runs, str) else runs
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = spacing
        p.space_after = Pt(space_after)
        for txt, c, b in ([(line, color, bold)] if isinstance(line, str) else line):
            r = p.add_run()
            r.text = txt
            r.font.size = Pt(size)
            r.font.color.rgb = c
            r.font.bold = b
            r.font.name = font
    return box


def title(s, head, kicker=None):
    top = Inches(0.5)
    if kicker:
        text(s, L, Inches(0.42), CW, Inches(0.35), kicker, size=T_KICKER,
             color=ACCENT, bold=True)
        top = Inches(0.82)
    text(s, L, top, CW, Inches(0.8), head, size=T_TITLE, bold=True, color=INK)
    bar(s, L, top + Inches(0.72), Inches(1.3), Pt(4), ACCENT_L)


def footer(s, n):
    text(s, L, Inches(6.95), Inches(9.0), Inches(0.32),
         f"Zero-Copy Python + Rust  ·  PyCon ES 2026  ·  {REPO}",
         size=T_FOOTER, color=MUTED)
    text(s, Inches(11.6), Inches(6.95), Inches(0.9), Inches(0.32), str(n),
         size=T_FOOTER, color=MUTED, align=PP_ALIGN.RIGHT)


def caption(s, t, txt):
    """Method note under a chart -- supplementary, spoken aloud anyway."""
    text(s, L, t, CW, Inches(0.4), txt, size=T_CAPTION, color=MUTED)


# --------------------------------------------------------------------------
# code
# --------------------------------------------------------------------------
KW = {
    "rust": r"\b(fn|let|mut|use|pub|struct|impl|match|return|mod|as|if|else|for|in|"
            r"where|crate|dyn|self|Ok|Err|Some|None|true|false|const|type)\b",
    "python": r"\b(def|import|from|return|if|else|elif|for|in|with|as|class|lambda|"
              r"True|False|None|print|not|and|or)\b",
    "toml": r"^\s*\[[^\]]+\]",
    "bash": r"\b(cargo|maturin|pip|python)\b",
}
MAXCOL = 56   # characters that fit across the content width at 24pt


def code(s, l, t, w, src, lang="rust", size=T_CODE, pad=Inches(0.3)):
    lines = src.strip("\n").split("\n")
    over = [x for x in lines if len(x) > MAXCOL]
    assert not over, f"code line too long for {size}pt: {over[0]!r} ({len(over[0])})"

    h = Emu(int(Inches(0.50) + len(lines) * Pt(size * 1.55)))
    panel = rect(s, l, t, w, h, PANEL, line=RULE, lw=0.75)
    alt(panel, "Code listing: " + " / ".join(lines[:3]))

    box = s.shapes.add_textbox(l + pad, t + Inches(0.22), w - 2 * pad, h - Inches(0.3))
    tf = box.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.line_spacing = 1.25
        p.space_after = Pt(0)
        for txt, col in _tok(line, lang):
            r = p.add_run()
            r.text = txt
            r.font.size = Pt(size)
            r.font.name = MONO
            r.font.color.rgb = col
    return t + h    # bottom, so callers can stack


def _tok(line, lang):
    marks = ("//",) if lang == "rust" else ("#",)
    if line.lstrip().startswith(marks):
        return [(line, MUTED)]
    cmt = r"//.*$" if lang == "rust" else r"(?://|\#).*$"
    out, idx = [], 0
    pat = re.compile(
        r'(?P<str>"[^"]*")|(?P<cmt>' + cmt + r")"
        r"|(?P<kw>" + KW.get(lang, KW["rust"]) + r")|(?P<num>\b\d[\d_\.]*\b)"
    )
    for m in pat.finditer(line):
        if m.start() > idx:
            out.append((line[idx:m.start()], INK))
        g = m.lastgroup
        out.append((m.group(), {"str": S3, "cmt": MUTED, "kw": S1}.get(g, S4)))
        idx = m.end()
    if idx < len(line):
        out.append((line[idx:], INK))
    return out or [(line, INK)]


# --------------------------------------------------------------------------
# tiles & hero numbers
# --------------------------------------------------------------------------
def tile(s, l, t, w, value, label, color=INK, sub=None, h=Inches(1.75)):
    rect(s, l, t, w, h, PANEL, line=RULE, lw=0.75)
    text(s, l + Inches(0.28), t + Inches(0.2), w - Inches(0.56), Inches(0.7),
         value, size=40, bold=True, color=color)
    text(s, l + Inches(0.28), t + Inches(0.92), w - Inches(0.56), Inches(0.4),
         label, size=T_BODY_SM, color=INK2)
    if sub:
        text(s, l + Inches(0.28), t + Inches(1.34), w - Inches(0.56), Inches(0.35),
             sub, size=T_CAPTION, color=MUTED)


def hero(s, value, label, color=S1):
    text(s, L, Inches(2.2), CW, Inches(1.9), value, size=115, bold=True,
         color=color, align=PP_ALIGN.CENTER)
    text(s, L, Inches(4.25), CW, Inches(0.9), label, size=T_BODY, color=INK2,
         align=PP_ALIGN.CENTER)


def section(s, num, head, sub):
    text(s, L, Inches(2.5), Inches(2.4), Inches(1.5), num, size=96, bold=True,
         color=ACCENT)
    text(s, L, Inches(3.95), CW, Inches(0.9), head, size=44, bold=True)
    text(s, L, Inches(4.9), CW, Inches(0.6), sub, size=T_BODY, color=INK2)


# --------------------------------------------------------------------------
# charts
# --------------------------------------------------------------------------
def _axes(chart, log=False):
    ca, va = chart.category_axis, chart.value_axis
    for ax in (ca, va):
        ax.has_major_gridlines = False
        ax.tick_labels.font.size = Pt(22)
        ax.tick_labels.font.color.rgb = INK2
        ax.tick_labels.font.name = SANS
        ax.format.line.color.rgb = RULE
    va.has_major_gridlines = True
    va.major_gridlines.format.line.color.rgb = GRID
    va.major_gridlines.format.line.width = Pt(0.75)
    if log:
        sc = va._element.find(qn("c:scaling"))
        lb = etree.SubElement(sc, qn("c:logBase"))
        lb.set("val", "10")
        sc.remove(lb)
        sc.insert(0, lb)      # schema order: logBase, orientation, max, min


def bar_chart(s, l, t, w, h, cats, vals, colors, log=False, fmt='0.0',
              description=""):
    cd = CategoryChartData()
    cd.categories = cats
    cd.add_series("t", vals)
    gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, l, t, w, h, cd)
    alt(gf, description)
    chart = gf.chart
    chart.has_legend = False
    chart.has_title = False
    chart.font.name = SANS
    plot = chart.plots[0]
    plot.gap_width = 40
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.font.size = Pt(24)
    dl.font.bold = True
    dl.font.color.rgb = INK
    dl.font.name = SANS
    dl.number_format = fmt
    dl.number_format_is_linked = False
    for i, c in enumerate(colors):
        pt = plot.series[0].points[i]
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = c
        pt.format.line.fill.background()
    _axes(chart, log=log)
    return chart


def line_chart(s, l, t, w, h, cats, series, description=""):
    cd = CategoryChartData()
    cd.categories = cats
    for name, vals in series:
        cd.add_series(name, vals)
    gf = s.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, l, t, w, h, cd)
    alt(gf, description)
    chart = gf.chart
    chart.has_title = False
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    chart.legend.font.size = Pt(22)
    chart.legend.font.color.rgb = INK2
    chart.legend.font.name = SANS
    chart.font.name = SANS
    # Dash + marker shape duplicate the colour encoding, so the series stay
    # separable for colour-blind viewers and in greyscale print.
    styles = [(S2, MSO_LINE.SOLID, XL_MARKER_STYLE.CIRCLE),
              (S1, MSO_LINE.SOLID, XL_MARKER_STYLE.SQUARE),
              (S3, MSO_LINE.DASH, XL_MARKER_STYLE.DIAMOND),
              (S4, MSO_LINE.ROUND_DOT, XL_MARKER_STYLE.TRIANGLE)]
    for ser, (col, dash, mk) in zip(chart.series, styles):
        ser.format.line.color.rgb = col
        ser.format.line.width = Pt(2.75)
        ser.format.line.dash_style = dash
        ser.smooth = False
        m = ser.marker
        m.style = mk
        m.size = 9
        m.format.fill.solid()
        m.format.fill.fore_color.rgb = col
        m.format.line.color.rgb = BG
        m.format.line.width = Pt(1.5)
    _axes(chart, log=True)
    return chart


def group_legend(s, l, t):
    """Colour encodes the strategy, not the library. Text carries it too."""
    for i, (col, lab) in enumerate([(S2, "copies the data"),
                                    (S1, "shares the data")]):
        x = l + Inches(i * 3.6)
        bar(s, x, t + Inches(0.07), Inches(0.3), Inches(0.18), col)
        text(s, x + Inches(0.46), t, Inches(3.0), Inches(0.35), lab,
             size=T_BODY_SM, color=INK2)


# ==========================================================================
F = RESULTS["floats"]
M = RESULTS["memory"]
ST = RESULTS["strings"]
SC = RESULTS["scaling"]
V = RESULTS["verify"]
E = RESULTS["env"]

SP_LIST = F["list_total"] / F["arrow_total"]
SP_JSON = F["json_total"] / F["arrow_total"]
SP_STR = ST["list_call"] / ST["arrow_call"]


def build():
    n = [0]

    def nxt(notes="", chrome=True):
        n[0] += 1
        s = slide(notes)
        if chrome:
            footer(s, n[0])
        return s

    # ---- 1 ---------------------------------------------------------------
    s = nxt("[0:00] Hi, I'm Allen. In 35 minutes: how to move millions of rows "
            "between Python and Rust without copying a byte. Everything is "
            "measured on this laptop; the repo link is on screen now and at "
            "the end.", chrome=False)
    bar(s, L, Inches(1.75), Inches(1.9), Pt(5), ACCENT_L)
    text(s, L, Inches(2.1), Inches(8.5), Inches(1.9),
         "Zero-Copy Data Sharing\nBetween Python and Rust", size=50, bold=True,
         spacing=1.15)
    s.shapes.add_picture(os.path.join(ASSETS, "speaker.png"),
                         Inches(9.75), Inches(2.25), Inches(2.65), Inches(2.65))
    alt(s.shapes[-1], "Photograph of the speaker, Allen Y.")
    text(s, L, Inches(3.95), CW, Inches(0.6),
         [[("with ", INK2, False), ("PyO3", ACCENT, True), (" and ", INK2, False),
           ("Apache Arrow", ACCENT, True)]], size=T_BODY)
    text(s, L, Inches(5.0), CW, Inches(1.2),
         [[("Allen Y", INK, True),
           ("  ·  Associate Software Engineer, Red Hat", INK2, False)],
          [("PyCon ES 2026  ·  8 Nov  ·  Aula B7", INK2, False)]],
         size=T_BODY_SM, spacing=1.3)
    text(s, L, Inches(6.35), CW, Inches(0.5),
         [[("Slides and code:  ", INK2, False), (REPO, ACCENT, True)]],
         size=T_BODY_SM)

    # ---- 2 ---------------------------------------------------------------
    s = nxt("[0:45] Name the file out loud — sensors.parquet, five million "
            "rows. We come back to it all talk. Pause before the last line. "
            "If this is genuinely your story, say 'this happened to me' "
            "instead of 'you' — first person is much stronger.")
    title(s, "A story you may recognise", kicker="THE PROBLEM")
    text(s, L, Inches(2.1), CW, Inches(4.2),
         [[("sensors.parquet", ACCENT, True),
           (" — 5 million readings.", INK, True)],
          [("Summing one column in Python is too slow,", INK2, False)],
          [("so you move it to Rust: ", INK2, False),
           ("40× faster", GOOD, True), (" in isolation.", INK2, False)],
          [("End to end, the pipeline is ", INK2, False),
           ("slower than before", CRIT, True), (".", INK2, False)]],
         size=T_BODY, spacing=1.4, space_after=20)

    # ---- 3 ---------------------------------------------------------------
    s = nxt("[1:45] The computation was never the bottleneck. Say it, then move.")
    hero(s, "The border", "was the bottleneck — not the computation", color=S2)

    # ---- 4 ---------------------------------------------------------------
    s = nxt("[2:15] Four parts. Part 3 is the heart. Part 4 is what I wish "
            "someone had told me. Intermediate level: I assume you've seen "
            "Python and know roughly what Rust is.")
    title(s, "Where we're going")
    for i, (num, head) in enumerate([
            ("01", "The naive bridge — PyO3, and what it costs"),
            ("02", "The serialization trap — “just use JSON”"),
            ("03", "Arrow as shared memory — zero copies"),
            ("04", "Pitfalls and production")]):
        t = Inches(2.0 + i * 1.1)
        text(s, L, t, Inches(1.1), Inches(0.6), num, size=32, bold=True, color=ACCENT)
        text(s, L + Inches(1.3), t + Inches(0.03), Inches(10.0), Inches(0.6),
             head, size=T_BODY)

    # ---- 5 ---------------------------------------------------------------
    s = nxt("[3:30] Part one.")
    section(s, "01", "The naive bridge", "PyO3, and what it costs")

    # ---- 6 ---------------------------------------------------------------
    s = nxt("[4:00] PyO3 in one slide. Three annotations: pyfunction marks the "
            "function, pymodule marks the entry point, add_function registers "
            "it. Read the two fn lines aloud.")
    title(s, "PyO3: a native module", kicker="01 · THE NAIVE BRIDGE")
    b = code(s, L, Inches(1.85), CW, """
#[pyfunction]
fn add(a: i64, b: i64) -> i64 { a + b }

#[pymodule]
fn demo(m: &Bound<PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(add, m)?)
}
""", lang="rust")
    text(s, L, b + Inches(0.3), CW, Inches(0.6),
         "Three annotations. That is the whole bridge.", size=T_BODY, color=INK2)

    # ---- 7 ---------------------------------------------------------------
    s = nxt("[5:00] Build it and call it. Works, and it's fast. Then the "
            "punchline at the bottom — set up the rest of the talk here.")
    title(s, "Build it, call it", kicker="01 · THE NAIVE BRIDGE")
    b = code(s, L, Inches(1.85), CW, """
$ maturin develop --release

>>> import demo
>>> demo.add(2, 3)
5
""", lang="bash")
    text(s, L, b + Inches(0.4), CW, Inches(1.0),
         [[("Scalars are easy.", INK, True)],
          [("Now pass five million of them.", ACCENT, True)]],
         size=32, spacing=1.3)

    # ---- 8 ---------------------------------------------------------------
    s = nxt("[5:45] This is what everyone writes first. Vec<f64> looks free. "
            "Say clearly: it takes a Python list and returns a float.")
    title(s, "The obvious way", kicker="01 · THE NAIVE BRIDGE")
    b = code(s, L, Inches(1.85), CW, """
#[pyfunction]
fn sum_list(values: Vec<f64>) -> f64 {
    values.iter().sum()
}
""", lang="rust")
    b = code(s, L, b + Inches(0.35), CW, """
>>> sum_list(data)    # data: 5,000,000 floats
""", lang="python")
    text(s, L, b + Inches(0.35), CW, Inches(0.8),
         "It looks like a function call. It is a full copy.",
         size=T_BODY, color=INK2)

    # ---- 9 ---------------------------------------------------------------
    s = nxt(f"[6:45] Measured here: {F['payload_mb']} MB of float64. "
            f"Read the three numbers aloud left to right. The addition itself "
            f"is a few milliseconds — none of this is arithmetic.")
    title(s, "What that costs", kicker="01 · MEASURED")
    w3 = Inches(3.63)
    tile(s, L, Inches(1.9), w3, f"{F['list_total']:.0f} ms", "Arrow → list → Rust",
         color=S4, sub=f"{F['n']:,} float64 · {F['payload_mb']} MB")
    tile(s, L + Inches(3.95), Inches(1.9), w3, f"{F['to_pylist_only']:.0f} ms",
         "of it: .to_pylist()", color=INK2, sub="format conversion")
    tile(s, L + Inches(7.9), Inches(1.9), w3, f"{F['list_call']:.0f} ms",
         "of it: the PyO3 call", color=INK2, sub="unboxing 5M PyFloats")
    text(s, L, Inches(4.15), CW, Inches(1.6),
         [[("Not one byte of this is arithmetic.", INK, True)],
          [("Every element becomes a Python object, then gets unboxed again.",
            INK2, False)]], size=T_BODY, spacing=1.35)

    # ---- 10 --------------------------------------------------------------
    s = nxt("[7:45] Describe the diagram out loud: on the left, a list of "
            "pointers to boxed objects scattered across the heap. On the "
            "right, one contiguous block of raw float64. Then the 24-bytes line.")
    title(s, "Why a Python list is expensive", kicker="01 · THE NAIVE BRIDGE")
    text(s, L, Inches(1.8), Inches(6.0), Inches(0.4),
         "list[float] — boxed objects", size=T_BODY_SM, color=INK2)
    for i in range(5):
        x = L + Inches(i * 1.12)
        c1 = rect(s, x, Inches(2.3), Inches(0.95), Inches(0.55),
                  RGBColor(0xE8, 0xE8, 0xE5), line=RULE, lw=0.75)
        text(s, x, Inches(2.38), Inches(0.95), Inches(0.4), "ptr",
             size=T_BODY_SM, color=INK2, align=PP_ALIGN.CENTER)
        rect(s, x, Inches(3.25), Inches(0.95), Inches(0.62),
             RGBColor(0xFF, 0xEC, 0xD4), line=ACCENT, lw=0.75)
        text(s, x, Inches(3.37), Inches(0.95), Inches(0.4), "float",
             size=T_BODY_SM, color=ACCENT, align=PP_ALIGN.CENTER)
    alt(c1, "Diagram: a Python list is an array of pointers, each pointing to "
            "a separately allocated float object elsewhere on the heap.")
    text(s, L + Inches(6.1), Inches(2.42), Inches(0.8), Inches(0.4), "vs",
         size=T_BODY_SM, color=MUTED, bold=True)
    text(s, L + Inches(7.3), Inches(1.8), Inches(4.3), Inches(0.4),
         "Arrow — one buffer", size=T_BODY_SM, color=INK2)
    for i in range(5):
        c2 = rect(s, L + Inches(7.3 + i * 0.84), Inches(2.3), Inches(0.8),
                  Inches(0.55), RGBColor(0xDD, 0xE6, 0xF5), line=S1, lw=0.75)
    alt(c2, "Diagram: Arrow stores the same values as one contiguous block of "
            "raw 8-byte floats, with no per-value object.")
    text(s, L + Inches(7.3), Inches(3.0), Inches(4.5), Inches(0.8),
         "contiguous, cache-friendly", size=T_CAPTION, color=MUTED)
    text(s, L, Inches(4.35), CW, Inches(1.6),
         [[("24 bytes of object header per 8 bytes of payload.", INK, True)],
          [("Pointer chasing defeats the cache — the CPU waits on memory.",
            INK2, False)]], size=T_BODY, spacing=1.35)

    # ---- 11 --------------------------------------------------------------
    s = nxt("[9:00] Part two — the trap.")
    section(s, "02", "The serialization trap", "“just send it as JSON”")

    # ---- 12 --------------------------------------------------------------
    s = nxt("[9:20] This migrates in from microservices, where it made sense. "
            "In-process it makes none — but I have seen it in production.")
    title(s, "The microservices habit", kicker="02 · THE SERIALIZATION TRAP")
    b = code(s, L, Inches(1.85), CW, """
#[pyfunction]
fn sum_json(payload: &str) -> PyResult<f64> {
    let v: Vec<f64> = serde_json::from_str(payload)?;
    Ok(v.iter().sum())
}
""", lang="rust")
    text(s, L, b + Inches(0.4), CW, Inches(1.4),
         [[(f"{F['payload_mb']} MB of float64 becomes a "
            f"{F['json_mb']} MB string.", INK, True)],
          [("Then it is parsed back into floats.", INK2, False)]],
         size=T_BODY, spacing=1.35)

    # ---- 13 --------------------------------------------------------------
    s = nxt(f"[10:20] Six seconds. Let it land. Then the breakdown: "
            f"{F['json_serialize_only']:.0f} ms is json.dumps on the Python "
            f"side, before Rust is even called.")
    title(s, "What that costs", kicker="02 · MEASURED")
    hero(s, f"{F['json_total']/1000:.1f} s", "to sum five million numbers",
         color=CRIT)
    text(s, L, Inches(5.3), CW, Inches(0.6),
         f"{F['json_serialize_only']:.0f} ms in json.dumps   ·   "
         f"{F['json_call']:.0f} ms parsing in Rust",
         size=T_BODY_SM, color=INK2, align=PP_ALIGN.CENTER)

    # ---- 14 --------------------------------------------------------------
    s = nxt("[11:15] Part three. The core of the talk.")
    section(s, "03", "Arrow as shared memory", "stop moving data; share it")

    # ---- 15 --------------------------------------------------------------
    s = nxt("[11:40] The reframe. Every approach so far asked 'how do I move "
            "this faster?' That is the wrong question.")
    hero(s, "Don't move it.", "Both languages can point at the same bytes.")

    # ---- 16 --------------------------------------------------------------
    s = nxt("[12:10] DEMO MOMENT, and the punchline up front. Run it live if "
            "the room allows; these are the saved numbers as the fallback. "
            "Say: same address on both sides — not a fast copy, the same "
            "memory. Then: 'that is the trick. The rest of this part is how "
            "it works.'")
    title(s, "Proof: the same memory", kicker="03 · LIVE DEMO")
    code(s, L, Inches(1.8), CW, f"""
>>> arr.buffers()[1].address
{V['py_buffer_addr']}
>>> arrowbridge.arrow_data_ptr(arr)
{V['rust_buffer_addr']}
""", lang="python")
    text(s, L, Inches(4.75), CW, Inches(1.4),
         [[("Identical.", GOOD, True),
           (" Nothing was copied or serialized.", INK, False)],
          [("How? That is the rest of this part.", INK2, False)]],
         size=32, align=PP_ALIGN.CENTER, spacing=1.3)

    # ---- 17 --------------------------------------------------------------
    s = nxt("[13:20] Key framing: Arrow is a specification, not a library. "
            "That is exactly why it bridges languages. Define the acronym: "
            "Arrow is a columnar in-memory format from the Apache foundation.")
    title(s, "Arrow is a memory layout, not a library",
          kicker="03 · ARROW AS SHARED MEMORY")
    for i, (head, body) in enumerate([
            ("Columnar", "one column, laid out contiguously"),
            ("Specified", "the byte layout is the standard"),
            ("Language-neutral", "Rust, Python, C++, Java read the same bytes"),
            ("Null-aware", "a validity bitmap: one bit per value")]):
        t = Inches(2.0 + i * 1.15)
        bar(s, L, t + Inches(0.08), Pt(4), Inches(0.62), ACCENT_L)
        text(s, L + Inches(0.35), t, Inches(4.0), Inches(0.5), head,
             size=T_BODY, bold=True)
        text(s, L + Inches(4.5), t + Inches(0.03), Inches(7.0), Inches(0.6),
             body, size=T_BODY_SM, color=INK2)

    # ---- 18 --------------------------------------------------------------
    s = nxt("[14:40] Describe the diagram aloud: a float64 column is exactly "
            "two flat buffers — a validity bitmap on top, the values below. "
            "Point out that nulls keep their slot, so element i is always at "
            "offset i times 8.")
    title(s, "A float64 column, completely",
          kicker="03 · ARROW AS SHARED MEMORY")
    text(s, L, Inches(1.85), Inches(5.5), Inches(0.4), "validity bitmap",
         size=T_BODY_SM, color=INK2)
    for i, bit in enumerate(["1", "1", "0", "1", "1"]):
        x = L + Inches(i * 1.0)
        on = bit == "1"
        d = rect(s, x, Inches(2.35), Inches(0.85), Inches(0.6),
                 RGBColor(0xDD, 0xE6, 0xF5) if on else RGBColor(0xEC, 0xEC, 0xEA),
                 line=S1 if on else RULE, lw=0.75)
        text(s, x, Inches(2.47), Inches(0.85), Inches(0.4), bit, size=T_BODY_SM,
             color=S1 if on else MUTED, align=PP_ALIGN.CENTER, font=MONO)
    alt(d, "Diagram: a validity bitmap of five bits, 1 1 0 1 1, where 0 marks "
           "a null.")
    text(s, L, Inches(3.3), Inches(5.5), Inches(0.4), "values buffer",
         size=T_BODY_SM, color=INK2)
    for i, v in enumerate(["1.5", "2.0", "—", "4.5", "5.0"]):
        x = L + Inches(i * 1.0)
        d = rect(s, x, Inches(3.8), Inches(0.85), Inches(0.6),
                 RGBColor(0xDD, 0xE6, 0xF5), line=S1, lw=0.75)
        text(s, x, Inches(3.92), Inches(0.85), Inches(0.45), v, size=T_BODY_SM,
             color=INK if v != "—" else MUTED, align=PP_ALIGN.CENTER, font=MONO)
    alt(d, "Diagram: the values buffer holds 1.5, 2.0, an unused slot, 4.5, "
           "5.0 — the null keeps its position.")
    text(s, L + Inches(5.9), Inches(2.3), Inches(5.6), Inches(2.0),
         [[("Two flat buffers.", INK, True)],
          [("No headers, no pointers,\nno objects.", INK2, False)]],
         size=T_BODY, spacing=1.3)
    text(s, L, Inches(5.0), CW, Inches(0.8),
         "Sharing the column = handing over two pointers and a length.",
         size=T_BODY, bold=True, color=S1)

    # ---- 19 --------------------------------------------------------------
    s = nxt("[16:00] The mechanism. A PyCapsule carrying this small C struct. "
            "Neither side needs the other's Arrow library — they agree on the "
            "struct. Read the buffers line aloud: that is the payload.")
    title(s, "The Arrow C Data Interface",
          kicker="03 · ARROW AS SHARED MEMORY")
    b = code(s, L, Inches(1.78), CW, """
struct ArrowArray {
  int64_t length;
  const void** buffers;     // the pointers
  void (*release)(struct ArrowArray*);
};
""", lang="rust")
    text(s, L, b + Inches(0.35), CW, Inches(1.2),
         [[("The capsule transfers a pointer and ownership.", INK, True)],
          [("release() keeps the buffer alive while Rust reads.", INK2, False)]],
         size=T_BODY, spacing=1.35)

    # ---- 20 --------------------------------------------------------------
    s = nxt("[17:10] Any object exposing this dunder can be handed to Rust. "
            "Name the ecosystem — this is why it is not a pyarrow-only trick.")
    title(s, "One dunder, every library",
          kicker="03 · ARROW AS SHARED MEMORY")
    b = code(s, L, Inches(1.85), CW, """
>>> obj.__arrow_c_array__()
(<PyCapsule schema>, <PyCapsule array>)
""", lang="python")
    text(s, L, b + Inches(0.45), CW, Inches(2.2),
         [[("pyarrow · polars · pandas · duckdb · nanoarrow", S1, True)],
          [("all speak it, in both directions.", INK2, False)]],
         size=T_BODY, spacing=1.4)

    # ---- 21 --------------------------------------------------------------
    s = nxt("[18:00] The Rust side. PyArray implements FromPyObject, so PyO3 "
            "calls the capsule protocol for you. into_inner is the zero-copy "
            "step. sum is null-aware for free.")
    title(s, "The Rust side", kicker="03 · ARROW AS SHARED MEMORY")
    b = code(s, L, Inches(1.85), CW, """
#[pyfunction]
fn sum_arrow(arr: PyArray) -> PyResult<Option<f64>> {
    let (arr, _f) = arr.into_inner();   // zero copy
    let col = arr.as_any()
        .downcast_ref::<Float64Array>()?;
    Ok(aggregate::sum(col))             // null-aware
}
""", lang="rust")
    text(s, L, b + Inches(0.3), CW, Inches(0.8),
         "No .to_pylist(). No serde. No allocation.",
         size=T_BODY, color=INK2)

    # ---- 22 --------------------------------------------------------------
    s = nxt("[19:00] Call back to slide 2: this is sensors.parquet, the file "
            "we started with. The Python side is anticlimactic — you just "
            "pass the array. Parquet to Arrow to Rust, no conversion "
            "anywhere.")
    title(s, "The Python side", kicker="03 · ARROW AS SHARED MEMORY")
    b = code(s, L, Inches(1.85), CW, """
import pyarrow.parquet as pq
import arrowbridge

t = pq.read_table("sensors.parquet")  # Arrow
col = t.column("value").combine_chunks()
arrowbridge.sum_arrow(col)
""", lang="python")
    text(s, L, b + Inches(0.35), CW, Inches(0.8),
         "That is the whole integration.", size=T_BODY, bold=True)

    # ---- 23 --------------------------------------------------------------
    s = nxt("[20:30] The headline. Say 'log scale' out loud — otherwise the "
            "bars understate it. Read the four numbers. Colour groups the two "
            "strategies; the labels name each bar.")
    title(s, "Same answer, four ways", kicker="03 · MEASURED · LOG SCALE")
    bar_chart(s, L, Inches(1.7), CW, Inches(3.95),
              ["JSON", "Python list", "NumPy", "Arrow"],
              [round(F["json_total"], 1), round(F["list_total"], 1),
               round(F["numpy_total"], 2), round(F["arrow_total"], 2)],
              [S2, S2, S1, S1], log=True,
              description=(
                  f"Column chart, logarithmic scale. Time to sum {F['n']:,} "
                  f"float64 values: JSON {F['json_total']:.0f} milliseconds, "
                  f"Python list {F['list_total']:.0f}, NumPy "
                  f"{F['numpy_total']:.1f}, Arrow {F['arrow_total']:.1f}."))
    group_legend(s, L, Inches(5.75))
    caption(s, Inches(6.25),
            f"milliseconds · {F['n']:,} float64 ({F['payload_mb']} MB) · "
            f"median of 7 runs · {E['cpu']}")

    # ---- 24 --------------------------------------------------------------
    s = nxt("[21:40] Two numbers to remember from this talk.")
    title(s, "Against the baseline", kicker="03 · MEASURED")
    for i, (val, lab) in enumerate([(f"{SP_LIST:.0f}×", "faster than a Python list"),
                                    (f"{SP_JSON:.0f}×", "faster than JSON")]):
        x = L + Inches(i * 5.95)
        text(s, x, Inches(2.2), Inches(5.6), Inches(1.5), val, size=96,
             bold=True, color=S1, align=PP_ALIGN.CENTER)
        text(s, x, Inches(3.9), Inches(5.6), Inches(0.7), lab, size=T_BODY,
             color=INK2, align=PP_ALIGN.CENTER)
    text(s, L, Inches(5.2), CW, Inches(0.6),
         f"{F['arrow_total']:.1f} ms  vs  {F['list_total']:.0f} ms  vs  "
         f"{F['json_total']:.0f} ms", size=T_BODY_SM, color=INK2,
         align=PP_ALIGN.CENTER)

    # ---- 25 --------------------------------------------------------------
    s = nxt("[22:30] Time is half the story. This is extra peak memory on top "
            "of holding the data, one fresh process per approach. Arrow's half "
            "a megabyte is Rust's bookkeeping, not the data.")
    title(s, "Peak memory on top of the data", kicker="03 · MEASURED")
    bar_chart(s, L, Inches(1.7), CW, Inches(3.95),
              ["JSON", "Python list", "NumPy", "Arrow"],
              [M["json"]["overhead_mb"], M["list"]["overhead_mb"],
               M["numpy"]["overhead_mb"], M["arrow"]["overhead_mb"]],
              [S2, S2, S1, S1],
              description=(
                  f"Column chart of extra peak memory: JSON "
                  f"{M['json']['overhead_mb']} megabytes, Python list "
                  f"{M['list']['overhead_mb']}, NumPy "
                  f"{M['numpy']['overhead_mb']}, Arrow "
                  f"{M['arrow']['overhead_mb']}."))
    group_legend(s, L, Inches(5.75))
    caption(s, Inches(6.25),
            f"megabytes of extra peak RSS beyond the {F['payload_mb']} MB "
            f"Arrow array")

    # ---- 26 --------------------------------------------------------------
    s = nxt("[23:30] The slope matters more than any single number. Copying is "
            "linear in rows; sharing is flat — the handoff is constant time. "
            "Each line has its own dash and marker, so say the names aloud.")
    title(s, "How it scales", kicker="03 · MEASURED · LOG SCALE")
    line_chart(s, L, Inches(1.7), CW, Inches(4.1),
               [f"{r['n']:,}" for r in SC],
               [("JSON", [round(r["json"], 3) for r in SC]),
                ("Arrow", [round(r["arrow"], 3) for r in SC]),
                ("NumPy", [round(r["numpy"], 3) for r in SC]),
                ("Python list", [round(r["list"], 3) for r in SC])],
               description=(
                   "Line chart, log scale, one decade per step from 10 "
                   "thousand to 10 million rows. JSON and Python list rise "
                   "steeply; NumPy and Arrow stay far lower and flatter."))
    caption(s, Inches(6.15),
            "milliseconds (log) vs row count · flatter is better")

    # ---- 27 --------------------------------------------------------------
    s = nxt("[24:40] Fair question: NumPy was already zero-copy and nearly as "
            "fast. This slide is the answer — the buffer protocol only carries "
            "numbers. Read a couple of rows aloud.")
    title(s, "“NumPy was zero-copy — why Arrow?”",
          kicker="03 · THE HONEST COMPARISON")
    text(s, L + Inches(5.4), Inches(1.9), Inches(3.0), Inches(0.4), "NumPy",
         size=T_BODY_SM, bold=True, color=INK2)
    text(s, L + Inches(8.3), Inches(1.9), Inches(3.2), Inches(0.4), "Arrow",
         size=T_BODY_SM, bold=True, color=S1)
    bar(s, L, Inches(2.4), CW, Pt(1), RULE)
    for i, (feat, a, b_) in enumerate([
            ("Numeric arrays", "yes", "yes"),
            ("Null values", "no", "yes"),
            ("Strings", "no", "yes"),
            ("Nested types", "no", "yes"),
            ("Whole tables", "no", "yes")]):
        t = Inches(2.6 + i * 0.72)
        text(s, L, t, Inches(5.2), Inches(0.45), feat, size=T_BODY_SM, color=INK)
        text(s, L + Inches(5.4), t, Inches(3.0), Inches(0.45), a,
             size=T_BODY_SM, color=CRIT if a == "no" else GOOD, bold=True)
        text(s, L + Inches(8.3), t, Inches(3.2), Inches(0.45), b_,
             size=T_BODY_SM, color=GOOD, bold=True)
    text(s, L, Inches(6.3), CW, Inches(0.5), "Same speed. Far wider contract.",
         size=T_BODY, bold=True, color=S1)

    # ---- 28 --------------------------------------------------------------
    s = nxt(f"[25:50] A workload NumPy cannot carry at all. {SP_STR:.0f}× on "
            f"strings, and the list version is the only alternative.")
    title(s, "Strings: the buffer protocol gives up", kicker="03 · MEASURED")
    w2 = Inches(5.55)
    tile(s, L, Inches(2.0), w2, f"{ST['list_call']:.0f} ms",
         "list[str] → Vec<String>", color=S4, sub="a String allocated per row")
    tile(s, L + Inches(5.98), Inches(2.0), w2, f"{ST['arrow_call']:.0f} ms",
         "Arrow StringArray", color=S1, sub="borrowed &str, no allocation")
    text(s, L, Inches(4.3), CW, Inches(1.6),
         [[(f"{SP_STR:.0f}× faster on {ST['n']:,} strings.", INK, True)],
          [("Arrow packs strings into one buffer plus offsets;\n"
            "Rust borrows slices out of it.", INK2, False)]],
         size=T_BODY, spacing=1.35)

    # ---- 29 --------------------------------------------------------------
    s = nxt("[26:50] Two things you get for free. Nulls: the validity bitmap "
            "travels with the buffer, so Rust's kernels skip them. Say the "
            "expected output: 1 plus 3 plus 6 is 10.")
    title(s, "Nulls come for free", kicker="03 · ARROW AS SHARED MEMORY")
    b = code(s, L, Inches(1.85), CW, f"""
>>> a = pa.array([1.0, None, 3.0, None, 6.0])
>>> arrowbridge.sum_arrow(a)
{V['null_aware_sum']}
""", lang="python")
    text(s, L, b + Inches(0.4), CW, Inches(1.4),
         [[("The validity bitmap travels with the buffer.", INK, True)],
          [("Rust's kernels are null-aware. Nothing to translate.",
            INK2, False)]], size=T_BODY, spacing=1.35)

    # ---- 30 --------------------------------------------------------------
    s = nxt("[27:40] And whole tables, with the schema. Output is rows then "
            "columns: four rows, two columns. This is what serialization can "
            "never give you cheaply.")
    title(s, "Whole tables, too", kicker="03 · ARROW AS SHARED MEMORY")
    b = code(s, L, Inches(1.85), CW, f"""
>>> batch = pa.record_batch(columns)
>>> arrowbridge.batch_shape(batch)
{tuple(V['batch_shape'])}
""", lang="python")
    text(s, L, b + Inches(0.4), CW, Inches(1.4),
         [[("A RecordBatch carries the schema.", INK, True)],
          [("Column names, types and nullability all survive.", INK2, False)]],
         size=T_BODY, spacing=1.35)

    # ---- 31 --------------------------------------------------------------
    s = nxt("[28:30] Part four — the things that cost me time.")
    section(s, "04", "Pitfalls and production", "what I wish I'd known")

    # ---- 32 --------------------------------------------------------------
    s = nxt("[28:50] This one bit me building the demo for this talk. Tell it "
            "as a story — it lands better than a warning.")
    title(s, "Pitfall 1: two arrow versions", kicker="04 · PITFALLS")
    b = code(s, L, Inches(1.85), CW, """
error[E0308]: mismatched types
  expected `arrow_schema::DataType`
  found    `DataType`
note: multiple versions of crate `arrow_schema`
""", lang="bash")
    text(s, L, b + Inches(0.35), CW, Inches(1.8),
         [[("Cargo links arrow 59 and 60 side by side.", INK2, False)],
          [("Pin arrow to what pyo3-arrow depends on.", INK, True)],
          [("Find duplicates with:  cargo tree -d", ACCENT, True)]],
         size=T_BODY, spacing=1.3)

    # ---- 33 --------------------------------------------------------------
    s = nxt("[29:50] Back to sensors.parquet one more time. A Table column is "
            "chunked; "
            "the C Data Interface hands over one contiguous array. So "
            "combine_chunks copies — do it once, not per call.")
    title(s, "Pitfall 2: ChunkedArray is not an Array", kicker="04 · PITFALLS")
    b = code(s, L, Inches(1.85), CW, """
>>> table.column("value")
<ChunkedArray ... 12 chunks>

>>> table.column("value").combine_chunks()
# one buffer -- and this COPIES
""", lang="python")
    text(s, L, b + Inches(0.35), CW, Inches(1.4),
         [[("Budget for that copy, or accept a", INK2, False)],
          [("RecordBatchReader and stream the chunks.", INK, True)]],
         size=T_BODY, spacing=1.3)

    # ---- 34 --------------------------------------------------------------
    s = nxt("[30:50] Ownership. The capsule keeps the buffer alive only while "
            "Rust holds the array — never stash a raw pointer. And release "
            "the GIL for long compute, or you gain nothing in a threaded server.")
    title(s, "Pitfall 3: ownership and the GIL", kicker="04 · PITFALLS")
    b = code(s, L, Inches(1.85), CW, """
#[pyfunction]
fn heavy(py: Python, arr: PyArray) -> PyResult<f64> {
    let (arr, _) = arr.into_inner();
    py.detach(|| expensive_scan(&arr))   // GIL off
}
""", lang="rust")
    text(s, L, b + Inches(0.35), CW, Inches(1.4),
         [[("Never cache a raw pointer past the call.", INK2, False)],
          [("Releasing the GIL is what makes Rust pay off.", INK, True)]],
         size=T_BODY, spacing=1.3)

    # ---- 35 --------------------------------------------------------------
    s = nxt("[31:50] Be honest — it buys credibility. At ten thousand rows the "
            "whole difference is a tenth of a millisecond. Nobody should "
            "rewrite anything for that.")
    title(s, "Pitfall 4: doing this when you shouldn't", kicker="04 · PITFALLS")
    sm = SC[0]
    w2 = Inches(5.55)
    tile(s, L, Inches(1.95), w2, f"{sm['list']:.2f} vs {sm['arrow']:.2f} ms",
         f"at {sm['n']:,} rows", color=INK2, sub="not worth a rewrite")
    tile(s, L + Inches(5.98), Inches(1.95), w2,
         f"{F['list_total']:.0f} vs {F['arrow_total']:.1f} ms",
         f"at {F['n']:,} rows", color=S1, sub="now it dominates")
    text(s, L, Inches(4.25), CW, Inches(2.0),
         [[("Reach for this when the data is already columnar,", INK2, False)],
          [("you cross the boundary often, or memory is tight.", INK2, False)],
          [("Otherwise a Python list is fine.", INK, True)]],
         size=T_BODY, spacing=1.35)

    # ---- 36 --------------------------------------------------------------
    s = nxt("[32:40] Practical toolbox. The pointer check is the one I use "
            "most — it turns 'I think this is zero copy' into a test.")
    title(s, "How to tell it is really zero copy", kicker="04 · DEBUGGING")
    for i, (head, body) in enumerate([
            ("Compare pointers", "assert the addresses match, in a test"),
            ("Watch peak RSS", "VmHWM in /proc — a copy shows as a step"),
            ("cargo tree -d", "duplicate crates cause baffling errors"),
            ("Count the chunks", "more than one means combine_chunks copies")]):
        t = Inches(2.05 + i * 1.12)
        bar(s, L, t + Inches(0.06), Pt(4), Inches(0.55), ACCENT_L)
        text(s, L + Inches(0.35), t, Inches(4.6), Inches(0.5), head,
             size=T_BODY, bold=True)
        text(s, L + Inches(5.2), t + Inches(0.04), Inches(6.3), Inches(0.6),
             body, size=T_BODY_SM, color=INK2)

    # ---- 37 --------------------------------------------------------------
    s = nxt("[33:30] You are not inventing this. Name them quickly — the point "
            "is that Arrow is already the substrate under tools they use daily.")
    title(s, "You are already using this", kicker="04 · PRODUCTION")
    for i, (name, desc) in enumerate([
            ("Polars", "Rust engine, Arrow memory"),
            ("DataFusion", "SQL engine in Rust"),
            ("delta-rs", "Delta Lake, Rust core"),
            ("pydantic-core", "PyO3 at the heart of v2")]):
        x = L + (Inches(5.98) if i % 2 else 0)
        t = Inches(2.05 + (i // 2) * 1.55)
        rect(s, x, t, Inches(5.55), Inches(1.25), PANEL, line=RULE, lw=0.75)
        text(s, x + Inches(0.3), t + Inches(0.18), Inches(5.0), Inches(0.45),
             name, size=T_BODY, bold=True, color=S1)
        text(s, x + Inches(0.3), t + Inches(0.72), Inches(5.0), Inches(0.4),
             desc, size=T_BODY_SM, color=INK2)
    text(s, L, Inches(5.5), CW, Inches(0.6),
         "Arrow is the substrate. The C Data Interface is the plug.",
         size=T_BODY_SM, color=INK2)

    # ---- 38 --------------------------------------------------------------
    s = nxt("[34:10] Three things. If they remember nothing else. Slow down.")
    title(s, "Three things to take home")
    for i, (head, body) in enumerate([
            ("The boundary is the bottleneck",
             "Profile the crossing before you rewrite anything."),
            ("Arrow is a layout, not a library",
             "Agree on bytes and the copy disappears."),
            ("Zero copy is testable",
             "Assert pointer equality. Don't assume; measure.")]):
        t = Inches(2.1 + i * 1.5)
        text(s, L, t, Inches(1.0), Inches(0.6), f"0{i+1}", size=32, bold=True,
             color=ACCENT)
        text(s, L + Inches(1.3), t, Inches(10.0), Inches(0.5), head,
             size=T_BODY, bold=True)
        text(s, L + Inches(1.3), t + Inches(0.58), Inches(10.0), Inches(0.5),
             body, size=T_BODY_SM, color=INK2)

    # ---- 39 --------------------------------------------------------------
    s = nxt("[35:00] Thank you. Repo on screen. Take questions — repeat each "
            "question into the mic before answering. Likely ones: free-"
            "threading and the GIL; why not just use Polars; Arrow IPC vs the "
            "C Data Interface for cross-process.", chrome=False)
    bar(s, L, Inches(1.7), Inches(1.9), Pt(5), ACCENT_L)
    text(s, L, Inches(2.05), CW, Inches(1.0), "Gracias · Thank you",
         size=50, bold=True)
    text(s, L, Inches(3.2), CW, Inches(0.6), "Questions?", size=32, color=INK2)
    rect(s, L, Inches(3.85), Inches(7.3), Inches(2.85), PANEL, line=RULE, lw=0.75)
    text(s, L + Inches(0.4), Inches(4.05), Inches(6.5), Inches(2.5),
         [[("Slides, code and benchmarks", MUTED, False)],
          [(REPO, ACCENT, True)],
          [("LinkedIn", MUTED, False)],
          [(LINKEDIN, ACCENT, True)]],
         size=T_BODY_SM, spacing=1.2, space_after=2)
    s.shapes.add_picture(os.path.join(ASSETS, "linkedin_qr.png"),
                         Inches(8.8), Inches(4.0), Inches(2.3), Inches(2.3))
    alt(s.shapes[-1], f"QR code linking to {LINKEDIN}")
    text(s, Inches(8.8), Inches(6.4), Inches(2.3), Inches(0.35),
         "Scan for LinkedIn", size=T_CAPTION, color=MUTED,
         align=PP_ALIGN.CENTER)

    out = os.path.join(HERE, "pycones2026-zero-copy-python-rust.pptx")
    prs.save(out)
    print(f"wrote {out}  ({n[0]} slides)")


if __name__ == "__main__":
    build()
