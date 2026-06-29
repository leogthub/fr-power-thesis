# -*- coding: utf-8 -*-
"""Generate Script_Soutenance.docx — 16 slides, ~17 minutes, spoken English."""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

NAVY   = RGBColor(0x00, 0x22, 0x55)
GOLD   = RGBColor(0xC8, 0xA0, 0x32)
GREY   = RGBColor(0x66, 0x66, 0x66)
RED    = RGBColor(0xC0, 0x39, 0x2B)
GREEN  = RGBColor(0x1A, 0x73, 0x4A)
LIGHT  = RGBColor(0xEE, 0xF3, 0xF9)

doc = Document()
for s in doc.sections:
    s.top_margin = s.bottom_margin = Cm(2.0)
    s.left_margin = s.right_margin = Cm(2.5)

# ── helpers ───────────────────────────────────────────────────
def border_bottom(p, color='002255', sz='8'):
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bot = OxmlElement('w:bottom')
    bot.set(qn('w:val'), 'single'); bot.set(qn('w:sz'), sz)
    bot.set(qn('w:space'), '4');    bot.set(qn('w:color'), color)
    pBdr.append(bot); pPr.append(pBdr)

def add_run(p, text, bold=False, italic=False, color=None, size=11):
    r = p.add_run(text)
    r.bold = bold; r.italic = italic
    r.font.size = Pt(size)
    if color: r.font.color.rgb = color
    return r

def h1(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after  = Pt(2)
    add_run(p, text, bold=True, color=NAVY, size=17)
    border_bottom(p, '002255', '10')

def slide_header(number, title, speaker, timing):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(16)
    p.paragraph_format.space_after  = Pt(4)
    add_run(p, f"SLIDE {number} — {title.upper()}   ", bold=True, color=NAVY, size=13)
    add_run(p, f"[{speaker}]", bold=True, color=GOLD, size=12)
    add_run(p, f"  · {timing}", italic=True, color=GREY, size=11)
    border_bottom(p, 'C8A032', '4')

def spoken(text, indent=0.5):
    """Main spoken text — body of the script."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(6)
    p.paragraph_format.left_indent  = Cm(indent)
    add_run(p, text, size=12)
    return p

def pause():
    """Visual empty line to indicate a natural pause."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)
    return p

def beat(text):
    """Key beat — bold, draws attention."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after  = Pt(6)
    p.paragraph_format.left_indent  = Cm(0.5)
    add_run(p, text, bold=True, color=NAVY, size=12)
    return p

def cue(text):
    """Stage direction / transition cue — italic grey."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after  = Pt(6)
    p.paragraph_format.left_indent  = Cm(0.5)
    add_run(p, text, italic=True, color=GREY, size=11)
    return p

def divider():
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after  = Pt(6)
    border_bottom(p, 'CCCCCC', '4')

def timing_table():
    rows_data = [
        ("1 — Title",            "LEO",       "20 s"),
        ("2 — Context",          "LEO",       "1 min 30 s"),
        ("3 — Research Q",       "LEO",       "1 min"),
        ("4 — Data",             "LEO",       "1 min"),
        ("5 — Methodology",      "LEO",       "1 min"),
        ("6 — Results stable",   "LEO",       "1 min 30 s"),
        ("7 — Central Finding",  "LEO",       "1 min"),
        ("8 — Feature Imp.",     "LYAM",      "1 min"),
        ("9 — Redundancy",       "LYAM",      "1 min 30 s"),
        ("10 — 2022 Crisis",     "LYAM",      "1 min 30 s"),
        ("11 — Why 2022",        "LYAM",      "1 min 30 s"),
        ("12 — Trading",         "LYAM",      "1 min"),
        ("13 — Robustness",      "LYAM",      "30 s"),
        ("14 — Limitations",     "LYAM",      "45 s"),
        ("15 — Conclusions",     "LYAM",      "1 min 30 s"),
        ("16 — Questions",       "BOTH",      "20 s"),
        ("TOTAL",                "",          "~17 min"),
    ]
    table = doc.add_table(rows=1+len(rows_data), cols=3)
    table.style = 'Table Grid'
    widths = [Cm(7.5), Cm(3.5), Cm(3.2)]
    for ci, h in enumerate(["Slide", "Speaker", "Target"]):
        cell = table.rows[0].cells[ci]
        cell.width = widths[ci]
        r = cell.paragraphs[0].add_run(h)
        r.bold = True; r.font.color.rgb = RGBColor(255,255,255); r.font.size = Pt(11)
        tc = cell._tc; tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'),'clear'); shd.set(qn('w:color'),'auto'); shd.set(qn('w:fill'),'002255')
        tcPr.append(shd)
    for ri, (s, sp, t) in enumerate(rows_data):
        row = table.rows[ri+1]
        is_total = s == "TOTAL"
        fill = 'EEF3F9' if ri % 2 == 0 else 'FFFFFF'
        for ci, txt in enumerate([s, sp, t]):
            cell = row.cells[ci]
            cell.width = widths[ci]
            r = cell.paragraphs[0].add_run(txt)
            r.bold = is_total; r.font.size = Pt(11)
            if is_total: r.font.color.rgb = NAVY
            tc = cell._tc; tcPr = tc.get_or_add_tcPr()
            shd = OxmlElement('w:shd')
            shd.set(qn('w:val'),'clear'); shd.set(qn('w:color'),'auto')
            shd.set(qn('w:fill'), 'D6E4F0' if is_total else fill)
            tcPr.append(shd)
    doc.add_paragraph()


# ══════════════════════════════════════════════════════════════
# DOCUMENT HEADER
# ══════════════════════════════════════════════════════════════
h1("ORAL DEFENSE SCRIPT")
p = doc.add_paragraph()
add_run(p, "The Role of Weather in French Day-Ahead Electricity Price Forecasting",
        bold=True, color=NAVY, size=12)
p = doc.add_paragraph()
add_run(p, "Leo Cambreleng & Lyam Oumedjeber  ·  EDHEC MSc DAAI  ·  June 2026",
        color=GREY, size=10)
p = doc.add_paragraph()
add_run(p, "LEO slides 1–7 (~7 min)  ·  LYAM slides 8–15 (~9 min 40 s)  ·  BOTH slide 16  ·  Total ~17 min",
        italic=True, color=GREY, size=10)
p = doc.add_paragraph()
add_run(p, "Rule: don't read the slide — comment on it. Guide attention. Add what isn't written.",
        italic=True, color=RED, size=10)
divider()

# ══════════════════════════════════════════════════════════════
# SLIDES
# ══════════════════════════════════════════════════════════════

slide_header(1, "Title", "LEO", "~20 s")
spoken("Good morning, Professor.")
pause()
spoken("I'm Leo — with Lyam, we spent the last year asking a very specific question about the French electricity market.")
pause()
spoken("I'll set it up, and Lyam will deliver the punchline.")
divider()

slide_header(2, "Context & Motivation", "LEO", "~1 min 30 s")
spoken("Before we get into the method — why does this market in particular make the question interesting?")
pause()
spoken("Three things make France unusual.")
pause()
spoken("Electricity can't be stored — so every mismatch between supply and demand shows up immediately in the price. That's why you get these violent spikes. It's structurally different from any financial market.")
pause()
spoken("France adds a layer on top: 70% nuclear, and most heating is electric. So a cold snap doesn't just increase demand — it hits a system with very little flexibility. 2.4 gigawatts per degree. The highest in Europe.")
pause()
spoken("And the practical reason we care: a one-euro improvement in forecast accuracy on a typical trading book is close to a million euros a year.")
pause()
spoken("So the question is: where does that improvement come from? Is weather one of the answers?")
divider()

slide_header(3, "Research Question", "LEO", "~1 min")
spoken("The question in the box is the one we set out to answer.")
pause()
spoken("But we didn't treat it as a yes/no. We broke it into four parts — you can see them on the slide.")
pause()
spoken("The first two are about forecast accuracy. The third — does the answer change depending on what the market is doing — is the one that turned out to be most interesting. The fourth connects it to something concrete: money.")
pause()
spoken("Those four questions are the backbone of everything that follows.")
divider()

slide_header(4, "Data", "LEO", "~1 min")
spoken("On the left, the four data sources — nothing exotic, all public.")
pause()
spoken("The interesting one is ERA5. It's not a weather forecast — it's a reanalysis. ECMWF takes actual observations and reconstructs the best possible version of past weather. We'll come back to why that distinction matters.")
pause()
spoken("The chart on the right is worth a look before we go further. Three periods: the flat pre-crisis years, the 2022 spike above 1,000 euros, then the return to normal. That's not just context — it's the core of our experiment.")
divider()

slide_header(5, "Methodology", "LEO", "~1 min")
spoken("The design is simple — and deliberately so.")
pause()
spoken("Four models. One controlled comparison.")
pause()
spoken("B and C are identical in every way except one: C gets the weather variables, B doesn't. That's the ablation. If C is better than B, weather helped. If not, the signal was already there without it.")
pause()
spoken("We test that difference formally with the Diebold-Mariano test — the standard in this literature for comparing forecast accuracy.")
divider()

slide_header(6, "Results — Stable Regime", "LEO", "~1 min 30 s")
spoken("Look at the table — specifically the last two rows.")
pause()
spoken("B and C. 16.95 and 16.94. One hundredth of a euro difference. On a baseline of 33 euros, that's noise.")
pause()
spoken("Now look at A versus B — 33 down to 16.9. That's the machine learning contribution. That's real.")
pause()
spoken("XGBoost underperforms here — we'll explain why that's actually informative.")
pause()
beat("But the headline: weather, in a calm market, doesn't move the needle.")
divider()

slide_header(7, "Central Finding", "LEO", "~1 min")
spoken("This slide formalises what we just saw.")
pause()
spoken("The top line — p = 0.572. We fail to reject equal accuracy between C and B. Weather is not statistically significant in the stable regime.")
pause()
spoken("Now look at the other rows. When we test ML against the naïve benchmark, the DM statistic is minus 53. That's enormous — essentially impossible to get by chance.")
pause()
spoken("So we have two very different stories in the same table: ML adds unambiguous value. Weather, right now, does not.")
pause()
cue("\"Lyam will take it from here.\"")
divider()

slide_header(8, "Feature Importance", "LYAM", "~1 min")
spoken("Thanks Leo. This chart answers the \"why\" before I even state it.")
pause()
spoken("Look at the bar lengths. The top five features are all price lags and rolling averages. Price history explains about 70% of what the model knows.")
pause()
spoken("Weather sits at the bottom. All variables combined — temperature, wind, solar, everything — about 2%.")
pause()
spoken("That's not a modelling flaw. It's telling us something real: weather's signal is already in the data — just encoded differently.")
divider()

slide_header(9, "Information Redundancy Hypothesis", "LYAM", "~1 min 30 s")
spoken("Three boxes on this slide — three reasons why weather is redundant.")
pause()
spoken("The first one is the most elegant. RTE publishes a load forecast every morning — and that forecast is built using temperature data. So when our model sees the load forecast, it's already implicitly seeing the weather. Adding raw temperature on top of that adds nothing.")
pause()
spoken("The second: TTF gas prices. Gas is the marginal fuel in Europe. When it's cold everywhere, TTF rises. The commodity market has already aggregated the weather signal for us.")
pause()
spoken("The third: nuclear dominates the French price so much that weather effects are small in comparison.")
pause()
cue("\"As long as they hold. In 2022, none of them did.\"")
divider()

slide_header(10, "2022 Crisis Reversal", "LYAM", "~1 min 30 s")
spoken("The two cards at the top tell the whole story.")
pause()
spoken("Same test. Same models. Different market.")
pause()
beat("p = 0.572 becomes p < 0.001.   DM = −0.565 becomes DM = −13.27.")
pause()
spoken("In this literature, a DM statistic of minus 3 or 4 is already considered strong. Minus 13 is almost never seen.")
pause()
spoken("The chart on the right shows why — 2022 is a completely different distribution. Models degrade badly across the board. But the gap between C and B, which was invisible before, becomes detectable and consistent.")
divider()

slide_header(11, "Why the Crisis Breaks Redundancy", "LYAM", "~1 min 30 s")
spoken("Each of the three channels — one by one, they break.")
pause()
spoken("TTF first. In 2022, gas prices were driven by the Russian supply shock, not by temperature. The link between weather and gas prices snapped. So TTF stopped encoding the weather signal.")
pause()
spoken("Nuclear second. 40% availability — half the normal level. In a well-functioning system, nuclear absorbs demand variations. At 40%, every cold day becomes a price event.")
pause()
spoken("And together, those two effects push price swings into the hundreds of euros — which is when the temperature signal finally becomes statistically measurable.")
pause()
beat("Weather isn't universally useful or useless. It depends on what the rest of the market is doing.")
divider()

slide_header(12, "Trading Backtest", "LYAM", "~1 min")
spoken("This slide answers a simple question: do the forecast improvements translate into anything real?")
pause()
spoken("Focus on the drawdown column.")
pause()
beat("Naïve: 2,367 euros of drawdown per megawatt.   RF models: 422. Five times smaller.")
pause()
spoken("That's the number that matters in practice. It's not about making more money in absolute terms — it's about how much you lose when you're wrong.")
pause()
spoken("And the strategy earns positive Sharpe ratios every single month of the test period. No lucky quarter hiding a bad year. Consistent.")
divider()

slide_header(13, "Robustness Checks", "LYAM", "~30 s")
spoken("One concern: do these results depend on the cost assumption we chose?")
pause()
spoken("The table shows three scenarios. The Sharpe barely moves — under two percent degradation from cheapest to most expensive.")
pause()
spoken("The conclusion is stable. It's not an artefact of our parameters.")
divider()

slide_header(14, "Limitations", "LYAM", "~45 s")
spoken("The most important one is the ERA5 point.")
pause()
spoken("ERA5 gives us the actual weather that happened. In production, you'd only have a forecast — with real errors. So our result is a ceiling on how much weather can help. Even the perfect version barely matters in a stable market. That actually strengthens the conclusion rather than weakening it.")
pause()
spoken("The other five — no walk-forward retraining, no carbon prices, XGBoost not tuned, one country, flat cost model — are genuine. We document them because they define the scope of what we can claim.")
divider()

slide_header(15, "Conclusions", "LYAM", "~1 min 30 s")
spoken("Four things to take away.")
pause()
spoken("ML works — clearly and robustly. That's not the surprise.")
pause()
spoken("The surprise is the second point. Weather matters — but only when the channels that normally encode it break down. In stable conditions, it's redundant. In a crisis, it's essential. That's a regime-dependent finding, and it's the core contribution.")
pause()
spoken("Random Forest holds up better than XGBoost under stress — a direct implication for model selection in volatile markets.")
pause()
spoken("And forecasts create real economic value — not just better numbers on a metric.")
pause()
beat("The practical takeaway: build a system that knows which regime it's in.")
pause()
spoken("Thank you.")
divider()

slide_header(16, "Questions", "LEO + LYAM", "~20 s")
cue("[Smile. Wait. Let them speak first.]")
cue("[If silence: \"Would you like to start with any particular aspect?\"]")
divider()

# ══════════════════════════════════════════════════════════════
# TIMING
# ══════════════════════════════════════════════════════════════
doc.add_paragraph()
h1("TIMING REFERENCE")
timing_table()

# ══════════════════════════════════════════════════════════════
# TRANSITIONS
# ══════════════════════════════════════════════════════════════
h1("KEY TRANSITIONS")
p = doc.add_paragraph()
add_run(p, "Slide 7 → 8  (LEO to LYAM):  ", bold=True, color=NAVY)
add_run(p, '"Lyam will take it from here."', italic=True, size=11)
p = doc.add_paragraph()
add_run(p, "Slide 9 → 10  (LYAM continues):  ", bold=True, color=NAVY)
add_run(p, '"As long as they hold. In 2022, none of them did."', italic=True, size=11)
p = doc.add_paragraph()
add_run(p, "End of slide 15:  ", bold=True, color=NAVY)
add_run(p, '"Thank you."', italic=True, size=11)

OUT = r"C:\Users\Public\fr-power-thesis\outputs\Script_Soutenance.docx"
doc.save(OUT)
print("Saved:", OUT)
