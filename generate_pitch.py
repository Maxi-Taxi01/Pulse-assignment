"""
generate_pitch.py
-----------------
Generates a short pitch presentation (PPTX) based on the research report
"Onderzoeksverslag Lab Persoonlijk.pdf".

The output includes:
  - Slide 1  : Title
  - Slide 2  : Aanleiding & Context
  - Slide 3  : Doel & Centrale Onderzoeksvraag
  - Slide 4  : Aanpak – 3 Testcycli
  - Slide 5  : Deelnemers & Dataverzameling
  - Slide 6  : Succescriteria & Deliverables
  - Slide 7  : Planning
  - Slide 8  : Conclusie & Aanbevelingen
  - Slide 9  : Concept Idee (empty template slide)

Usage:
    python generate_pitch.py
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import pdfplumber

# ---------------------------------------------------------------------------
# Colour palette (matches the existing design)
# ---------------------------------------------------------------------------
C_PURPLE      = RGBColor(0x6B, 0x21, 0xA8)   # deep purple  #6B21A8
C_PURPLE_MID  = RGBColor(0x7C, 0x3A, 0xED)   # mid purple   #7C3AED
C_PURPLE_LIGHT= RGBColor(0xED, 0xE9, 0xFE)   # lavender     #EDE9FE
C_WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
C_DARK        = RGBColor(0x1E, 0x1B, 0x4B)   # dark indigo  #1E1B4B
C_GRAY        = RGBColor(0x6B, 0x72, 0x80)
C_ACCENT      = RGBColor(0xF5, 0x9E, 0x0B)   # amber        #F59E0B

# ---------------------------------------------------------------------------
# Slide dimensions  (16:9  widescreen)
# ---------------------------------------------------------------------------
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

# ---------------------------------------------------------------------------
# Helper utilities
# ---------------------------------------------------------------------------

def add_rect(slide, left, top, width, height, fill_color, transparency=0):
    """Add a solid-filled rectangle shape."""
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    fill = shape.fill
    fill.solid()
    fill.fore_color.rgb = fill_color
    shape.line.fill.background()
    if transparency:
        fill.fore_color.theme_color  # noop – transparency not exposed in python-pptx
    return shape


def add_textbox(slide, text, left, top, width, height,
                font_size=18, bold=False, color=C_DARK,
                align=PP_ALIGN.LEFT, wrap=True):
    """Add a text box."""
    txBox = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    return txBox


def add_header_bar(slide, title_text, slide_num, total=9):
    """Purple top bar with slide title and page number."""
    add_rect(slide, 0, 0, 13.333, 1.0, C_PURPLE)
    add_textbox(slide, title_text,
                left=0.35, top=0.15, width=11.5, height=0.7,
                font_size=24, bold=True, color=C_WHITE)
    add_textbox(slide, f"{slide_num} / {total}",
                left=11.8, top=0.28, width=1.2, height=0.45,
                font_size=13, bold=False, color=C_WHITE, align=PP_ALIGN.RIGHT)


def add_footer(slide, text):
    """Light footer bar."""
    add_rect(slide, 0, 7.1, 13.333, 0.4, C_PURPLE_LIGHT)
    add_textbox(slide, text,
                left=0.35, top=7.13, width=12.6, height=0.35,
                font_size=9, color=C_GRAY)


def add_bullet_box(slide, title, bullets, left, top, width, height,
                   accent_color=C_PURPLE):
    """A card with a coloured title bar and bullet list."""
    # card background
    card = slide.shapes.add_shape(1,
        Inches(left), Inches(top), Inches(width), Inches(height))
    card.fill.solid()
    card.fill.fore_color.rgb = C_PURPLE_LIGHT
    card.line.color.rgb = accent_color
    card.line.width = Pt(1.5)

    # title strip
    title_h = 0.4
    t_strip = slide.shapes.add_shape(1,
        Inches(left), Inches(top), Inches(width), Inches(title_h))
    t_strip.fill.solid()
    t_strip.fill.fore_color.rgb = accent_color
    t_strip.line.fill.background()
    add_textbox(slide, title,
                left=left + 0.08, top=top + 0.04,
                width=width - 0.16, height=title_h - 0.06,
                font_size=11, bold=True, color=C_WHITE)

    # bullets
    bullet_top = top + title_h + 0.08
    bullet_h = (height - title_h - 0.16) / max(len(bullets), 1)
    for i, b in enumerate(bullets):
        add_textbox(slide, f"▸  {b}",
                    left=left + 0.12,
                    top=bullet_top + i * bullet_h,
                    width=width - 0.24,
                    height=bullet_h,
                    font_size=10, color=C_DARK)


def add_pill(slide, text, left, top, width=2.2, height=0.45,
             bg=C_PURPLE_MID, fg=C_WHITE):
    """Rounded pill / tag shape."""
    pill = slide.shapes.add_shape(
        5,  # ROUNDED_RECTANGLE
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    pill.fill.solid()
    pill.fill.fore_color.rgb = bg
    pill.line.fill.background()
    tf = pill.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = text
    run.font.size = Pt(11)
    run.font.bold = True
    run.font.color.rgb = fg


# ---------------------------------------------------------------------------
# Individual slide builders
# ---------------------------------------------------------------------------

def slide_title(prs):
    """Slide 1 – Title."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank

    # Background gradient simulation – two rectangles
    add_rect(slide, 0, 0, 13.333, 7.5, C_PURPLE)
    add_rect(slide, 0, 4.5, 13.333, 3.0, C_PURPLE_MID)

    # Decorative accent bar
    add_rect(slide, 0, 3.3, 13.333, 0.06, C_ACCENT)

    # Main title
    add_textbox(slide, "LAB PERSOONLIJK",
                left=1.0, top=1.2, width=11.3, height=1.2,
                font_size=52, bold=True, color=C_WHITE,
                align=PP_ALIGN.CENTER)

    # Subtitle
    add_textbox(slide, "Pitch Presentatie – Onderzoeksverslag",
                left=1.0, top=2.5, width=11.3, height=0.7,
                font_size=22, color=C_WHITE, align=PP_ALIGN.CENTER)

    # Three pill tags
    tags = [("Eenzaamheid ↓", 1.8), ("Veerkracht ↑", 4.8), ("Zorgdruk ↓", 7.8)]
    for label, x in tags:
        add_pill(slide, label, x, 3.6, width=2.4, height=0.5)

    # Footer info
    add_textbox(slide,
                "Develop & Experiment  |  Pulse-project  |  Januari – Juni 2026",
                left=1.0, top=6.6, width=11.3, height=0.5,
                font_size=12, color=C_WHITE, align=PP_ALIGN.CENTER)


def slide_context(prs):
    """Slide 2 – Aanleiding & Context."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header_bar(slide, "Aanleiding & Context", 2)

    cards = [
        ("😰  Het probleem",
         ["Formele zorgcontacten zijn te kort, taakgericht of belast",
          "Geen ruimte voor diepgaande interactie"],
         C_PURPLE),
        ("💡  De kans",
         ["Jongeren, ouderen én professionals willen authentiek contact",
          "Maar de 'kapstok' om dat gesprek te starten ontbreekt"],
         C_PURPLE_MID),
        ("🎯  Ons antwoord",
         ["Lab Persoonlijk test laagdrempelige creatieve interventies",
          "Spelvormen, gespreksstarters en maakopdrachten als katalysator"],
         RGBColor(0x05, 0x96, 0x69)),
    ]
    for i, (title, bullets, color) in enumerate(cards):
        add_bullet_box(slide, title, bullets,
                       left=0.4 + i * 4.3,
                       top=1.3, width=4.0, height=2.2,
                       accent_color=color)

    # Context block
    add_rect(slide, 0.4, 3.8, 12.5, 2.8, C_PURPLE_LIGHT)
    add_textbox(slide, "Pulse-programma context",
                left=0.6, top=3.9, width=8, height=0.4,
                font_size=12, bold=True, color=C_PURPLE)
    context_text = (
        "Pulse ontwikkelt via co-creatieve, ontwerpgerichte aanpak nieuwe verbindingsvormen "
        "tussen jongeren, ouderen en zorgprofessionals. Binnen de fase Develop & Experiment "
        "(jan–jun 2026) worden drie labs uitgevoerd. Lab Persoonlijk richt zich op directe, "
        "laagdrempelige creatieve interventies in de dagelijkse zorg- en leefomgeving."
    )
    add_textbox(slide, context_text,
                left=0.6, top=4.35, width=12.1, height=2.1,
                font_size=11, color=C_DARK)

    add_footer(slide, "Pulse  ·  Co-creatief  ·  Design-Based Research  ·  Lab Persoonlijk")


def slide_goal(prs):
    """Slide 3 – Doel & Centrale Onderzoeksvraag."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header_bar(slide, "Doel & Centrale Onderzoeksvraag", 3)

    # Goal box
    add_rect(slide, 0.4, 1.2, 12.5, 1.9, C_PURPLE_LIGHT)
    add_rect(slide, 0.4, 1.2, 0.06, 1.9, C_PURPLE)
    add_textbox(slide, "🎯  Hoofddoel",
                left=0.6, top=1.28, width=5, height=0.4,
                font_size=13, bold=True, color=C_PURPLE)
    add_textbox(slide,
                "Ontwerpen, testen en verfijnen van laagdrempelige creatieve interventies die "
                "aantoonbaar helpen om:\n"
                "① Betekenisvolle interactie te starten\n"
                "② Ervaren verbinding en ondersteuning te vergroten",
                left=0.6, top=1.7, width=12.1, height=1.3,
                font_size=11, color=C_DARK)

    # Research question box
    add_rect(slide, 0.4, 3.3, 12.5, 1.7, C_PURPLE_LIGHT)
    add_rect(slide, 0.4, 3.3, 0.06, 1.7, C_PURPLE_MID)
    add_textbox(slide, "❓  Centrale Onderzoeksvraag",
                left=0.6, top=3.38, width=7, height=0.4,
                font_size=13, bold=True, color=C_PURPLE_MID)
    add_textbox(slide,
                "Hoe kunnen laagdrempelige creatieve interventies in de dagelijkse context van "
                "ouderen, jongeren en zorgprofessionals betekenisvolle ontmoetingen en "
                "wederzijdse ondersteuning op gang brengen en versterken?",
                left=0.6, top=3.8, width=12.1, height=1.1,
                font_size=11, color=C_DARK)

    # Sub-questions pills
    add_textbox(slide, "STURENDE DEELVRAGEN:",
                left=0.4, top=5.2, width=4, height=0.4,
                font_size=11, bold=True, color=C_DARK)
    deelvragen = [
        "Drempels & Triggers", "Werkzame Elementen",
        "Inpasbaarheid", "Effect & Betekenis", "Opschalingsvoorwaarden"
    ]
    for i, dv in enumerate(deelvragen):
        add_pill(slide, dv, left=0.4 + i * 2.55, top=5.7, width=2.4, height=0.45)

    add_footer(slide, "Pulse  ·  Co-creatief  ·  Design-Based Research  ·  Lab Persoonlijk")


def slide_approach(prs):
    """Slide 4 – Aanpak: 3 testcycli."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header_bar(slide, "Aanpak – 3 Testcycli (Design-Based Research)", 4)

    # Iteratie-pijl achtergrond
    add_rect(slide, 0.4, 2.9, 12.5, 0.25, C_PURPLE_LIGHT)

    cycles = [
        ("TM0", "Start & Afbakening", "Week 1–2",
         ["Kick-off · Contextscan", "Behoeftepeiling", "Ontwerpcriteria"]),
        ("TM1", "Concept Try-out", "Week 3–4",
         ["3–5 ruwe prototypes", "Directe respons meten", "Besluit: stoppen/bijstellen?"]),
        ("TM2", "In-Context Pilot", "Week 6–8",
         ["2–3 verbeterde versies", "Realistische setting", "Inpasbaarheid & werkdruk"]),
        ("TM3", "Overdrachtstest", "Week 10–12",
         ["Near-final + handleiding", "Test door nieuwe prof.", "Overdraagbaarheid check"]),
    ]

    for i, (badge, title, week, bullets) in enumerate(cycles):
        x = 0.4 + i * 3.2
        # card
        add_rect(slide, x, 1.2, 2.9, 4.8, C_PURPLE_LIGHT)
        add_rect(slide, x, 1.2, 2.9, 0.45, C_PURPLE)
        # badge circle (oval)
        oval = slide.shapes.add_shape(9,  # OVAL
            Inches(x + 1.1), Inches(1.1), Inches(0.7), Inches(0.7))
        oval.fill.solid()
        oval.fill.fore_color.rgb = C_ACCENT
        oval.line.fill.background()
        tf = oval.text_frame
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = badge
        run.font.size = Pt(10)
        run.font.bold = True
        run.font.color.rgb = C_WHITE

        add_textbox(slide, title,
                    left=x + 0.1, top=1.2, width=2.7, height=0.45,
                    font_size=12, bold=True, color=C_WHITE)
        add_textbox(slide, week,
                    left=x + 0.1, top=1.7, width=2.7, height=0.35,
                    font_size=10, color=C_GRAY)
        for j, b in enumerate(bullets):
            add_textbox(slide, f"· {b}",
                        left=x + 0.15, top=2.1 + j * 0.5,
                        width=2.6, height=0.48,
                        font_size=10, color=C_DARK)

    add_textbox(slide,
                "↺  Ontwerpen → Testen → Leren → Aanpassen",
                left=3.0, top=6.3, width=7.0, height=0.5,
                font_size=14, bold=True, color=C_PURPLE,
                align=PP_ALIGN.CENTER)

    add_footer(slide, "Pulse  ·  Co-creatief  ·  Design-Based Research  ·  Lab Persoonlijk")


def slide_participants(prs):
    """Slide 5 – Deelnemers & Dataverzameling."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header_bar(slide, "Deelnemers & Dataverzameling", 5)

    # Pilot group
    add_textbox(slide, "👥  Pilotgroep",
                left=0.4, top=1.2, width=6, height=0.45,
                font_size=14, bold=True, color=C_PURPLE)
    groups = [
        ("👴👵  Ouderen (woonzorg)", "6–10 personen"),
        ("🧑‍🤝‍🧑  Jongeren", "6–10 personen"),
        ("👩‍⚕️  Zorgprofessionals", "4–6 personen"),
    ]
    for i, (label, count) in enumerate(groups):
        add_rect(slide, 0.4 + i * 2.2, 1.75, 2.0, 1.2, C_PURPLE_LIGHT)
        add_textbox(slide, label,
                    left=0.55 + i * 2.2, top=1.85, width=1.8, height=0.4,
                    font_size=10, bold=True, color=C_PURPLE)
        add_textbox(slide, count,
                    left=0.55 + i * 2.2, top=2.25, width=1.8, height=0.4,
                    font_size=10, color=C_DARK)

    add_textbox(slide,
                "Werving via coalitiepartners & implementatiepartners (bijv. Laurens)",
                left=0.4, top=3.1, width=6.5, height=0.4,
                font_size=10, color=C_GRAY)

    # Instruments
    add_textbox(slide, "📊  Meetinstrumenten",
                left=7.2, top=1.2, width=5.7, height=0.45,
                font_size=14, bold=True, color=C_PURPLE)

    add_bullet_box(slide, "Kwalitatief (kern)",
                   ["Observaties interactiepatronen",
                    "Korte interviews na sessie (5–10 min)",
                    "Reflectiekaarten: wat raakte je?",
                    "Deelnemersverhalen (narratief)"],
                   left=7.2, top=1.75, width=5.7, height=2.2,
                   accent_color=C_PURPLE)

    add_bullet_box(slide, "Kwantitatief – Pulse-Check (0–10)",
                   ["💜  Verbondenheid – voor en na elke sessie",
                    "⚡  Energie / Stemming – voor en na elke sessie",
                    "🧘  Spanning / Stress – voor en na elke sessie"],
                   left=7.2, top=4.1, width=5.7, height=2.0,
                   accent_color=C_PURPLE_MID)

    add_footer(slide, "Pulse  ·  Co-creatief  ·  Design-Based Research  ·  Lab Persoonlijk")


def slide_criteria(prs):
    """Slide 6 – Succescriteria & Deliverables."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header_bar(slide, "Succescriteria & Deliverables", 6)

    add_textbox(slide, "✅  Succescriteria",
                left=0.4, top=1.2, width=6, height=0.45,
                font_size=14, bold=True, color=C_PURPLE)

    criteria = [
        ("⏱️", "< 15 min uitvoerbaar", "Minstens één interventie zonder voorbereiding"),
        ("👍", "≥ 70% waardering", "Deelnemers vinden het 'prettig / waardevol'"),
        ("💬", "≥ 2 van 3 sessies", "Aantoonbaar verdiepend moment (observatie)"),
        ("🏥", "Werkdruk-neutraal", "Professionals: haalbaar binnen routine"),
    ]
    for i, (icon, label, desc) in enumerate(criteria):
        y = 1.75 + i * 0.9
        add_rect(slide, 0.4, y, 6.2, 0.78, C_PURPLE_LIGHT)
        add_textbox(slide, icon,
                    left=0.55, top=y + 0.12, width=0.6, height=0.55,
                    font_size=20, color=C_PURPLE)
        add_textbox(slide, label,
                    left=1.2, top=y + 0.05, width=2.5, height=0.38,
                    font_size=11, bold=True, color=C_PURPLE)
        add_textbox(slide, desc,
                    left=1.2, top=y + 0.42, width=5.2, height=0.35,
                    font_size=10, color=C_DARK)

    # Deliverables
    add_bullet_box(slide, "📦  Deliverables",
                   ["1–3 gevalideerde interventies",
                    "Stappenplan + varianten per context",
                    "Do's & don'ts + veiligheidsnotes",
                    "Pulse-check + reflectiekaarten",
                    "Korte evaluatierapportage",
                    "Aanbevelingen Make & Deliver"],
                   left=7.0, top=1.2, width=5.9, height=5.4,
                   accent_color=C_PURPLE)

    add_footer(slide, "Pulse  ·  Co-creatief  ·  Design-Based Research  ·  Lab Persoonlijk")


def slide_planning(prs):
    """Slide 7 – Planning."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header_bar(slide, "Planning – 12 Weken", 7)

    timeline = [
        ("Week 1–2",  "Kick-off & Contextscan",                C_GRAY),
        ("Week 3–4",  "🧪 Testmoment 1: Concept try-out",       C_PURPLE),
        ("Week 5",    "Redesign Sprint",                        C_GRAY),
        ("Week 6–8",  "🧪 Testmoment 2: In-context pilot",      C_PURPLE_MID),
        ("Week 9",    "Redesign Sprint",                        C_GRAY),
        ("Week 10–12","🧪 Testmoment 3: Overdrachtstest",       C_ACCENT),
    ]

    # Horizontal bar background
    add_rect(slide, 0.4, 3.0, 12.5, 0.2, C_PURPLE_LIGHT)
    week_positions = [0.4, 2.45, 4.5, 6.55, 8.6, 10.65, 12.7]
    labels = ["W0", "W2", "W4", "W6", "W8", "W10", "W12"]
    for pos, lbl in zip(week_positions, labels):
        add_rect(slide, pos, 2.95, 0.05, 0.3, C_PURPLE)
        add_textbox(slide, lbl,
                    left=pos - 0.15, top=3.3, width=0.5, height=0.3,
                    font_size=9, color=C_GRAY, align=PP_ALIGN.CENTER)

    for i, (week, label, color) in enumerate(timeline):
        y = 1.3 + i * 0.85
        add_rect(slide, 0.4, y, 1.7, 0.65, C_PURPLE_LIGHT)
        add_textbox(slide, week,
                    left=0.5, top=y + 0.12, width=1.5, height=0.42,
                    font_size=11, bold=True, color=C_PURPLE)
        add_rect(slide, 2.2, y + 0.18, 9.5, 0.35, color)
        add_textbox(slide, label,
                    left=2.35, top=y + 0.18, width=9.3, height=0.35,
                    font_size=11, bold=(color != C_GRAY), color=C_WHITE)

    add_footer(slide, "Pulse  ·  Co-creatief  ·  Design-Based Research  ·  Lab Persoonlijk")


def slide_conclusion(prs):
    """Slide 8 – Conclusie & Aanbevelingen."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header_bar(slide, "Conclusie & Aanbevelingen", 8)

    # Conclusion text
    add_rect(slide, 0.4, 1.2, 12.5, 2.0, C_PURPLE_LIGHT)
    add_rect(slide, 0.4, 1.2, 0.06, 2.0, C_PURPLE)
    add_textbox(slide, "Conclusie",
                left=0.6, top=1.28, width=5, height=0.4,
                font_size=13, bold=True, color=C_PURPLE)
    add_textbox(slide,
                "Lab Persoonlijk biedt een gestructureerd maar flexibel onderzoekstraject met "
                "drie iteratieve testcycli. De design-based research aanpak maakt snel leren "
                "mogelijk over wat werkt, voor wie en onder welke condities. De Theory of Change "
                "fungeert als kompas om aannames doorlopend te toetsen.",
                left=0.6, top=1.68, width=12.1, height=1.4,
                font_size=11, color=C_DARK)

    # Recommendations
    add_bullet_box(slide, "Aanbevelingen",
                   ["Plan voldoende buffers: co-creatie met kwetsbare doelgroepen vraagt meer afstemming",
                    "Betrek professionals als mede-ontwerpers van meet af aan, niet alleen als uitvoerders",
                    "Documenteer ook mislukte prototypes: negatieve lessen zijn minstens zo waardevol",
                    "Plan de overdrachtstest (TM3) vroegtijdig zodat ook 'outsiders' weten hoe het werkt",
                    "Koppel bevindingen terug aan de Pulse-brede Theory of Change"],
                   left=0.4, top=3.35, width=12.5, height=3.4,
                   accent_color=C_PURPLE)

    add_footer(slide, "Pulse  ·  Co-creatief  ·  Design-Based Research  ·  Lab Persoonlijk")


def slide_concept_idea(prs):
    """Slide 9 – Concept Idee (empty template for a new idea)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Background
    add_rect(slide, 0, 0, 13.333, 7.5, C_PURPLE_LIGHT)

    # Top banner
    add_rect(slide, 0, 0, 13.333, 1.0, C_PURPLE)
    add_pill(slide, "💡  CONCEPT IDEE", left=5.2, top=0.2, width=2.9, height=0.55,
             bg=C_ACCENT, fg=C_DARK)

    # Title placeholder
    add_rect(slide, 1.0, 1.2, 11.3, 0.9, C_WHITE)
    add_rect(slide, 1.0, 1.2, 0.06, 0.9, C_ACCENT)
    add_textbox(slide, "[Naam van jouw concept]",
                left=1.2, top=1.3, width=11.0, height=0.7,
                font_size=22, bold=True, color=C_GRAY,
                align=PP_ALIGN.CENTER)

    # Tagline
    add_rect(slide, 2.0, 2.25, 9.3, 0.55, C_WHITE)
    add_textbox(slide,
                "[ Korte pakkende tagline – één zin die je concept omschrijft ]",
                left=2.0, top=2.3, width=9.3, height=0.48,
                font_size=13, color=C_GRAY, align=PP_ALIGN.CENTER)

    # Four quadrant cards
    cards = [
        ("🔍\nHet Probleem",
         "Welk probleem lost dit concept op?\nVoor wie?",
         0.5, 3.05),
        ("✨\nDe Oplossing",
         "Wat is jouw idee in één alinea?\nWat maakt het uniek?",
         6.9, 3.05),
        ("👥\nDoelgroep",
         "Wie zijn de gebruikers?\nWat is hun context?",
         0.5, 5.1),
        ("🚀\nVolgende Stap",
         "Wat heb je nodig om te testen?\nWat is de eerste actie?",
         6.9, 5.1),
    ]
    for title, body, x, y in cards:
        add_rect(slide, x, y, 5.9, 1.85, C_WHITE)
        add_rect(slide, x, y, 0.06, 1.85, C_PURPLE_MID)
        add_textbox(slide, title,
                    left=x + 0.15, top=y + 0.1,
                    width=1.4, height=0.8,
                    font_size=11, bold=True, color=C_PURPLE)
        add_textbox(slide, body,
                    left=x + 1.6, top=y + 0.15,
                    width=4.1, height=1.55,
                    font_size=10, color=C_GRAY)

    # Footer
    add_footer(slide,
               "Pulse Lab Persoonlijk  ·  Concept Idee Slide  ·  Vul in met jouw idee")
    add_textbox(slide, "9 / 9",
                left=12.0, top=7.13, width=1.1, height=0.35,
                font_size=9, color=C_GRAY, align=PP_ALIGN.RIGHT)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def build_presentation(output_path="Pitch Presentatie Lab Persoonlijk.pptx"):
    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H

    if len(prs.slide_layouts) < 7:
        raise RuntimeError(
            f"Expected at least 7 slide layouts, found {len(prs.slide_layouts)}. "
            "Ensure you are using the default python-pptx theme."
        )

    slide_title(prs)
    slide_context(prs)
    slide_goal(prs)
    slide_approach(prs)
    slide_participants(prs)
    slide_criteria(prs)
    slide_planning(prs)
    slide_conclusion(prs)
    slide_concept_idea(prs)   # empty concept-idea slide

    prs.save(output_path)
    print(f"✅  Presentatie opgeslagen als: {output_path}")
    print(f"   Aantal slides: {len(prs.slides)}")


if __name__ == "__main__":
    build_presentation()
