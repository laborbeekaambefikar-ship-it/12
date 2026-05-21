from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def apply_line_spacing(para, line_val='360'):
    """Apply 1.5 line spacing to a paragraph via XML."""
    pPr = para._p.get_or_add_pPr()
    existing = pPr.find(qn('w:spacing'))
    if existing is not None:
        pPr.remove(existing)
    lSpacing = OxmlElement('w:spacing')
    lSpacing.set(qn('w:line'), line_val)
    lSpacing.set(qn('w:lineRule'), 'auto')
    pPr.append(lSpacing)


def set_run_font(run, size_pt=12, bold=False):
    """Set Times New Roman font on a run."""
    run.font.name = 'Times New Roman'
    run.font.size = Pt(size_pt)
    run.bold = bold
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:ascii'), 'Times New Roman')
    rFonts.set(qn('w:hAnsi'), 'Times New Roman')
    rFonts.set(qn('w:cs'),    'Times New Roman')


def add_mixed_paragraph(doc, segments, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
                         space_after=12, size_pt=12):
    """
    Add a paragraph with mixed bold/normal runs.
    segments: list of (text, bold) tuples
    """
    para = doc.add_paragraph()
    para.alignment = align
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after  = Pt(space_after)
    apply_line_spacing(para)

    for text, bold in segments:
        run = para.add_run(text)
        set_run_font(run, size_pt=size_pt, bold=bold)

    return para


def create_acknowledgement_doc(output_path):
    doc = Document()

    # ── Page setup: A4, 1-inch (2.54 cm) margins ──────────────────────────────
    section = doc.sections[0]
    section.page_width  = Cm(21.0)
    section.page_height = Cm(29.7)
    margin = Cm(2.54)
    section.top_margin    = margin
    section.bottom_margin = margin
    section.left_margin   = margin
    section.right_margin  = margin

    # ── Default Normal style ───────────────────────────────────────────────────
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:ascii'), 'Times New Roman')
    rFonts.set(qn('w:hAnsi'), 'Times New Roman')
    rFonts.set(qn('w:cs'),    'Times New Roman')

    # ── Heading: ACKNOWLEDGEMENT ───────────────────────────────────────────────
    heading_para = doc.add_paragraph()
    heading_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    heading_para.paragraph_format.space_before = Pt(0)
    heading_para.paragraph_format.space_after  = Pt(12)
    apply_line_spacing(heading_para)

    run = heading_para.add_run('ACKNOWLEDGEMENT')
    set_run_font(run, size_pt=14, bold=True)

    # ── Body paragraphs ────────────────────────────────────────────────────────
    # Each paragraph is a list of (text, is_bold) tuples.
    paragraphs = [
        # Para 1 – opening line (bold)
        [("In the name of Allah, the Most Gracious, the Most Merciful.", True)],

        # Para 2
        [
            ("First and foremost, the author expresses sincere gratitude to ", False),
            ("Allah (SWT)", True),
            (" for His infinite mercy, countless blessings, guidance, and strength "
             "throughout this academic journey. By His grace, the author was able to "
             "overcome challenges and successfully complete this project.", False),
        ],

        # Para 3
        [
            ("The author extends heartfelt appreciation to ", False),
            ("Mr. Rashid Akber", True),
            (" and ", False),
            ("Mrs. Kaniz Sakeena Rashid", True),
            (" for their unwavering love, sacrifices, prayers, encouragement, and "
             "constant support. Their guidance, patience, and belief have been a source "
             "of inspiration and strength throughout the completion of this work.", False),
        ],

        # Para 4
        [
            ("Special thanks are extended to ", False),
            ("Ms. Rida Fatima", True),
            (" for her patience, understanding, unwavering moral support, and "
             "encouragement. Her positive influence and constant motivation were "
             "invaluable during the course of this project.", False),
        ],

        # Para 5
        [
            ("The author is deeply grateful to ", False),
            ("Prof. Mohammad Muzammil", True),
            (" for his invaluable guidance, insightful suggestions, constructive "
             "feedback, and continuous encouragement. His mentorship and expertise "
             "played a crucial role in shaping and successfully completing this project.", False),
        ],

        # Para 6
        [
            ("Sincere appreciation is also extended to ", False),
            ("Mr. Sheikh Noorul", True),
            (", project partner, for his cooperation, dedication, and collaborative "
             "efforts throughout the development of this project. His support, teamwork, "
             "and valuable contributions greatly facilitated the successful accomplishment "
             "of the project objectives.", False),
        ],

        # Para 7
        [
            ("The author would also like to thank all ", False),
            ("faculty members of the Department of Mechanical Engineering", True),
            (" for their guidance, encouragement, and support throughout this academic "
             "endeavor.", False),
        ],

        # Para 8
        [
            ("Finally, heartfelt thanks are extended to family members, friends, and "
             "well-wishers for their encouragement, understanding, and support, which "
             "provided constant motivation throughout this journey.", False),
        ],
    ]

    for segments in paragraphs:
        add_mixed_paragraph(doc, segments)

    doc.save(output_path)
    print(f"Document saved to: {output_path}")


if __name__ == '__main__':
    create_acknowledgement_doc('/projects/sandbox/12/Acknowledgement.docx')
