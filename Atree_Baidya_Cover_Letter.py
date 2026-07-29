from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ── Colors ────────────────────────────────────────────────────
NAVY      = RGBColor(0x1B, 0x2A, 0x4A)
LT_NAVY   = RGBColor(0x34, 0x4E, 0x8A)
MID_GREY  = RGBColor(0x6B, 0x6B, 0x6B)
BODY      = RGBColor(0x22, 0x22, 0x22)

def ar(para, text, size=11, bold=False, italic=False, clr=None):
    if clr is None:
        clr = BODY
    run = para.add_run(text)
    run.font.name      = 'Calibri'
    run.font.size      = Pt(size)
    run.font.bold      = bold
    run.font.italic    = italic
    run.font.color.rgb = clr
    return run

def sp(para, before=0, after=0):
    pf = para.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after  = Pt(after)

def build_cover_letter(output_path):
    doc = Document()
    doc.styles['Normal'].font.name = 'Calibri'
    doc.styles['Normal'].font.size = Pt(11)
    doc.styles['Normal'].paragraph_format.space_before = Pt(0)
    doc.styles['Normal'].paragraph_format.space_after  = Pt(6)

    sec = doc.sections[0]
    sec.page_width     = Cm(21)
    sec.page_height    = Cm(29.7)
    sec.top_margin     = Cm(1.8)
    sec.bottom_margin  = Cm(1.8)
    sec.left_margin    = Cm(2.2)
    sec.right_margin   = Cm(2.2)

    # ── SENDER'S ADDRESS ──────────────────────────────────────
    p = doc.add_paragraph()
    sp(p, before=0, after=0)
    ar(p, 'Atree Baidya', size=11, bold=True, clr=NAVY)
    
    p = doc.add_paragraph()
    sp(p, before=0, after=0)
    ar(p, 'Dubai, UAE', size=11, clr=BODY)
    
    p = doc.add_paragraph()
    sp(p, before=0, after=0)
    ar(p, '+971-554463371', size=11, clr=BODY)
    
    p = doc.add_paragraph()
    sp(p, before=0, after=6)
    ar(p, 'atreebaidya@gmail.com', size=11, clr=BODY)

    # ── DATE ──────────────────────────────────────────────────
    p = doc.add_paragraph()
    sp(p, before=0, after=12)
    ar(p, '28 July 2026', size=11, clr=BODY)

    # ── RECIPIENT'S ADDRESS ───────────────────────────────────
    p = doc.add_paragraph()
    sp(p, before=0, after=0)
    ar(p, 'Hiring Manager', size=11, clr=BODY)
    
    p = doc.add_paragraph()
    sp(p, before=0, after=0)
    ar(p, '[Company Name]', size=11, clr=BODY)
    
    p = doc.add_paragraph()
    sp(p, before=0, after=0)
    ar(p, '[Company Address]', size=11, clr=BODY)
    
    p = doc.add_paragraph()
    sp(p, before=0, after=0)
    ar(p, 'Dubai, UAE', size=11, clr=BODY)
    
    p = doc.add_paragraph()
    sp(p, before=0, after=12)
    ar(p, '[Postal Code]', size=11, clr=BODY)

    # ── SALUTATION ────────────────────────────────────────────
    p = doc.add_paragraph()
    sp(p, before=0, after=12)
    ar(p, 'Dear Hiring Manager,', size=11, clr=BODY)

    # ── OPENING PARAGRAPH ─────────────────────────────────────
    p = doc.add_paragraph()
    sp(p, before=0, after=12)
    ar(p,
       'I am writing to express my strong interest in the HR Assistant position at [Company Name]. '
       'With my MBA in Human Resource Management, combined with hands-on experience in recruitment, '
       'employee onboarding, HR operations, and compliance management, I am confident in my ability to '
       'contribute meaningfully to your HR department and support your organizational objectives. '
       'I am particularly attracted to this opportunity because of your company\'s commitment to excellence '
       'and people-centric culture, which aligns with my professional values and career aspirations.',
       size=11, clr=BODY)

    # ── BODY PARAGRAPH 1: RELEVANT EXPERIENCE ─────────────────
    p = doc.add_paragraph()
    sp(p, before=0, after=12)
    ar(p,
       'In my current role as HR Assistant at Sensia Ventures Group since October 2025, I have successfully '
       'managed the full recruitment lifecycle—from job posting and CV screening to interview coordination and '
       'candidate communication. I have consistently prepared employment documentation including offer letters, '
       'contracts, NDAs, and onboarding agreements with meticulous attention to accuracy and compliance with UAE '
       'labor regulations. My responsibilities extend to maintaining employee records, HR databases, and confidential '
       'documentation while ensuring data security and organizational compliance. Additionally, I coordinate employee '
       'onboarding and offboarding processes, manage attendance and leave tracking, and prepare HR reports using advanced '
       'Excel functions—all critical aspects of the HR Assistant role you are seeking to fill.',
       size=11, clr=BODY)

    # ── BODY PARAGRAPH 2: SKILLS & COMPETENCIES ───────────────
    p = doc.add_paragraph()
    sp(p, before=0, after=12)
    ar(p,
       'Beyond my technical competencies in HR systems, MS Excel (Pivot Tables, VLOOKUP), and SQL database management, '
       'I bring a strong commitment to confidentiality, data protection, and ethical HR practices. My experience across '
       'multiple organizations—including Websol Energy Systems Limited, National HRD Network, and Chevron—has equipped me '
       'with a deep understanding of diverse HR operations, employee engagement initiatives, and cross-departmental coordination. '
       'I am particularly skilled at handling sensitive employee information with discretion, supporting HR policy implementation, '
       'and assisting in organizational projects while maintaining the highest standards of professionalism and compliance.',
       size=11, clr=BODY)

    # ── BODY PARAGRAPH 3: PERSONAL QUALITIES ──────────────────
    p = doc.add_paragraph()
    sp(p, before=0, after=12)
    ar(p,
       'What distinguishes me as a candidate is my combination of formal HR education (MBA in Human Resources), practical '
       'experience in fast-paced multinational environments, and a genuine passion for supporting employee experience and '
       'organizational growth. I am detail-oriented, proactive, and committed to continuous learning—qualities that enable me '
       'to anticipate HR needs, solve problems efficiently, and support the department\'s strategic initiatives. My ability to '
       'manage multiple priorities, communicate effectively with employees and internal stakeholders, and maintain confidentiality '
       'positions me well to deliver immediate value to your organization.',
       size=11, clr=BODY)

    # ── CLOSING PARAGRAPH ──────────────────────────────────────
    p = doc.add_paragraph()
    sp(p, before=0, after=12)
    ar(p,
       'I am excited about the opportunity to bring my expertise, dedication, and enthusiasm to [Company Name]. '
       'I would welcome the chance to discuss how my skills in HR operations, recruitment, compliance, and employee '
       'management can contribute to your team\'s success. Thank you for considering my application. I am available at '
       'your earliest convenience for an interview and look forward to the possibility of joining your organization.',
       size=11, clr=BODY)

    # ── SIGN-OFF ───────────────────────────────────────────────
    p = doc.add_paragraph()
    sp(p, before=0, after=2)
    ar(p, 'Sincerely,', size=11, clr=BODY)

    # ── SIGNATURE SPACE ────────────────────────────────────────
    doc.add_paragraph()
    doc.add_paragraph()

    # ── TYPED NAME ─────────────────────────────────────────────
    p = doc.add_paragraph()
    sp(p, before=0, after=0)
    ar(p, 'Atree Baidya', size=11, bold=True, clr=NAVY)

    # ── CONTACT INFO ──────────────────────────────────────────
    p = doc.add_paragraph()
    sp(p, before=0, after=0)
    ar(p, '+971-554463371  |  atreebaidya@gmail.com  |  linkedin.com/in/atreebaidya', size=9, clr=MID_GREY)

    doc.save(output_path)
    print(f'Saved: {output_path}')

# ── Run ─────────────────────────────────────────────────────
BASE = 'C:/xampp/htdocs/resume/'
build_cover_letter(BASE + 'Atree_Baidya_Cover_Letter.docx')
print('\n✓ Professional cover letter created successfully!')
