from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ── Colors ────────────────────────────────────────────────────
NAVY      = RGBColor(0x1B, 0x2A, 0x4A)
LT_NAVY   = RGBColor(0x34, 0x4E, 0x8A)
MID_GREY  = RGBColor(0x6B, 0x6B, 0x6B)
DK_GREY   = RGBColor(0x40, 0x40, 0x40)
BODY      = RGBColor(0x22, 0x22, 0x22)

# ── Helpers ───────────────────────────────────────────────────

def sp(para, before=0, after=0):
    pf = para.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after  = Pt(after)

def add_bottom_border(para, color='1B2A4A', sz=8):
    p   = para._p
    pPr = p.get_or_add_pPr()
    for old in pPr.findall(qn('w:pBdr')):
        pPr.remove(old)
    pBdr = OxmlElement('w:pBdr')
    bot  = OxmlElement('w:bottom')
    bot.set(qn('w:val'),   'single')
    bot.set(qn('w:sz'),    str(sz))
    bot.set(qn('w:space'), '1')
    bot.set(qn('w:color'), color)
    pBdr.append(bot)
    pPr.append(pBdr)

def set_right_tab(para, pos_in=6.3):
    pPr  = para._p.get_or_add_pPr()
    tabs = OxmlElement('w:tabs')
    tab  = OxmlElement('w:tab')
    tab.set(qn('w:val'), 'right')
    tab.set(qn('w:pos'), str(int(pos_in * 1440)))
    tabs.append(tab)
    pPr.append(tabs)

def ar(para, text, size=10, bold=False, italic=False, clr=None):
    if clr is None:
        clr = BODY
    run = para.add_run(text)
    run.font.name      = 'Calibri'
    run.font.size      = Pt(size)
    run.font.bold      = bold
    run.font.italic    = italic
    run.font.color.rgb = clr
    return run

def heading(doc, text, before=10, after=4):
    p = doc.add_paragraph()
    sp(p, before=before, after=after)
    ar(p, text.upper(), size=10, bold=True, clr=NAVY)
    add_bottom_border(p)
    return p

def blt(doc, text, before=1, after=1):
    p = doc.add_paragraph()
    sp(p, before=before, after=after)
    p.paragraph_format.left_indent       = Inches(0.22)
    p.paragraph_format.first_line_indent = Inches(-0.17)
    ar(p, u'•  ', size=10, bold=True, clr=LT_NAVY)
    ar(p, text, size=10, clr=BODY)

def company(doc, name, dates, before=8):
    p = doc.add_paragraph()
    sp(p, before=before, after=0)
    set_right_tab(p, 6.3)
    ar(p, name, size=11, bold=True, clr=NAVY)
    ar(p, '\t', size=10, clr=BODY)
    ar(p, dates, size=9.5, clr=MID_GREY)

def build(output_path, with_photo=False):
    doc = Document()
    doc.styles['Normal'].font.name = 'Calibri'
    doc.styles['Normal'].font.size = Pt(10)
    doc.styles['Normal'].paragraph_format.space_before = Pt(0)
    doc.styles['Normal'].paragraph_format.space_after  = Pt(0)

    sec = doc.sections[0]
    sec.page_width     = Cm(21)
    sec.page_height    = Cm(29.7)
    sec.top_margin     = Cm(1.8)
    sec.bottom_margin  = Cm(1.8)
    sec.left_margin    = Cm(2.2)
    sec.right_margin   = Cm(2.2)

    # ── HEADER ──────────────────────────────────────────────
    if with_photo:
        tbl = doc.add_table(rows=1, cols=2)
        tbl.allow_autofit = False
        tbl.columns[0].width = Cm(13.2)
        tbl.columns[1].width = Cm(3.4)

        tblPr = tbl._tbl.tblPr
        tblBorders = OxmlElement('w:tblBorders')
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            el = OxmlElement('w:' + side)
            el.set(qn('w:val'), 'none')
            tblBorders.append(el)
        tblPr.append(tblBorders)

        left_cell  = tbl.rows[0].cells[0]
        right_cell = tbl.rows[0].cells[1]
        right_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        p_name = left_cell.paragraphs[0]
        sp(p_name, before=0, after=2)
        ar(p_name, 'JOY KAR', size=26, bold=True, clr=NAVY)

        p_hl = left_cell.add_paragraph()
        sp(p_hl, before=0, after=5)
        add_bottom_border(p_hl, sz=6)
        ar(p_hl, 'Electrical QC Engineer  |  Infrastructure  |  Team Leadership  |  DEWA & Dubai Municipality Approved', size=10.5, clr=LT_NAVY)

        p_ct = left_cell.add_paragraph()
        sp(p_ct, before=4, after=0)
        ar(p_ct, 'Dubai, UAE  |  +971-504203242  |  joykar2010@gmail.com', size=9, clr=MID_GREY)
        p_ct2 = left_cell.add_paragraph()
        sp(p_ct2, before=1, after=0)
        ar(p_ct2, 'linkedin.com/in/joy-kar', size=9, clr=MID_GREY)

        p_photo = right_cell.paragraphs[0]
        p_photo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_img = p_photo.add_run()
        run_img.add_picture('C:/xampp/htdocs/resume/JOY_KAR.png', width=Cm(3.0))

        doc.add_paragraph()

    else:
        p = doc.add_paragraph()
        sp(p, before=0, after=2)
        ar(p, 'JOY KAR', size=28, bold=True, clr=NAVY)

        p = doc.add_paragraph()
        sp(p, before=0, after=5)
        add_bottom_border(p, sz=6)
        ar(p, 'Electrical Quality Control Engineer  ·  Infrastructure Projects  ·  Team Leadership  ·  DEWA & Dubai Municipality Approved', size=11, clr=LT_NAVY)

        p = doc.add_paragraph()
        sp(p, before=4, after=8)
        ar(p, 'Dubai, UAE  ·  +971-504203242  ·  joykar2010@gmail.com  ·  linkedin.com/in/joy-kar', size=9, clr=MID_GREY)

    # ── PROFESSIONAL SUMMARY ────────────────────────────────
    heading(doc, 'Professional Summary', before=0)
    p = doc.add_paragraph()
    sp(p, before=4, after=0)
    ar(p, 'Results-driven Electrical QC Engineer with 11+ years expertise in quality assurance, installation, and commissioning on large-scale infrastructure projects (Dubai Metro 2020, Expo Link, municipal utilities). DEWA and Dubai Municipality Certified Practitioner with exclusive authority to sign off on critical installations. Proven leader managing 10-30+ professionals. Specialized in LV/MV systems, cable routing, testing protocols, and regulatory compliance. UAE-based.', size=10, clr=BODY)

    # ── CORE COMPETENCIES ────────────────────────────────────
    heading(doc, 'Core Competencies & Technical Skills')
    for label, content in [
        ('Electrical QC & Testing:', 'Quality Control & Inspection  ·  LV/MV Systems  ·  Cable Routing  ·  Grounding Systems  ·  Hi-Pot & Insulation Testing  ·  Commissioning  ·  DEWA/Dubai Municipality Compliance'),
        ('Standards & Quality Management:', 'IEC & BS Standards  ·  NCR/IR Management  ·  ITPs  ·  CAPA  ·  Document Control (Aconex/Procore)  ·  Third-Party Coordination'),
        ('Tools & Systems:', 'AutoCAD  ·  Primavera P6  ·  MS Office  ·  Generative AI (Defect Detection)  ·  Megger Testing Equipment'),
        ('Leadership:', 'Team Management (10-30+ professionals)  ·  Mentoring  ·  Subcontractor Coordination  ·  Budget & Schedule Management'),
    ]:
        p = doc.add_paragraph()
        sp(p, before=2, after=2)
        ar(p, label + '  ', size=10, bold=True, clr=NAVY)
        ar(p, content, size=10, clr=BODY)

    # ── CERTIFICATIONS ──────────────────────────────────────
    heading(doc, 'Certifications & Affiliations')
    for c in [
        'DEWA Approved Electrical Engineer (Reg. 114376776)',
        'Dubai Municipality Approved Electrical Engineer (Reg. 114376776)',
        'Member, Society of Engineers, UAE (2011/64590/2022)',
        'Introduction to Generative AI – Ali & Sons Holding LLC (2026)',
    ]:
        blt(doc, c)

    # ── PROFESSIONAL EXPERIENCE ──────────────────────────────
    heading(doc, 'Professional Experience')

    company(doc, 'Emarat Aloula Contracting LLC (EAC)  |  Dubai Municipality Nazwa Sewerage & Stormwater Project', 'October 2025 – Present', before=6)
    p = doc.add_paragraph()
    sp(p, before=2, after=2)
    ar(p, 'Electrical Quality Control Engineer', size=10, bold=True, clr=LT_NAVY)
    for b in [
        'Lead comprehensive QC of electrical installations: LV panels, cable trays, conduits, power/control cabling, lighting, and grounding systems',
        'Perform material inspections, coordinate approvals, conduct testing (insulation resistance, continuity, grounding verification, functional tests)',
        'Identify non-conformities, prepare NCRs, monitor corrective actions ensuring zero-defect handover',
        'Maintain quality documentation, inspection records, test certificates per project QMS',
        'Coordinate with construction, QA/QC, and project teams to resolve technical issues and maintain schedule',
        'Apply DEWA/Dubai Municipality approvals to enable accelerated regulatory approvals',
    ]:
        blt(doc, b)

    company(doc, 'BK Gulf LLC  |  JA Sports Center & Shooting Club, Dubai', 'May 2025 – September 2025', before=8)
    p = doc.add_paragraph()
    sp(p, before=2, after=2)
    ar(p, 'Electrical Quality Control Engineer  |  Site Leadership', size=10, bold=True, clr=LT_NAVY)
    for b in [
        'Led 30+ member site execution team delivering electrical installations for high-profile sports facility on accelerated schedule',
        'Conducted quality inspections for panels, cabling, grounding, lighting systems; ensured compliance with drawings, specifications, DEWA requirements',
        'Coordinated NCR resolution and implemented corrective actions to maintain quality standards',
        'Managed site execution, safety protocols, progress tracking to deliver on schedule without quality compromises',
        'Resolved technical conflicts between consultants, clients, subcontractors; supervised team training on quality procedures',
    ]:
        blt(doc, b)

    company(doc, 'Concord Corporation Pte Ltd. (Hollysys Group)  |  Dubai Metro Route 2020', 'January 2018 – April 2025', before=8)
    p = doc.add_paragraph()
    sp(p, before=2, after=2)
    ar(p, 'Electrical Quality Control Engineer', size=10, bold=True, clr=LT_NAVY)
    for b in [
        'Led QC on Dubai Metro expansion managing electrical systems for 500,000+ daily commuters (Expo Link Tie-In, Rashidiya Depot Automation, Signaling Extension)',
        'Conducted inspections for LV/MV systems, cable trays, panels, lighting, earthing; performed Hi-Pot testing, insulation resistance, continuity verification per IEC/BS standards',
        'Managed third rail terminations, power cable installation (1CX400Sqmm XLPE systems), SCADA systems (ERTU cubicles, HSCB installations)',
        'Prepared IRs, NCRs, quality records via Aconex/Procore; coordinated with consultants, clients, third-party inspectors for approvals',
        'Generated as-built drawings, test certificates; achieved zero safety incidents through proactive compliance monitoring',
    ]:
        blt(doc, b)

    company(doc, 'Websol Energy System Ltd.  |  India', 'February 2012 – October 2017', before=8)
    p = doc.add_paragraph()
    sp(p, before=2, after=2)
    ar(p, 'Quality Control Engineer & Production Engineer', size=10, bold=True, clr=LT_NAVY)
    for b in [
        'Implemented QC programs achieving 30% reduction in product rejection rates—significant cost savings and efficiency improvement',
        'Performed inspection and testing of photovoltaic components; conducted supplier verification and quality enhancements',
        'Supervised solar cell manufacturing operations monitoring electrical parameters and production quality',
        'Conducted field inspections and process improvements; mentored production and QC teams on quality standards',
    ]:
        blt(doc, b)

    # ── EDUCATION ────────────────────────────────────────────
    heading(doc, 'Education')
    p = doc.add_paragraph()
    sp(p, before=4, after=0)
    ar(p, 'Bachelor of Technology (B.Tech) in Electrical Engineering', size=10, bold=True, clr=BODY)
    ar(p, '  |  ', size=10, clr=BODY)
    ar(p, 'Hooghly Engineering and Technology College (WBUT/MAKAUT), India', size=9, italic=True, clr=MID_GREY)
    p_gpa = doc.add_paragraph()
    sp(p_gpa, before=0, after=4)
    ar(p_gpa, 'DGPA: 7.38  |  Graduated: 2011', size=9, clr=MID_GREY)

    # ── TECHNICAL PROFICIENCY ───────────────────────────────
    heading(doc, 'Technical Tools & Proficiency')
    p = doc.add_paragraph()
    sp(p, before=4, after=0)
    ar(p, 'AutoCAD  ·  Primavera P6  ·  Aconex/Procore  ·  Generative AI  ·  MS Office Suite  ·  C Programming  ·  Windows', size=10, clr=BODY)

    # ── ADDITIONAL INFORMATION ──────────────────────────────
    heading(doc, 'Additional Information')
    p = doc.add_paragraph()
    sp(p, before=4, after=0)
    ar(p, 'Nationality:  ', size=10, bold=True, clr=NAVY)
    ar(p, 'Indian  |  Current Location: Dubai, UAE  |  Work Authorization: Valid', size=10, clr=BODY)

    p = doc.add_paragraph()
    sp(p, before=2, after=0)
    ar(p, 'Languages:  ', size=10, bold=True, clr=NAVY)
    ar(p, 'English (Professional)  ·  Hindi (Fluent)  ·  Bengali (Native)', size=10, clr=BODY)

    p = doc.add_paragraph()
    sp(p, before=2, after=0)
    ar(p, 'Licenses & Availability:  ', size=10, bold=True, clr=NAVY)
    ar(p, 'UAE & India Driving License  |  Available for UAE-based roles', size=10, clr=BODY)

    doc.save(output_path)
    print(f'Saved: {output_path}')

BASE = 'C:/xampp/htdocs/resume/'
build(BASE + 'Joy_Kar_Electrical_QC_Engineer_FINAL_ATS.docx', with_photo=False)
build(BASE + 'Joy_Kar_Electrical_QC_Engineer_FINAL_With_Photo.docx', with_photo=True)
print('\n✓ Condensed professional ATS resumes (under 3 pages) created successfully!')
