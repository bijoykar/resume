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

        p_role = left_cell.add_paragraph()
        sp(p_role, before=0, after=1)
        ar(p_role, 'ELECTRICAL ENGINEER  ||  ELECTRICAL QC ENGINEER', size=12, bold=True, clr=NAVY)

        p_hl = left_cell.add_paragraph()
        sp(p_hl, before=0, after=5)
        add_bottom_border(p_hl, sz=6)
        ar(p_hl, 'Infrastructure & Building Services  |  Installation, Testing & Commissioning  |  Team Leadership  |  DEWA & Dubai Municipality Approved', size=9.5, clr=LT_NAVY)

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
        sp(p, before=0, after=1)
        ar(p, 'ELECTRICAL ENGINEER  ||  ELECTRICAL QC ENGINEER', size=13, bold=True, clr=NAVY)

        p = doc.add_paragraph()
        sp(p, before=0, after=5)
        add_bottom_border(p, sz=6)
        ar(p, 'Infrastructure & Building Services  ·  Installation, Testing & Commissioning  ·  Team Leadership  ·  DEWA & Dubai Municipality Approved', size=10.5, clr=LT_NAVY)

        p = doc.add_paragraph()
        sp(p, before=4, after=8)
        ar(p, 'Dubai, UAE  ·  +971-504203242  ·  joykar2010@gmail.com  ·  linkedin.com/in/joy-kar', size=9, clr=MID_GREY)

    # ── PROFESSIONAL SUMMARY ────────────────────────────────
    heading(doc, 'Professional Summary', before=0)
    p = doc.add_paragraph()
    sp(p, before=4, after=0)
    ar(p, 'Results-driven Electrical Engineer with 11+ years of combined site engineering and quality control experience across large-scale infrastructure and building services projects (Dubai Metro Route 2020, Expo Link, Dubai Municipality sewerage & stormwater utilities, sports facilities). Hands-on across the full delivery cycle — drawing and specification review, LV/MV installation supervision, cable routing and termination, earthing systems, testing, commissioning and handover — combined with formal QA/QC ownership of inspections, ITPs, NCRs and test documentation. DEWA and Dubai Municipality Approved Electrical Engineer authorised to sign off on critical installations. Proven site leader managing teams of 10–30+ engineers, technicians and subcontractors to IEC/BS standards. UAE-based, immediately available.', size=10, clr=BODY)

    # ── CORE COMPETENCIES ────────────────────────────────────
    heading(doc, 'Core Competencies & Technical Skills')
    for label, content in [
        ('Electrical Engineering & Execution:', 'LV/MV Systems  ·  Distribution Panels & Switchgear  ·  Cable Sizing, Routing & Termination  ·  Earthing & Lightning Protection  ·  Lighting & Small Power  ·  Containment (Trays/Trunking/Conduits)  ·  SCADA & Control Systems  ·  Installation Supervision  ·  Testing & Commissioning  ·  Shop Drawing & Specification Review  ·  As-Built Documentation'),
        ('Quality Control & Assurance:', 'Quality Control & Inspection  ·  ITPs & Inspection Requests  ·  NCR Management & CAPA  ·  Hi-Pot, Insulation Resistance & Continuity Testing  ·  Material Submittal Approvals  ·  Third-Party & Authority Inspections  ·  DEWA / Dubai Municipality Compliance'),
        ('Standards & Regulatory:', 'IEC & BS Standards  ·  DEWA Regulations  ·  Dubai Municipality Requirements  ·  HSE & Permit-to-Work  ·  Project QMS  ·  Document Control (Aconex / Procore)'),
        ('Tools & Systems:', 'AutoCAD  ·  Primavera P6  ·  Aconex / Procore  ·  Megger & Hi-Pot Test Equipment  ·  MS Office  ·  Generative AI (Defect Detection)'),
        ('Leadership & Coordination:', 'Site Team Management (10–30+ professionals)  ·  Subcontractor & Vendor Coordination  ·  Consultant / Client Interface  ·  Mentoring & Training  ·  Progress, Budget & Schedule Management'),
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
    ar(p, 'Electrical Engineer  |  Electrical Quality Control Engineer', size=10, bold=True, clr=LT_NAVY)
    for b in [
        'Own electrical delivery and quality control for pumping station and utility works: LV panels, cable trays, conduits, power and control cabling, lighting, and earthing systems',
        'Review shop drawings, layouts, cable schedules and material submittals against specifications and DEWA / Dubai Municipality requirements prior to installation',
        'Supervise site installation and termination works, resolving field routing and coordination clashes with civil, mechanical and instrumentation disciplines',
        'Execute and witness testing — insulation resistance, continuity, earth resistance and functional tests — and lead pre-commissioning through to handover',
        'Raise and close out NCRs, drive corrective actions, and maintain ITPs, inspection records and test certificates under the project QMS for zero-defect handover',
        'Leverage DEWA / Dubai Municipality approval authority to sign off installations and accelerate regulatory clearance',
    ]:
        blt(doc, b)

    company(doc, 'BK Gulf LLC  |  JA Sports Center & Shooting Club, Dubai', 'May 2025 – September 2025', before=8)
    p = doc.add_paragraph()
    sp(p, before=2, after=2)
    ar(p, 'Electrical Engineer  |  QC & Site Leadership', size=10, bold=True, clr=LT_NAVY)
    for b in [
        'Led a 30+ member site team delivering the complete electrical installation for a high-profile sports facility on an accelerated programme',
        'Planned and sequenced daily execution — manpower, materials and containment/cabling activities — against the construction schedule and inspection milestones',
        'Inspected and signed off panels, cabling, earthing and lighting systems for compliance with drawings, specifications and DEWA requirements',
        'Coordinated NCR resolution and corrective actions, keeping quality standards intact under schedule pressure',
        'Resolved technical conflicts between consultants, client and subcontractors; trained the team on installation and quality procedures',
        'Enforced site safety protocols and progress tracking, delivering on schedule with no quality compromises',
    ]:
        blt(doc, b)

    company(doc, 'Concord Corporation Pte Ltd. (Hollysys Group)  |  Dubai Metro Route 2020', 'January 2018 – April 2025', before=8)
    p = doc.add_paragraph()
    sp(p, before=2, after=2)
    ar(p, 'Electrical Engineer  |  Electrical Quality Control Engineer', size=10, bold=True, clr=LT_NAVY)
    for b in [
        'Delivered electrical works and quality control across the Dubai Metro expansion serving 500,000+ daily commuters — Expo Link Tie-In, Rashidiya Depot Automation and Signaling Extension packages',
        'Executed and supervised installation of LV/MV systems, cable trays, panels, lighting and earthing; managed third rail terminations and 1C×400 sqmm XLPE power cable installation',
        'Engineered and commissioned SCADA and traction control interfaces including ERTU cubicles and HSCB installations, coordinating with signalling and systems teams',
        'Performed Hi-Pot, insulation resistance and continuity testing to IEC/BS standards; witnessed and cleared inspections with consultants, client and third-party inspectors',
        'Raised IRs and NCRs and maintained quality records via Aconex / Procore; produced as-built drawings and test certificate dossiers for handover',
        'Achieved zero safety incidents across seven years through proactive compliance monitoring and method statement discipline',
    ]:
        blt(doc, b)

    company(doc, 'Websol Energy System Ltd.  |  India', 'February 2012 – October 2017', before=8)
    p = doc.add_paragraph()
    sp(p, before=2, after=2)
    ar(p, 'Electrical / Production Engineer  |  Quality Control Engineer', size=10, bold=True, clr=LT_NAVY)
    for b in [
        'Implemented QC programmes that cut product rejection rates by 30%, delivering significant cost and efficiency gains',
        'Supervised solar cell and photovoltaic module manufacturing operations, monitoring electrical parameters and production output',
        'Inspected and tested photovoltaic components; conducted supplier verification and drove process quality enhancements',
        'Carried out field inspections and process improvements; mentored production and QC teams on electrical and quality standards',
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
build(BASE + 'Joy_Kar_Electrical_Engineer_QC_DUAL_ROLE_ATS.docx', with_photo=False)
build(BASE + 'Joy_Kar_Electrical_Engineer_QC_DUAL_ROLE_With_Photo.docx', with_photo=True)
print('\nDone: Dual-role (Electrical Engineer || Electrical QC Engineer) resumes created.')
