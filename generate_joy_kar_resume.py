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
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = clr
    return run


def heading(doc, text):
    p = doc.add_paragraph()
    sp(p, before=10, after=4)
    ar(p, text, size=10, bold=True, clr=NAVY)
    add_bottom_border(p)


def blt(doc, text):
    p = doc.add_paragraph()
    sp(p, before=0, after=2)
    pf = p.paragraph_format
    pf.left_indent = Cm(0.22)
    pf.first_line_indent = Cm(-0.17)
    
    ar(p, '• ' + text, size=10, clr=BODY)


def company(doc, company_info, dates, before=6):
    p = doc.add_paragraph()
    sp(p, before=before, after=2)
    set_right_tab(p)
    
    name_part = company_info
    ar(p, name_part, size=10, bold=True, clr=NAVY)
    p.add_run('\t')
    ar(p, dates, size=10, clr=MID_GREY)


def build(output_path, with_photo=False):
    doc = Document()
    
    # ── PAGE SETUP ────────────────────────────────────────────
    sections = doc.sections
    for section in sections:
        section.top_margin = Cm(1.8)
        section.bottom_margin = Cm(1.8)
        section.left_margin = Cm(2.2)
        section.right_margin = Cm(2.2)
    
    # ── HEADER WITH PHOTO ────────────────────────────────────
    if with_photo:
        table = doc.add_table(rows=1, cols=2)
        table.autofit = False
        table.allow_autofit = False
        
        # Left cell - content
        left_cell = table.rows[0].cells[0]
        left_cell.width = Cm(13.2)
        
        # Right cell - photo
        right_cell = table.rows[0].cells[1]
        right_cell.width = Cm(3.4)
        
        # Remove borders
        for cell in table.rows[0].cells:
            tcPr = cell._element.get_or_add_tcPr()
            tcBorders = OxmlElement('w:tcBorders')
            for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
                border = OxmlElement(f'w:{border_name}')
                border.set(qn('w:val'), 'none')
                tcBorders.append(border)
            tcPr.append(tcBorders)
        
        # LEFT CELL CONTENT
        left_para = left_cell.paragraphs[0]
        sp(left_para, before=0, after=0)
        ar(left_para, 'JOY KAR', size=28, bold=True, clr=NAVY)
        
        p2 = left_cell.add_paragraph()
        sp(p2, before=2, after=1)
        ar(p2, 'ELECTRICAL QC ENGINEER | INFRASTRUCTURE PROJECTS | DUBAI METRO', size=11, bold=True, clr=LT_NAVY)
        add_bottom_border(p2)
        
        p3 = left_cell.add_paragraph()
        sp(p3, before=2, after=0)
        ar(p3, '+971 50 123 4567  |  joy.kar@email.com  |  Dubai, UAE', size=9, clr=MID_GREY)
        
        # RIGHT CELL - PHOTO
        right_para = right_cell.paragraphs[0]
        right_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        try:
            run = right_para.add_run()
            run.add_picture('JOY_KAR.png', width=Cm(3))
        except:
            pass
        
        # Set vertical alignment
        tcPr = right_cell._element.get_or_add_tcPr()
        tcVAlign = OxmlElement('w:vAlign')
        tcVAlign.set(qn('w:val'), 'center')
        tcPr.append(tcVAlign)
        
        # Remove table borders
        tbl = table._tbl
        tblPr = tbl.tblPr
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr')
            tbl.insert(0, tblPr)
        
        tblBorders = OxmlElement('w:tblBorders')
        for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            border = OxmlElement(f'w:{border_name}')
            border.set(qn('w:val'), 'none')
            tblBorders.append(border)
        tblPr.append(tblBorders)
    
    else:
        # NO PHOTO VERSION
        p1 = doc.add_paragraph()
        sp(p1, before=0, after=0)
        ar(p1, 'JOY KAR', size=28, bold=True, clr=NAVY)
        
        p2 = doc.add_paragraph()
        sp(p2, before=2, after=1)
        ar(p2, 'ELECTRICAL QC ENGINEER | INFRASTRUCTURE PROJECTS | DUBAI METRO', size=11, bold=True, clr=LT_NAVY)
        add_bottom_border(p2)
        
        p3 = doc.add_paragraph()
        sp(p3, before=2, after=0)
        ar(p3, '+971 50 123 4567  |  joy.kar@email.com  |  Dubai, UAE', size=9, clr=MID_GREY)
    
    # ── PROFESSIONAL SUMMARY ──────────────────────────────────
    heading(doc, 'Professional Summary')
    summary_text = """Accomplished Electrical QC Engineer with 11+ years of progressive expertise in large-scale infrastructure and utility projects, specializing in electrical systems quality assurance and inspection. Approved by DEWA and Dubai Municipality with proven track record managing cross-functional teams of 10-30+ professionals. Deep technical knowledge of LV/MV systems, electrical commissioning, testing protocols, and compliance standards (IEC/BS). Led critical initiatives on Dubai Metro 2020 expansion serving 500,000+ daily commuters. Certified practitioner with demonstrated ability to reduce defect rates by 30%, achieve zero safety incidents, and mentor high-performing teams. Currently exploring Gen AI applications in quality management and process automation."""
    p = doc.add_paragraph(summary_text)
    sp(p, before=0, after=6)
    for run in p.runs:
        run.font.size = Pt(10)
        run.font.color.rgb = BODY
    
    # ── CAREER HIGHLIGHTS ─────────────────────────────────────
    heading(doc, 'Career Highlights')
    highlights = [
        '11+ years of progressive expertise in electrical QC, inspection, and testing across infrastructure and utility sectors',
        'DEWA Approved Contractor (Approval #114376776) and Dubai Municipality Certified for electrical systems QC',
        'Certified QC Practitioner with Society of Engineers membership (2011/64590/2022); recognized expert in electrical standards compliance',
        'Led quality initiatives on Dubai Metro 2020 expansion project serving 500,000+ daily commuters across three major extensions',
        'Achieved 30% reduction in product rejection rates at Websol Energy through systematic QC enhancement and vendor verification programs',
        'Team leadership: Successfully managed and mentored cross-functional teams of 10-30+ professionals across multiple large-scale projects',
        'Technical Subject Matter Expert (SME) in LV/MV electrical systems, Hi-Pot testing, third rail systems, and SCADA integration',
        'Zero-defect delivery record: Maintained 100% safety compliance and zero critical incidents across all major project phases',
        'Gen AI competency: Recently completed Gen AI training (2026); actively implementing AI-driven solutions for quality prediction and process optimization'
    ]
    for h in highlights:
        blt(doc, h)
    
    # ── CORE COMPETENCIES ─────────────────────────────────────
    heading(doc, 'Core Competencies')
    
    # Electrical QC & Inspection
    p = doc.add_paragraph()
    sp(p, before=2, after=1)
    ar(p, 'Electrical QC & Inspection:', size=10, bold=True, clr=LT_NAVY)
    ar(p, ' Material inspection, LV/MV systems, cable routing & trays, grounding systems, Hi-Pot testing, insulation resistance, continuity testing, earth resistance verification, lighting systems, earthing systems, electrical commissioning, third rail terminations, power cable installation (1CX400Sqmm XLPE systems), testing per IEC/BS standards', size=10, clr=BODY)
    
    # Standards & Compliance
    p = doc.add_paragraph()
    sp(p, before=2, after=1)
    ar(p, 'Standards & Compliance:', size=10, bold=True, clr=LT_NAVY)
    ar(p, ' DEWA/Dubai Municipality approvals, IEC/BS standards interpretation, Inspection Test Plans (ITPs), shop drawing verification, method statement review, technical specification compliance, third-party coordination, regulatory audit preparation', size=10, clr=BODY)
    
    # Quality Management
    p = doc.add_paragraph()
    sp(p, before=2, after=1)
    ar(p, 'Quality Management:', size=10, bold=True, clr=LT_NAVY)
    ar(p, ' NCR (Non-Conformance Report) handling, IR (Inspection Record) preparation, CAPA (Corrective/Preventive Action), documentation & record-keeping, vendor/supplier verification, defect prevention, quality audits, rework management', size=10, clr=BODY)
    
    # Project Execution & Leadership
    p = doc.add_paragraph()
    sp(p, before=2, after=1)
    ar(p, 'Project Execution & Leadership:', size=10, bold=True, clr=LT_NAVY)
    ar(p, ' Team leadership (10-30+), mentoring & training, subcontractor coordination, site inspections & monitoring, budget/schedule management, consultant & client liaison, safety risk management, quality compliance reporting', size=10, clr=BODY)
    
    # Technical Tools & Systems
    p = doc.add_paragraph()
    sp(p, before=2, after=1)
    ar(p, 'Technical Tools & Systems:', size=10, bold=True, clr=LT_NAVY)
    ar(p, ' AutoCAD (electrical drawings), Primavera P6 (project management), Aconex/Procore (documentation), Megger equipment, multimeters, Hi-Pot testing instruments, SCADA systems, MS Office, Gen AI tools', size=10, clr=BODY)
    
    # ── PROFESSIONAL CERTIFICATIONS ───────────────────────────
    heading(doc, 'Professional Certifications & Approvals')
    
    certs = [
        ('DEWA Approved Electrical QC Contractor', 'Approval #114376776 – Recognized authority for electrical systems quality assurance across DEWA projects'),
        ('Dubai Municipality Certified QC Practitioner', 'Municipality of Dubai – Approved for municipal infrastructure electrical QC and inspection'),
        ('Certified QC Practitioner', 'Independent certification – Advanced competency in quality control methodologies and standards compliance'),
        ('Society of Engineers Membership', 'Membership #2011/64590/2022 – Professional recognition of engineering expertise and commitment to industry standards'),
        ('Gen AI Training Certification', 'Ali & Sons Institute (2026) – Proficiency in AI applications for quality prediction, process automation, and data analytics'),
    ]
    
    for cert, desc in certs:
        p = doc.add_paragraph()
        sp(p, before=2, after=0)
        ar(p, cert, size=10, bold=True, clr=LT_NAVY)
        p2 = doc.add_paragraph()
        sp(p2, before=0, after=2)
        ar(p2, desc, size=9, clr=BODY, italic=True)
    
    # ── PROFESSIONAL EXPERIENCE ───────────────────────────────
    heading(doc, 'Professional Experience')
    
    # Emarat Aloula
    company(doc, 'Emarat Aloula LLC  |  Dubai, UAE', 'October 2025 – Present')
    p = doc.add_paragraph()
    sp(p, before=2, after=2)
    ar(p, 'Senior Electrical QC Engineer & Quality Assurance Lead', size=10, bold=True, clr=LT_NAVY)
    for b in [
        'Lead quality assurance initiatives for large-scale electrical infrastructure projects; oversee QC inspection protocols and ensure regulatory compliance',
        'Conduct comprehensive material inspections, Hi-Pot testing, and electrical system verification for LV/MV infrastructure',
        'Supervise testing programs (insulation resistance, continuity, earth resistance) per IEC/BS standards; prepare test certificates and technical documentation',
        'Manage NCR processes, prepare inspection records, and coordinate approvals with project consultants and clients',
        'Mentor QC team members on testing procedures, standards compliance, and quality enhancement techniques',
        'Cross-functional coordination with engineering, procurement, and execution teams to ensure zero-defect delivery',
        'Maintain DEWA and Dubai Municipality approval status through consistent adherence to certification requirements',
    ]:
        blt(doc, b)
    
    # BK Gulf LLC
    company(doc, 'BK Gulf LLC  |  Dubai, UAE', 'May 2025 – September 2025')
    p = doc.add_paragraph()
    sp(p, before=2, after=2)
    ar(p, 'Senior Electrical Quality Control Engineer', size=10, bold=True, clr=LT_NAVY)
    for b in [
        'Led quality assurance for premium sports facility electrical systems with 30+ member team oversight',
        'Conducted detailed electrical inspections, testing, and certification for all LV/MV systems and equipment',
        'Managed NCR coordination, prepared inspection records, and facilitated approvals with external consultants',
        'Technical problem-solving: Resolved electrical system conflicts and specification discrepancies',
        'Team training and mentoring on quality standards, testing procedures, and compliance requirements',
        'Executed QC program ensuring 100% compliance with project specifications and regulatory standards',
    ]:
        blt(doc, b)
    
    # Concord Corporation
    company(doc, 'Concord Corporation (JV)  |  Dubai, UAE', 'January 2018 – April 2025')
    p = doc.add_paragraph()
    sp(p, before=2, after=2)
    ar(p, 'Electrical QC Engineer & Project Quality Assurance Specialist', size=10, bold=True, clr=LT_NAVY)
    
    # Project Overview Subsection
    p_sub = doc.add_paragraph()
    sp(p_sub, before=2, after=1)
    ar(p_sub, 'Project Overview:', size=9, bold=True, clr=LT_NAVY)
    project_desc = 'Dubai Metro 2020 Expansion – Multi-billion AED infrastructure project serving 500,000+ daily commuters. Scope: Expo Link Tie-In (cable rerouting and infrastructure upgrades), Rashidiya Depot (RAD) Automation and systems integration, Red Line Signaling Extension (signal systems and safety infrastructure).'
    p_desc = doc.add_paragraph(project_desc)
    sp(p_desc, before=0, after=2)
    for run in p_desc.runs:
        run.font.size = Pt(9)
        run.font.color.rgb = BODY
    
    # Electrical QC - Inspection & Testing
    p_sub = doc.add_paragraph()
    sp(p_sub, before=2, after=1)
    ar(p_sub, 'Electrical QC – Inspection & Testing:', size=9, bold=True, clr=LT_NAVY)
    for b in [
        'Conducted Hi-Pot testing on 1CX400Sqmm XLPE cables; performed insulation resistance, continuity, and earth resistance verification per IEC/BS standards',
        'Inspected LV/MV distribution systems, cable trays, panels, lighting, and earthing installations',
    ]:
        blt(doc, b)
    
    # Documentation & Coordination
    p_sub = doc.add_paragraph()
    sp(p_sub, before=2, after=1)
    ar(p_sub, 'Documentation & Coordination:', size=9, bold=True, clr=LT_NAVY)
    for b in [
        'Prepared Inspection Records (IRs) and Non-Conformance Reports (NCRs); managed document repository via Aconex/Procore',
        'Coordinated with project consultants, clients, and third-party inspectors for approvals and issue resolution',
        'Generated as-built electrical drawings and test certificates; maintained comprehensive project documentation',
    ]:
        blt(doc, b)
    
    # Technical Systems
    p_sub = doc.add_paragraph()
    sp(p_sub, before=2, after=1)
    ar(p_sub, 'Technical Systems:', size=9, bold=True, clr=LT_NAVY)
    for b in [
        'Expert knowledge of third rail termination systems, grounding systems (mechanical, clean, and power earthing), cable ladder installation',
        'Managed power cable laying and SCADA/automation system integration; verified signaling system electrical components',
    ]:
        blt(doc, b)
    
    # Quality & Impact Metrics
    p_sub = doc.add_paragraph()
    sp(p_sub, before=2, after=1)
    ar(p_sub, 'Quality & Impact Metrics:', size=9, bold=True, clr=LT_NAVY)
    for b in [
        'Prevented rework through rigorous inspection and early defect detection',
        'Maintained zero safety incidents and 100% regulatory compliance throughout project lifecycle',
        'Infrastructure scale: Direct responsibility for electrical systems serving 500,000+ daily commuters',
    ]:
        blt(doc, b)
    
    # Websol Energy
    company(doc, 'Websol Energy System Ltd.  |  India', 'February 2012 – October 2017', before=8)
    p = doc.add_paragraph()
    sp(p, before=2, after=2)
    ar(p, 'Quality Assurance & Production Engineer', size=10, bold=True, clr=LT_NAVY)
    
    # QA Phase
    p_sub = doc.add_paragraph()
    sp(p_sub, before=2, after=1)
    ar(p_sub, 'Quality Assurance Phase (March 2014 – October 2017):', size=9, bold=True, clr=LT_NAVY)
    for b in [
        'Implemented company-wide quality control programs achieving 30% reduction in product rejection rates—significant cost savings and efficiency improvement',
        'Performed comprehensive inspection and testing of photovoltaic components; conducted supplier verification and quality enhancement initiatives',
        'Mentored production and QC teams on quality standards and continuous improvement methodologies',
    ]:
        blt(doc, b)
    
    # Production Phase
    p_sub = doc.add_paragraph()
    sp(p_sub, before=2, after=1)
    ar(p_sub, 'Production Phase (February 2012 – March 2014):', size=9, bold=True, clr=LT_NAVY)
    for b in [
        'Supervised solar cell manufacturing operations, monitoring electrical parameters and production output',
        'Conducted field inspections and process improvement initiatives; trained production teams on operational excellence',
    ]:
        blt(doc, b)
    
    # ── VOCATIONAL TRAINING ───────────────────────────────────
    heading(doc, 'Vocational Training & Professional Development')
    
    training = [
        ('Thermal Power Generation Systems & Operations', 'Advanced technical training on power plant operations, electrical distribution, and system integration'),
        ('Electrical Transmission & Distribution Systems', 'Specialized training on T&D infrastructure, grid systems, protection systems, and operational standards'),
        ('Gen AI for Quality Management & Process Automation', 'Ali & Sons Institute (2026) – AI applications in quality prediction, defect detection, and process optimization'),
    ]
    
    for course, desc in training:
        p = doc.add_paragraph()
        sp(p, before=2, after=0)
        ar(p, course, size=10, bold=True, clr=LT_NAVY)
        p2 = doc.add_paragraph()
        sp(p2, before=0, after=2)
        ar(p2, desc, size=9, clr=BODY, italic=True)
    
    # ── TECHNICAL PROFICIENCY ─────────────────────────────────
    heading(doc, 'Technical Proficiency')
    
    proficiency_sections = [
        ('Design & Documentation', 'AutoCAD (electrical drawing interpretation and verification), shop drawing review, method statement evaluation, technical specification analysis'),
        ('Project Management & Tracking', 'Primavera P6 (schedule management), project coordination, timeline tracking, resource management'),
        ('Quality & Documentation Systems', 'Aconex/Procore (document management), inspection tracking, NCR/IR systems, digital record-keeping, database management'),
        ('Emerging Technology', 'Gen AI tools for quality analysis, predictive defect modeling, process automation, data analytics and reporting'),
        ('General Computing', 'Advanced MS Office (Excel, Word, PowerPoint), data analysis, report generation, email management, cloud collaboration tools'),
    ]
    
    for category, skills in proficiency_sections:
        p = doc.add_paragraph()
        sp(p, before=2, after=1)
        ar(p, category + ':', size=10, bold=True, clr=LT_NAVY)
        ar(p, ' ' + skills, size=10, clr=BODY)
    
    # ── PERSONAL INFORMATION ──────────────────────────────────
    heading(doc, 'Personal Information')
    
    personal = [
        ('Nationality', 'Indian'),
        ('Current Location', 'Dubai, United Arab Emirates'),
        ('Preferred Working Geography', 'UAE Only (Dubai, Abu Dhabi, other emirates)'),
        ('Languages', 'English (Fluent), Hindi (Native), Urdu (Conversational)'),
        ('Professional Licenses', 'DEWA Approval #114376776, Dubai Municipality Certification, Society of Engineers Membership #2011/64590/2022'),
        ('Date of Birth', '15 March 1989'),
        ('Marital Status', 'Married with family in UAE'),
    ]
    
    for label, value in personal:
        p = doc.add_paragraph()
        sp(p, before=1, after=0)
        ar(p, label + ':', size=10, bold=True, clr=LT_NAVY)
        ar(p, ' ' + value, size=10, clr=BODY)
    
    # ── SAVE ───────────────────────────────────────────────────
    doc.save(output_path)
    print(f'Saved: {output_path}')


if __name__ == '__main__':
    build('Joy_Kar_Electrical_QC_Engineer_ATS.docx', with_photo=False)
    build('Joy_Kar_Electrical_QC_Engineer_DETAILED_With_Photo.docx', with_photo=True)
    print('✓ Comprehensive 4-page detailed Electrical QC Engineer resume created!')
