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


def subheader(doc, text, before=4, after=1):
    p = doc.add_paragraph()
    sp(p, before=before, after=after)
    ar(p, text, size=10, bold=True, clr=NAVY)


# ── Build Document ────────────────────────────────────────────
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

    # ── HEADER (with or without photo) ──────────────────────
    if with_photo:
        tbl = doc.add_table(rows=1, cols=2)
        tbl.allow_autofit = False
        tbl.columns[0].width = Cm(13.2)
        tbl.columns[1].width = Cm(3.4)

        # Remove all table borders
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

        # Name in left cell
        p_name = left_cell.paragraphs[0]
        sp(p_name, before=0, after=2)
        ar(p_name, 'ATREE BAIDYA', size=26, bold=True, clr=NAVY)

        # Headline in left cell
        p_hl = left_cell.add_paragraph()
        sp(p_hl, before=0, after=5)
        add_bottom_border(p_hl, sz=6)
        ar(p_hl, 'HR Professional  |  Recruitment  |  HR Operations  |  Employee Engagement', size=10.5, clr=LT_NAVY)

        # Contact in left cell
        p_ct = left_cell.add_paragraph()
        sp(p_ct, before=4, after=0)
        ar(p_ct, 'Dubai, UAE  |  +971-554463371  |  atreebaidya@gmail.com', size=9, clr=MID_GREY)
        p_ct2 = left_cell.add_paragraph()
        sp(p_ct2, before=1, after=0)
        ar(p_ct2, 'linkedin.com/in/atreebaidya', size=9, clr=MID_GREY)

        # Photo in right cell
        p_photo = right_cell.paragraphs[0]
        p_photo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_img = p_photo.add_run()
        run_img.add_picture('C:/xampp/htdocs/resume/atree.jpg', width=Cm(3.0))

        doc.add_paragraph()  # spacer

    else:
        # NAME
        p = doc.add_paragraph()
        sp(p, before=0, after=2)
        ar(p, 'ATREE BAIDYA', size=28, bold=True, clr=NAVY)

        # HEADLINE
        p = doc.add_paragraph()
        sp(p, before=0, after=5)
        add_bottom_border(p, sz=6)
        ar(p, 'HR Professional  ·  Recruitment & Selection  ·  HR Operations  ·  Employee Engagement  ·  Compliance', size=11, clr=LT_NAVY)

        # CONTACT
        p = doc.add_paragraph()
        sp(p, before=4, after=8)
        ar(p, 'Dubai, UAE  ·  +971-554463371  ·  atreebaidya@gmail.com  ·  linkedin.com/in/atreebaidya', size=9, clr=MID_GREY)

    # ── EXECUTIVE PROFILE ────────────────────────────────────
    heading(doc, 'Professional Summary', before=6 if with_photo else 0)
    p = doc.add_paragraph()
    sp(p, before=4, after=0)
    ar(p,
       'Strategic HR professional with MBA in Human Resource Management and 2+ years of progressive experience '
       'in recruitment, employee onboarding, HR operations and compliance across multinational organizations. '
       'Proven expertise in full-cycle recruitment, HR process optimization and building scalable onboarding frameworks. '
       'Skilled in HR systems administration, data analytics (SQL, Excel), and employee engagement initiatives. '
       'Strong background in administrative operations and cross-functional coordination with demonstrated ability '
       'to enhance HR efficiency through process improvements and compliance management. Seeking HR Specialist or '
       'Recruitment Coordinator role in Dubai to drive organizational growth through talent acquisition and retention.',
       size=10, clr=BODY)

    # ── CORE COMPETENCIES & TECHNICAL SKILLS ─────────────────
    heading(doc, 'Core Competencies & Skills')
    for label, content in [
        ('HR Expertise:',
         'Full-Cycle Recruitment  ·  CV Screening & Interview Coordination  ·  Employee Onboarding  ·  HR Operations  ·  Payroll & Attendance Management  ·  HR Compliance  ·  Employee Engagement  ·  HR Documentation'),
        ('Technical & Data Skills:',
         'MS Excel (Pivot Tables, VLOOKUP, Advanced Functions)  ·  SQL (Database Queries, Data Analysis)  ·  HR Metrics & Reporting  ·  HR Management Software  ·  MS Office Suite  ·  Google Workspace'),
        ('Soft Skills & Leadership:',
         'Communication  ·  Interpersonal Relations  ·  Problem-Solving  ·  Conflict Resolution  ·  Time Management  ·  Team Coordination  ·  Stakeholder Management'),
    ]:
        p = doc.add_paragraph()
        sp(p, before=2, after=2)
        ar(p, label + '  ', size=10, bold=True, clr=NAVY)
        ar(p, content, size=10, clr=BODY)

    # ── PROFESSIONAL EXPERIENCE ──────────────────────────────
    heading(doc, 'Professional Experience')

    # Sensia Ventures Group
    company(doc, 'Sensia Ventures Group  |  World Trade Center, Dubai, UAE', 'October 2025 – Present', before=6)
    p = doc.add_paragraph()
    sp(p, before=2, after=2)
    ar(p, 'HR Assistant', size=10, bold=True, clr=LT_NAVY)
    for b in [
        'Support daily HR operations and administrative activities, coordinating recruitment, onboarding, and employee data management across the organization',
        'Manage full recruitment lifecycle: job posting, CV screening, candidate communication, interview scheduling and coordination with hiring managers',
        'Prepare comprehensive employment documentation including offer letters, employment contracts, NDAs and onboarding agreements with 100% accuracy and compliance',
        'Maintain employee records, HR databases and confidential documentation to ensure data security, accuracy and compliance with UAE labor regulations',
        'Support employee onboarding and offboarding processes: coordinate pre-entry activities, new hire integration, exit interviews and final documentation',
        'Coordinate with internal departments and employees regarding HR queries, leave requests, benefits questions and other HR-related matters',
        'Manage attendance records, leave tracking, employee data entry and payroll information to ensure accurate HR database maintenance',
        'Prepare HR reports and analytics using MS Excel (Pivot Tables, VLOOKUP) to support HR decision-making and organizational planning',
        'Assist in HR policy implementation, employee engagement activities and special HR projects to support organizational culture and compliance',
        'Handle and maintain confidentiality of sensitive employee information, ensuring proper data protection and compliance with privacy regulations',
    ]:
        blt(doc, b)

    # Websol Energy Systems Limited
    company(doc, 'Websol Energy Systems Limited  |  Kolkata, India', 'June 2024 – September 2025', before=8)
    p = doc.add_paragraph()
    sp(p, before=2, after=2)
    ar(p, 'Administrative Assistant', size=10, bold=True, clr=LT_NAVY)
    for b in [
        'Supported daily administrative and HR operations, ensuring smooth office coordination and efficient workflow across departments',
        'Assisted with recruitment activities including candidate tracking, interview coordination and scheduling with internal stakeholders',
        'Maintained employee records, personnel files and HR documentation to ensure accuracy, organization and compliance with company policies',
        'Coordinated with HR department on onboarding, employee queries and HR-related matters to support seamless employee experience',
        'Handled comprehensive documentation management: prepared reports, managed email records and maintained organized company files',
        'Supported attendance tracking, leave management and employee data entry to maintain accurate HR databases and payroll information',
        'Coordinated meetings, calls and internal communications between departments, management and external vendors',
        'Assisted with HR projects, administrative initiatives and employee engagement activities to support organizational objectives',
    ]:
        blt(doc, b)

    # HR Internship
    company(doc, 'National HRD Network (NHRDN)  |  Summer Internship Program, India', 'Summer 2023', before=8)
    p = doc.add_paragraph()
    sp(p, before=2, after=2)
    ar(p, 'HR Research Intern  |  Employee Engagement Study', size=10, bold=True, clr=LT_NAVY)
    for b in [
        'Conducted comprehensive research on post-COVID employee engagement trends across multiple business sectors and industries',
        'Analyzed key variables impacting engagement: leadership effectiveness, communication channels, recognition & rewards programs and work-life balance',
        'Identified emerging HR challenges and organizational behavioral patterns through data analysis and stakeholder interviews',
        'Prepared detailed research findings and recommendations for organizational leadership',
    ]:
        blt(doc, b)

    # Chevron
    company(doc, 'Chevron  |  Chittagong, Bangladesh', 'February 2020 – December 2021', before=8)
    p = doc.add_paragraph()
    sp(p, before=2, after=2)
    ar(p, 'Administrative Officer', size=10, bold=True, clr=LT_NAVY)
    for b in [
        'Supported daily HR and administrative operations including coordination of recruitment, onboarding and employee documentation processes',
        'Maintained employee records, attendance documentation, leave tracking and confidential personnel files in secure database systems',
        'Coordinated with HR department on recruitment activities, onboarding procedures, offboarding processes and employee documentation requirements',
        'Prepared HR reports and documentation while ensuring strict adherence to data protection, confidentiality and compliance regulations',
        'Managed visitor coordination, telephone communications, appointment scheduling and internal communications across departments',
        'Handled employee queries, requests and administrative matters with professionalism and confidentiality',
    ]:
        blt(doc, b)

    # ── EDUCATION ────────────────────────────────────────────
    heading(doc, 'Education')
    
    p = doc.add_paragraph()
    sp(p, before=4, after=0)
    ar(p, 'Master of Business Administration (MBA) in Human Resources', size=10, bold=True, clr=BODY)
    ar(p, '  |  ', size=10, clr=BODY)
    ar(p, 'Indian Institute of Social Welfare & Business Management (IISWBM) / University of Calcutta, India', size=9, italic=True, clr=MID_GREY)
    p_gpa = doc.add_paragraph()
    sp(p_gpa, before=0, after=4)
    ar(p_gpa, 'GPA: 60.22%  |  2022-2024', size=9, clr=MID_GREY)

    p = doc.add_paragraph()
    sp(p, before=0, after=0)
    ar(p, 'Bachelor of Science (BSc) in Psychology', size=10, bold=True, clr=BODY)
    ar(p, '  |  ', size=10, clr=BODY)
    ar(p, 'Chittagong Government Women\'s College (NU), Bangladesh', size=9, italic=True, clr=MID_GREY)
    p_gpa2 = doc.add_paragraph()
    sp(p_gpa2, before=0, after=4)
    ar(p_gpa2, 'CGPA: 3.32  |  2020', size=9, clr=MID_GREY)

    p = doc.add_paragraph()
    sp(p, before=0, after=0)
    ar(p, 'Higher Secondary (Business Studies)', size=10, bold=True, clr=BODY)
    ar(p, '  |  ', size=10, clr=BODY)
    ar(p, 'Chittagong Government City College, Bangladesh', size=9, italic=True, clr=MID_GREY)
    p_gpa3 = doc.add_paragraph()
    sp(p_gpa3, before=0, after=4)
    ar(p_gpa3, 'Score: 76.60%  |  2015', size=9, clr=MID_GREY)

    p = doc.add_paragraph()
    sp(p, before=0, after=0)
    ar(p, 'Secondary (Science)', size=10, bold=True, clr=BODY)
    ar(p, '  |  ', size=10, clr=BODY)
    ar(p, 'Aparna Charan City Corporation Girls\' High School, Bangladesh', size=9, italic=True, clr=MID_GREY)
    p_gpa4 = doc.add_paragraph()
    sp(p_gpa4, before=0, after=4)
    ar(p_gpa4, 'Score: 90.00%  |  2013', size=9, clr=MID_GREY)

    # ── CERTIFICATIONS ───────────────────────────────────────
    heading(doc, 'Certifications & Training')
    for c in [
        'Diploma in Computer Studies, Grade A – Chittagong City Corporation Computer Institute, 2016',
        'HR Management & Operations Training – National HRD Network (NHRDN), 2023',
    ]:
        blt(doc, c)

    # ── ADDITIONAL INFORMATION ───────────────────────────────
    heading(doc, 'Additional Information')
    
    p = doc.add_paragraph()
    sp(p, before=4, after=2)
    ar(p, 'Languages:  ', size=10, bold=True, clr=NAVY)
    ar(p, 'English (Professional)  ·  Hindi (Fluent)  ·  Bengali (Native)', size=10, clr=BODY)

    p = doc.add_paragraph()
    sp(p, before=0, after=2)
    ar(p, 'Nationality:  ', size=10, bold=True, clr=NAVY)
    ar(p, 'Bangladeshi  |  Work Authorization: UAE Valid', size=10, clr=BODY)

    p = doc.add_paragraph()
    sp(p, before=0, after=2)
    ar(p, 'Location:  ', size=10, bold=True, clr=NAVY)
    ar(p, 'Dubai, UAE  |  Available for Office and Work-from-Home roles', size=10, clr=BODY)

    p = doc.add_paragraph()
    sp(p, before=0, after=0)
    ar(p, 'Areas of Interest:  ', size=10, bold=True, clr=NAVY)
    ar(p, 'HR Operations, Talent Acquisition, Employee Relations, Public Relations, Professional Development', size=10, clr=BODY)

    doc.save(output_path)
    print(f'Saved: {output_path}')


# ── Run both versions ─────────────────────────────────────────
BASE = 'C:/xampp/htdocs/resume/'
build(BASE + 'Atree_Baidya_HR_Professional_FINAL_ATS.docx',   with_photo=False)
build(BASE + 'Atree_Baidya_HR_Professional_FINAL_With_Photo.docx', with_photo=True)
print('\n✓ JD-optimized resume versions created successfully with atree.jpg!')
