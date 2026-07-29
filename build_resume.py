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
        ar(p_name, 'DEBABRATA MANDAL', size=26, bold=True, clr=NAVY)

        # Headline in left cell
        p_hl = left_cell.add_paragraph()
        sp(p_hl, before=0, after=5)
        add_bottom_border(p_hl, sz=6)
        ar(p_hl, 'SAP Quality Engineering Manager  |  S/4HANA  |  SAP Fiori  |  SAP BTP  |  Global Delivery', size=10.5, clr=LT_NAVY)

        # Contact in left cell
        p_ct = left_cell.add_paragraph()
        sp(p_ct, before=4, after=0)
        ar(p_ct, 'Kolkata, India  |  +91-8335866633  |  debabratamandal.ju@gmail.com', size=9, clr=MID_GREY)
        p_ct2 = left_cell.add_paragraph()
        sp(p_ct2, before=1, after=0)
        ar(p_ct2, 'linkedin.com/in/debabrata-mandal-84b93256  |  Open to Relocation – Germany / DACH', size=9, clr=MID_GREY)

        # Photo in right cell
        p_photo = right_cell.paragraphs[0]
        p_photo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_img = p_photo.add_run()
        run_img.add_picture('C:/xampp/htdocs/resume/unnamed.png', width=Cm(3.0))

        doc.add_paragraph()  # spacer

    else:
        # NAME
        p = doc.add_paragraph()
        sp(p, before=0, after=2)
        ar(p, 'DEBABRATA MANDAL', size=28, bold=True, clr=NAVY)

        # HEADLINE
        p = doc.add_paragraph()
        sp(p, before=0, after=5)
        add_bottom_border(p, sz=6)
        ar(p, 'SAP Quality Engineering Manager  ·  S/4HANA  ·  SAP Fiori  ·  SAP BTP  ·  Global Delivery', size=11, clr=LT_NAVY)

        # CONTACT
        p = doc.add_paragraph()
        sp(p, before=4, after=8)
        ar(p, 'Kolkata, India  ·  +91-8335866633  ·  debabratamandal.ju@gmail.com  ·  linkedin.com/in/debabrata-mandal-84b93256  ·  Open to Relocation – Germany / DACH', size=9, clr=MID_GREY)

    # ── EXECUTIVE PROFILE ────────────────────────────────────
    heading(doc, 'Executive Profile', before=6 if with_photo else 0)
    p = doc.add_paragraph()
    sp(p, before=4, after=0)
    ar(p,
       'Results-driven SAP Quality Engineering Manager with 14+ years of IT experience, including 10+ years '
       'specialising in SAP Testing and Quality Governance. Currently Manager at PwC India, leading enterprise '
       'SAP Test Management for Bayer’s global S/4HANA transformation programme. One of a select group of '
       'India-based SAP QA professionals with verified on-site delivery experience in Germany — having led '
       'quality engineering at Bayer (Germany) across two separate programmes and coordinated the SAP RISE cloud '
       'migration across 50+ product teams. Expert in Test Strategy, Risk-Based Testing, Release Governance, '
       'SIT, UAT and Hypercare, with a proven track record of protecting go-live integrity through data-driven '
       'quality decisions and executive stakeholder management.',
       size=10, clr=BODY)

    # ── CAREER HIGHLIGHTS ────────────────────────────────────
    heading(doc, 'Career Highlights')
    for h in [
        '14+ years in IT; 10+ years specialising in SAP Testing and Quality Engineering at Big 4 and Tier-1 consulting firms',
        'SAP Test / QA Manager responsibilities since 2022 across enterprise S/4HANA, SAP Fiori and SAP BTP programmes',
        'Led and managed distributed QA teams of up to 15 professionals across onsite and offshore delivery models',
        'On-site programme delivery in Germany (Bayer ×2, SAP Goes RISE) and Brazil (Petrobras) — direct DACH enterprise culture exposure',
        'Coordinated testing readiness across 50+ independent product teams for a single SAP RISE cloud migration programme',
        'Executive Go/No-Go decision support and quality dashboard presentation to programme directors',
        'Coordinated emergency P1 hotfix testing for a post-go-live financial processing defect at a major UK retail client',
        'Conducted 20+ technical interviews; mentored QA engineers across Accenture and TCS delivery organisations',
    ]:
        blt(doc, h)

    # ── CORE EXPERTISE ───────────────────────────────────────
    heading(doc, 'Core Expertise')
    for label, content in [
        ('SAP Technology:',
         'SAP S/4HANA  ·  SAP SD  ·  SAP MM  ·  SAP Fiori  ·  SAP BTP  ·  SAP HCM'),
        ('Quality Engineering:',
         'Test Strategy  ·  Test Planning  ·  SIT  ·  UAT  ·  Regression Testing  ·  Defect Management  ·  Hypercare  ·  Risk-Based Testing'),
        ('Leadership & Governance:',
         'Release Readiness  ·  Go/No-Go Support  ·  Test Governance  ·  Executive Reporting  ·  Stakeholder Management  ·  Team Leadership'),
        ('Tools & Platforms:',
         'Jira  ·  HP ALM  ·  SAP Solution Manager  ·  Azure DevOps  ·  Tricentis Tosca (L1 & L2 Certified)'),
    ]:
        p = doc.add_paragraph()
        sp(p, before=2, after=2)
        ar(p, label + '  ', size=10, bold=True, clr=NAVY)
        ar(p, content, size=10, clr=BODY)

    # ── PROFESSIONAL EXPERIENCE ──────────────────────────────
    heading(doc, 'Professional Experience')

    # PwC India
    company(doc, 'PwC India', 'Apr 2026 – Present', before=6)
    p = doc.add_paragraph()
    sp(p, before=2, after=2)
    ar(p, 'Manager  |  SAP Test Manager', size=10, bold=True, clr=LT_NAVY)
    ar(p, '   |   Bayer AG (Germany)', size=9.5, italic=True, clr=MID_GREY)
    for b in [
        'Lead enterprise SAP Test Management for Bayer’s global S/4HANA transformation programme.',
        'Define and govern Test Strategy, Test Plans, quality gates and release readiness criteria.',
        'Manage distributed QA team of up to 15 members across onsite and offshore delivery.',
        'Drive end-to-end quality execution: SIT, UAT, Regression Testing, Hypercare and production validation.',
        'Present executive quality dashboards, programme risks and Go/No-Go recommendations to programme directors.',
        'Partner with SAP functional teams, business stakeholders and programme leadership to ensure controlled releases.',
    ]:
        blt(doc, b)

    # Accenture
    company(doc, 'Accenture', 'Nov 2022 – Apr 2026', before=8)
    p = doc.add_paragraph()
    sp(p, before=2, after=3)
    ar(p, 'Quality Engineering Associate Manager', size=10, bold=True, clr=LT_NAVY)

    subheader(doc, 'Bayer AG (Germany)  |  SAP Test Manager')
    for b in [
        'Led SAP S/4HANA, SAP Fiori and SAP BTP testing and quality governance.',
        'Owned test execution oversight, quality governance framework and release stakeholder communication.',
        'Delivered risk-based quality recommendations to support programme Go/No-Go decisions.',
    ]:
        blt(doc, b)

    subheader(doc, 'SAP Goes RISE (Germany)  |  Test Coordinator', before=5)
    for b in [
        'Coordinated testing readiness across 50+ independent product teams for SAP’s enterprise migration to RISE.',
        'Operated a programme office model: daily quality standups, release tollgates and cross-workstream dependency tracking.',
        'Managed release readiness reporting and supported production cutover and cloud migration validation.',
    ]:
        blt(doc, b)

    subheader(doc, 'Petrobras (Brazil)  |  SAP Test Manager', before=5)
    for b in [
        'Directed SAP Test Management for Petrobras’s enterprise SAP transformation initiative.',
        'Established Test Strategy, test planning, governance frameworks and executive release reporting.',
        'Led SIT, UAT, Regression and Hypercare cycles; mentored QA team and coordinated cross-functional stakeholders.',
    ]:
        blt(doc, b)

    # TCS
    company(doc, 'Tata Consultancy Services', 'Jul 2012 – Nov 2022', before=8)
    p = doc.add_paragraph()
    sp(p, before=2, after=3)
    ar(p, 'IT Analyst  |  SAP QA Specialist', size=10, bold=True, clr=LT_NAVY)

    subheader(doc, 'Kingfisher (UK)  |  SAP Test Lead')
    blt(doc, 'Led SAP SD/MM testing: planning, execution, defect management and business validation across retail business processes.')

    subheader(doc, 'ASDA (UK)  |  UAT Coordinator', before=5)
    blt(doc, 'Managed UAT planning, business sign-off, defect governance and cross-functional stakeholder communication.')

    subheader(doc, 'Production Incident Response  |  Major UK Retail Client', before=5)
    blt(doc, 'Coordinated emergency hotfix testing for a P1 post-go-live defect causing financial processing errors; managed rapid validation and safe return to production.')

    # ── CERTIFICATIONS ───────────────────────────────────────
    heading(doc, 'Certifications')
    for c in [
        'Microsoft Certified: Azure Fundamentals (AZ-900)',
        'Microsoft Certified: Azure Data Fundamentals (DP-900)',
        'Microsoft Certified: Security, Compliance & Identity Fundamentals (SC-900)',
        'Tricentis Tosca Automation Specialist – Level 1 & Level 2',
        'Claude Certified Architect – Foundations (AI & Automation)',
    ]:
        blt(doc, c)

    # ── EDUCATION ────────────────────────────────────────────
    heading(doc, 'Education')
    p = doc.add_paragraph()
    sp(p, before=4, after=2)
    ar(p, 'Bachelor of Engineering (Pharmacy)', size=10, bold=True, clr=BODY)
    ar(p, '  |  Jadavpur University, Kolkata', size=10, clr=DK_GREY)

    # ── LANGUAGES ────────────────────────────────────────────
    heading(doc, 'Languages')
    p = doc.add_paragraph()
    sp(p, before=4, after=0)
    ar(p, 'English – Professional  |  Bengali – Native  |  Hindi – Fluent', size=10, clr=BODY)

    doc.save(output_path)
    print(f'Saved: {output_path}')


# ── Run both versions ─────────────────────────────────────────
BASE = 'C:/xampp/htdocs/resume/'
build(BASE + 'Debabrata_Mandal_DACH_SAP_QA_Manager_ATS.docx',   with_photo=False)
build(BASE + 'Debabrata_Mandal_DACH_SAP_QA_Manager_Photo.docx', with_photo=True)
