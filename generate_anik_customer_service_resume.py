"""Anik Raj Sarkar - Customer Service / Remote Support Associate resume.

Tailored to a work-from-home customer service JD (rotational shifts, phone/chat/
email, case management, ticketing tools, store-operations exposure preferred).

Source of truth: AnikRajSarkar.pdf. No experience is invented - the customer
service angle is built by reframing real training, project and language assets.
"""

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
BODY      = RGBColor(0x22, 0x22, 0x22)

PHOTO = 'C:/xampp/htdocs/resume/anik.png'

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
    ar(p, u'\u2022  ', size=10, bold=True, clr=LT_NAVY)
    ar(p, text, size=10, clr=BODY)


def role(doc, title, dates, before=8):
    p = doc.add_paragraph()
    sp(p, before=before, after=0)
    set_right_tab(p, 6.3)
    ar(p, title, size=10.5, bold=True, clr=NAVY)
    ar(p, '\t', size=10, clr=BODY)
    ar(p, dates, size=9.5, clr=MID_GREY)


def org(doc, text, before=1, after=2):
    p = doc.add_paragraph()
    sp(p, before=before, after=after)
    ar(p, text, size=9.5, italic=True, clr=LT_NAVY)


def labelled(doc, label, content, before=2, after=2):
    p = doc.add_paragraph()
    sp(p, before=before, after=after)
    ar(p, label + '  ', size=10, bold=True, clr=NAVY)
    ar(p, content, size=10, clr=BODY)


# ── Content ───────────────────────────────────────────────────
NAME     = 'ANIK RAJ SARKAR'
HEADLINE = ('Customer Service Associate  \u00b7  Remote Support  \u00b7  '
            'Case Management  \u00b7  Technical Troubleshooting')
CONTACT  = ('Kolkata, India  \u00b7  +91 98748 18998  \u00b7  ankraz009@gmail.com  '
            '\u00b7  Open to rotational shifts & work-from-home')

SUMMARY = (
    'Customer-focused Computer Science graduate (B.Tech CSE, Diploma CSE) seeking an entry-level '
    'Customer Service Associate role supporting a global enterprise across phone, chat and email '
    'channels. Trained in incident identification, analysis and documentation through a live '
    'cybersecurity operations internship, where accurate case logging and process discipline are '
    'daily requirements. Combines hands-on comfort with computer applications, databases and '
    'digital communication platforms with the patience to simplify technical detail into clear, '
    'empathetic explanations for non-technical users. Trilingual in English, Hindi and Bengali, '
    'enabling confident service across diverse customer and cultural contexts. Available for '
    'rotational shifts and remote working, and motivated to build a long-term career in customer '
    'experience within a performance-metric-driven environment.'
)

COMPETENCIES = [
    ('Customer Service & Communication:',
     'Customer Centricity  \u00b7  Inbound & Outbound Phone, Chat and Email Support  \u00b7  '
     'Active Listening  \u00b7  Empathetic Explanation of Products, Services & Policies  \u00b7  '
     'De-escalation & Composure  \u00b7  Clear Verbal & Written Communication  \u00b7  '
     'Customer Satisfaction Focus'),
    ('Case & Process Management:',
     'Case Handling from First Contact to Closure  \u00b7  Interaction Documentation & Summaries  \u00b7  '
     'Data Accuracy & Record Keeping  \u00b7  Escalation to Internal Support Teams  \u00b7  '
     'Knowledge Base Research  \u00b7  Root-Cause Problem Solving  \u00b7  '
     'Repeat-Contact Reduction  \u00b7  SLA & Performance Metric Awareness'),
    ('Digital Tools & Technical Skills:',
     'Ticketing & Case-Tracking Tools  \u00b7  CRM / Knowledge Base Navigation  \u00b7  '
     'MS Office & Google Workspace  \u00b7  Digital Communication Platforms (Email, Chat, Video)  \u00b7  '
     'Python (data handling, scripting)  \u00b7  MySQL & SQLite (queries, records, data integrity)  \u00b7  '
     'HTML, CSS, JavaScript (basic)  \u00b7  Basic Hardware/Software Troubleshooting'),
    ('Professional Competencies:',
     'Build Trust, Collaboration & Transparency  \u00b7  Drive Results with Accountability  \u00b7  '
     'Drive Change & Innovation  \u00b7  Promote Engagement  \u00b7  Strategic & Enterprise Mindset  \u00b7  '
     'Data Privacy, Security & Compliance Discipline  \u00b7  Continuous Learning  \u00b7  '
     'Remote Self-Management & Productivity'),
]

EXPERIENCE = [
    dict(
        title='Trainee Intern  \u2014  Cybersecurity Operations',
        org='Jain Cyber Solutions  |  Kolkata, India',
        dates='Jul 2026 \u2013 Present',
        bullets=[
            'Observe and analyse real-world security events within a live operations environment, '
            'applying structured procedures to identify what an incident is, who it affects and '
            'what action it requires \u2014 the same intake-to-resolution discipline used in customer '
            'case management.',
            'Log incident details, investigation steps and outcomes accurately into tracking systems, '
            'maintaining data quality standards that allow supervisors to review, analyse and improve '
            'recurring issue handling.',
            'Follow defined escalation paths for issues beyond first-line scope, handing off to senior '
            'analysts with complete context and following through until the case reaches closure.',
            'Research known issue patterns using internal documentation and knowledge references to '
            'resolve repeat scenarios faster and reduce duplicated effort across the team.',
            'Operate strictly within confidentiality, privacy and security guidelines when handling '
            'sensitive client information, protecting both customer data and company reputation.',
            'Translate technical findings into plain-language notes and verbal updates for '
            'non-specialist stakeholders, building the habit of explaining complex detail simply and calmly.',
            'Participate in ongoing coaching, tooling walkthroughs and knowledge updates, adopting new '
            'procedures and system changes as they are rolled out.',
        ],
    ),
    dict(
        title='Full Stack Web Development with AI  \u2014  Professional Training',
        org='Internshala Training  |  Certificate No. dilxwnhnrdo',
        dates='May 2026 \u2013 Jul 2026',
        bullets=[
            'Completed an intensive practical programme covering HTML, CSS, Bootstrap, JavaScript, DOM, '
            'PHP, React and DBMS, building day-to-day fluency across the browser-based applications and '
            'web tools that underpin modern remote service desks.',
            'Delivered a final project end to end under defined milestones and review checkpoints, '
            'demonstrating ability to manage a workload independently against a schedule.',
            'Applied AI-assisted development workflows to research solutions and troubleshoot faults '
            'quickly \u2014 directly transferable to knowledge-base-driven customer issue resolution.',
        ],
    ),
    dict(
        title='Programming in Python with AI  \u2014  Professional Training',
        org='Internshala Training  |  Certificate No. pc1y13vfw1_',
        dates='May 2026 \u2013 Jun 2026',
        bullets=[
            'Built Python fundamentals, Object-Oriented Programming and SQLite database connectivity, '
            'strengthening comfort with structured data entry, retrieval and record accuracy.',
            'Developed a PyQt desktop GUI application, gaining an applied understanding of how end users '
            'experience software and where they typically get stuck \u2014 useful context when guiding '
            'customers through digital tools.',
            'Practised systematic debugging: reproduce the problem, isolate the cause, apply a fix, '
            'confirm the result \u2014 a root-cause approach applied to customer-reported issues.',
        ],
    ),
]

PROJECTS = [
    dict(
        title='Budget Master \u2014 Budget Management Web Application',
        org='College Major Group Project  |  SQL, JSON, PHP',
        dates='Feb 2023 \u2013 May 2023',
        bullets=[
            'Built a transaction-handling application covering expense tracking, income management and '
            'reporting \u2014 practical exposure to order/transaction records, balance queries and '
            'statement-style reporting of the kind raised in retail and point-of-service enquiries.',
            'Designed the database and validation rules to keep financial records accurate and '
            'consistent, reinforcing a high standard for data entry and customer record integrity.',
            'Collaborated within a team following structured development practices, dividing ownership, '
            'handing off work and resolving blockers together to deliver on schedule.',
        ],
    ),
    dict(
        title='Python Learning Internship',
        org='College Summer Internship',
        dates='Apr 2021 \u2013 May 2021',
        bullets=[
            'Learned Python programming fundamentals and practised logic building, scripting and '
            'step-by-step problem solving.',
        ],
    ),
]

EDUCATION = [
    ('B.Tech in Computer Science & Technology',
     'Camellia School of Engineering & Technology, WBSCTE', '2026'),
    ('Diploma in Computer Science & Technology',
     'Regent Institute of Science & Technology, WBSCTE', '2023'),
]

CERTIFICATIONS = [
    'Full Stack Web Development with AI \u2013 Internshala Training (Jul 2026)  |  Cert. No. dilxwnhnrdo',
    'Programming in Python with AI \u2013 Internshala Training (Jun 2026)  |  Cert. No. pc1y13vfw1_',
    'Node.js and MongoDB Training (80 Hours) \u2013 Topstack (Apr 2023)',
    'Certificate in Information Technology Application \u2013 Govt. of West Bengal (2015)',
    'Refresher English \u2013 British Institute (2015)',
    'Certified Customer Service Professional (CCSP) \u2013 willing to complete as part of role onboarding',
]


# ── Build Document ────────────────────────────────────────────
def build(output_path, with_photo=False):
    doc = Document()
    doc.styles['Normal'].font.name = 'Calibri'
    doc.styles['Normal'].font.size = Pt(10)
    doc.styles['Normal'].paragraph_format.space_before = Pt(0)
    doc.styles['Normal'].paragraph_format.space_after  = Pt(0)

    sec = doc.sections[0]
    sec.page_width    = Cm(21)
    sec.page_height   = Cm(29.7)
    sec.top_margin    = Cm(1.6)
    sec.bottom_margin = Cm(1.6)
    sec.left_margin   = Cm(2.0)
    sec.right_margin  = Cm(2.0)

    # ── HEADER ───────────────────────────────────────────────
    if with_photo:
        tbl = doc.add_table(rows=1, cols=2)
        tbl.allow_autofit = False
        tbl.columns[0].width = Cm(13.6)
        tbl.columns[1].width = Cm(3.4)

        tblBorders = OxmlElement('w:tblBorders')
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            el = OxmlElement('w:' + side)
            el.set(qn('w:val'), 'none')
            tblBorders.append(el)
        tbl._tbl.tblPr.append(tblBorders)

        left_cell  = tbl.rows[0].cells[0]
        right_cell = tbl.rows[0].cells[1]
        right_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        p_name = left_cell.paragraphs[0]
        sp(p_name, before=0, after=1)
        ar(p_name, NAME, size=24, bold=True, clr=NAVY)

        p_cred = left_cell.add_paragraph()
        sp(p_cred, before=0, after=4)
        add_bottom_border(p_cred, sz=6)
        ar(p_cred, HEADLINE, size=10, clr=LT_NAVY)

        p_ct = left_cell.add_paragraph()
        sp(p_ct, before=4, after=0)
        ar(p_ct, 'Kolkata, India  |  +91 98748 18998  |  ankraz009@gmail.com', size=9, clr=MID_GREY)
        p_ct2 = left_cell.add_paragraph()
        sp(p_ct2, before=1, after=0)
        ar(p_ct2, 'B.Tech (CSE), Diploma (CSE)  |  Open to rotational shifts & work-from-home',
           size=9, clr=MID_GREY)

        p_photo = right_cell.paragraphs[0]
        p_photo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_photo.add_run().add_picture(PHOTO, width=Cm(3.0))

    else:
        p = doc.add_paragraph()
        sp(p, before=0, after=1)
        ar(p, NAME, size=26, bold=True, clr=NAVY)

        p = doc.add_paragraph()
        sp(p, before=0, after=4)
        add_bottom_border(p, sz=6)
        ar(p, HEADLINE, size=10.5, clr=LT_NAVY)

        p = doc.add_paragraph()
        sp(p, before=4, after=0)
        ar(p, CONTACT, size=9, clr=MID_GREY)

        p = doc.add_paragraph()
        sp(p, before=1, after=6)
        ar(p, 'B.Tech (CSE), Diploma (CSE)', size=9, clr=MID_GREY)

    # ── PROFESSIONAL SUMMARY ─────────────────────────────────
    heading(doc, 'Professional Summary', before=8 if with_photo else 0)
    p = doc.add_paragraph()
    sp(p, before=4, after=0)
    ar(p, SUMMARY, size=10, clr=BODY)

    # ── CORE COMPETENCIES ────────────────────────────────────
    heading(doc, 'Core Competencies & Service Skills')
    for label, content in COMPETENCIES:
        labelled(doc, label, content)

    # ── EXPERIENCE & PROFESSIONAL TRAINING ───────────────────
    heading(doc, 'Experience & Professional Training')
    for i, item in enumerate(EXPERIENCE):
        role(doc, item['title'], item['dates'], before=6 if i == 0 else 8)
        org(doc, item['org'])
        for b in item['bullets']:
            blt(doc, b)

    # ── PROJECTS ─────────────────────────────────────────────
    heading(doc, 'Projects')
    for i, item in enumerate(PROJECTS):
        role(doc, item['title'], item['dates'], before=6 if i == 0 else 8)
        org(doc, item['org'])
        for b in item['bullets']:
            blt(doc, b)

    # ── EDUCATION ────────────────────────────────────────────
    heading(doc, 'Education')
    for degree, institute, year in EDUCATION:
        p = doc.add_paragraph()
        sp(p, before=4, after=0)
        set_right_tab(p, 6.5)
        ar(p, degree, size=10, bold=True, clr=BODY)
        ar(p, '\t', size=10, clr=BODY)
        ar(p, year, size=9.5, clr=MID_GREY)
        p2 = doc.add_paragraph()
        sp(p2, before=0, after=2)
        ar(p2, institute, size=9, italic=True, clr=MID_GREY)

    # ── CERTIFICATIONS ───────────────────────────────────────
    heading(doc, 'Certifications & Training')
    for c in CERTIFICATIONS:
        blt(doc, c)

    # ── ADDITIONAL INFORMATION ───────────────────────────────
    heading(doc, 'Additional Information')
    labelled(doc, 'Languages:',
             'English (Fluent \u2013 verbal & written)  \u00b7  Hindi (Second native)  \u00b7  '
             'Bengali (Native)  \u2014 able to support customers across diverse linguistic '
             'and cultural contexts',
             before=4)
    labelled(doc, 'Availability:',
             'Fully available for rotational and night shifts, weekends and continuous-service '
             'rosters  \u00b7  Comfortable working from home  \u00b7  Immediate joiner',
             before=0)
    labelled(doc, 'Strengths:',
             'Calm under pressure  \u00b7  Detail-oriented documentation  \u00b7  Quick to learn '
             'new systems and procedures  \u00b7  Comfortable being measured on quality, '
             'resolution and satisfaction targets',
             before=0)
    labelled(doc, 'Location:', 'Kolkata, India', before=0)

    doc.save(output_path)
    print('Saved: ' + output_path)


BASE = 'C:/xampp/htdocs/resume/'
build(BASE + 'Anik_Raj_Sarkar_Customer_Service_ATS.docx',        with_photo=False)
build(BASE + 'Anik_Raj_Sarkar_Customer_Service_With_Photo.docx', with_photo=True)
print('\nDone - both versions created using anik.png')
