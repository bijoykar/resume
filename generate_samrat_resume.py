"""Samrat Chakraborty - Engineering Software Development Manager resume (ATS).

Source of truth: SamratChakraborty.docx (header contact line, profile summary,
employment table, technological skills, certifications, education, 6 project
tables + TCS training, awards, extra-curricular, personal information).
Nothing is invented - the content is reorganised into a single-column,
ATS-friendly layout, the mojibake characters from the original are repaired,
wording is tightened, and Passport / PAN numbers are intentionally left out.

Outputs:
  Samrat_Chakraborty_Software_Dev_Manager_ATS.docx         (plain, ATS-safe)
  Samrat_Chakraborty_Software_Dev_Manager_With_Photo.docx  (photo in header)
"""

from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# -- Colors ----------------------------------------------------
NAVY     = RGBColor(0x1B, 0x2A, 0x4A)
LT_NAVY  = RGBColor(0x34, 0x4E, 0x8A)
MID_GREY = RGBColor(0x6B, 0x6B, 0x6B)
BODY     = RGBColor(0x22, 0x22, 0x22)

BASE  = 'C:/xampp/htdocs/resume/'
PHOTO = BASE + 'samrat.png'


# -- Helpers ---------------------------------------------------

def sp(para, before=0, after=0):
    pf = para.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after  = Pt(after)


def add_bottom_border(para, color='1B2A4A', sz=8):
    pPr = para._p.get_or_add_pPr()
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


def no_borders(tbl):
    tblBorders = OxmlElement('w:tblBorders')
    for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        el = OxmlElement('w:' + side)
        el.set(qn('w:val'), 'none')
        tblBorders.append(el)
    tbl._tbl.tblPr.append(tblBorders)


def ar(para, text, size=10, bold=False, italic=False, clr=None):
    run = para.add_run(text)
    run.font.name      = 'Calibri'
    run.font.size      = Pt(size)
    run.font.bold      = bold
    run.font.italic    = italic
    run.font.color.rgb = BODY if clr is None else clr
    return run


def heading(doc, text, before=10, after=4):
    p = doc.add_paragraph()
    sp(p, before=before, after=after)
    ar(p, text.upper(), size=10, bold=True, clr=NAVY)
    add_bottom_border(p)
    return p


def blt(doc, text, before=1, after=1, size=9.5):
    p = doc.add_paragraph()
    sp(p, before=before, after=after)
    p.paragraph_format.left_indent       = Inches(0.22)
    p.paragraph_format.first_line_indent = Inches(-0.17)
    ar(p, u'\u2022  ', size=size, bold=True, clr=LT_NAVY)
    ar(p, text, size=size, clr=BODY)
    return p


def role(doc, title, dates, before=8):
    p = doc.add_paragraph()
    sp(p, before=before, after=0)
    set_right_tab(p, 6.6)
    ar(p, title, size=10.5, bold=True, clr=NAVY)
    ar(p, '\t')
    ar(p, dates, size=9.5, clr=MID_GREY)


def org(doc, text, before=1, after=2):
    p = doc.add_paragraph()
    sp(p, before=before, after=after)
    ar(p, text, size=9.5, italic=True, clr=LT_NAVY)


def labelled(doc, label, content, before=2, after=2, size=9.5):
    p = doc.add_paragraph()
    sp(p, before=before, after=after)
    ar(p, label + '  ', size=size, bold=True, clr=NAVY)
    ar(p, content, size=size, clr=BODY)
    return p


# -- Content ---------------------------------------------------
NAME     = 'SAMRAT CHAKRABORTY'
HEADLINE = ('Engineering Software Development Manager  \u00b7  Tech Lead  \u00b7  '
            'Solution Designer  \u00b7  Scrum Master  \u00b7  Java\u00a0/\u00a0Spring\u00a0Boot  \u00b7  '
            'Telecom\u00a0OSS/BSS')
CONTACT1 = '+91-8017499587  |  meetsamrat08@gmail.com'
CONTACT2 = ('linkedin.com/in/samrat-chakraborty-41b14446  |  AWS Certified Cloud Practitioner  |  '
            'Professional Scrum Master')

SUMMARY = (
    'Engineering Software Development Manager, senior solution designer, tech lead and Scrum '
    'Master with 15 years of experience across enterprise application development, technology '
    'consulting, project management and telecom OSS/BSS systems. Deep hands-on expertise in Core '
    'Java, Spring Boot microservices, REST APIs and AWS, with a strong track record designing order '
    'management, orchestration, provisioning and service-assurance workflows for leading telcos in '
    'the Nordics, Denmark and the USA. Combines business acumen with consulting experience to '
    'drive digital transformation of legacy systems, from landscape assessment and roadmap through '
    'to implementation, and leads onshore-offshore teams across the full software development '
    'lifecycle.'
)

HIGHLIGHTS = [
    'Drove digital transformation and reinvention programmes, covering digital landscape assessment, '
    'roadmap and implementation, for legacy systems at the largest telco clients in the Nordic '
    'region and a leading telco in the USA.',
    'Managed end-to-end Software Development Lifecycle across projects, enterprise solutions and '
    'system design.',
    'Led a team of 10+ direct developers as Scrum Master for a leading Danish telecom client.',
    'Spent 1.5 years onsite in Denmark as Project Lead, managing teams of 8+ developers, analysts '
    'and designers in an onshore-offshore delivery model.',
    'Brings in-depth development, design and business-analysis experience across Telecom OSS/BSS '
    'modules, including provisioning, fulfilment and order management.',
    'Established new code repository procedures that protected the customer\u2019s intellectual '
    'property.',
    'Supported implementation of SAP Business ByDesign, a global ERP product, for SAP customers '
    'worldwide.',
    'Practising AI adopter who uses tools such as Claude and Microsoft Copilot in day-to-day '
    'engineering work.',
    'Authored articles and knowledge documents for internal and external audiences; the team\u2019s '
    'go-to person and a mentor to developers and team members.',
]

CORE_COMPETENCIES = (
    'Software Development Management  \u00b7  Technical Leadership  \u00b7  Solution Design & '
    'Architecture  \u00b7  Microservices  \u00b7  Digital Transformation  \u00b7  Legacy '
    'Modernisation  \u00b7  Cloud Migration (AWS)  \u00b7  Agile / Scrum  \u00b7  Stakeholder '
    'Management  \u00b7  Requirements Analysis  \u00b7  Onshore-Offshore Delivery  \u00b7  '
    'Telecom Order Management & Orchestration  \u00b7  Service Provisioning & Assurance  \u00b7  '
    'Mentoring'
)

SKILLS = [
    ('Languages & Frameworks:',
     'Core Java (Java 8 / 17)  \u00b7  J2EE  \u00b7  Spring Boot  \u00b7  Spring Cloud (Netflix '
     'OSS)  \u00b7  Hibernate  \u00b7  REST APIs  \u00b7  Microservices  \u00b7  BPMN  \u00b7  '
     'EJB  \u00b7  Struts  \u00b7  SQL  \u00b7  JSON  \u00b7  XML'),
    ('Telecom Products:',
     'Sigma / Hansen Technologies OSS COTS products: Catalog, Service Order Management (SOM), '
     'SMP; order management and catalog mapping'),
    ('Cloud & DevOps:',
     'AWS (EC2, EKS, Athena, PostgreSQL, S3)  \u00b7  Docker  \u00b7  Maven  \u00b7  Git  \u00b7  '
     'CI/CD'),
    ('Databases:', 'Oracle SQL  \u00b7  PostgreSQL  \u00b7  MongoDB'),
    ('Web & Servers:',
     'HTML  \u00b7  CSS  \u00b7  JavaScript  \u00b7  jQuery  \u00b7  WebLogic  \u00b7  Apache '
     'Tomcat'),
    ('Tools:',
     'Eclipse  \u00b7  STS  \u00b7  IntelliJ IDEA  \u00b7  SQL Developer & Data Modeler  \u00b7  '
     'Git Bash  \u00b7  PuTTY  \u00b7  WinSCP  \u00b7  MS Office'),
    ('Methodologies & OS:',
     'Agile / Scrum  \u00b7  TDD  \u00b7  CI/CD  \u00b7  Windows  \u00b7  Linux  \u00b7  macOS'),
]

EMPLOYERS = [
    dict(
        employer='CAPGEMINI',
        title='Engineering Software Development Manager',
        dates='Sep 2022 \u2013 Present',
        projects=[
            dict(
                name='Telecom Proactive Assurance Platform \u2014 Development, Enhancement & '
                     'Modernisation',
                client='Leading Telco in the USA  |  Sep 2022 \u2013 Present  |  Team: 15+',
                role='Tech Lead, Solution Designer, Senior Developer & Quality Sign-off Engineer',
                bullets=[
                    'Work on a proactive assurance application that detects inconsistencies and '
                    'automatically syncs a telecom service with the service inventory, the network '
                    'and other OSS modules.',
                    'Analyse customer requirements, break them into user stories and deliver them '
                    'in order of business criticality.',
                    'Design and develop new features and scale existing functionality to give the '
                    'client\u2019s end customers more accurate service parameters.',
                    'Led the technology upgrade and migration of the platform to AWS, including '
                    'the required services (EKS, PostgreSQL).',
                    'Own unit testing, deployment and other service lifecycle management (SLM) '
                    'activities, as well as quality sign-off for releases.',
                    'Provide development support and data-quality verification; define and '
                    'document processes for users.',
                ],
                stack='Core Java 8/17, Spring Boot microservices, Spring Cloud (Netflix OSS), '
                      'AWS (EKS, PostgreSQL), Hibernate, Maven, REST APIs, Docker, Git, SQL, JSON',
            ),
        ],
    ),
    dict(
        employer='TATA CONSULTANCY SERVICES (TCS)',
        title='Assistant Consultant',
        dates='Mar 2011 \u2013 Aug 2022',
        projects=[
            dict(
                name='Telecom Order Management, Orchestration & OSS Integration',
                client='Leading Telco in the Nordic Region  |  Jun 2020 \u2013 Aug 2022  |  '
                       'Team: 10+',
                role='Solution Designer & Senior Developer',
                bullets=[
                    'Worked on the Order Management and Order Orchestration applications that '
                    'enable telecom services such as internet, TV and VoIP on the client\u2019s '
                    'network.',
                    'Designed and developed workflows on Sigma (Hansen) COTS products and '
                    'integrated them with external and third-party applications via REST APIs to '
                    'provision services.',
                    'Analysed customer requirements, broke them into user stories and delivered '
                    'them in order of business criticality.',
                    'Reviewed order-fallout scenarios regularly with end users to make the systems '
                    'more efficient.',
                    'Owned unit testing, deployment, SLM activities and technology upgrades; '
                    'provided development support, verified data quality and documented processes '
                    'for users.',
                ],
                stack='Core Java 8, Spring Boot microservices, Spring Cloud (Netflix OSS), Hansen '
                      'COTS products, Hibernate, MongoDB, Maven, REST APIs, Docker, Git, SQL, JSON',
            ),
            dict(
                name='Digital Transformation of a Legacy Network Resource Management System',
                client='Leading Telco in Denmark  |  Sep 2018 \u2013 May 2020  |  Team: 8+',
                role='Application / DB Designer, Scrum Master, Functional Analyst & Senior '
                     'Developer',
                bullets=[
                    'Migrated a network resource management system (VLANs, DHCP servers, etc.) '
                    'from Java EJBs to Spring Boot, with a simpler data model, a better UI and '
                    'stronger security.',
                    'As application owner, modelled the database structure and helped the client '
                    'and other stakeholders shape the design after analysing the existing '
                    'functionality.',
                    'As Scrum Master, led the team in breaking complex requirements into user '
                    'stories, tasks and sub-tasks, and kept incremental delivery on a regular '
                    'cadence.',
                    'Responsible for coding, unit and integration testing, DevOps operations, '
                    'development support and end-user documentation.',
                ],
                stack='Core Java 8+, Spring Boot, Hibernate, Maven, REST APIs, Apache Tomcat, Git, '
                      'SQL, JSON, XML',
            ),
            dict(
                name='Telecom Service Provisioning Applications \u2014 Development & Maintenance',
                client='Leading Telco in Denmark  |  Jan 2014 \u2013 Aug 2018  |  Team: 8+  |  '
                       'Includes 1.5 years onsite in Denmark',
                role='Project Lead & Developer',
                bullets=[
                    'Worked on logical capacity management and provisioning applications in the '
                    'OSS layer that interact with the network and other external systems.',
                    'Analysed customer requirements and was responsible for the design, coding, '
                    'testing, deployment and SLM of features.',
                    'Led the team in delivering business-critical functionality with significant '
                    'revenue impact.',
                ],
                stack='Core Java, EJB, WebLogic Server, Git, SQL',
            ),
            dict(
                name='SAP Business ByDesign Product',
                client='SAP AG  |  May 2012 \u2013 Dec 2013',
                role='Functional Consultant',
                bullets=[
                    'Performed functional analysis of business-integration products (CRM, SCM, '
                    'SRM and HCM) and provided solutions to customers.',
                    'Worked with SAP customers worldwide to support their implementation of the '
                    'product.',
                    'Reported status to SAP Level 2 Product Management.',
                ],
                stack='SAP PI, SAP business modules, SAP tools, Java, Linux, cloud concepts (SaaS, '
                      'PaaS, IaaS)',
            ),
            dict(
                name='Incident Management & Portal Management',
                client='One of the Largest Banks in the USA  |  Sep 2011 \u2013 Apr 2012',
                role='Java / J2EE Developer',
                bullets=[
                    'Developed an incident management portal for the bank\u2019s Priority 1 and '
                    'Priority 2 issues to support continuous customer relationships; contributed '
                    'to design, testing and stabilisation.',
                ],
                stack='Java, J2EE (Servlets, Struts, JSP), JavaScript, jQuery, HTML, CSS, XML, SQL',
            ),
            dict(
                name='TCS Initial Learning Program',
                client='TCS Internal  |  Mar 2011 \u2013 Jul 2011',
                role='Trainee',
                bullets=[
                    'Trained in Java/J2EE (Struts), HTML, CSS, Linux, SQL, soft skills and '
                    'professional etiquette.',
                ],
                stack=None,
            ),
        ],
    ),
]

CERTIFICATIONS = [
    'AWS Certified Cloud Practitioner \u2014 Amazon Web Services, 2022 (renewed 2025)',
    'Microsoft Certified: Azure Fundamentals (AZ-900) \u2014 Microsoft, 2021',
    'Professional Scrum Master (PSM) \u2014 Scrum.org, 2018',
    'Sigma (Hansen Technologies) OSS COTS Products Training (Catalog, SOM, SMP) \u2014 Sigma '
    'Systems, 2018',
    'Oracle Database 11g: SQL Fundamentals \u2014 Oracle, 2017',
    'Oracle Certified Java Professional \u2014 Oracle, 2014',
]

EDUCATION = [
    ('Bachelor of Technology (B.Tech)', 'Electronics & Communication Engineering',
     'West Bengal University of Technology', '7.78 DGPA'),
    ('Class XII (H.S.C.)', 'Science', '', '86.1%'),
    ('Class X (S.S.C.)', 'General', '', '81.1%'),
]

AWARDS = [
    'Received On the Spot awards for delivering ahead of estimated timelines with quality.',
    'Received 3 Service Commitment awards.',
    'Earned a high Customer Satisfaction Index; recognised and awarded by customers for the '
    'quality of work delivered.',
]

ACTIVITIES = [
    'Active member of TCS Maitree, TCS\u2019s social initiative: took part in events such as '
    '\u201cSafety First\u201d, which promotes employees\u2019 physical and mental safety, and '
    'supported underprivileged children\u2019s education by distributing books and supplies.',
    'Organised and took part in cultural events within the project.',
]

INTERESTS = 'Sports and outdoor events  \u00b7  Singing  \u00b7  Teaching and mentoring'

PERSONAL = [
    ('Date of Birth', '28 May 1988'),
    ('Total Experience', '15 years'),
]


# -- Build -----------------------------------------------------
def build(output_path, with_photo=False):
    doc = Document()
    doc.styles['Normal'].font.name = 'Calibri'
    doc.styles['Normal'].font.size = Pt(10)
    doc.styles['Normal'].paragraph_format.space_before = Pt(0)
    doc.styles['Normal'].paragraph_format.space_after  = Pt(0)

    sec = doc.sections[0]
    sec.page_width    = Cm(21)
    sec.page_height   = Cm(29.7)
    sec.top_margin    = Cm(1.4)
    sec.bottom_margin = Cm(1.4)
    sec.left_margin   = Cm(1.8)
    sec.right_margin  = Cm(1.8)

    # -- HEADER ------------------------------------------------
    if with_photo:
        tbl = doc.add_table(rows=1, cols=2)
        tbl.allow_autofit    = False
        tbl.columns[0].width = Cm(14.4)
        tbl.columns[1].width = Cm(3.0)
        no_borders(tbl)

        left, right = tbl.rows[0].cells
        left.width, right.width = Cm(14.4), Cm(3.0)
        left.vertical_alignment  = WD_ALIGN_VERTICAL.CENTER
        right.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        host = left
        p = left.paragraphs[0]
    else:
        host = doc
        p = doc.add_paragraph()

    sp(p, before=0, after=1)
    ar(p, NAME, size=23 if with_photo else 24, bold=True, clr=NAVY)

    p = host.add_paragraph()
    sp(p, before=0, after=4)
    add_bottom_border(p, sz=6)
    ar(p, HEADLINE, size=8.5 if with_photo else 9.5, clr=LT_NAVY)

    p = host.add_paragraph()
    sp(p, before=4, after=0)
    ar(p, CONTACT1, size=8.5 if with_photo else 9, clr=MID_GREY)
    p = host.add_paragraph()
    sp(p, before=1, after=2)
    ar(p, CONTACT2, size=8.5 if with_photo else 9, clr=MID_GREY)

    if with_photo:
        p = right.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        sp(p, before=0, after=0)
        p.add_run().add_picture(PHOTO, width=Cm(2.8))

    # -- PROFESSIONAL SUMMARY ----------------------------------
    heading(doc, 'Professional Summary', before=6)
    p = doc.add_paragraph()
    sp(p, before=4, after=2)
    ar(p, SUMMARY, size=9.5, clr=BODY)
    for h in HIGHLIGHTS:
        blt(doc, h)

    # -- CORE COMPETENCIES -------------------------------------
    heading(doc, 'Core Competencies')
    p = doc.add_paragraph()
    sp(p, before=3, after=0)
    ar(p, CORE_COMPETENCIES, size=9.5, clr=BODY)

    # -- TECHNICAL SKILLS --------------------------------------
    heading(doc, 'Technical Skills')
    for label, content in SKILLS:
        labelled(doc, label, content)

    # -- PROFESSIONAL EXPERIENCE -------------------------------
    heading(doc, 'Professional Experience')
    for ei, emp in enumerate(EMPLOYERS):
        role(doc, emp['employer'] + '  \u2014  ' + emp['title'], emp['dates'],
             before=6 if ei == 0 else 12)
        for pi, proj in enumerate(emp['projects']):
            p = doc.add_paragraph()
            sp(p, before=6 if pi else 5, after=0)
            ar(p, proj['name'], size=10, bold=True, clr=LT_NAVY)
            org(doc, proj['client'], before=1, after=1)

            p = doc.add_paragraph()
            sp(p, before=0, after=2)
            p.paragraph_format.left_indent = Inches(0.05)
            ar(p, 'Role:  ', size=9.5, bold=True, clr=NAVY)
            ar(p, proj['role'], size=9.5, clr=BODY)

            for b in proj['bullets']:
                blt(doc, b)

            if proj['stack']:
                p = doc.add_paragraph()
                sp(p, before=3, after=0)
                p.paragraph_format.left_indent = Inches(0.05)
                ar(p, 'Tech Stack:  ', size=9, bold=True, clr=NAVY)
                ar(p, proj['stack'], size=9, italic=True, clr=MID_GREY)

    # -- CERTIFICATIONS ----------------------------------------
    heading(doc, 'Certifications')
    for i, c in enumerate(CERTIFICATIONS):
        blt(doc, c, before=3 if i == 0 else 1)

    # -- EDUCATION ---------------------------------------------
    heading(doc, 'Education')
    for i, (degree, spec, board, grade) in enumerate(EDUCATION):
        p = doc.add_paragraph()
        sp(p, before=4 if i == 0 else 2, after=0)
        set_right_tab(p, 6.6)
        ar(p, degree, size=10, bold=True, clr=NAVY)
        ar(p, '  \u2014  ' + spec, size=9.5, clr=BODY)
        ar(p, '\t')
        ar(p, grade, size=9.5, clr=MID_GREY)
        if board:
            org(doc, board, before=0, after=0)

    # -- AWARDS ------------------------------------------------
    heading(doc, 'Awards & Recognition')
    for i, a in enumerate(AWARDS):
        blt(doc, a, before=3 if i == 0 else 1)

    # -- EXTRA-CURRICULAR --------------------------------------
    heading(doc, 'Extra-Curricular Activities')
    for i, a in enumerate(ACTIVITIES):
        blt(doc, a, before=3 if i == 0 else 1)

    # -- INTERESTS ---------------------------------------------
    heading(doc, 'Interests')
    p = doc.add_paragraph()
    sp(p, before=3, after=0)
    ar(p, INTERESTS, size=9.5, clr=BODY)

    # -- PERSONAL DETAILS --------------------------------------
    heading(doc, 'Personal Details')
    p = doc.add_paragraph()
    sp(p, before=4, after=0)
    for i, (k, v) in enumerate(PERSONAL):
        if i:
            ar(p, '    |    ', size=9, clr=MID_GREY)
        ar(p, k + ': ', size=9, bold=True, clr=NAVY)
        ar(p, v, size=9, clr=BODY)

    doc.save(output_path)
    print('Saved: ' + output_path)


build(BASE + 'Samrat_Chakraborty_Software_Dev_Manager_ATS.docx',        with_photo=False)
build(BASE + 'Samrat_Chakraborty_Software_Dev_Manager_With_Photo.docx', with_photo=True)
