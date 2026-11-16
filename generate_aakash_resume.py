"""Aakash Agarwal - Senior Java / Backend Developer resume.

Source of truth: aakash.docx (career summary, employment profile, 4 project
tables, technical skills, academics, leadership skills, personal vitae).
Nothing is invented - the content is reorganised into a modern layout, the
mojibake characters from the original are repaired, and the profile photo
(aakash.png) plus the AWS Certified Cloud Practitioner badge extracted from
the original document are placed in the header.

Project date ranges are derived from the durations stated in aakash.docx
(27 / 36 / 40 / 24 months) chained against the employment dates.

Outputs:
  Aakash_Agarwal_Senior_Java_Developer_ATS.docx        (plain, ATS-safe)
  Aakash_Agarwal_Senior_Java_Developer_With_Photo.docx (photo + badge)
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
PHOTO = BASE + 'aakash.png'
BADGE = BASE + 'aws_cloud_practitioner_badge.png'


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
    set_right_tab(p, 6.4)
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
NAME     = 'AAKASH AGARWAL'
HEADLINE = ('Senior Java Backend Developer  \u00b7  Tech Lead  \u00b7  '
            'Microservices & Spring Boot  \u00b7  Cloud / DevOps')
CONTACT1 = ('Kolkata, India  |  +91-89602 50109 / +91-79801 36865  |  '
            'agarwal.aakaash@gmail.com')
CONTACT2 = ('linkedin.com/in/aakash-agarwal-0ab8a8b0  |  '
            'AWS Certified Cloud Practitioner  |  B.Tech (CSE)')

SUMMARY = (
    'Senior Java Developer with 10.5+ years of IT experience building applications on '
    'microservice architecture, backed by a degree in Computer Science Engineering. Two years of '
    'on-site experience working closely with clients and interacting directly with business users. '
    'Experienced across all phases of the application development lifecycle \u2014 development, '
    'testing, maintenance and debugging, along with documentation and deployment of the '
    'application. Skilled at analysing requirements against the current implementation and '
    'building solutions and designs for the application. Currently a Tech Manager at Deloitte, '
    'leading back-end delivery, revamping legacy architectures, mentoring engineers and owning '
    'production-critical systems end to end.'
)

SKILLS = [
    ('Primary Stack:',
     'Java  \u00b7  Spring Boot  \u00b7  Microservices  \u00b7  REST API / SOAP  \u00b7  AWS  '
     '\u00b7  Docker  \u00b7  Linux  \u00b7  SQL  \u00b7  JPA / Hibernate  \u00b7  '
     'JUnit / Mockito'),
    ('Databases:',
     'Oracle DB  \u00b7  PostgreSQL  \u00b7  MariaDB  \u00b7  MySQL  \u00b7  MongoDB  \u00b7  '
     'Couchbase  \u00b7  Liquibase / Flyway'),
    ('Middleware & Tooling:',
     'RabbitMQ / Kafka  \u00b7  Redis  \u00b7  Nginx  \u00b7  Jenkins  \u00b7  SonarQube  \u00b7  '
     'ELK  \u00b7  Postman  \u00b7  Git  \u00b7  Gradle / Maven  \u00b7  Kubernetes  \u00b7  '
     'Jacoco'),
    ('Platforms & Practices:',
     'IntelliJ  \u00b7  Eclipse / STS IDE  \u00b7  Apache Tomcat  \u00b7  DroneCI  \u00b7  '
     'ArgoCD  \u00b7  Helm Charts  \u00b7  ServiceNow / HPSM  \u00b7  JIRA / Rally  \u00b7  '
     'Confluence  \u00b7  Agile (Scrum, Kanban)'),
]

EMPLOYERS = [
    dict(
        employer='DELOITTE TOUCHE TOHMATSU INDIA LLP',
        title='Tech Manager',
        dates='May 2021 \u2013 Present',
        projects=[
            dict(
                name='Connected Worker Safety Platform  \u2014  Senior Back-End Lead',
                client='American Connected Worker Safety Client  |  Role: Tech Lead  |  '
                       '27 Months (Jun 2024 \u2013 Present)',
                summary=(
                    'Built a functional system enabling an Industrial Internet of People via an '
                    'Industrial IoT platform through which data is collected, analysed and used in '
                    'industrial workplaces, incorporating a growing ecosystem of wearable devices '
                    'and integrations combined to deliver solutions across various industries.'),
                bullets=[
                    'Developed and tested new functionalities as per specified requirements, and '
                    'debugged and fixed defects raised on existing functionalities.',
                    'Applied Liquibase integration for existing PostgreSQL DB migrations and '
                    'supported DB data migration activity, fixing data in databases through SQL '
                    'and NoSQL commands as needed.',
                    'Added Caffeine / Spring Cache for data caching and expiry, and identified and '
                    'resolved cache-related issues on the application.',
                    'Implemented FeignClient and Spring RestTemplate for consuming internal '
                    'services\u2019 REST endpoints.',
                    'Implemented WebSockets for publishing time-sensitive data to other '
                    'applications, and Kafka for faster event processing and publishing.',
                    'Created the unit testing setup and added tests to increase code coverage, '
                    'integrated with SonarQube via Jacoco for coverage reports.',
                    'Configured HTTPD and NGINX servers for reverse-proxy and load-balancing to the '
                    'back-end application, and containerised the application to run smoothly on the '
                    'servers as per specifications.',
                    'Identified and fixed issues related to application stabilisation and '
                    'scalability, including transaction-related issues in the JPA / Hibernate '
                    'implementation.',
                    'Automated application boot-ups through scripts to minimise downtime and manual '
                    'intervention after server reboots, and worked extensively with Linux shell '
                    'scripts for running jobs and specific functionality.',
                    'Revamped the application\u2019s existing architecture to address identified '
                    'drawbacks, preventing potential future issues.',
                    'Performed efficient code walk-throughs to understand existing microservice '
                    'functionality and enhance it on request, and conducted code reviews on merge '
                    'requests to verify no existing functionality breaks.',
                    'Performed load testing through JMeter on the APIs to identify thresholds and '
                    'addressed them via auto-scaling.',
                    'Implemented spring-profiling and centralised logging via ELK in compliance '
                    'with GDPR guidelines.',
                    'Worked on Lightspeed Enterprise and Harness setup to create CI/CD pipelines '
                    'along with Ansible scripts on VMs and OpenShift clusters.',
                    'Documented architecture, sequence and block diagrams for existing workflows '
                    'for better understanding of flow and to identify and resolve loopholes; '
                    'documented missing knowledge and created runbooks for CHG implementation in '
                    'PROD and DR environments.',
                    'Estimated effort for work and responsibly delivered on time, provided '
                    'workaround solutions to business-critical problems, and mentored juniors on '
                    'Agile ways of working, best practices and industry standards.',
                    'Contributed to a POC on Prompt Engineering for GitHub Copilot.',
                ],
                stack=[
                    'Java, Spring Boot, Microservices, JPA/Hibernate, REST/SOAP, Kafka, Gradle, '
                    'Git, OAuth, JWT, JUnit/Mockito, Docker, MongoDB, PostgreSQL, SQL, Linux',
                    'AWS EC2, S3, Liquibase, ActiveMQ/MQTT, Caffeine/Spring Cache, Jenkins, '
                    'SonarQube, Feign, Jacoco, ELK, WebSockets, Redis',
                    'Apache Tomcat, Apache HTTPD, NGINX, IntelliJ IDE, Postman, Jenkins, GitHub, '
                    'JIRA, Confluence, Agile \u2013 Scrum, Kanban',
                ],
            ),
            dict(
                name='Personal Finance Management Platform  \u2014  Senior Back-End Developer',
                client='Indian Bank Client  |  Role: Tech Lead  |  '
                       '36 Months (May 2021 \u2013 May 2024)',
                summary=(
                    'Built a functional system for personal finance management enabling customers '
                    'to take care of their investment needs and desires without needing to visit '
                    'the bank physically.'),
                bullets=[
                    'Actively developed and delivered feature requests and bug fixes in '
                    'coordination with other teams, while handling daily BAU operations alongside '
                    'assigned development work.',
                    'Created the sandbox environment and a Disaster Recovery site for the '
                    'production application from scratch.',
                    'Reviewed and refactored code and guided team members on appropriate changes '
                    'in pull requests.',
                    'Established API documentation with Swagger for convenient sharing of '
                    'endpoints and models.',
                    'Created, integrated and tested REST / SOAP APIs and enhanced endpoints when '
                    'needed.',
                    'Built regular expressions and fixes for pattern matching used in report '
                    'generation.',
                    'Renewed the SSL certificate and configured the application to use the renewed '
                    'certificate.',
                    'Analysed and applied the fixes highlighted in the application\u2019s VAPT '
                    'testing.',
                    'Implemented Redis cache on the application while identifying and resolving '
                    'cache issues.',
                    'Built a range of utilities and helper functions to reduce code redundancy.',
                    'Redesigned the application\u2019s legacy architecture around its drawbacks and '
                    'loopholes, preventing further incidents, and anticipated failures before they '
                    'happen by designing around them.',
                    'Developed, scheduled and scripted CRON jobs through expressions.',
                    'Integrated AWS S3 for persisting generated reports via the AWS SDK and fixed '
                    'AWS Security Group configuration for connectivity issues.',
                    'Fixed bad data in the database via NoSQL commands for MongoDB.',
                    'Implemented spring-profiling and logging in compliance with PII guidelines.',
                    'Applied workaround solutions to business-critical problems to reduce impact.',
                    'Created architecture, sequence and flow diagrams for current workflows to '
                    'enhance understanding of functionality and to identify and address gaps.',
                    'Managed and mentored the team including work distribution, conducted Agile '
                    'ceremonies and demonstrated the work delivered through each Sprint.',
                    'Coordinated and provided active support to dependent teams for testing and '
                    'seamless production delivery.',
                    'Created documentation and knowledge articles and provided knowledge '
                    'transitions wherever needed.',
                    'Produced several Requests for Proposal (RFP) for new incoming projects by '
                    'analysing requirement documents and existing functionality.',
                ],
                stack=[
                    'Java, Spring Boot, Microservices, JPA/Hibernate, REST/SOAP, Swagger, '
                    'Gradle/Maven, Git, MVC, Schedulers, OAuth, JWT, JUnit/Mockito, Docker, '
                    'MongoDB, Oracle DB, SQL, Linux, Redis',
                    'Apache Tomcat, Apache HTTPD, NGINX, IntelliJ IDE, Postman, Bitbucket, '
                    'GitLab-CI, GitLab, JIRA, Confluence, Agile \u2013 Scrum, Kanban',
                ],
            ),
        ],
    ),
    dict(
        employer='TATA CONSULTANCY SERVICES LIMITED',
        title='Systems Engineer',
        dates='Feb 2016 \u2013 May 2021',
        projects=[
            dict(
                name='OSS-I Order Fulfilment Integration  \u2014  Back-End Developer',
                client='Danish Telecommunications Client  |  Role: Team Member  |  '
                       '40 Months (Feb 2018 \u2013 May 2021)',
                summary=(
                    'OSS-I is the integration layer between the BSS and OSS domains. The project '
                    'developed a functional integration system for fulfilment, able to orchestrate '
                    'orders from BSS to the network and network-related systems in OSS.'),
                bullets=[
                    'On-site exposure working closely with clients for 2 years, providing many '
                    'on-the-go solutions to business-critical problems that saved time and money '
                    'for the business.',
                    'Performed efficient code walk-throughs and enhancements for microservices as '
                    'per specified requirements.',
                    'Created extendable REST API endpoints and enhanced them when required.',
                    'Integrated with third-party vendor APIs using OAuth, including payload '
                    'analysis.',
                    'Developed and deployed the application as a microservice after successful IFT '
                    'and EFT testing.',
                    'Wrote unit tests with JUnit and Mockito to protect the code from future '
                    'changes.',
                    'Handled configuration for different environments through multiple profiles '
                    'via spring-profiles, and implemented centralised configuration with '
                    'spring-config.',
                    'Added logging and comments with proper messages for better code readability '
                    'through logging libraries, keeping GDPR in consideration.',
                    'Resolved defects identified in testing and debugged errors and failures to '
                    'find the RCA, providing hot fixes to ensure minimum business impact.',
                    'Developed CRON jobs along with scheduling and scripting, and contributed '
                    'Python and Puppet scripts for application and schedule monitoring.',
                    'Performed data migration in databases and fixed bad data via SQL commands '
                    'whenever required.',
                    'Handled batch processing and load testing through JMeter.',
                    'Built logical solutions from the user\u2019s point of view, covering as many '
                    'edge cases as possible.',
                    'Redesigned the legacy model of a business feature around its drawbacks and '
                    'loopholes, preventing further incidents on the application.',
                    'Created sequence and block diagrams (HLD and LLD) for incoming work for '
                    'better understanding of flow and to identify and resolve loopholes.',
                    'Executed weekly releases and deployments with utmost care inside the downtime '
                    'window, with a proper rollback plan.',
                ],
                stack=[
                    'Java, Spring Boot, Microservices, JPA, Hibernate, Maven, Git, MVC, JUnit, '
                    'Mockito, RabbitMQ, Docker, Kubernetes, ELK, Loki-Grafana, JMeter, MongoDB, '
                    'MariaDB, SQL',
                    'STS / IntelliJ IDE, Postman, GitHub, GitLab, JIRA, Agile \u2013 Scrum, '
                    'Kanban, DevOps / GitOps',
                ],
            ),
            dict(
                name='Order Fulfilment & Invoicing System  \u2014  Developer',
                client='American Education Publisher Client  |  Role: Team Member  |  '
                       '24 Months (Feb 2016 \u2013 Jan 2018)',
                summary=(
                    'Built and scaled a system for fulfilling received orders for shipment of '
                    'consignments from inventories to large institutions and universities, retail '
                    'stores and the online market across the globe, consequently generating '
                    'invoices used for payment and revenue purposes.'),
                bullets=[
                    'Handled development, testing and troubleshooting for service requests and the '
                    'enhancements and customisations provided by the customer on a regular basis.',
                    'Analysed the technical and operational feasibility of use cases in the '
                    'requirements to implement business logic as requested.',
                    'Performed data and process migration for jobs and schedules from one '
                    'environment to another.',
                    'Developed and tested new jobs and schedules prior to movement into production.',
                    'Identified the RCA and troubleshot reported bugs, working co-efficiently with '
                    'other teams to rectify and resolve issues.',
                    'Generated several workflow-driven service operations and processes along with '
                    'multiple knowledge articles to be reused by the team.',
                ],
                stack=[
                    'Core Java, JSP, Servlets, JDBC, Collections, Groovy and Python for scripts',
                    'HTML, CSS, JavaScript, JSON/XML, SQL, Shell Scripts',
                    'IDE \u2013 Eclipse, Server \u2013 Apache Tomcat',
                ],
            ),
        ],
    ),
]

CERTIFICATIONS = [
    'AWS Certified Cloud Practitioner \u2013 Foundational (Amazon Web Services)',
]

EDUCATION = [
    ('B. Tech. (Computer Science Engineering)', '2015',
     'Dr. A.P.J. Abdul Kalam Technical University', '65%'),
    ('Class XII', '2011', 'Indian Certificate of Secondary Education (ICSE)', '72%'),
    ('Class X',   '2009', 'Indian Certificate of Secondary Education (ICSE)', '76%'),
]

LEADERSHIP = [
    'Strength to work in a team-oriented, collaborative environment.',
    'Strong interpersonal and client handling skills.',
    'Excellent communication and documentation skills with a strong customer service orientation.',
    'Good aptitude for picking up enough of the business and application logic to effectively and '
    'efficiently tune, or make tuning recommendations.',
    'Highly organised and trained to work with multiple people across different aspects of the '
    'work while managing a larger set of deliverables.',
    'Strong analytical skills and experience identifying and solving performance issues.',
]

PERSONAL_SKILLS = [
    'Self-motivated with proficient and effective communication.',
    'Works independently with minimum guidance.',
    'Ability to multi-task in a fast-paced, deadline-driven environment.',
    'Regards team spirit as an integral part of productivity.',
    'Self-confidence with fast learning capability, and smart work with a risk-taking mentality.',
    'Leadership quality and innovative thinking.',
    'Excellent analytical and programming skills with a willingness to learn and adapt to new '
    'technologies.',
    'Ability to solve real-world problems while gaining insight into the bigger picture.',
]

PERSONAL_VITAE = [
    ('Name', 'Aakash Agarwal'),
    ('Date of Birth', '18-03-1993'),
    ('Gender', 'Male'),
    ('Marital Status', 'Unmarried'),
    ('Languages Known', 'English, Hindi & Bengali'),
    ('Nationality', 'Indian'),
]

DECLARATION = (
    'I hereby declare that the above-mentioned information is correct to the best of my knowledge '
    'and I bear the responsibility for the correctness of the particulars furnished above.'
)


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
        tbl.allow_autofit   = False
        tbl.columns[0].width = Cm(13.9)
        tbl.columns[1].width = Cm(3.5)
        no_borders(tbl)

        left, right = tbl.rows[0].cells
        right.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        p = left.paragraphs[0]
        sp(p, before=0, after=1)
        ar(p, NAME, size=23, bold=True, clr=NAVY)

        p = left.add_paragraph()
        sp(p, before=0, after=4)
        add_bottom_border(p, sz=6)
        ar(p, HEADLINE, size=9.5, clr=LT_NAVY)

        p = left.add_paragraph()
        sp(p, before=4, after=0)
        ar(p, CONTACT1, size=9, clr=MID_GREY)
        p = left.add_paragraph()
        sp(p, before=1, after=0)
        ar(p, CONTACT2, size=9, clr=MID_GREY)

        p = right.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sp(p, before=0, after=2)
        p.add_run().add_picture(PHOTO, width=Cm(3.1))
        p = right.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sp(p, before=0, after=0)
        p.add_run().add_picture(BADGE, width=Cm(1.5))
    else:
        p = doc.add_paragraph()
        sp(p, before=0, after=1)
        ar(p, NAME, size=25, bold=True, clr=NAVY)

        p = doc.add_paragraph()
        sp(p, before=0, after=4)
        add_bottom_border(p, sz=6)
        ar(p, HEADLINE, size=10, clr=LT_NAVY)

        p = doc.add_paragraph()
        sp(p, before=4, after=0)
        ar(p, CONTACT1, size=9, clr=MID_GREY)
        p = doc.add_paragraph()
        sp(p, before=1, after=2)
        ar(p, CONTACT2, size=9, clr=MID_GREY)

    # -- CAREER SUMMARY ----------------------------------------
    heading(doc, 'Career Summary', before=8 if with_photo else 6)
    p = doc.add_paragraph()
    sp(p, before=4, after=0)
    ar(p, SUMMARY, size=9.5, clr=BODY)

    # -- TECHNICAL SKILLS --------------------------------------
    heading(doc, 'Technical Skills')
    for label, content in SKILLS:
        labelled(doc, label, content)

    # -- PROFESSIONAL EXPERIENCE -------------------------------
    heading(doc, 'Professional Experience')
    for ei, emp in enumerate(EMPLOYERS):
        role(doc, emp['employer'] + '  \u2014  ' + emp['title'], emp['dates'],
             before=6 if ei == 0 else 10)
        for pi, proj in enumerate(emp['projects']):
            p = doc.add_paragraph()
            sp(p, before=6 if pi else 5, after=0)
            ar(p, proj['name'], size=10, bold=True, clr=LT_NAVY)
            org(doc, proj['client'], before=1, after=2)

            p = doc.add_paragraph()
            sp(p, before=0, after=2)
            p.paragraph_format.left_indent = Inches(0.05)
            ar(p, 'Project:  ', size=9.5, bold=True, clr=NAVY)
            ar(p, proj['summary'], size=9.5, clr=BODY)

            for b in proj['bullets']:
                blt(doc, b)

            p = doc.add_paragraph()
            sp(p, before=3, after=0)
            p.paragraph_format.left_indent = Inches(0.05)
            ar(p, 'Technical Skills:  ', size=9, bold=True, clr=NAVY)
            ar(p, '  \u00b7  '.join(proj['stack']), size=9, italic=True, clr=MID_GREY)

    # -- CERTIFICATIONS ----------------------------------------
    heading(doc, 'Certifications')
    for c in CERTIFICATIONS:
        blt(doc, c, before=4)

    # -- ACADEMIC PROFILE --------------------------------------
    heading(doc, 'Academic Profile')
    tbl = doc.add_table(rows=1, cols=4)
    tbl.style = 'Table Grid'
    tbl.allow_autofit = False
    widths = [Cm(7.2), Cm(2.2), Cm(6.2), Cm(1.8)]
    for i, w in enumerate(widths):
        tbl.columns[i].width = w
    hdr = ['Course', 'Year of Passing', 'Board / University', 'Marks']
    for i, h in enumerate(hdr):
        cell = tbl.rows[0].cells[i]
        cell.width = widths[i]
        cp = cell.paragraphs[0]
        sp(cp, before=2, after=2)
        ar(cp, h, size=9, bold=True, clr=NAVY)
    for course, year, board, marks in EDUCATION:
        row = tbl.add_row()
        for i, val in enumerate((course, year, board, marks)):
            cell = row.cells[i]
            cell.width = widths[i]
            cp = cell.paragraphs[0]
            sp(cp, before=2, after=2)
            ar(cp, val, size=9, clr=BODY)

    # -- LEADERSHIP SKILLS -------------------------------------
    heading(doc, 'Leadership Skills')
    for item in LEADERSHIP:
        blt(doc, item, before=2 if item is LEADERSHIP[0] else 1)

    # -- PERSONAL SKILLS ---------------------------------------
    heading(doc, 'Personal Skills')
    for item in PERSONAL_SKILLS:
        blt(doc, item, before=2 if item is PERSONAL_SKILLS[0] else 1)

    # -- PERSONAL VITAE ----------------------------------------
    heading(doc, 'Personal Vitae')
    p = doc.add_paragraph()
    sp(p, before=4, after=0)
    for i, (k, v) in enumerate(PERSONAL_VITAE):
        if i:
            ar(p, '    |    ', size=9, clr=MID_GREY)
        ar(p, k + ': ', size=9, bold=True, clr=NAVY)
        ar(p, v, size=9, clr=BODY)

    # -- DECLARATION -------------------------------------------
    heading(doc, 'Declaration')
    p = doc.add_paragraph()
    sp(p, before=4, after=0)
    ar(p, DECLARATION, size=9.5, clr=BODY)

    p = doc.add_paragraph()
    sp(p, before=8, after=0)
    set_right_tab(p, 6.6)
    ar(p, 'Place: Kolkata', size=9.5, clr=BODY)
    ar(p, '\t')
    ar(p, 'Aakash Agarwal', size=9.5, bold=True, clr=NAVY)

    doc.save(output_path)
    print('Saved: ' + output_path)


build(BASE + 'Aakash_Agarwal_Senior_Java_Developer_ATS.docx',        with_photo=False)
build(BASE + 'Aakash_Agarwal_Senior_Java_Developer_With_Photo.docx', with_photo=True)
print('\nDone - both versions created using aakash.png + AWS badge')
