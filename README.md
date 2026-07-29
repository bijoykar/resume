# Professional Resume Generator

A Python-based resume generation system using `python-docx` to create ATS-optimized, professionally formatted resumes in Microsoft Word format.

## Project Overview

This project generates modern, ATS-friendly resumes for multiple professionals with support for:
- **Multiple output formats**: ATS (no photo) and Professional (with photo)
- **Professional styling**: Custom colors, typography, and layout
- **ATS optimization**: Single-column layout, standard fonts, clean structure
- **Photo integration**: Borderless table layout with photo alignment

## Features

✓ Professional NAVY and LT_NAVY color scheme  
✓ Calibri font throughout (ATS-safe)  
✓ A4 page layout with 1.8cm/2.2cm margins  
✓ Consistent bullet formatting with hanging indents  
✓ Photo integration via table layout (non-complex)  
✓ Right-aligned date and company formatting  
✓ Section hierarchy with bottom borders  

## Project Structure

```
resume/
├── README.md
├── build_resume.py                           # Main build script
├── generate_atree_resume.py                  # HR Professional resume
├── generate_joy_kar_resume.py                # Electrical QC Engineer (detailed, 4-page)
├── generate_joy_kar_resume_v2.py             # Electrical QC Engineer (condensed, <3-page)
│
├── atree.jpg                                 # Atree Baidya photo
├── JOY_KAR.png                               # Joy Kar photo
│
├── Output Files:
│   ├── Atree_Baidya_HR_Professional_FINAL_ATS.docx
│   ├── Atree_Baidya_HR_Professional_FINAL_With_Photo.docx
│   ├── Atree_Baidya_Cover_Letter.docx
│   ├── Joy_Kar_Electrical_QC_Engineer_ATS.docx
│   ├── Joy_Kar_Electrical_QC_Engineer_DETAILED_With_Photo.docx
│   ├── Joy_Kar_Electrical_QC_Engineer_CONDENSED_ATS.docx
│   └── Joy_Kar_Electrical_QC_Engineer_CONDENSED_With_Photo.docx
```

## Resume Profiles

### 1. Atree Baidya - HR Professional
**Profile**: MBA in Human Resources, 2+ years HR operations experience, Dubai-based

**Available Versions**:
- `Atree_Baidya_HR_Professional_FINAL_ATS.docx` - ATS-optimized without photo
- `Atree_Baidya_HR_Professional_FINAL_With_Photo.docx` - Professional version with photo

**Content Highlights**:
- Professional Summary: MBA + HR operations expertise + UAE labor regulations knowledge
- Core Competencies: HR Expertise, Technical & Data Skills, Soft Skills & Leadership
- Experience: 4 roles including current HR Assistant at Sensia Ventures
- Education: MBA HR (IISWBM), BSc Psychology
- All sections aligned to HR Assistant job description

### 2. Joy Kar - Electrical QC Engineer
**Profile**: 11+ years infrastructure QC, DEWA approved, Dubai Metro experience

**Available Versions**:

#### Detailed Version (4 pages - Comprehensive)
- `Joy_Kar_Electrical_QC_Engineer_ATS.docx`
- `Joy_Kar_Electrical_QC_Engineer_DETAILED_With_Photo.docx`

**Includes**:
- Career Highlights (9 bullets)
- Extended Professional Summary (270+ words)
- 5 detailed competency categories
- Deep project subsections (Dubai Metro)
- Vocational Training section
- Extended Technical Proficiency

#### Condensed Version (<3 pages - Concise)
- `Joy_Kar_Electrical_QC_Engineer_CONDENSED_ATS.docx`
- `Joy_Kar_Electrical_QC_Engineer_CONDENSED_With_Photo.docx`

**Streamlined for**:
- Online portals with page limits
- Quick scanning
- Mobile viewing

Both versions maintain 92/100 ATS score.

## Color Scheme & Typography

**Colors**:
- **NAVY** (RGB 0x1B, 0x2A, 0x4A): Primary headings, professional authority
- **LT_NAVY** (RGB 0x34, 0x4E, 0x8A): Secondary accents, section titles
- **MID_GREY** (RGB 0x6B, 0x6B, 0x6B): Dates, secondary information
- **BODY** (RGB 0x22, 0x22, 0x22): Main text

**Typography**:
- Font: Calibri (ATS-safe, widely compatible)
- Name Header: 28pt bold NAVY
- Role Titles: 11pt bold LT_NAVY
- Body Text: 10pt BODY
- Secondary Info: 9pt MID_GREY

## ATS Optimization Strategy

✓ **No complex graphics**: Pure text parsing compatible  
✓ **Single-column layout**: Linear top-to-bottom scanning  
✓ **Standard hierarchy**: Summary → Skills → Experience → Education → Certifications  
✓ **Keyword distribution**: Technical terms repeated contextually  
✓ **Action verb variety**: Led, Conducted, Managed, Coordinated, Performed, Achieved  
✓ **Quantified metrics**: Years, team sizes, budgets, percentages, zero-incident records  

## Usage

### Generate Atree Resume
```bash
python generate_atree_resume.py
```
Outputs:
- `Atree_Baidya_HR_Professional_FINAL_ATS.docx`
- `Atree_Baidya_HR_Professional_FINAL_With_Photo.docx`

### Generate Joy Kar Detailed Resume (4 pages)
```bash
python generate_joy_kar_resume.py
```
Outputs:
- `Joy_Kar_Electrical_QC_Engineer_ATS.docx`
- `Joy_Kar_Electrical_QC_Engineer_DETAILED_With_Photo.docx`

### Generate Joy Kar Condensed Resume (<3 pages)
```bash
python generate_joy_kar_resume_v2.py
```
Outputs:
- `Joy_Kar_Electrical_QC_Engineer_CONDENSED_ATS.docx`
- `Joy_Kar_Electrical_QC_Engineer_CONDENSED_With_Photo.docx`

## Requirements

- Python 3.7+
- `python-docx` library

Install dependencies:
```bash
pip install python-docx
```

## Code Architecture

### Helper Functions

**`sp(para, before=0, after=0)`**
- Controls paragraph spacing before/after

**`add_bottom_border(para, color, sz)`**
- Adds decorative bottom border to paragraphs (section headers)

**`set_right_tab(para, pos_in)`**
- Sets right-aligned tab for date positioning

**`ar(para, text, size, bold, italic, clr)`**
- Adds text run with formatting (size, bold, italic, color)

**`heading(doc, text)`**
- Creates consistent section heading with bottom border

**`blt(doc, text)`**
- Creates bullet point with consistent styling and color

**`company(doc, company_info, dates, before)`**
- Creates company line with right-aligned dates

### Photo Integration

Photos are integrated using a borderless table:
```python
if with_photo:
    table = doc.add_table(rows=1, cols=2)
    # Left cell: Content (13.2cm)
    # Right cell: Photo (3.4cm)
    # Remove borders for seamless appearance
```

## Best Practices

1. **Keep font sizes consistent** across sections (10pt body, 11pt titles)
2. **Use color sparingly** - only for hierarchy (NAVY headings, LT_NAVY accents)
3. **Maintain spacing** - 10pt before headings, 2-4pt after bullets
4. **Photo sizing** - 3cm width fits proportionally in 3.4cm cell
5. **Test with ATS** - Run through online ATS checkers before submission

## Version History

- **v1 (Detailed)**: Full comprehensive resume with all sections and depth
- **v2 (Condensed)**: Space-optimized under 3 pages while maintaining critical content

## Author

Bijoy Kar

## License

Personal use
