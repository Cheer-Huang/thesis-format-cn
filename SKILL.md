---
name: thu-thesis-format
description: "Format Word (.docx) documents according to Tsinghua University Graduate Thesis Writing Guidelines (清华大学研究生学位论文写作指南). This skill should be used when users need to format academic theses, dissertations, or course project reports to comply with Tsinghua's formatting standards, including page margins, heading styles, body text, figure/table captions, three-line tables, references, page numbers, and table of contents. It is also useful for auditing existing documents for compliance. Triggers include mentions of 清华, 论文格式, 学位论文, 毕业论文, thesis format, 三线表, 论文排版, or requests to format/audit Word documents according to university standards."
---

# Tsinghua University Thesis Format Skill

## Purpose

Format Word (.docx) documents to comply with the Tsinghua University Graduate
Thesis Writing Guidelines (清华大学研究生学位论文写作指南, 202605 edition). Provides
both one-click formatting and detailed compliance auditing.

## When to Use

- Formatting a thesis, dissertation, or course project report to Tsinghua standards
- Auditing an existing document for format compliance
- Fixing specific format issues (fonts, margins, three-line tables, captions, etc.)

## Prerequisites

- Python 3.8+ with `python-docx` installed (`pip install python-docx`)
- The input file must be a valid .docx file

## Workflow

### Step 1: Audit the Document (Optional but Recommended)

Before formatting, audit the document to understand current state:

```bash
python scripts/audit_thu_thesis.py <input.docx>
```

This produces a detailed report showing which elements comply and which need fixing.

### Step 2: Format the Document

Apply all Tsinghua formatting standards in one pass:

```bash
python scripts/format_thu_thesis.py <input.docx> [output.docx]
```

If `output.docx` is omitted, the input file is modified in place.

The script performs these operations in order:
1. **Page setup**: margins 3cm, header/footer 2.2cm, A4
2. **Style colors**: all styles set to black (#000000)
3. **Run colors**: all text runs explicitly set to black
4. **Headings**: chapter titles (16pt bold HeiTi/Arial, centered), section titles (14pt HeiTi/Arial, left)
5. **Body text**: 12pt SongTi/TNR, justified, 2-char indent, 20pt fixed line spacing
6. **Figure captions**: 11pt SongTi/TNR, centered, chapter-numbered (图X.Y)
7. **Table captions**: 11pt SongTi/TNR, centered, chapter-numbered (表X.Y)
8. **Three-line tables**: 1.5pt top/bottom borders, 1pt header border, no side borders
9. **References**: 10.5pt SongTi/TNR, 16pt fixed spacing, hanging indent
10. **Header/Footer**: 10.5pt centered header, PAGE field footer
11. **Table of contents**: TOC field inserted before abstract

### Step 3: Post-Processing (Manual)

After running the script, the user must:
1. Open the document in Microsoft Word
2. Right-click the table of contents area and select "Update Field" (更新域)
3. Review the document for any content-specific adjustments

### Step 4: Re-Audit (Verification)

Run the audit script again to verify all formatting is correct:

```bash
python scripts/audit_thu_thesis.py <output.docx>
```

## Critical Workflow Rule

Formatting is the **last step** before submission. The correct workflow is:

1. Edit content (add/remove paragraphs, modify text, insert figures/tables)
2. Confirm content is final
3. Run formatting script
4. Update TOC in Word
5. Submit

**Do NOT edit the document after formatting.** AI agents that insert new paragraphs or modify text after formatting may inherit default Word styles, breaking the carefully set fonts, indents, line spacing, and table borders. If content changes are needed, make them first, then re-run the formatting script.

## Format Specifications

Detailed format specifications are in `references/thu_format_spec.md`. Key specs:

| Element | Font (CN/EN) | Size | Align | Line Spacing | Before/After |
|---------|-------------|------|-------|-------------|-------------|
| Chapter title | 黑体/Arial | 16pt | Center | Single | 24pt/18pt |
| Section title | 黑体/Arial | 14pt | Left | 20pt fixed | 24pt/6pt |
| Body text | 宋体/TNR | 12pt | Justify | 20pt fixed | 0/0 |
| Figure caption | 宋体/TNR | 11pt | Center | Single | 6pt/12pt |
| Table caption | 宋体/TNR | 11pt | Center | Single | 12pt/6pt |
| References | 宋体/TNR | 10.5pt | Justify | 16pt fixed | 3pt/0 |

## Important Notes

- **Font color**: Word's built-in Heading styles have blue colors (365F91, 4F81BD). The script overrides both style-level and run-level colors to black.
- **Chapter naming**: The script converts "一、问题" format to "第1章　问题" format automatically.
- **Figure/table numbering**: Captions should use chapter-based numbering (图1.1, 表2.3, etc.). The script formats existing captions but does not auto-number them.
- **Source lines**: Figure source citations (来源：xxx) should be removed; references are cited in body text via [1], [2] notation.
- **TOC**: The table of contents is a Word TOC field that requires manual update in Word.
- **Page breaks**: Each chapter automatically gets a page break before it.
- **Three-line tables**: The script formats all tables as three-line tables. If a table needs a different format, format it manually after running the script.

## Script API

The formatting script (`scripts/format_thu_thesis.py`) exposes individual functions
that can be imported and called selectively:

```python
from format_thu_thesis import (
    fix_page_setup, fix_style_colors, fix_all_run_colors,
    fix_headings, fix_body_text, fix_figure_captions,
    fix_table_captions, fix_three_line_tables, fix_references,
    fix_cover, fix_header_footer, add_toc, add_table_titles,
    format_thu_thesis
)

# Format everything at once
format_thu_thesis("input.docx", "output.docx")

# Or call individual functions
doc = Document("input.docx")
fix_page_setup(doc)
fix_three_line_tables(doc)
doc.save("output.docx")
```
