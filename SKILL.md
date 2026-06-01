---
name: "literature-review"
description: "Generate formatted literature review documents (.docx) from a set of PDF papers. Use when a user provides PDF files and a Word template (or format requirements) and needs a structured review/survey paper synthesized from those sources."
---

# Literature Review Generator

## When to use
- The user provides a folder of PDF research papers and a format requirement (Word template or description).
- The user needs a formatted .docx review/survey article, not just a summary.
- The task involves: reading multiple PDFs into analyzing content into writing a structured review into outputting a formatted Word document.

## Workflow

### 1. Understand the requirements
Read the Word template / format requirements first. Identify:
- Paper title format
- Required sections (usually: Title, Abstract, Keywords, Introduction, Body sections, Summary/Outlook, References)
- Font, size, and spacing conventions (e.g. Chinese Song/Hei fonts, heading styles, reference format)

### 2. Read the PDF papers
Extract text from PDFs using available tools. Preferred methods (in order):
- Node.js: pdfjs-dist (already in project if installed)
- Python: pdfplumber, pypdf, or PyMuPDF
- Fallback: Raw PDF text extraction (less reliable)

For each paper, extract:
- Title, authors, journal, year
- Abstract
- Key findings / methods / results
- Conclusion

Store extracted text in a temporary directory for reference.

### 3. Analyze and synthesize
Organize the papers by topic. Identify:
- Common themes and patterns
- Contradictions or differing findings
- Chronological development
- Gaps in the literature

Map each paper contribution to the planned sections of the review.

### 4. Generate the review document
Create a docx file using the OpenXML approach.

**Manual ZIP + XML method (no external libraries needed):**
1. Build the content as raw XML following the OpenXML WordprocessingML schema
2. Include these files inside the ZIP: [Content_Types].xml, _rels/.rels, word/document.xml, word/styles.xml, word/_rels/document.xml.rels
3. Zip the files together with a .docx extension
4. Use UTF-8 encoding without BOM for all XML files

### 5. Document formatting rules
- Title: Hei font (黑体), Bold, 36pt, centered
- Heading 1: Hei font (黑体), Bold, 28pt, left-aligned
- Heading 2: Hei font (黑体), Bold, 26pt, left-aligned
- Body text: Song font (宋体), 24pt (12pt), 1.5x line spacing
- References: Times New Roman, 21pt (10.5pt), hanging indent 480 twips
- Spacing: heading before 240/after 120 (H1), before 200/after 100 (H2)

### 6. Reference formatting
Format each reference as:
[N] Author A, Author B, et al. Title of paper[J]. Journal Name, Year, Volume: Pages.

Only include papers that were actually provided by the user. Do NOT add extra background references unless explicitly requested.

### 7. Content quality
- Each section must cite at least one paper from the provided set
- Write in formal academic language (as required by the template)
- Include in-text citations like [1], [2-3]
- Abstract should cover: background, scope, methods reviewed, key findings, conclusion
- The review must be original synthesis, not just a list of paper summaries

### 8. ZIP file structure for .docx
Create a ZIP archive containing these entries (order matters):
```
[Content_Types].xml
_rels/.rels
word/_rels/document.xml.rels
word/document.xml
word/styles.xml
```

### 9. Verification
- After generating, verify the docx can be opened (check ZIP structure)
- Verify Chinese characters are present and not escaped as unicode sequences
- Count references = number of provided papers
- Verify citation numbers match the reference list
- Check content with a quick scan of the XML text
