---
name: "literature-review"
description: "Complete literature review pipeline: extract PDFs, analyze papers, generate formatted .docx with DOI clickable citations + EndNote import files"
license: MIT
metadata:
  author: "Literature Review Generator"
  version: "2.0.0"
---

# Literature Review Generator v2

## When to use

User provides PDF research papers and needs:
- A structured academic review in .docx format
- Clickable DOI citations (Ctrl+click [N] opens paper in browser)
- EndNote-compatible import files (.ris and .xml)
- Academic-quality analysis following standard paper review framework

## Full Pipeline

PDF papers folder (*.pdf) + Cover template (.docx, optional)
Step 1: Extract PDF text (pdfjs-dist)
Step 2: Extract metadata (title, authors, journal, year, DOI, volume, pages)
Step 3: Detect figures per page (image count + caption text)
Step 4: Analyze papers (academic-researcher framework) - RQ, methodology, findings, implications, limitations
Step 5: Write structured review - Title, Abstract, Introduction, Thematic sections, Conclusion, References
Step 6: Generate EndNote import files - references.ris + references.xml
Step 7: Output - review.docx with hyperlinked citations, references.ris, references.xml

## Paper Analysis Framework

Analyze each paper addressing 5 dimensions:
1. Research Question and Significance - core RQ, why it matters, gap filled
2. Methodology - design, sample, variables, techniques, limitations
3. Key Findings - main results with quantitative data, statistical significance
4. Interpretation and Implications - authors interpretation, applications, connections
5. Limitations and Future Directions - remaining questions, suggested work

## Literature Review Structure

# Title (HeiTi 36pt centered)
## Abstract - concise summary
## 1 Introduction - research question, significance, scope
## 2 Theme 1 (e.g. Preparation Methods)
### 2.1 Subtopic
### 2.2 Subtopic
## 3 Theme 2 (e.g. Structure and Properties)
## 4 Theme 3 (e.g. Applications)
## 5 Conclusion and Outlook - key insights, implications, directions
## References - full DOI hyperlinks

## DOCX Hyperlink Mechanism

In word/_rels/document.xml.rels:
<Relationship Id="rIdLink1" Type="...hyperlink" Target="https://doi.org/10.1016/..." TargetMode="External"/>

In word/document.xml:
<w:hyperlink r:id="rIdLink1" w:history="1">
  <w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr>
  <w:t xml:space="preserve">[1]</w:t></w:r>
</w:hyperlink>

Both in-text [N] and reference entries link to the same DOI. Ctrl+click opens browser.

## Document Formatting

| Element | Font | Size | Bold | Alignment |
| Title | HeiTi | 36pt | Yes | Center |
| Heading 1 | HeiTi | 28pt | Yes | Left |
| Heading 2 | HeiTi | 26pt | Yes | Left |
| Body | SongTi | 24pt | No | Justified, indent |
| Citation [N] | SongTi+Hyperlink | 24pt | No | Inline |
| References | Times New Roman | 21pt | Hyperlink | Hanging indent |

## EndNote Import

RIS format - universal, works in EndNote:
TY - JOUR / TI - Title / AU - Author / PY - Year / JF - Journal / VL - Vol / SP - Pages / DO - DOI / UR - URL / ER -

EndNote XML format - richer metadata, PDF paths:
<record><ref-type>17</ref-type><contributors><authors><author>Name</author></authors></contributors>
<titles><title>Title</title><secondary-title>Journal</secondary-title></titles>
<dates><year>2025</year></dates><volume>118</volume><pages>118396</pages>
<electronic-resource-num>10.1016/...</electronic-resource-num>
<urls><web-url>https://doi.org/...</web-url></urls>
<pdf-urls><url>file:///path/to/paper.pdf</url></pdf-urls></record>

## Verification Checklist
- [ ] docx opens correctly in Word
- [ ] Hyperlinks are blue + underlined
- [ ] Ctrl+click [N] opens browser to DOI page
- [ ] Reference list entries are clickable hyperlinks
- [ ] .ris file imports into EndNote (File, Import, RIS Reference)
- [ ] .xml file imports into EndNote (File, Import, EndNote Import)
- [ ] Citations in text match reference numbering
- [ ] Cover page template preserved if provided