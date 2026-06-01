// review-helper.js — PDF extraction, clickable-citation docx builder, RIS export
// Usage: const { extractPDFs, buildClickableDocx, ... } = require("./review-helper.js");

const fs = require("fs");
const path = require("path");

// ===== 1. PDF Text Extraction =====
async function extractPDFs(folderPath, tmpDir) {
  tmpDir = tmpDir || "tmp/papers";
  const files = fs.readdirSync(folderPath).filter(f => f.toLowerCase().endsWith(".pdf"));
  const results = [];
  fs.mkdirSync(tmpDir, { recursive: true });

  try {
    const pdfjsLib = require("pdfjs-dist/legacy/build/pdf.mjs");
    for (const file of files) {
      const fullPath = path.join(folderPath, file);
      console.log("Reading:", file);
      try {
        const pdf = await pdfjsLib.getDocument(fullPath).promise;
        let text = "";
        for (let i = 1; i <= Math.min(pdf.numPages, 25); i++) {
          const page = await pdf.getPage(i);
          const content = await page.getTextContent();
          text += "---Page " + i + "---\n" + content.items.map(item => item.str).join(" ") + "\n\n";
        }
        const txtPath = path.join(tmpDir, file.replace(".pdf", ".txt"));
        fs.writeFileSync(txtPath, text, "utf8");
        results.push({ file, pages: pdf.numPages, text, txtPath });
        console.log("  OK: " + pdf.numPages + " pages, " + text.length + " chars");
      } catch (e) { console.log("  FAIL: " + file + " - " + e.message); }
    }
  } catch (e) { console.log("pdfjs-dist unavailable: " + e.message); }
  return results;
}

// ===== 2. Figure Detection =====
async function detectFigures(pdfPath) {
  const results = [];
  try {
    const pdfjsLib = require("pdfjs-dist/legacy/build/pdf.mjs");
    const doc = await pdfjsLib.getDocument(pdfPath).promise;
    for (let p = 1; p <= doc.numPages; p++) {
      const page = await doc.getPage(p);
      const ops = await page.getOperatorList();
      const textContent = await page.getTextContent();
      const pageText = textContent.items.map(i => i.str).join(" ");
      let imgCount = 0;
      for (const fn of ops.fnArray) {
        if (fn === pdfjsLib.OPS.paintImageXObject || fn === pdfjsLib.OPS.paintInlineImageXObject) imgCount++;
      }
      if (imgCount > 0) {
        let caption = "";
        const m = pageText.match(/(?:Fig\.?\s*\d+[^.]*\.|Figure\s*\d+[^.]*\.|图\s*\d+[^.]*[。\.])/g);
        if (m) caption = m.join("; ");
        results.push({ page: p, imageCount: imgCount, caption: caption || "(auto-detected)" });
      }
    }
  } catch (e) { console.log("Figure detection error: " + e.message); }
  return results;
}

// ===== 3. Clickable-Citation DOCX Builder =====

/**
 * Build a docx with hyperlinked citations.
 *
 * @param {Array} sections  — [{type: "h1"|"h2"|"p", text: "..."}]
 * @param {Array} refs      — ["[1] Author...", "[2] Author...", ...]
 * @param {string} title    — Document title
 * @param {Array} mediaFiles — [{name, localPath}] (optional)
 * @param {string} outputPath — Where to save the .docx
 */
function buildClickableDocx(sections, refs, title, mediaFiles, outputPath) {
  const hei = "黑体", song = "宋体";
  const ESC = s => String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");

  let body = "";

  // --- Title ---
  if (title) {
    body += `<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="200"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="${hei}" w:hAnsi="Times New Roman" w:eastAsia="${hei}"/><w:b/><w:sz w:val="36"/></w:rPr><w:t xml:space="preserve">${ESC(title)}</w:t></w:r></w:p>`;
  }

  // --- Body sections ---
  // First pass: collect all citation markers [1], [2-3], [1,3,5] etc.
  const citationRefs = new Set();
  for (const sec of sections) {
    const matches = sec.text.match(/\[[\d,\-\s]+\]/g) || [];
    for (const m of matches) {
      const nums = m.replace(/[\[\]\s]/g, "").split(",");
      for (const n of nums) {
        if (n.includes("-")) {
          const [a, b] = n.split("-").map(Number);
          for (let i = a; i <= b; i++) citationRefs.add(i);
        } else {
          citationRefs.add(parseInt(n));
        }
      }
    }
  }

  // Replace [N] citations with hyperlinks
  function renderText(text) {
    return text.replace(/\[([\d,\-\s]+)\]/g, (match, nums) => {
      const parts = nums.split(",").map(s => s.trim());
      const linked = parts.map(p => {
        if (p.includes("-")) {
          const [a, b] = p.split("-").map(Number);
          return a + "-" + b;
        }
        const n = parseInt(p);
        if (!isNaN(n)) {
          const bookmarkName = `_Ref${n}`;
          return `<w:hyperlink w:anchor="${bookmarkName}" w:history="1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr><w:t xml:space="preserve">[${n}]</w:t></w:r></w:hyperlink>`;
        }
        return match;
      });
      return linked.join(",");
    });
  }

  for (const sec of sections) {
    const t = renderText(ESC(sec.text));
    if (sec.type === "h1") {
      body += `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="${hei}" w:hAnsi="Times New Roman" w:eastAsia="${hei}"/><w:b/><w:sz w:val="28"/></w:rPr><w:t xml:space="preserve">${ESC(sec.text)}</w:t></w:r></w:p>`;
    } else if (sec.type === "h2") {
      body += `<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="${hei}" w:hAnsi="Times New Roman" w:eastAsia="${hei}"/><w:b/><w:sz w:val="26"/></w:rPr><w:t xml:space="preserve">${ESC(sec.text)}</w:t></w:r></w:p>`;
    } else if (sec.type === "figure") {
      body += `<w:p><w:r><w:rPr><w:rFonts w:ascii="${song}" w:hAnsi="Times New Roman" w:eastAsia="${song}"/><w:sz w:val="21"/><w:i/></w:rPr><w:t xml:space="preserve">${ESC(sec.text)}</w:t></w:r></w:p>`;
    } else {
      // For paragraphs with hyperlinks, we need to handle mixed content
      body += `<w:p><w:r><w:rPr><w:rFonts w:ascii="${song}" w:hAnsi="Times New Roman" w:eastAsia="${song}"/><w:sz w:val="24"/></w:rPr>${t}</w:r></w:p>`;
    }
  }

  // --- References with bookmarks ---
  body += `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="${hei}" w:hAnsi="Times New Roman" w:eastAsia="${hei}"/><w:b/><w:sz w:val="28"/></w:rPr><w:t xml:space="preserve">参考文献</w:t></w:r></w:p>`;

  for (let i = 0; i < refs.length; i++) {
    const refNum = i + 1;
    const bookmarkName = `_Ref${refNum}`;
    const bookmarkId = i;
    body += `<w:p><w:pPr><w:ind w:hanging="480" w:left="480"/><w:spacing w:line="300" w:lineRule="auto" w:after="60"/></w:pPr><w:bookmarkStart w:id="${bookmarkId}" w:name="${bookmarkName}"/><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="21"/></w:rPr><w:t xml:space="preserve">${ESC(refs[i])}</w:t></w:r><w:bookmarkEnd w:id="${bookmarkId}"/></w:p>`;
  }

  const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;

  // Styles with Hyperlink character style
  const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/><w:pPr><w:spacing w:line="360" w:lineRule="auto" w:after="120"/></w:pPr><w:rPr><w:rFonts w:ascii="宋体" w:hAnsi="Times New Roman" w:eastAsia="宋体"/><w:sz w:val="24"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="240" w:after="120"/><w:keepNext/><w:keepLines/></w:pPr><w:rPr><w:rFonts w:ascii="黑体" w:hAnsi="Times New Roman" w:eastAsia="黑体"/><w:b/><w:sz w:val="28"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="200" w:after="100"/><w:keepNext/><w:keepLines/></w:pPr><w:rPr><w:rFonts w:ascii="黑体" w:hAnsi="Times New Roman" w:eastAsia="黑体"/><w:b/><w:sz w:val="26"/></w:rPr></w:style><w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style></w:styles>`;

  buildDocx(docXml, stylesXml, mediaFiles || [], outputPath);
}

// ===== Low-level DOCX Builder =====
function buildDocx(docXml, stylesXml, mediaFiles, outputPath) {
  const hasMedia = mediaFiles.length > 0;
  const contentTypes = buildContentTypes(hasMedia);
  const rels = buildRels();
  const wordRels = buildWordRels(mediaFiles);

  const files = [
    { name: "[Content_Types].xml", data: contentTypes },
    { name: "_rels/.rels", data: rels },
    { name: "word/document.xml", data: docXml },
    { name: "word/styles.xml", data: stylesXml },
    { name: "word/_rels/document.xml.rels", data: wordRels }
  ];

  for (const mf of mediaFiles) {
    const imgData = fs.readFileSync(mf.localPath);
    files.push({ name: "word/media/" + mf.name, data: imgData.toString("binary"), isBinary: true });
  }

  const zipData = buildZip(files);
  fs.writeFileSync(outputPath, Buffer.from(zipData, "binary"), { encoding: null });
  const size = fs.statSync(outputPath).size;
  console.log("DOCX created: " + outputPath + " (" + (size / 1024).toFixed(0) + " KB)");
  return size;
}

function buildContentTypes(hasMedia) {
  let mediaTypes = "";
  if (hasMedia) {
    mediaTypes += '<Override PartName="/word/media/fig1.png" ContentType="image/png"/>\n';
  }
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' + mediaTypes + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>';
}

function buildRels() {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>';
}

function buildWordRels(mediaFiles) {
  let extra = "";
  let idx = 2;
  for (const mf of mediaFiles) {
    const ext = path.extname(mf.name).toLowerCase();
    const mimeType = ext === ".png" ? "image/png" : "image/jpeg";
    extra += `<Relationship Id="rIdImg${idx}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${mf.name}"/>\n`;
    idx++;
  }
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' + extra + '</Relationships>';
}

// ===== RIS Generator =====
function generateRIS(refs) {
  const parsed = refs.map(r => parseRef(r));
  let ris = "";
  for (const r of parsed) {
    ris += "TY  - JOUR\n";
    for (const a of r.authors) ris += "AU  - " + a + "\n";
    if (r.year) ris += "PY  - " + r.year + "\n";
    if (r.title) ris += "TI  - " + r.title + "\n";
    if (r.journal) ris += "JF  - " + r.journal + "\n";
    if (r.volume) ris += "VL  - " + r.volume + "\n";
    if (r.pages) ris += "SP  - " + r.pages + "\n";
    if (r.doi) ris += "DO  - " + r.doi + "\n";
    ris += "ER  - \n\n";
  }
  return ris;
}

function parseRef(refStr) {
  const result = { authors: [], year: "", title: "", journal: "", volume: "", pages: "", doi: "" };

  // Extract [N] prefix
  const clean = refStr.replace(/^\[\d+\]\s*/, "");

  // Split by period for structured parsing
  const parts = clean.split(/\.\s+/);

  // First part: Authors. Title[J]
  if (parts[0]) {
    // Try to extract title between authors and [J]
    const authorTitleMatch = parts[0].match(/^(.+?)\.\s+(.+?)\s*\[J\]/);
    if (authorTitleMatch) {
      result.authors = authorTitleMatch[1].split(/[,;]\s+/).filter(a => a && !a.match(/^et\s+al/i));
      result.title = authorTitleMatch[2].trim();
    }
  }

  // Second part: Journal Name, Year, Volume: Pages.
  for (const p of parts) {
    const journalMatch = p.match(/([^,]+),\s*(\d{4}),\s*(\d+):\s*(.+)/);
    if (journalMatch) {
      result.journal = journalMatch[1].trim();
      result.year = journalMatch[2];
      result.volume = journalMatch[3];
      result.pages = journalMatch[4].replace(/\.$/, "").trim();
    }
    // DOI
    const doiMatch = p.match(/https?:\/\/doi\.org\/(.+)/);
    if (doiMatch) result.doi = doiMatch[1];
  }

  return result;
}

// ===== ZIP Builder =====
function crc32(buf) {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i++) { let c = i; for (let j = 0; j < 8; j++) c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1); table[i] = c; }
  let crc = 0xffffffff;
  for (const b of buf) crc = table[(crc ^ b) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

function buildZip(files) {
  const lh = []; const ch = []; let offset = 0;
  for (const f of files) {
    const nb = Buffer.from(f.name, "utf8");
    const db = f.isBinary ? Buffer.from(f.data, "binary") : Buffer.from(f.data, "utf8");
    const crc = crc32(db);
    const local = Buffer.alloc(30 + nb.length);
    local.writeUInt32LE(0x04034b50, 0); local.writeUInt16LE(20, 4); local.writeUInt16LE(0, 8);
    local.writeUInt32LE(crc, 14); local.writeUInt32LE(db.length, 18); local.writeUInt32LE(db.length, 22);
    local.writeUInt16LE(nb.length, 26); nb.copy(local, 30);
    lh.push({ h: local, d: db }); offset += local.length + db.length;
    const cent = Buffer.alloc(46 + nb.length);
    cent.writeUInt32LE(0x02014b50, 0); cent.writeUInt16LE(20, 4); cent.writeUInt16LE(20, 6);
    cent.writeUInt32LE(crc, 16); cent.writeUInt32LE(db.length, 20); cent.writeUInt32LE(db.length, 24);
    cent.writeUInt16LE(nb.length, 28); cent.writeUInt32LE(lh.length - 1, 42); nb.copy(cent, 46);
    ch.push(cent);
  }
  let off = 0;
  for (let i = 0; i < lh.length; i++) { ch[i].writeUInt32LE(off, 42); off += lh[i].h.length + lh[i].d.length; }
  const co = off; const cs = ch.reduce((s, h) => s + h.length, 0);
  const e = Buffer.alloc(22); e.writeUInt32LE(0x06054b50, 0);
  e.writeUInt16LE(ch.length, 8); e.writeUInt16LE(ch.length, 10); e.writeUInt32LE(cs, 12); e.writeUInt32LE(co, 16);
  const total = co + cs + 22; const result = Buffer.alloc(total); let pos = 0;
  for (const x of lh) { x.h.copy(result, pos); pos += x.h.length; x.d.copy(result, pos); pos += x.d.length; }
  for (const x of ch) { x.copy(result, pos); pos += x.length; }
  e.copy(result, pos); return result;
}


/**
 * Build a docx with DOI hyperlinks (Ctrl+click opens the paper in browser).
 *
 * @param {Array} sections — [{type:"h1"|"h2"|"p", text:"..."}]
 * @param {Array} refsWithDOI — [{text:"[1] Author...", url:"https://doi.org/..."}]
 * @param {string} title
 * @param {Array} mediaFiles — [{name, localPath}] (optional)
 * @param {string} outputPath
 */
function buildDOIClickableDocx(sections, refsWithDOI, title, mediaFiles, outputPath) {
  const hei = "黑体", song = "宋体";
  const ESC = s => String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
  let body = "";

  function addP(extra, content) { body += "<w:p>" + extra + content + "</w:p>"; }

  // Title
  if (title) {
    addP('<w:pPr><w:jc w:val="center"/><w:spacing w:after="200"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="' + hei + '" w:hAnsi="Times New Roman" w:eastAsia="' + hei + '"/><w:b/><w:sz w:val="36"/></w:rPr><w:t xml:space="preserve">' + ESC(title) + '</w:t></w:r>');
  }

  // Body sections
  for (const sec of sections) {
    if (sec.type === "h1") {
      addP('<w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="' + hei + '" w:hAnsi="Times New Roman" w:eastAsia="' + hei + '"/><w:b/><w:sz w:val="28"/></w:rPr><w:t xml:space="preserve">' + ESC(sec.text) + '</w:t></w:r>');
    } else if (sec.type === "h2") {
      addP('<w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="' + hei + '" w:hAnsi="Times New Roman" w:eastAsia="' + hei + '"/><w:b/><w:sz w:val="26"/></w:rPr><w:t xml:space="preserve">' + ESC(sec.text) + '</w:t></w:r>');
    } else {
      // Replace [N] with DOI hyperlinks
      const parts = sec.text.split(/(\[\d+\])/g);
      let xml = '<w:r><w:rPr><w:rFonts w:ascii="' + song + '" w:hAnsi="Times New Roman" w:eastAsia="' + song + '"/><w:sz w:val="24"/></w:rPr>';
      for (const p of parts) {
        const m = p.match(/^\[(\d+)\]$/);
        if (m) {
          const n = m[1];
          const idx = parseInt(n) - 1;
          const url = (refsWithDOI[idx] && refsWithDOI[idx].url) ? refsWithDOI[idx].url : "#";
          xml += '</w:r><w:hyperlink r:id="rIdLink' + n + '" w:history="1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr><w:t xml:space="preserve">[' + n + ']</w:t></w:r></w:hyperlink><w:r><w:rPr><w:rFonts w:ascii="' + song + '" w:hAnsi="Times New Roman" w:eastAsia="' + song + '"/><w:sz w:val="24"/></w:rPr>';
        } else {
          xml += '<w:t xml:space="preserve">' + ESC(p) + '</w:t>';
        }
      }
      xml += '</w:r>';
      addP("", xml);
    }
  }

  // References heading
  addP('<w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="' + hei + '" w:hAnsi="Times New Roman" w:eastAsia="' + hei + '"/><w:b/><w:sz w:val="28"/></w:rPr><w:t xml:space="preserve">参考文献</w:t></w:r>');

  // References with DOI hyperlinks
  for (let i = 0; i < refsWithDOI.length; i++) {
    const n = i + 1;
    const url = refsWithDOI[i].url || "#";
    addP('<w:pPr><w:ind w:hanging="480" w:left="480"/><w:spacing w:line="300" w:lineRule="auto" w:after="60"/></w:pPr><w:hyperlink r:id="rIdLink' + n + '" w:history="1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:color w:val="0563C1"/><w:u w:val="single"/><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="21"/></w:rPr><w:t xml:space="preserve">' + ESC(refsWithDOI[i].text) + '</w:t></w:r></w:hyperlink>');
  }

  const docXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>' + body + '</w:body></w:document>';

  // Build external hyperlink relationships
  let extRels = "";
  for (let i = 0; i < refsWithDOI.length; i++) {
    extRels += '<Relationship Id="rIdLink' + (i+1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="' + (refsWithDOI[i].url || "#") + '" TargetMode="External"/>';
  }
  const wordRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' + extRels + '</Relationships>';

  const stylesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/><w:pPr><w:spacing w:line="360" w:lineRule="auto" w:after="120"/></w:pPr><w:rPr><w:rFonts w:ascii="宋体" w:hAnsi="Times New Roman" w:eastAsia="宋体"/><w:sz w:val="24"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="240" w:after="120"/><w:keepNext/><w:keepLines/></w:pPr><w:rPr><w:rFonts w:ascii="黑体" w:hAnsi="Times New Roman" w:eastAsia="黑体"/><w:b/><w:sz w:val="28"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="200" w:after="100"/><w:keepNext/><w:keepLines/></w:pPr><w:rPr><w:rFonts w:ascii="黑体" w:hAnsi="Times New Roman" w:eastAsia="黑体"/><w:b/><w:sz w:val="26"/></w:rPr></w:style><w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style></w:styles>';

  buildDocx(docXml, stylesXml, mediaFiles || [], outputPath);
}


// ===== 5. Metadata Extraction from PDF Text =====
function extractMetadata(text, pdfPath) {
  const result = {
    title: "",
    authors: [],
    journal: "",
    year: "",
    volume: "",
    issue: "",
    pages: "",
    doi: "",
    pdfPath: pdfPath || ""
  };

  // DOI
  const doiMatch = text.match(/10\.\d{4,}\/[^\s\"]+/);
  if (doiMatch) result.doi = doiMatch[0].replace(/[\.\)\]]$/, "");

  // Year
  const yearMatch = text.match(/\b(19\d{2}|20\d{2})\b/);
  if (yearMatch) result.year = yearMatch[1];

  // Journal - look for common journal name patterns
  const journalPatterns = [
    /(?:Journal of|Advanced|Applied|ACS|Nature|Science|Chemical|Materials|Nano|Small|Energy|Acta|Carbon|Polymer|Composites|Ceramics)\s[^\n,]{5,80}(?=\n|,|\.)/,
    /Contents lists available at\s+ScienceDirect\s+\n\s*([^\n]+)/,
    /journal homepage:\s*[^\n]+\n\s*([^\n]+)/
  ];
  for (const p of journalPatterns) {
    const m = text.match(p);
    if (m) { result.journal = m[1] || m[0]; break; }
  }
  result.journal = result.journal.replace(/^\s+|\s+$/g, "");

  // Volume
  const volMatch = text.match(/(?:Vol\.|Volume|v\.|vol)\s*[\.\s]*(\d+)/i);
  if (volMatch) result.volume = volMatch[1];

  // Pages
  const pageMatch = text.match(/(?:\d+)\s*\((\d+)\)\s*(\d+[–\-]\d+)/);
  if (pageMatch) result.pages = pageMatch[2];

  // Authors - extract from first page header
  const lines = text.split("\n");
  let authorLine = "";
  for (let i = 0; i < Math.min(lines.length, 30); i++) {
    const l = lines[i].trim();
    // Look for author pattern: Last, First pattern or comma-separated names
    if (/^[A-Z][a-z]+.*,\s*[A-Z]\./.test(l) && l.length > 10 && l.length < 300) {
      authorLine = l;
      break;
    }
  }
  if (authorLine) {
    // Split by ", " and look for name patterns
    const parts = authorLine.split(",");
    for (let i = 0; i < parts.length - 1; i += 2) {
      const lastName = parts[i].trim();
      const firstPart = parts[i + 1] ? parts[i + 1].trim() : "";
      if (lastName && firstPart && !lastName.match(/^\d+$/) && !firstPart.match(/^\d+$/)) {
        const initial = firstPart.split(" ")[0];
        result.authors.push(lastName + ", " + initial);
      }
    }
    // Clean up
    result.authors = result.authors.filter(a => a.length > 3 && a.length < 50);
  }

  // Title - from first few lines
  for (let i = 0; i < Math.min(lines.length, 8); i++) {
    const l = lines[i].trim();
    if (l.length > 20 && l.length < 300 && !l.includes("http") && !l.includes("@") &&
        !l.includes("ARTICLE") && !l.includes("Contents") && !l.includes("journal")) {
      // Check if it looks like a title (no author keywords)
      if (!/[A-Z][a-z]+,\s*[A-Z]\./.test(l) && !l.match(/^\d+$/)) {
        result.title = l;
        break;
      }
    }
  }

  return result;
}

// ===== 6. EndNote XML Generator =====
function generateEndNoteXML(refs) {
  let xml = '<?xml version="1.0" encoding="utf-8"?>\n';
  xml += '<xml>\n';
  xml += '  <records>\n';

  for (let i = 0; i < refs.length; i++) {
    const r = refs[i];
    xml += '    <record>\n';
    xml += '      <database name="PubMed">n/a</database>\n';
    xml += '      <source-app name="EndNote">Literature Review Generator</source-app>\n';
    xml += '      <rec-number>' + (i + 1) + '</rec-number>\n';
    xml += '      <ref-type name="Journal Article">17</ref-type>\n';

    // Contributors (authors)
    xml += '      <contributors>\n';
    xml += '        <authors>\n';
    if (r.authors && r.authors.length > 0) {
      for (const author of r.authors) {
        xml += '          <author>' + escXML(author) + '</author>\n';
      }
    }
    xml += '        </authors>\n';
    xml += '      </contributors>\n';

    // Titles
    xml += '      <titles>\n';
    xml += '        <title>' + escXML(r.title || "") + '</title>\n';
    xml += '        <secondary-title>' + escXML(r.journal || "") + '</secondary-title>\n';
    xml += '      </titles>\n';

    xml += '      <section>Article</section>\n';

    // Dates
    xml += '      <dates>\n';
    xml += '        <year>' + escXML(r.year || "") + '</year>\n';
    xml += '      </dates>\n';

    if (r.volume) xml += '      <volume>' + escXML(r.volume) + '</volume>\n';
    if (r.issue) xml += '      <number>' + escXML(r.issue) + '</number>\n';
    if (r.pages) xml += '      <pages>' + escXML(r.pages) + '</pages>\n';
    if (r.doi) xml += '      <electronic-resource-num>' + escXML(r.doi) + '</electronic-resource-num>\n';

    // URL
    xml += '      <urls>\n';
    const url = r.doi ? 'https://doi.org/' + r.doi : "";
    xml += '        <web-url>' + escXML(url) + '</web-url>\n';
    xml += '      </urls>\n';

    // PDF attachment path
    if (r.pdfPath) {
      xml += '      <pdf-urls>\n';
      xml += '        <url>' + escXML(r.pdfPath) + '</url>\n';
      xml += '      </pdf-urls>\n';
    }

    xml += '    </record>\n';
  }

  xml += '  </records>\n';
  xml += '</xml>\n';
  return xml;
}

function escXML(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

// ===== 7. Full Pipeline: PDF folder → all outputs =====
async function generateFullReview(pdfFolder, outputDir, title, sections, options) {
  options = options || {};
  const fs = require("fs");
  const path = require("path");
  fs.mkdirSync(outputDir, { recursive: true });

  // Step 1: Extract text from PDFs
  console.log("=== Step 1: Extract PDFs ===");
  const papers = await extractPDFs(pdfFolder, path.join(outputDir, "tmp"));

  // Step 2: Metadata extraction
  console.log("\n=== Step 2: Extract Metadata ===");
  const refs = [];
  for (const p of papers) {
    const meta = extractMetadata(p.text, p.file);
    console.log("  " + p.file + " → Title: " + (meta.title || "(auto-detected)").substring(0, 60));
    refs.push({
      text: formatRefText(meta, p.file),
      url: meta.doi ? "https://doi.org/" + meta.doi : "",
      meta: meta
    });
  }

  // Step 3: Generate EndNote files
  console.log("\n=== Step 3: Generate EndNote Files ===");
  const ris = generateRIS(refs.map(r => ({
    authors: r.meta.authors,
    year: r.meta.year,
    title: r.meta.title,
    journal: r.meta.journal,
    volume: r.meta.volume,
    pages: r.meta.pages,
    doi: r.meta.doi
  })));
  fs.writeFileSync(path.join(outputDir, "references.ris"), ris, "utf8");
  console.log("  references.ris -> " + (ris.length / 1024).toFixed(0) + " KB");

  const enXml = generateEndNoteXML(refs.map(r => ({
    authors: r.meta.authors,
    year: r.meta.year,
    title: r.meta.title,
    journal: r.meta.journal,
    volume: r.meta.volume,
    issue: r.meta.issue,
    pages: r.meta.pages,
    doi: r.meta.doi,
    pdfPath: r.meta.pdfPath ? path.resolve(pdfFolder, r.meta.pdfPath) : ""
  })));
  fs.writeFileSync(path.join(outputDir, "references.xml"), enXml, "utf8");
  console.log("  references.xml -> " + (enXml.length / 1024).toFixed(0) + " KB");

  // Step 4: Detect figures
  console.log("\n=== Step 4: Detect Figures ===");
  const allFigures = [];
  for (const p of papers) {
    const fullPath = path.join(pdfFolder, p.file);
    const figs = await detectFigures(fullPath);
    if (figs.length > 0) {
      allFigures.push({ file: p.file, figures: figs });
      console.log("  " + p.file + ": " + figs.length + " pages with figures");
    }
  }

  // Step 5: Build docx with DOI links
  console.log("\n=== Step 5: Build DOCX ===");
  const docxPath = path.join(outputDir, "review.docx");
  buildDOIClickableDocx(sections, refs, title, [], docxPath);

  // Summary
  console.log("\n=== Done ===");
  console.log("Output folder: " + path.resolve(outputDir));
  console.log("Files:");
  console.log("  review.docx      — Main review with DOI hyperlinks");
  console.log("  references.ris   — EndNote RIS import");
  console.log("  references.xml   — EndNote XML import (with PDF paths)");

  return { papers, refs, figures: allFigures };
}

function formatRefText(meta, filename) {
  const authors = meta.authors.length > 0 ? meta.authors.slice(0, 3).map(a => a.replace(/, $/, "")).join(", ") + (meta.authors.length > 3 ? ", et al." : "") : "[Authors]";
  const year = meta.year || "[Year]";
  const title = meta.title || "[Title]";
  const journal = meta.journal || "[Journal]";
  const vol = meta.volume || "";
  const pages = meta.pages || "";
  return authors + " " + title + "[J]. " + journal + ", " + year + (vol ? ", " + vol : "") + (pages ? ": " + pages : "") + ".";
}
module.exports = {
  extractMetadata,
  generateEndNoteXML,
  generateFullReview,
  extractPDFs,
  detectFigures,
  buildClickableDocx,
  buildDOIClickableDocx,
  generateRIS,
  parseRef
};
