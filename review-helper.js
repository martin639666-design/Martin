// literature-review helper: extract text from PDFs and build docx
// Usage: node review-helper.js <pdf-folder> <output-path>

const fs = require("fs");
const path = require("path");

// ===== PDF Text Extraction =====
async function extractPDFs(folderPath) {
  const files = fs.readdirSync(folderPath).filter(f => f.toLowerCase().endsWith(".pdf"));
  const results = [];

  try {
    const pdfjsLib = require("pdfjs-dist/legacy/build/pdf.mjs");
    for (const file of files) {
      const fullPath = path.join(folderPath, file);
      console.log("Reading:", file);
      try {
        const loadingTask = pdfjsLib.getDocument(fullPath);
        const pdf = await loadingTask.promise;
        let text = "";
        const maxPages = Math.min(pdf.numPages, 20);
        for (let i = 1; i <= maxPages; i++) {
          const page = await pdf.getPage(i);
          const content = await page.getTextContent();
          text += "---Page " + i + "---\n" + content.items.map(item => item.str).join(" ") + "\n\n";
        }
        results.push({ file, pages: pdf.numPages, text });
        console.log("  OK: " + pdf.numPages + " pages, " + text.length + " chars");
      } catch (e) {
        console.log("  FAIL: " + e.message);
      }
    }
  } catch (e) {
    console.log("pdfjs-dist not available, trying pdf-parse...");
    try {
      const pdfParse = require("pdf-parse");
      for (const file of files) {
        const fullPath = path.join(folderPath, file);
        const buf = fs.readFileSync(fullPath);
        try {
          const data = await pdfParse(buf);
          results.push({ file, pages: data.numpages, text: data.text });
          console.log("  OK: " + file + " (" + data.numpages + " pages)");
        } catch (e2) {
          console.log("  FAIL: " + file + " - " + e2.message);
        }
      }
    } catch (e2) {
      console.log("No PDF library available. Install with: npm install pdfjs-dist");
    }
  }

  return results;
}

// ===== DOCX Builder (ZIP + XML) =====
function buildDocx(docXml, stylesXml, outputPath) {
  const contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>';
  const rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>';
  const wordRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>';

  const files = [
    { name: "[Content_Types].xml", data: contentTypes },
    { name: "_rels/.rels", data: rels },
    { name: "word/document.xml", data: docXml },
    { name: "word/styles.xml", data: stylesXml },
    { name: "word/_rels/document.xml.rels", data: wordRels }
  ];

  // Build ZIP manually
  const zipData = buildZip(files);
  fs.writeFileSync(outputPath, zipData, { encoding: null });
  const size = fs.statSync(outputPath).size;
  console.log("DOCX created: " + outputPath + " (" + (size / 1024).toFixed(0) + " KB)");
  return size;
}

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
    const nb = Buffer.from(f.name, "utf8"); const db = Buffer.from(f.data, "utf8"); const crc = crc32(db);
    const local = Buffer.alloc(30 + nb.length);
    local.writeUInt32LE(0x04034b50, 0); local.writeUInt16LE(20, 4); local.writeUInt16LE(0, 8);
    local.writeUInt32LE(crc, 14); local.writeUInt32LE(db.length, 18); local.writeUInt32LE(db.length, 22);
    local.writeUInt16LE(nb.length, 26); nb.copy(local, 30);
    lh.push({ h: local, d: db, o: offset }); offset += local.length + db.length;
    const cent = Buffer.alloc(46 + nb.length);
    cent.writeUInt32LE(0x02014b50, 0); cent.writeUInt16LE(20, 4); cent.writeUInt16LE(20, 6);
    cent.writeUInt32LE(crc, 16); cent.writeUInt32LE(db.length, 20); cent.writeUInt32LE(db.length, 24);
    cent.writeUInt16LE(nb.length, 28); cent.writeUInt32LE(lh[lh.length - 1].o, 42); nb.copy(cent, 46);
    ch.push(cent);
  }
  const co = offset; const cs = ch.reduce((s, h) => s + h.length, 0);
  const e = Buffer.alloc(22); e.writeUInt32LE(0x06054b50, 0);
  e.writeUInt16LE(ch.length, 8); e.writeUInt16LE(ch.length, 10); e.writeUInt32LE(cs, 12); e.writeUInt32LE(co, 16);
  const total = co + cs + 22; const result = Buffer.alloc(total); let pos = 0;
  for (const x of lh) { x.h.copy(result, pos); pos += x.h.length; x.d.copy(result, pos); pos += x.d.length; }
  for (const x of ch) { x.copy(result, pos); pos += x.length; }
  e.copy(result, pos); return result;
}

// ===== Default Styles =====
function getDefaultStyles() {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/><w:pPr><w:spacing w:line="360" w:lineRule="auto" w:after="120"/></w:pPr><w:rPr><w:rFonts w:ascii="宋体" w:hAnsi="Times New Roman" w:eastAsia="宋体"/><w:sz w:val="24"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="240" w:after="120"/><w:keepNext/><w:keepLines/></w:pPr><w:rPr><w:rFonts w:ascii="黑体" w:hAnsi="Times New Roman" w:eastAsia="黑体"/><w:b/><w:sz w:val="28"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="200" w:after="100"/><w:keepNext/><w:keepLines/></w:pPr><w:rPr><w:rFonts w:ascii="黑体" w:hAnsi="Times New Roman" w:eastAsia="黑体"/><w:b/><w:sz w:val="26"/></w:rPr></w:style></w:styles>';
}

// ===== XML Builder Helpers =====
function esc(s) { return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); }

function makeDefaultDocXml(sections, refs, title) {
  let body = "";
  const hei = "黑体";
  const song = "宋体";

  function addP(text, extra) {
    body += "<w:p>" + (extra || "") + "<w:r><w:t xml:space=\"preserve\">" + esc(text) + "</w:t></w:r></w:p>\n";
  }

  function h1(t) { addP(t, '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="' + hei + '" w:hAnsi="Times New Roman" w:eastAsia="' + hei + '"/><w:b/><w:sz w:val="28"/></w:rPr>'); }
  function h2(t) { addP(t, '<w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="' + hei + '" w:hAnsi="Times New Roman" w:eastAsia="' + hei + '"/><w:b/><w:sz w:val="26"/></w:rPr>'); }
  function p(t) { addP(t, '<w:r><w:rPr><w:rFonts w:ascii="' + song + '" w:hAnsi="Times New Roman" w:eastAsia="' + song + '"/><w:sz w:val="24"/></w:rPr>'); }

  // Title
  if (title) {
    addP(title, '<w:pPr><w:jc w:val="center"/><w:spacing w:after="200"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="' + hei + '" w:hAnsi="Times New Roman" w:eastAsia="' + hei + '"/><w:b/><w:sz w:val="36"/></w:rPr>');
  }

  // Sections
  for (const sec of sections) {
    if (sec.type === "h1") h1(sec.text);
    else if (sec.type === "h2") h2(sec.text);
    else p(sec.text);
  }

  // References
  h1("参考文献");
  for (const r of refs) {
    body += '<w:p><w:pPr><w:ind w:hanging="480" w:left="480"/><w:spacing w:line="300" w:lineRule="auto" w:after="60"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="21"/></w:rPr><w:t xml:space="preserve">' + esc(r) + '</w:t></w:r></w:p>\n';
  }

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + body + '</w:body></w:document>';
}

module.exports = { extractPDFs, buildDocx, getDefaultStyles, makeDefaultDocXml, buildZip };
