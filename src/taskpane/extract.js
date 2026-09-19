/* global JSZip, DOMParser, TextDecoder */

// Bundled in ./vendor (pinned: JSZip 3.10.1, PDF.js 6.3.289 legacy build) so no third-party CDN ever sees
// or can tamper with attachment content. Paths resolve relative to taskpane.html.
const LIBS = {
  jszip: "vendor/jszip.min.js?v=11",
  pdf: "./vendor/pdf.min.mjs?v=11",
  pdfWorker: "vendor/pdf.worker.min.mjs?v=11",
};

const TEXT_EXTS = new Set(["txt", "csv", "tsv", "md", "json", "xml", "log", "html", "htm"]);
const OOXML_EXTS = new Set(["docx", "xlsx", "pptx"]);
const MAX_INDEX_BYTES = 25 * 1024 * 1024;
const MAX_PDF_PAGES = 150;

const scriptPromises = {};
function loadScript(url) {
  if (!scriptPromises[url]) {
    scriptPromises[url] = new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = url;
      s.onload = resolve;
      s.onerror = () => reject(new Error("Could not load " + url));
      document.head.appendChild(s);
    });
  }
  return scriptPromises[url];
}

function fileExt(name) {
  const parts = name.split(".");
  return parts.length > 1 ? parts.pop().toLowerCase() : "";
}

function canExtractText(name, size) {
  if (size && size > MAX_INDEX_BYTES) return false;
  const ext = fileExt(name);
  return TEXT_EXTS.has(ext) || OOXML_EXTS.has(ext) || ext === "pdf";
}

function base64ToBytes(b64) {
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes;
}

// Joins the text runs (<w:t>, <a:t>, <t>) inside each grouping element (paragraph / shared string)
// without spaces, since Word and PowerPoint often split a single word across several runs.
function groupedXmlText(xml, groupTag) {
  const doc = new DOMParser().parseFromString(xml, "text/xml");
  const parts = [];
  const groups = doc.getElementsByTagNameNS("*", groupTag);
  for (let i = 0; i < groups.length; i++) {
    const runs = groups[i].getElementsByTagNameNS("*", "t");
    let line = "";
    for (let j = 0; j < runs.length; j++) line += runs[j].textContent;
    if (line) parts.push(line);
  }
  return parts.join(" ");
}

async function zipTexts(bytes, pathPattern, groupTag) {
  await loadScript(LIBS.jszip);
  const zip = await JSZip.loadAsync(bytes);
  const files = zip.file(pathPattern);
  const chunks = [];
  for (const f of files) chunks.push(groupedXmlText(await f.async("string"), groupTag));
  return { zip, text: chunks.join(" ") };
}

async function extractDocx(bytes) {
  const { text } = await zipTexts(bytes, /^word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$/, "p");
  return text;
}

async function extractPptx(bytes) {
  const { text } = await zipTexts(bytes, /^ppt\/(slides\/slide|notesSlides\/notesSlide)\d+\.xml$/, "p");
  return text;
}

async function extractXlsx(bytes) {
  const { zip, text } = await zipTexts(bytes, /^xl\/sharedStrings\.xml$/, "si");
  const wb = zip.file("xl/workbook.xml");
  let sheetNames = "";
  if (wb) {
    const doc = new DOMParser().parseFromString(await wb.async("string"), "text/xml");
    const sheets = doc.getElementsByTagNameNS("*", "sheet");
    sheetNames = Array.from(sheets).map((s) => s.getAttribute("name")).join(" ");
  }
  return sheetNames + " " + text;
}

let pdfjsPromise;
function loadPdfjs() {
  if (!pdfjsPromise) {
    pdfjsPromise = import(LIBS.pdf).then((lib) => {
      lib.GlobalWorkerOptions.workerSrc = LIBS.pdfWorker;
      return lib;
    });
  }
  return pdfjsPromise;
}

async function extractPdf(bytes) {
  const pdfjsLib = await loadPdfjs();
  const task = pdfjsLib.getDocument({ data: bytes, isEvalSupported: false });
  try {
    const pdf = await task.promise;
    const pages = Math.min(pdf.numPages, MAX_PDF_PAGES);
    const chunks = [];
    for (let p = 1; p <= pages; p++) {
      const page = await pdf.getPage(p);
      const content = await page.getTextContent();
      chunks.push(content.items.map((it) => it.str).join(" "));
    }
    return chunks.join(" ");
  } finally {
    await task.destroy();
  }
}

function extractHtml(bytes) {
  const doc = new DOMParser().parseFromString(new TextDecoder("utf-8").decode(bytes), "text/html");
  doc.querySelectorAll("script, style").forEach((n) => n.remove());
  return doc.body ? doc.body.textContent : "";
}

// Resolves to plain text (whitespace collapsed). Throws if the file is unreadable/corrupt.
async function extractText(name, bytes) {
  const ext = fileExt(name);
  let text;
  if (ext === "docx") text = await extractDocx(bytes);
  else if (ext === "pptx") text = await extractPptx(bytes);
  else if (ext === "xlsx") text = await extractXlsx(bytes);
  else if (ext === "pdf") text = await extractPdf(bytes);
  else if (ext === "html" || ext === "htm") text = extractHtml(bytes);
  else text = new TextDecoder("utf-8").decode(bytes);
  return text.replace(/\s+/g, " ").trim();
}
