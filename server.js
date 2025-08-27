import express from "express";
import cors from "cors";
import multer from "multer";
import fs from "fs";
import { PDFDocument } from "pdf-lib";
import AdmZip from "adm-zip";
import { lookup as mimeLookup } from "mime-types";
import axios from "axios";

const app = express();
app.set("trust proxy", 1);

// CORS permissif pour la bêta gratuite
app.use(cors({ origin: true }));
app.options("*", cors({ origin: true }));

const upload = multer({ dest: "/tmp" });

/* ===== Config ===== */
const PORT = process.env.PORT || 8080;
const ALLOWED = (process.env.ALLOWED_EXT || "pdf,docx,pptx,xlsx,txt")
  .split(",").map(s=>s.trim().toLowerCase());
const MAX_MB = parseInt(process.env.MAX_FILE_MB || "25", 10);
const LIMIT_PER_DAY = parseInt(process.env.RATE_LIMIT_PER_DAY || "20", 10);

/* ===== Limiteur naïf (bêta) ===== */
const usage = new Map();
function canUse(ip){
  const now = Date.now();
  let u = usage.get(ip);
  if(!u || now > u.reset){ u = { count: 0, reset: now + 24*60*60*1000 }; }
  if(u.count >= LIMIT_PER_DAY) return false;
  u.count++; usage.set(ip, u); return true;
}

/* ===== Utils ===== */
const extOf = n => (n.split(".").pop() || "").toLowerCase();
const baseNoExt = n => n.replace(/\.[^.]+$/, "");

/* ===== Correcteur simple + LanguageTool (auto FR/EN) ===== */
async function correctWithLanguageTool(text, lang="auto") {
  try {
    if(!text || !text.trim()) return text;
    const res = await axios.post("https://api.languagetool.org/v2/check", null, {
      params: { text, language: lang }
    });
    let corrected = text;
    let shift = 0;
    for (const m of res.data.matches || []) {
      if (!m.replacements || m.replacements.length === 0) continue;
      const repl = m.replacements[0].value ?? "";
      const start = m.offset + shift;
      const end = start + m.length;
      corrected = corrected.slice(0, start) + repl + corrected.slice(end);
      shift += repl.length - m.length;
    }
    return corrected;
  } catch (e) {
    console.error("LanguageTool error:", e?.response?.status || e.message);
    return text; // fallback si rate limit
  }
}

/* ===== PDF : neutralisation métadonnées ===== */
async function cleanPDF(inputBuf){
  const pdf = await PDFDocument.load(inputBuf, { updateMetadata: true });
  try { pdf.setTitle(""); } catch {}
  try { pdf.setAuthor(""); } catch {}
  try { pdf.setSubject(""); } catch {}
  try { pdf.setKeywords([]); } catch {}
  try { pdf.setProducer(""); } catch {}
  try { pdf.setCreator(""); } catch {}
  try { pdf.setCreationDate(new Date(0)); } catch {}
  try { pdf.setModificationDate(new Date()); } catch {}
  const out = await pdf.save({ useObjectStreams: false });
  return {
    buffer: Buffer.from(out),
    report: "PDF: métadonnées neutralisées (titre/auteur/sujet/mots-clés/créateur/producer/dates)."
  };
}

/* ===== OOXML (PPTX/XLSX) : suppression docProps ===== */
function cleanOOXML_removeDocProps(inputBuf){
  const zip = new AdmZip(inputBuf);
  ["docProps/core.xml","docProps/app.xml","docProps/custom.xml"].forEach(p=>{
    const e = zip.getEntry(p);
    if(e) zip.deleteFile(p);
  });
  const coreXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
 xmlns:dc="http://purl.org/dc/elements/1.1/"
 xmlns:dcterms="http://purl.org/dc/terms/"
 xmlns:dcmitype="http://purl.org/dc/dcmitype/"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title></dc:title>
  <dc:creator></dc:creator>
  <cp:lastModifiedBy></cp:lastModifiedBy>
  <cp:revision>1</cp:revision>
</cp:coreProperties>`;
  zip.addFile("docProps/core.xml", Buffer.from(coreXml, "utf8"));
  return { buffer: zip.toBuffer(), report: "Office: propriétés supprimées/neutralisées (docProps)." };
}

/* ===== DOCX : docProps + correction <w:t> ===== */
async function cleanDOCX_contentAndProps(inputBuf){
  const zip = new AdmZip(inputBuf);
  let changedChars = 0;
  let ltCorrections = 0;

  // Métadonnées
  ["docProps/core.xml","docProps/app.xml","docProps/custom.xml"].forEach(p=>{
    const e = zip.getEntry(p);
    if(e) zip.deleteFile(p);
  });
  const coreXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
 xmlns:dc="http://purl.org/dc/elements/1.1/"
 xmlns:dcterms="http://purl.org/dc/terms/"
 xmlns:dcmitype="http://purl.org/dc/dcmitype/"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title></dc:title>
  <dc:creator></dc:creator>
  <cp:lastModifiedBy></cp:lastModifiedBy>
  <cp:revision>1</cp:revision>
</cp:coreProperties>`;
  zip.addFile("docProps/core.xml", Buffer.from(coreXml, "utf8"));

  // Fichiers texte Word à traiter
  const targets = ["word/document.xml","word/footnotes.xml","word/endnotes.xml","word/comments.xml"];
  zip.getEntries().forEach(e=>{
    if (e.entryName.startsWith("word/header") && e.entryName.endsWith(".xml")) targets.push(e.entryName);
    if (e.entryName.startsWith("word/footer") && e.entryName.endsWith(".xml")) targets.push(e.entryName);
  });

  // Normalisation & grammaire
  const fixTypo = async (text) => {
    const before = text.length;
    let t = text;
    t = t.replace(/\u00A0/g, " ");      // nbsp -> espace
    t = t.replace(/[ \t]{2,}/g, " ");   // espaces multiples
    t = t.replace(/\n{3,}/g, "\n\n");   // trop de sauts de ligne
    t = t.replace(/ *([:;!?])/g, "$1");// pas d’espace avant :;!?
    t = t.replace(/, +/g, ", ");        // espace après virgule
    t = t.replace(/\.{4,}/g, "...");    // …. -> …
    const t2 = await correctWithLanguageTool(t, "auto"); // FR/EN
    if (t2 !== t) ltCorrections++;
    const after = t2.length;
    changedChars += Math.max(0, before - after);
    return t2;
  };

  for (const path of targets) {
    const entry = zip.getEntry(path);
    if (!entry) continue;
    const xml = entry.getData().toString("utf8");

    const fixed = await (async () => {
      const parts = [];
      let lastIndex = 0;
      const regex = /(<w:t\b[^>]*>)([\s\S]*?)(<\/w:t>)/g;
      let m;
      while ((m = regex.exec(xml)) !== null) {
        const [full, open, inner, close] = m;
        parts.push(xml.slice(lastIndex, m.index));
        const cleaned = await fixTypo(inner);
        parts.push(open + cleaned + close);
        lastIndex = m.index + full.length;
      }
      parts.push(xml.slice(lastIndex));
      return parts.join("");
    })();

    if (fixed !== xml) {
      zip.updateFile(path, Buffer.from(fixed, "utf8"));
    }
  }

  return {
    buffer: zip.toBuffer(),
    report: `DOCX: métadonnées neutralisées + contenu corrigé (espaces/ponctuation/grammaire), corrections LT: ${ltCorrections}, Δ${changedChars} chars.`
  };
}

/* ===== TXT : normalisation + LT ===== */
async function cleanTXT_spaces(buf){
  let s = buf.toString("utf8");
  const before = s.length;
  s = s.replace(/\u00A0/g, " ");
  s = s.replace(/[ \t]{2,}/g, " ");
  s = s.replace(/\n{3,}/g, "\n\n");
  s = s.replace(/ *([:;!?])/g, "$1");
  s = s.replace(/, +/g, ", ");
  s = s.replace(/\.{4,}/g, "...");
  const corrected = await correctWithLanguageTool(s, "auto");
  const after = corrected.length;
  return { buffer: Buffer.from(corrected, "utf8"), report: `TXT: normalisation + corrections LT (Δ${before-after} chars).` };
}

/* ===== Routes ===== */
app.get("/health", (_req, res) => res.json({ ok: true, message: "Backend is running ✅" }));

app.get("/", (_req, res) => {
  res.type("text/plain").send("DocSafe API (beta). Try GET /health or POST /api/clean (multipart/form-data, field 'file').");
});

app.post("/api/clean", upload.single("file"), async (req, res) => {
  try {
    const ip = req.ip || "unknown";
    if(!canUse(ip)) return res.status(429).json({ error: "Daily beta limit reached (20 docs)." });
    if (!req.file) return res.status(400).json({ error: "No file uploaded" });

    const sizeMB = req.file.size / (1024 * 1024);
    const ext = extOf(req.file.originalname);
    if (!ALLOWED.includes(ext)){
      try{ fs.unlinkSync(req.file.path); }catch{}
      return res.status(400).json({ error: "Unsupported file type" });
    }
    if (sizeMB > MAX_MB){
      try{ fs.unlinkSync(req.file.path); }catch{}
      return res.status(400).json({ error: "File too large" });
    }

    const input = fs.readFileSync(req.file.path);
    try{ fs.unlinkSync(req.file.path); }catch{}

    let outBuf = input, report = "No change.";
    if(ext === "pdf"){
      ({ buffer: outBuf, report } = await cleanPDF(input));
    } else if (ext === "docx"){
      ({ buffer: outBuf, report } = await cleanDOCX_contentAndProps(input));
    } else if (["pptx","xlsx"].includes(ext)){
      ({ buffer: outBuf, report } = cleanOOXML_removeDocProps(input));
    } else if (ext === "txt"){
      ({ buffer: outBuf, report } = await cleanTXT_spaces(input));
    } else {
      ({ buffer: outBuf, report } = cleanOOXML_removeDocProps(input));
    }

    const ctype = mimeLookup(ext) || "application/octet-stream";
    res.setHeader("Content-Type", ctype);
    res.setHeader("X-DocSafe-Report", report);
    const base = baseNoExt(req.file.originalname);
    res.setHeader("Content-Disposition", `attachment; filename="${base}_cleaned.${ext}"`);
    return res.send(outBuf);
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: "Server error" });
  }
});

/* ===== Boot ===== */
app.listen(PORT, () => console.log(`DocSafe API Beta V1 listening on ${PORT}`));

