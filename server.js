import express from "express";
import cors from "cors";
import multer from "multer";
import fs from "fs";
import path from "path";
import { PDFDocument } from "pdf-lib";
import AdmZip from "adm-zip";
import mime from "mime-types";

const app = express();
app.set("trust proxy", 1);
app.use(cors({ origin: true }));
app.options("*", cors({ origin: true }));

const upload = multer({ dest: "/tmp" });

const PORT = process.env.PORT || 8080;
const ALLOWED = (process.env.ALLOWED_EXT || "pdf,docx,pptx,xlsx,txt").split(",").map(s=>s.trim().toLowerCase());
const MAX_MB = parseInt(process.env.MAX_FILE_MB || "25", 10);
const LIMIT_PER_DAY = parseInt(process.env.RATE_LIMIT_PER_DAY || "50", 10); // on assouplit pour test

// —————————————————— utils
const usage = new Map();
function canUse(ip){
  const now = Date.now();
  let u = usage.get(ip);
  if(!u || now > u.reset){ u = { count: 0, reset: now + 24*60*60*1000 }; }
  if(u.count >= LIMIT_PER_DAY) return false;
  u.count++; usage.set(ip, u); return true;
}
function baseNameNoExt(name){
  return name.replace(/\.[^.]+$/, "");
}
function extOf(name){
  return (name.split(".").pop() || "").toLowerCase();
}

// —————————————————— cleaners
async function cleanPDF(inputBuf){
  const pdf = await PDFDocument.load(inputBuf);
  // Supprime métadonnées principales
  pdf.setTitle("");
  pdf.setAuthor("");
  pdf.setSubject("");
  pdf.setKeywords([]);
  pdf.setProducer("");
  pdf.setCreator("");
  pdf.setCreationDate(undefined);
  pdf.setModificationDate(new Date());
  // Supprime champs XMP si présents (best effort)
  // pdf-lib ne donne pas l’API directe pour tout XMP → au moins les infos ci-dessus
  const out = await pdf.save({ useObjectStreams: false }); // compat large
  return { buffer: Buffer.from(out), report: "PDF: métadonnées vidées (titre/auteur/sujet/mots-clés/créateur/producer/dates)." };
}

function cleanOOXML_removeDocProps(inputBuf){
  // Pour DOCX/PPTX/XLSX : on ouvre le zip, on retire/blank docProps/core.xml & app.xml
  const zip = new AdmZip(inputBuf);
  const entries = zip.getEntries().map(e=>e.entryName);

  // Supprimer les fichiers de propriétés si existants
  ["docProps/core.xml","docProps/app.xml","docProps/custom.xml"].forEach(p=>{
    const e = zip.getEntry(p);
    if(e) zip.deleteFile(p);
  });

  // Réécrire un core.xml minimal vide (optionnel)
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

  const out = zip.toBuffer();
  return { buffer: out, report: "Office: propriétés supprimées/neutralisées (docProps)." };
}

function cleanTXT_spaces(buf){
  let s = buf.toString("utf8");
  const beforeLen = s.length;

  // Remplacements simples visibles
  s = s.replace(/\u00A0/g, " ");           // nbsp → espace
  s = s.replace(/[ \t]{2,}/g, " ");        // espaces multiples → 1
  s = s.replace(/\n{3,}/g, "\n\n");        // >2 lignes vides → 1 ligne vide
  s = s.replace(/ *([:;!?])/g, "$1");      // espace avant : ; ! ? (FR simplifiée)
  s = s.replace(/, +/g, ", ");             // virgules espacées correctement
  s = s.replace(/\.{4,}/g, "...");         // points de suspension
  const afterLen = s.length;

  return { buffer: Buffer.from(s, "utf8"),
           report: `TXT: espaces multiples/ponctuation normalisés (Δ${beforeLen-afterLen} chars).` };
}

// —————————————————— routes
app.get("/health", (_req, res) => res.json({ ok: true }));

app.post("/api/clean", upload.single("file"), async (req, res) => {
  try {
    const ip = req.ip || "unknown";
    if(!canUse(ip)) return res.status(429).json({ error: "Daily beta limit reached." });

    if (!req.file) return res.status(400).json({ error: "No file uploaded" });
    const sizeMB = req.file.size / (1024 * 1024);
    const ext = extOf(req.file.originalname);

    if (!ALLOWED.includes(ext)) { fs.unlinkSync(req.file.path); return res.status(400).json({ error: "Unsupported file type" }); }
    if (sizeMB > MAX_MB) { fs.unlinkSync(req.file.path); return res.status(400).json({ error: "File too large" }); }

    const input = fs.readFileSync(req.file.path);
    fs.unlinkSync(req.file.path);

    let outBuf = input;
    let report = "No change.";
    if(ext === "pdf"){
      const r = await cleanPDF(input); outBuf = r.buffer; report = r.report;
    } else if (["docx","pptx","xlsx"].includes(ext)){
      const r = cleanOOXML_removeDocProps(input); outBuf = r.buffer; report = r.report;
    } else if (ext === "txt"){
      const r = cleanTXT_spaces(input); outBuf = r.buffer; report = r.report;
    }

    // Type MIME approximatif
    const ctype = mime.lookup(ext) || "application/octet-stream";
    res.setHeader("Content-Type", ctype);
    const base = baseNameNoExt(req.file.originalname);
    // On glisse un petit rapport en header pour l’instant
    res.setHeader("X-DocSafe-Report", report);
    res.setHeader("Content-Disposition", `attachment; filename="${base}_cleaned.${ext}"`);
    return res.send(outBuf);
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: "Server error" });
  }
});

app.listen(PORT, () => console.log(`DocSafe API V2 listening on ${PORT}`));
