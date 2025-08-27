import express from "express";
import cors from "cors";
import multer from "multer";
import fs from "fs";
import { PDFDocument } from "pdf-lib";
import AdmZip from "adm-zip";
import { lookup as mimeLookup } from "mime-types";

const app = express();
app.set("trust proxy", 1);
app.use(cors({ origin: true }));
app.options("*", cors({ origin: true }));

const upload = multer({ dest: "/tmp" });

const PORT = process.env.PORT || 8080;
const ALLOWED = (process.env.ALLOWED_EXT || "pdf,docx,pptx,xlsx,txt")
  .split(",").map(s=>s.trim().toLowerCase());
const MAX_MB = parseInt(process.env.MAX_FILE_MB || "25", 10);
const LIMIT_PER_DAY = parseInt(process.env.RATE_LIMIT_PER_DAY || "50", 10);

// ——— rate limit naïf (mémoire)
const usage = new Map();
function canUse(ip){
  const now = Date.now();
  let u = usage.get(ip);
  if(!u || now > u.reset){ u = { count: 0, reset: now + 24*60*60*1000 }; }
  if(u.count >= LIMIT_PER_DAY) return false;
  u.count++; usage.set(ip, u); return true;
}
const extOf = n => (n.split(".").pop() || "").toLowerCase();
const baseNoExt = n => n.replace(/\.[^.]+$/, "");

// ——— cleaners
async function cleanPDF(inputBuf){
  const pdf = await PDFDocument.load(inputBuf, { updateMetadata: true });

  // Vider métadonnées classiques (sans passer undefined)
  try { pdf.setTitle(""); } catch {}
  try { pdf.setAuthor(""); } catch {}
  try { pdf.setSubject(""); } catch {}
  try { pdf.setKeywords([]); } catch {}
  try { pdf.setProducer(""); } catch {}
  try { pdf.setCreator(""); } catch {}
  // Dates : on évite undefined → on met une date neutre/présente
  try { pdf.setCreationDate(new Date(0)); } catch {}
  try { pdf.setModificationDate(new Date()); } catch {}

  const out = await pdf.save({ useObjectStreams: false });
  return {
    buffer: Buffer.from(out),
    report: "PDF: métadonnées neutralisées (titre/auteur/sujet/mots-clés/créateur/producer/dates)."
  };
}

function cleanOOXML_removeDocProps(inputBuf){
  const zip = new AdmZip(inputBuf);
  // Supprime les props si présentes
  ["docProps/core.xml","docProps/app.xml","docProps/custom.xml"].forEach(p=>{
    const e = zip.getEntry(p);
    if(e) zip.deleteFile(p);
  });
  // Ajoute un core.xml minimal (révision 1)
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

  return {
    buffer: zip.toBuffer(),
    report: "Office: propriétés supprimées/neutralisées (docProps)."
  };
}

function cleanTXT_spaces(buf){
  let s = buf.toString("utf8");
  const before = s.length;
  s = s.replace(/\u00A0/g, " ");        // nbsp → espace
  s = s.replace(/[ \t]{2,}/g, " ");     // espaces multiples → 1
  s = s.replace(/\n{3,}/g, "\n\n");     // >2 sauts → 1
  s = s.replace(/ *([:;!?])/g, "$1");   // espace avant : ; ! ?
  s = s.replace(/, +/g, ", ");          // espace après virgule
  s = s.replace(/\.{4,}/g, "...");      // … nettoyés
  const after = s.length;
  return {
    buffer: Buffer.from(s, "utf8"),
    report: `TXT: espaces/ponctuation normalisés (Δ${before-after} chars).`
  };
}

// ——— routes
app.get("/health", (_req, res) => res.json({ ok: true }));

app.post("/api/clean", upload.single("file"), async (req, res) => {
  try {
    const ip = req.ip || "unknown";
    if(!canUse(ip)) return res.status(429).json({ error: "Daily beta limit reached." });

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
    } else if (["docx","pptx","xlsx"].includes(ext)){
      ({ buffer: outBuf, report } = cleanOOXML_removeDocProps(input));
    } else if (ext === "txt"){
      ({ buffer: outBuf, report } = cleanTXT_spaces(input));
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

app.listen(PORT, () => console.log(`DocSafe API V2 listening on ${PORT}`));
