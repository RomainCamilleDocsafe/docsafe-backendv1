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

/* ========= Config ========= */
const PORT = process.env.PORT || 8080;
const ALLOWED = (process.env.ALLOWED_EXT || "pdf,docx,pptx,xlsx,txt")
  .split(",").map(s => s.trim().toLowerCase());
const MAX_MB = parseInt(process.env.MAX_FILE_MB || "25", 10);
const LIMIT_PER_DAY = parseInt(process.env.RATE_LIMIT_PER_DAY || "20", 10);

/* ========= Limiteur simple (bêta) ========= */
const usage = new Map();
function canUse(ip) {
  const now = Date.now();
  let u = usage.get(ip);
  if (!u || now > u.reset) u = { count: 0, reset: now + 24 * 60 * 60 * 1000 };
  if (u.count >= LIMIT_PER_DAY) return false;
  u.count++; usage.set(ip, u); return true;
}

/* ========= Helpers ========= */
const extOf = n => (n.split(".").pop() || "").toLowerCase();
const baseNoExt = n => n.replace(/\.[^.]+$/, "");

/* ========= PDF : neutraliser métadonnées ========= */
async function cleanPDF(inputBuf) {
  const pdf = await PDFDocument.load(inputBuf, { updateMetadata: true });
  try { pdf.setTitle(""); } catch {}
  try { pdf.setAuthor(""); } catch {}
  try { pdf.setSubject(""); } catch {}
  try { pdf.setKeywords([]); } catch {}
  try { pdf.setCreator(""); } catch {}
  try { pdf.setProducer(""); } catch {}
  try { pdf.setCreationDate(new Date(0)); } catch {}
  try { pdf.setModificationDate(new Date()); } catch {}
  const out = await pdf.save({ useObjectStreams: false });
  return {
    buffer: Buffer.from(out),
    report: "PDF: métadonnées neutralisées (titre, auteur, sujet, mots-clés, créateur, producteur, dates)."
  };
}

/* ========= OOXML (PPTX/XLSX) : supprimer docProps ========= */
function cleanOOXML_removeDocProps(inputBuf) {
  const zip = new AdmZip(inputBuf);
  ["docProps/core.xml", "docProps/app.xml", "docProps/custom.xml"].forEach(p => {
    const e = zip.getEntry(p);
    if (e) zip.deleteFile(p);
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
  return { buffer: zip.toBuffer(), report: "Office: docProps supprimés/neutralisés." };
}

/* ========= DOCX : docProps + correction simple du contenu ========= */
function cleanDOCX_contentAndProps(inputBuf) {
  const zip = new AdmZip(inputBuf);

  // 1) Métadonnées
  ["docProps/core.xml", "docProps/app.xml", "docProps/custom.xml"].forEach(p => {
    const e = zip.getEntry(p);
    if (e) zip.deleteFile(p);
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

  // 2) Fichiers XML Word à traiter
  const targets = new Set([
    "word/document.xml",
    "word/footnotes.xml",
    "word/endnotes.xml",
    "word/comments.xml"
  ]);
  zip.getEntries().forEach(e => {
    if (e.entryName.startsWith("word/header") && e.entryName.endsWith(".xml")) targets.add(e.entryName);
    if (e.entryName.startsWith("word/footer") && e.entryName.endsWith(".xml")) targets.add(e.entryName);
  });

  let changedBlocks = 0;

  // Normalisation très sûre (pas de grammaire externe)
  const fixInline = (s) => {
    const before = s;
    let t = s;

    // Espaces insécables -> espaces
    t = t.replace(/\u00A0/g, " ");
    // Plusieurs espaces -> un
    t = t.replace(/[ \t]{2,}/g, " ");
    // Avant ponctuation forte : retirer espace
    t = t.replace(/ *([:;!?])/g, "$1");
    // Espace avant point/virgule
    t = t.replace(/ \./g, ".");
    t = t.replace(/ ,/g, ",");
    // Points de suspension normalisés
    t = t.replace(/\.{4,}/g, "...");

    if (t !== before) changedBlocks++;
    return t;
  };

  // Remplacer le texte dans <w:t>…</w:t> uniquement
  for (const path of targets) {
    const entry = zip.getEntry(path);
    if (!entry) continue;
    const xml = entry.getData().toString("utf8");

    const out = [];
    let last = 0;
    const re = /(<w:t\b[^>]*>)([\s\S]*?)(<\/w:t>)/g;
    let m;
    while ((m = re.exec(xml)) !== null) {
      const [full, open, inner, close] = m;
      out.push(xml.slice(last, m.index));
      out.push(open + fixInline(inner) + close);
      last = m.index + full.length;
    }
    out.push(xml.slice(last));
    const fixed = out.join("");

    if (fixed !== xml) {
      zip.updateFile(path, Buffer.from(fixed, "utf8"));
    }
  }

  return {
    buffer: zip.toBuffer(),
    report: `DOCX: docProps neutralisés + correction simple (espaces/ponctuation) sur contenu. Sections modifiées: ${changedBlocks}.`
  };
}

/* ========= TXT : normalisation ========= */
function cleanTXT_spaces(buf) {
  let s = buf.toString("utf8");
  const before = s.length;
  s = s.replace(/\u00A0/g, " ");
  s = s.replace(/[ \t]{2,}/g, " ");
  s = s.replace(/\n{3,}/g, "\n\n");
  s = s.replace(/ *([:;!?])/g, "$1");
  s = s.replace(/ \./g, ".");
  s = s.replace(/ ,/g, ",");
  s = s.replace(/\.{4,}/g, "...");

  const after = s.length;
  return { buffer: Buffer.from(s, "utf8"), report: `TXT: normalisation (Δ${before - after} chars).` };
}

/* ========= Routes ========= */
app.get("/health", (_req, res) => res.json({ ok: true, message: "Backend is running ✅" }));

app.get("/", (_req, res) => {
  res.type("text/plain").send("DocSafe API (beta). Try GET /health or POST /api/clean (multipart/form-data, field 'file').");
});

app.post("/api/clean", upload.single("file"), async (req, res) => {
  try {
    const ip = req.ip || "unknown";
    if (!canUse(ip)) return res.status(429).json({ error: "Daily beta limit reached (20 docs)." });
    if (!req.file) return res.status(400).json({ error: "No file uploaded" });

    const sizeMB = req.file.size / (1024 * 1024);
    const ext = extOf(req.file.originalname);

    if (!ALLOWED.includes(ext)) {
      try { fs.unlinkSync(req.file.path); } catch {}
      return res.status(400).json({ error: "Unsupported file type" });
    }
    if (sizeMB > MAX_MB) {
      try { fs.unlinkSync(req.file.path); } catch {}
      return res.status(400).json({ error: "File too large" });
    }

    const input = fs.readFileSync(req.file.path);
    try { fs.unlinkSync(req.file.path); } catch {}

    let outBuf = input, report = "No change";
    if (ext === "pdf") {
      ({ buffer: outBuf, report } = await cleanPDF(input));
    } else if (ext === "docx") {
      ({ buffer: outBuf, report } = cleanDOCX_contentAndProps(input));
    } else if (["pptx", "xlsx"].includes(ext)) {
      ({ buffer: outBuf, report } = cleanOOXML_removeDocProps(input));
    } else if (ext === "txt") {
      ({ buffer: outBuf, report } = cleanTXT_spaces(input));
    } else {
      ({ buffer: outBuf, report } = cleanOOXML_removeDocProps(input));
    }

    const ctype = mimeLookup(ext) || "application/octet-stream";
    res.setHeader("Content-Type", ctype);
    res.setHeader("X-DocSafe-Report", report);
    const base = baseNoExt(req.file.originalname);
    res.setHeader("Content-Disposition", `attachment; filename="${base}_cleaned.${ext}"`);
    res.send(outBuf);
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: "Server error" });
  }
});

/* ========= Boot ========= */
app.listen(PORT, () => console.log(`DocSafe API Beta V1 listening on ${PORT}`));
