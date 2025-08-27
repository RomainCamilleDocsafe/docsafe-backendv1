import express from "express";
import cors from "cors";
import multer from "multer";
import fs from "fs";

const app = express();
app.set("trust proxy", 1);
app.use(cors({ origin: true }));
app.options("*", cors({ origin: true }));

const upload = multer({ dest: "/tmp" });

const PORT = process.env.PORT || 8080;
const ALLOWED = (process.env.ALLOWED_EXT || "pdf,docx,pptx,xlsx").split(",").map(s=>s.trim().toLowerCase());
const MAX_MB = parseInt(process.env.MAX_FILE_MB || "25", 10);

app.get("/health", (_req, res) => res.json({ ok: true }));

app.post("/api/clean", upload.single("file"), (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "No file uploaded" });

    const sizeMB = req.file.size / (1024 * 1024);
    const ext = (req.file.originalname.split(".").pop() || "").toLowerCase();
    if (!ALLOWED.includes(ext)) { fs.unlinkSync(req.file.path); return res.status(400).json({ error: "Unsupported file type" }); }
    if (sizeMB > MAX_MB) { fs.unlinkSync(req.file.path); return res.status(400).json({ error: "File too large" }); }

    const buf = fs.readFileSync(req.file.path);
    fs.unlinkSync(req.file.path);

    res.setHeader("Content-Type", "application/octet-stream");
    const base = req.file.originalname.replace(/\.[^.]+$/, "");
    res.setHeader("Content-Disposition", `attachment; filename="${base}_cleaned.${ext}"`);
    return res.send(buf);
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: "Server error" });
  }
});

app.listen(PORT, () => console.log(`DocSafe API listening on ${PORT}`));
