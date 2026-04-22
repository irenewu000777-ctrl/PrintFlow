import express from "express";
import multer from "multer";
import { randomUUID } from "crypto";
import { spawn } from "child_process";
import { promises as fs } from "fs";
import os from "os";
import path from "path";

const app = express();
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 }
});

const PORT = process.env.PORT || 8080;
const API_KEY = process.env.CONVERTER_API_KEY || "";
const CORS_ORIGIN = process.env.CORS_ORIGIN || "*";

app.use((req, res, next) => {
  res.setHeader("Access-Control-Allow-Origin", CORS_ORIGIN);
  res.setHeader("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type,Authorization");
  if (req.method === "OPTIONS") {
    res.status(204).end();
    return;
  }
  next();
});

function requireAuth(req, res, next) {
  if (!API_KEY) return next();
  const auth = req.header("authorization") || "";
  const expected = `Bearer ${API_KEY}`;
  if (auth !== expected) {
    res.status(401).send("Unauthorized");
    return;
  }
  next();
}

function runSoffice(inputPath, outputDir) {
  return new Promise((resolve, reject) => {
    const cmd = process.platform === "win32" ? "soffice.exe" : "soffice";
    const child = spawn(cmd, [
      "--headless",
      "--nologo",
      "--nofirststartwizard",
      "--convert-to",
      "pdf",
      "--outdir",
      outputDir,
      inputPath
    ]);

    let stderr = "";
    child.stderr.on("data", (d) => {
      stderr += String(d);
    });
    child.on("error", reject);
    child.on("close", (code) => {
      if (code === 0) {
        resolve();
        return;
      }
      reject(new Error(stderr || `soffice exit code ${code}`));
    });
  });
}

app.get("/health", (_req, res) => {
  res.json({ ok: true });
});

app.post("/convert", requireAuth, upload.single("file"), async (req, res) => {
  try {
    if (!req.file) {
      res.status(400).send("Missing file");
      return;
    }

    const original = req.file.originalname.toLowerCase();
    if (!original.endsWith(".ppt") && !original.endsWith(".pptx")) {
      res.status(400).send("Only .ppt/.pptx supported");
      return;
    }

    const workDir = await fs.mkdtemp(path.join(os.tmpdir(), "pptx-convert-"));
    try {
      const ext = original.endsWith(".pptx") ? ".pptx" : ".ppt";
      const base = randomUUID();
      const inputPath = path.join(workDir, `${base}${ext}`);
      const outputPath = path.join(workDir, `${base}.pdf`);

      await fs.writeFile(inputPath, req.file.buffer);
      await runSoffice(inputPath, workDir);
      const pdf = await fs.readFile(outputPath);

      res.setHeader("Content-Type", "application/pdf");
      res.status(200).send(pdf);
    } finally {
      await fs.rm(workDir, { recursive: true, force: true });
    }
  } catch (err) {
    const msg = err instanceof Error ? err.message : "convert failed";
    res.status(500).send(msg);
  }
});

app.listen(PORT, () => {
  console.log(`converter-service listening on :${PORT}`);
});
