/**
 * Express server with file upload and translation routes
 */

import "dotenv/config";
import express, { Request, Response } from "express";
import multer from "multer";
import path from "path";
import fs from "fs";
import { v4 as uuidv4 } from "uuid";

import {
  createJob,
  getJob,
  cancelJob,
  updateJob,
  finishJob,
  getElapsedSeconds,
  JobState,
} from "./jobs";
import { convertPdfToDocx } from "./services/adobe";
import { parseDocx, writeDocx } from "./services/docx";
import { translateSegments, qaAndRetranslate } from "./services/translator";
import { analyzeCbAndGenerateCnsDocx } from "./services/cb-to-cns";

const app = express();
const PORT = process.env.PORT || 3000;

// Ensure directories exist
const UPLOAD_DIR = path.join(__dirname, "..", "uploads");
const WORK_DIR = path.join(__dirname, "..", "work");
const OUTPUT_DIR = path.join(__dirname, "..", "output");

[UPLOAD_DIR, WORK_DIR, OUTPUT_DIR].forEach((dir) => {
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
});

// Configure multer for file uploads
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, UPLOAD_DIR);
  },
  filename: (req, file, cb) => {
    const uniqueName = `${Date.now()}-${uuidv4()}${path.extname(file.originalname)}`;
    cb(null, uniqueName);
  },
});

const upload = multer({
  storage,
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (ext === ".pdf" || ext === ".docx") {
      cb(null, true);
    } else {
      cb(new Error("Only .pdf and .docx files are allowed"));
    }
  },
  limits: {
    fileSize: 50 * 1024 * 1024, // 50MB limit
  },
});

// Serve static frontend
app.use(express.static(path.join(__dirname, "..", "public")));
app.use(express.json());

/**
 * Process a job in the background
 * @param mode - "translate" (default) or "cb-to-cns"
 */
async function processJob(
  job: JobState,
  uploadedFilePath: string,
  mode: "translate" | "cb-to-cns" = "translate"
): Promise<void> {
  const ext = path.extname(job.fileName).toLowerCase();
  const baseName = path.basename(job.fileName, ext);

  try {
    job.startedAt = Date.now();
    let workingDocxPath: string;

    // CB → CNS mode: analyze PDF and generate CNS DOCX first
    if (mode === "cb-to-cns") {
      if (ext !== ".pdf") {
        throw new Error("CB → CNS 模式只支援 PDF 檔案");
      }

      updateJob(job, {
        status: "converting",
        stepMessage: "正在分析 CB 報告並生成 CNS 結構...",
        progress: 2,
      });

      workingDocxPath = path.join(WORK_DIR, `${uuidv4()}-cns.docx`);
      const { pageRanges } = await analyzeCbAndGenerateCnsDocx(
        job,
        uploadedFilePath,
        workingDocxPath
      );

      console.log("CB analysis complete. Page ranges:", pageRanges);

      if (job.cancelled) {
        console.log("Job cancelled after CB analysis");
        return;
      }

      // Continue with translation pipeline below...
    }
    // Standard translate mode: Convert PDF to DOCX if needed
    else if (ext === ".pdf") {
      updateJob(job, {
        status: "converting",
        stepMessage: "正在將 PDF 轉換為 DOCX...",
        progress: 5,
      });

      workingDocxPath = path.join(WORK_DIR, `${uuidv4()}.docx`);
      await convertPdfToDocx(uploadedFilePath, workingDocxPath);

      if (job.cancelled) {
        console.log("Job cancelled after PDF conversion");
        return;
      }
    } else {
      // DOCX file - copy to work directory
      workingDocxPath = path.join(WORK_DIR, `${uuidv4()}.docx`);
      fs.copyFileSync(uploadedFilePath, workingDocxPath);
    }

    // Step 2: Parse DOCX
    updateJob(job, {
      status: "parsing-docx",
      stepMessage: "正在解析 DOCX 文件...",
      progress: 15,
    });

    const parsed = await parseDocx(workingDocxPath);
    job.totalSegments = parsed.segments.length;

    if (job.cancelled) {
      console.log("Job cancelled after parsing");
      return;
    }

    // Step 3: Translate
    updateJob(job, {
      status: "translating",
      stepMessage: "翻譯中...",
      progress: 20,
    });

    await translateSegments(job, parsed.segments, {
      sourceLang: "English",
      targetLang: "Traditional Chinese",
    });

    if (job.cancelled) {
      console.log("Job cancelled after translation");
      return;
    }

    // Step 4: QA and retranslate
    await qaAndRetranslate(job, parsed.segments, {
      sourceLang: "English",
      targetLang: "Traditional Chinese",
    });

    if (job.cancelled) {
      console.log("Job cancelled after QA");
      return;
    }

    // Step 5: Pack output DOCX
    updateJob(job, {
      status: "packing",
      stepMessage: "正在打包翻譯後的文件...",
      progress: 95,
    });

    const outputSuffix = mode === "cb-to-cns" ? "-cns-translated.docx" : "-translated.docx";
    const outputPath = path.join(OUTPUT_DIR, `${baseName}${outputSuffix}`);
    await writeDocx(parsed, outputPath);

    // Done
    job.outputPath = outputPath;
    finishJob(job, "done");
    updateJob(job, {
      stepMessage: "完成！",
    });

    console.log(`Job ${job.id} completed successfully`);

    // Cleanup: remove uploaded file and working file
    try {
      fs.unlinkSync(uploadedFilePath);
      fs.unlinkSync(workingDocxPath);
    } catch (e) {
      // Ignore cleanup errors
    }
  } catch (error: any) {
    if (job.cancelled) {
      console.log(`Job ${job.id} was cancelled`);
      return;
    }

    console.error(`Job ${job.id} failed:`, error);
    updateJob(job, {
      status: "error",
      errorMessage: error.message || "Unknown error",
      stepMessage: "處理失敗",
      finishedAt: Date.now(),
    });
  }
}

/**
 * Fix multer filename encoding (Latin-1 -> UTF-8)
 * Multer stores originalname as Latin-1 bytes, but browsers send UTF-8
 */
function fixFilenameEncoding(filename: string): string {
  try {
    // Convert Latin-1 string to UTF-8
    const bytes = Buffer.from(filename, "latin1");
    return bytes.toString("utf8");
  } catch {
    return filename;
  }
}

/**
 * POST /api/upload
 * Upload a PDF or DOCX file and start processing
 * Query param: mode=translate (default) or mode=cb-to-cns
 */
app.post("/api/upload", upload.single("file"), (req: Request, res: Response) => {
  if (!req.file) {
    res.status(400).json({ error: "No file uploaded" });
    return;
  }

  // Get mode from query parameter or form field
  const mode = (req.query.mode || req.body?.mode || "translate") as "translate" | "cb-to-cns";

  // Validate mode
  if (mode !== "translate" && mode !== "cb-to-cns") {
    res.status(400).json({ error: "Invalid mode. Use 'translate' or 'cb-to-cns'" });
    return;
  }

  // CB→CNS mode only supports PDF
  const ext = path.extname(req.file.originalname).toLowerCase();
  if (mode === "cb-to-cns" && ext !== ".pdf") {
    res.status(400).json({ error: "CB → CNS 模式只支援 PDF 檔案" });
    return;
  }

  const jobId = uuidv4();
  // Fix filename encoding from Latin-1 to UTF-8
  const originalName = fixFilenameEncoding(req.file.originalname);
  const job = createJob(jobId, originalName);

  // Respond immediately with job ID
  res.json({ jobId, mode });

  // Process in background
  (async () => {
    await processJob(job, req.file!.path, mode);
  })();
});

/**
 * GET /api/status/:jobId
 * Get the current status of a job
 */
app.get("/api/status/:jobId", (req: Request, res: Response) => {
  const job = getJob(req.params.jobId);

  if (!job) {
    res.status(404).json({ error: "Job not found" });
    return;
  }

  const elapsedSeconds = getElapsedSeconds(job);
  const downloadable = job.status === "done" && !!job.outputPath;

  res.json({
    id: job.id,
    fileName: job.fileName,
    status: job.status,
    progress: job.progress,
    stepMessage: job.stepMessage,
    errorMessage: job.errorMessage,
    elapsedSeconds,
    usage: job.usage,
    costUSD: job.costUSD,
    downloadable,
  });
});

/**
 * POST /api/stop/:jobId
 * Cancel a running job
 */
app.post("/api/stop/:jobId", (req: Request, res: Response) => {
  const job = getJob(req.params.jobId);

  if (!job) {
    res.status(404).json({ error: "Job not found" });
    return;
  }

  cancelJob(req.params.jobId);
  res.json({ ok: true });
});

/**
 * GET /api/download/:jobId
 * Download the translated DOCX file
 */
app.get("/api/download/:jobId", (req: Request, res: Response) => {
  const job = getJob(req.params.jobId);

  if (!job || !job.outputPath || job.status !== "done") {
    res.status(404).json({ error: "File not available" });
    return;
  }

  if (!fs.existsSync(job.outputPath)) {
    res.status(404).json({ error: "Output file not found" });
    return;
  }

  const baseName = path.basename(job.fileName, path.extname(job.fileName));
  const downloadName = `${baseName}-translated.docx`;

  // Properly encode filename for Content-Disposition header (RFC 5987)
  // filename: ASCII fallback (replace non-ASCII with underscore)
  // filename*: UTF-8 encoded version for modern browsers
  const asciiFallback = downloadName.replace(/[^\x00-\x7F]/g, "_");
  const encodedName = encodeURIComponent(downloadName).replace(/'/g, "%27");
  res.setHeader(
    "Content-Disposition",
    `attachment; filename="${asciiFallback}"; filename*=UTF-8''${encodedName}`
  );
  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  );

  res.sendFile(job.outputPath);
});

// Error handling middleware
app.use((err: any, req: Request, res: Response, next: any) => {
  console.error("Error:", err);
  res.status(500).json({ error: err.message || "Internal server error" });
});

// Start server
app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
  console.log("Ready to accept file uploads for translation");
});
