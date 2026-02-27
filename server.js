require("dotenv").config();
const express = require("express");
const multer = require("multer");
const fs = require("fs");
const path = require("path");
const pdfParse = require("pdf-parse");
const mammoth = require("mammoth");
const XLSX = require("xlsx");

const app = express();
const uploadDir = path.join(__dirname, "uploads");
if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir);

const storage = multer.diskStorage({
  destination: uploadDir,
  filename: (_req, file, cb) => cb(null, `${Date.now()}-${file.originalname}`),
});
const upload = multer({ storage, limits: { fileSize: 50 * 1024 * 1024 } });

app.use(express.json({ limit: "50mb" }));
app.use(express.static("public"));

// --------------- email helpers ---------------

function parseHeaders(raw) {
  const headers = {};
  const lines = raw.split(/\r?\n/);
  let key = null;

  for (const line of lines) {
    if (!line.trim()) continue;
    if (/^\s/.test(line) && key) {
      headers[key] = `${headers[key]} ${line.trim()}`.trim();
      continue;
    }
    const idx = line.indexOf(":");
    if (idx === -1) continue;
    key = line.slice(0, idx).trim().toLowerCase();
    headers[key] = line.slice(idx + 1).trim();
  }
  return headers;
}

function decodeQP(text) {
  return text.replace(/=\r?\n/g, "").replace(/=([A-Fa-f0-9]{2})/g, (_m, h) =>
    String.fromCharCode(parseInt(h, 16))
  );
}

function decodeEncodedWords(text) {
  return text.replace(/=\?([^?]+)\?([bBqQ])\?([^?]+)\?=/g, (_m, cs, enc, data) => {
    try {
      if (enc.toUpperCase() === "B")
        return Buffer.from(data, "base64").toString(/utf-8/i.test(cs) ? "utf8" : "latin1");
      const qp = data.replace(/_/g, " ");
      return decodeQP(qp);
    } catch (_e) {
      return data;
    }
  });
}

function sanitize(text) {
  return text
    .replace(/\u0000/g, "")
    .replace(/[^\x09\x0A\x0D\x20-\x7E\u00A0-\u024F]/g, " ")
    .replace(/[ \t]{2,}/g, " ")
    .trim();
}

function stripHtml(html) {
  return html
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<[^>]+>/g, " ");
}

function splitHeaderBody(text) {
  for (const sep of ["\r\n\r\n", "\n\n"]) {
    const idx = text.indexOf(sep);
    if (idx !== -1) {
      return { rawHeaders: text.slice(0, idx), body: text.slice(idx + sep.length) };
    }
  }
  return { rawHeaders: text, body: "" };
}

function decodeBody(body, headers) {
  const enc = (headers["content-transfer-encoding"] || "").toLowerCase();
  try {
    if (enc.includes("base64"))
      return Buffer.from(body.replace(/\s+/g, ""), "base64").toString("utf8");
    if (enc.includes("quoted-printable"))
      return Buffer.from(decodeQP(body), "latin1").toString("utf8");
  } catch (_e) {}
  return body;
}

function getBoundary(ct) {
  const m = ct.match(/boundary="?([^";]+)"?/i);
  return m ? m[1] : "";
}

function extractText(body, headers) {
  const ct = (headers["content-type"] || "").toLowerCase();
  const boundary = getBoundary(ct);

  if (!ct.includes("multipart/") || !boundary) {
    const decoded = decodeBody(body, headers);
    return ct.includes("text/html") ? stripHtml(decoded) : decoded;
  }

  const marker = `--${boundary}`;
  const parts = body.split(marker).filter((p) => p.trim() && p.trim() !== "--");
  const texts = [];

  for (const raw of parts) {
    const part = raw.replace(/--$/, "").trim();
    if (!part) continue;

    const { rawHeaders: ph, body: pb } = splitHeaderBody(part);
    const pHeaders = parseHeaders(ph);
    const pType = (pHeaders["content-type"] || "text/plain").toLowerCase();

    if (pType.includes("multipart/")) {
      const nested = extractText(pb, pHeaders);
      if (nested.trim()) texts.push(nested);
      continue;
    }

    if (!pType.includes("text/plain") && !pType.includes("text/html")) continue;

    const decoded = decodeBody(pb, pHeaders);
    const clean = pType.includes("text/html") ? stripHtml(decoded) : decoded;
    if (clean.trim()) texts.push(clean);
  }

  return texts.length ? texts.join("\n\n") : decodeBody(body, headers);
}

function parseEml(content, id, filename) {
  const { rawHeaders, body } = splitHeaderBody(content);
  const headers = parseHeaders(rawHeaders);
  const bodyText = extractText(body, headers);

  return {
    id,
    filename,
    from: sanitize(decodeEncodedWords(headers.from || "")),
    to: sanitize(decodeEncodedWords(headers.to || "")),
    subject: sanitize(decodeEncodedWords(headers.subject || "")),
    date: sanitize(headers.date || ""),
    body_preview: sanitize(bodyText).slice(0, 1500),
  };
}

// --------------- attachment helpers ---------------

async function extractPdfText(filePath) {
  const buffer = fs.readFileSync(filePath);
  const data = await pdfParse(buffer);
  return data.text || "";
}

async function extractDocxText(filePath) {
  const result = await mammoth.extractRawText({ path: filePath });
  return result.value || "";
}

function extractXlsxText(filePath) {
  const workbook = XLSX.readFile(filePath);
  const sheets = [];
  for (const name of workbook.SheetNames) {
    const sheet = workbook.Sheets[name];
    const csv = XLSX.utils.sheet_to_csv(sheet);
    if (csv.trim()) sheets.push(`[Sheet: ${name}]\n${csv}`);
  }
  return sheets.join("\n\n");
}

function extractDocText(filePath) {
  const buf = fs.readFileSync(filePath);
  const text = buf.toString("utf8").replace(/[^\x09\x0A\x0D\x20-\x7E]/g, " ").replace(/\s{3,}/g, "\n");
  return text.trim();
}

const ATTACHMENT_EXTS = new Set([".pdf", ".docx", ".doc", ".xlsx", ".pptx"]);

async function parseAttachment(filePath, id, filename) {
  const ext = path.extname(filename).toLowerCase();
  let text = "";

  try {
    if (ext === ".pdf") {
      text = await extractPdfText(filePath);
    } else if (ext === ".docx" || ext === ".pptx") {
      text = await extractDocxText(filePath);
    } else if (ext === ".xlsx") {
      text = extractXlsxText(filePath);
    } else if (ext === ".doc") {
      text = extractDocText(filePath);
    }
  } catch (_e) {
    text = `[Could not extract text from ${filename}]`;
  }

  return {
    id,
    type: "attachment",
    filename,
    file_type: ext.replace(".", "").toUpperCase(),
    body_preview: sanitize(text).slice(0, 2000),
  };
}

// --------------- folder scan ---------------

async function loadAllFolder(dirPath) {
  const allFiles = fs.readdirSync(dirPath).sort();
  const records = [];
  let id = 0;

  for (const file of allFiles) {
    const ext = path.extname(file).toLowerCase();
    const fullPath = path.join(dirPath, file);

    try {
      if (ext === ".eml") {
        id++;
        const content = fs.readFileSync(fullPath, "utf8");
        const record = parseEml(content, id, file);
        if (record.from || record.to || record.subject || record.body_preview) {
          records.push(record);
        }
      } else if (ATTACHMENT_EXTS.has(ext)) {
        id++;
        const record = await parseAttachment(fullPath, id, file);
        if (record.body_preview && record.body_preview.length > 20) {
          records.push(record);
        }
      }
    } catch (_e) {
      // Skip unreadable files.
    }
  }

  return records;
}

// --------------- routes ---------------

const ALL_DIR = path.join(__dirname, "All");

app.get("/api/status", (_req, res) => {
  const hasAll = fs.existsSync(ALL_DIR);
  let emlCount = 0;
  let attachmentCount = 0;
  if (hasAll) {
    const files = fs.readdirSync(ALL_DIR);
    emlCount = files.filter((f) => f.toLowerCase().endsWith(".eml")).length;
    attachmentCount = files.filter((f) => {
      const ext = path.extname(f).toLowerCase();
      return ATTACHMENT_EXTS.has(ext);
    }).length;
  }
  res.json({ hasAll, emlCount, attachmentCount });
});

app.post("/api/convert-folder", async (_req, res) => {
  try {
    if (!fs.existsSync(ALL_DIR)) {
      return res.status(400).json({ error: "All/ folder not found next to server." });
    }
    const records = await loadAllFolder(ALL_DIR);
    if (records.length === 0) {
      return res.status(400).json({ error: "No readable files found in All/." });
    }
    const emails = records.filter((r) => r.type !== "attachment");
    const attachments = records.filter((r) => r.type === "attachment");
    const jsonl = records.map((r) => JSON.stringify(r)).join("\n");
    return res.json({
      count: records.length,
      emailCount: emails.length,
      attachmentCount: attachments.length,
      records,
      jsonl,
    });
  } catch (error) {
    return res.status(500).json({ error: error.message || "Convert failed." });
  }
});

app.post("/api/convert", upload.single("mbox"), (req, res) => {
  let filePath;
  try {
    if (!req.file) return res.status(400).json({ error: "No file uploaded." });
    filePath = req.file.path;
    const content = fs.readFileSync(filePath, "utf8");
    const record = parseEml(content, 1, req.file.originalname);
    const records = [record];
    const jsonl = JSON.stringify(record);
    return res.json({ count: 1, records, jsonl });
  } catch (error) {
    return res.status(500).json({ error: error.message || "Convert failed." });
  } finally {
    if (filePath) fs.unlink(filePath, () => {});
  }
});

function makePrompt(records) {
  const capped = records.slice(0, 80);
  return [
    "You are reviewing email data converted to JSONL.",
    "Provide a concise summary with:",
    "1) Main topics and themes across the emails",
    "2) Key people involved and their roles",
    "3) Any potential risks, disputes, or items needing attention",
    "4) Timeline of key events",
    "5) Actionable next steps",
    "",
    "Here are the email records:",
    ...capped.map((r) => JSON.stringify(r)),
  ].join("\n");
}

async function reviewWithOpenAI(records) {
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    return {
      mode: "mock",
      review:
        "Mock review (no OPENAI_API_KEY set):\n\n" +
        `Loaded ${records.length} emails.\n` +
        "Set OPENAI_API_KEY in .env to get a real analysis.\n\n" +
        "Top subjects:\n" +
        records.slice(0, 10).map((r) => `  - ${r.subject || "(no subject)"}`).join("\n"),
    };
  }

  const prompt = makePrompt(records);
  const resp = await fetch("https://api.openai.com/v1/chat/completions", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`,
    },
    body: JSON.stringify({
      model: "gpt-4o-mini",
      messages: [
        { role: "system", content: "You are a careful email analyst." },
        { role: "user", content: prompt },
      ],
      temperature: 0.2,
    }),
  });

  if (!resp.ok) {
    const text = await resp.text();
    throw new Error(`LLM request failed: ${resp.status} ${text}`);
  }

  const data = await resp.json();
  return {
    mode: "live",
    review: data.choices?.[0]?.message?.content || "No review returned.",
  };
}

app.post("/api/ask", async (req, res) => {
  try {
    const { question, records } = req.body || {};
    if (!question) return res.status(400).json({ error: "No question provided." });
    if (!Array.isArray(records) || records.length === 0) {
      return res.status(400).json({ error: "No records loaded." });
    }

    const apiKey = process.env.OPENAI_API_KEY;
    if (!apiKey) {
      return res.json({
        answer:
          "No OPENAI_API_KEY set in .env\n\n" +
          `${records.length} records loaded. Set the key to query them.`,
      });
    }

    const capped = records.slice(0, 80);
    const context = capped.map((r) => JSON.stringify(r)).join("\n");

    const resp = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`,
      },
      body: JSON.stringify({
        model: "gpt-4o-mini",
        messages: [
          {
            role: "system",
            content:
              "You are an expert email analyst. The user has loaded a set of emails and attachments as JSONL records. " +
              "Answer the user's question based ONLY on the provided records. Be specific, cite filenames/subjects/senders when relevant.\n\n" +
              "RECORDS:\n" + context,
          },
          { role: "user", content: question },
        ],
        temperature: 0.2,
      }),
    });

    if (!resp.ok) {
      const text = await resp.text();
      throw new Error(`LLM error: ${resp.status} ${text}`);
    }

    const data = await resp.json();
    return res.json({
      answer: data.choices?.[0]?.message?.content || "No answer returned.",
    });
  } catch (error) {
    return res.status(500).json({ error: error.message || "Ask failed." });
  }
});

const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log(`Mbox JSONL reviewer running at http://localhost:${port}`);
});
