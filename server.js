require("dotenv").config();
const express = require("express");
const multer = require("multer");
const fs = require("fs");
const path = require("path");
const os = require("os");
const pdfParse = require("pdf-parse");
const mammoth = require("mammoth");
const XLSX = require("xlsx");
const AdmZip = require("adm-zip");

const app = express();
const zipUpload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 500 * 1024 * 1024 } });
app.use(express.json({ limit: "10mb" }));
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
      return decodeQP(data.replace(/_/g, " "));
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
    if (idx !== -1)
      return { rawHeaders: text.slice(0, idx), body: text.slice(idx + sep.length) };
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
    const pH = parseHeaders(ph);
    const pType = (pH["content-type"] || "text/plain").toLowerCase();
    if (pType.includes("multipart/")) {
      const nested = extractText(pb, pH);
      if (nested.trim()) texts.push(nested);
      continue;
    }
    if (!pType.includes("text/plain") && !pType.includes("text/html")) continue;
    const decoded = decodeBody(pb, pH);
    const clean = pType.includes("text/html") ? stripHtml(decoded) : decoded;
    if (clean.trim()) texts.push(clean);
  }
  return texts.length ? texts.join("\n\n") : decodeBody(body, headers);
}

function parseEml(content, id, filename) {
  const { rawHeaders, body } = splitHeaderBody(content);
  const headers = parseHeaders(rawHeaders);
  const bodyText = extractText(body, headers);
  const fullBody = sanitize(bodyText);
  return {
    id, filename,
    type: "email",
    from: sanitize(decodeEncodedWords(headers.from || "")),
    to: sanitize(decodeEncodedWords(headers.to || "")),
    subject: sanitize(decodeEncodedWords(headers.subject || "")),
    date: sanitize(headers.date || ""),
    body: fullBody,
  };
}

// --------------- attachment helpers ---------------

async function extractPdfText(fp) {
  const data = await pdfParse(fs.readFileSync(fp));
  return data.text || "";
}

async function extractDocxText(fp) {
  const r = await mammoth.extractRawText({ path: fp });
  return r.value || "";
}

function extractXlsxText(fp) {
  const wb = XLSX.readFile(fp);
  const sheets = [];
  for (const name of wb.SheetNames) {
    const csv = XLSX.utils.sheet_to_csv(wb.Sheets[name]);
    if (csv.trim()) sheets.push(`[Sheet: ${name}]\n${csv}`);
  }
  return sheets.join("\n\n");
}

function extractDocText(fp) {
  return fs.readFileSync(fp).toString("utf8").replace(/[^\x09\x0A\x0D\x20-\x7E]/g, " ").replace(/\s{3,}/g, "\n").trim();
}

const ATTACHMENT_EXTS = new Set([".pdf", ".docx", ".doc", ".xlsx", ".pptx"]);

async function parseAttachment(fp, id, filename) {
  const ext = path.extname(filename).toLowerCase();
  let text = "";
  try {
    if (ext === ".pdf") text = await extractPdfText(fp);
    else if (ext === ".docx" || ext === ".pptx") text = await extractDocxText(fp);
    else if (ext === ".xlsx") text = extractXlsxText(fp);
    else if (ext === ".doc") text = extractDocText(fp);
  } catch (_e) {
    text = `[Could not extract text from ${filename}]`;
  }
  return {
    id, filename,
    type: "attachment",
    file_type: ext.replace(".", "").toUpperCase(),
    body: sanitize(text),
  };
}

// --------------- buffer-based parsers (for zip upload) ---------------

async function parseAttachmentFromBuffer(buffer, id, filename) {
  const ext = path.extname(filename).toLowerCase();
  let text = "";
  try {
    if (ext === ".pdf") {
      const data = await pdfParse(buffer);
      text = data.text || "";
    } else if (ext === ".docx" || ext === ".pptx") {
      const tmpPath = path.join(os.tmpdir(), `upload-${Date.now()}-${filename}`);
      fs.writeFileSync(tmpPath, buffer);
      const r = await mammoth.extractRawText({ path: tmpPath });
      text = r.value || "";
      fs.unlinkSync(tmpPath);
    } else if (ext === ".xlsx") {
      const wb = XLSX.read(buffer, { type: "buffer" });
      const sheets = [];
      for (const name of wb.SheetNames) {
        const csv = XLSX.utils.sheet_to_csv(wb.Sheets[name]);
        if (csv.trim()) sheets.push(`[Sheet: ${name}]\n${csv}`);
      }
      text = sheets.join("\n\n");
    } else if (ext === ".doc") {
      text = buffer.toString("utf8").replace(/[^\x09\x0A\x0D\x20-\x7E]/g, " ").replace(/\s{3,}/g, "\n").trim();
    }
  } catch (_e) {
    text = `[Could not extract text from ${filename}]`;
  }
  return {
    id, filename,
    type: "attachment",
    file_type: ext.replace(".", "").toUpperCase(),
    body: sanitize(text),
  };
}

async function loadFromZipBuffer(zipBuffer) {
  const zip = new AdmZip(zipBuffer);
  const entries = zip.getEntries();
  const records = [];
  let id = 0;

  const sorted = entries
    .filter((e) => !e.isDirectory && !e.entryName.startsWith("__MACOSX"))
    .sort((a, b) => a.entryName.localeCompare(b.entryName));

  for (const entry of sorted) {
    const filename = path.basename(entry.entryName);
    const ext = path.extname(filename).toLowerCase();
    const buf = entry.getData();

    try {
      if (ext === ".eml") {
        id++;
        const content = buf.toString("utf8");
        const record = parseEml(content, id, filename);
        if (record.from || record.to || record.subject || record.body) records.push(record);
      } else if (ATTACHMENT_EXTS.has(ext)) {
        id++;
        const record = await parseAttachmentFromBuffer(buf, id, filename);
        if (record.body && record.body.length > 20) records.push(record);
      }
    } catch (_e) {}
  }

  return records;
}

// --------------- in-memory store ---------------

let ALL_RECORDS = [];
let CHUNKS = [];

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
        if (record.from || record.to || record.subject || record.body) records.push(record);
      } else if (ATTACHMENT_EXTS.has(ext)) {
        id++;
        const record = await parseAttachment(fullPath, id, file);
        if (record.body && record.body.length > 20) records.push(record);
      }
    } catch (_e) {}
  }
  return records;
}

const MAX_CHUNK_CHARS = 80000;

function buildChunks(records) {
  const chunks = [];
  let current = [];
  let currentSize = 0;

  for (const r of records) {
    const line = JSON.stringify(r);
    if (currentSize + line.length > MAX_CHUNK_CHARS && current.length > 0) {
      chunks.push(current);
      current = [];
      currentSize = 0;
    }
    current.push(line);
    currentSize += line.length;
  }
  if (current.length > 0) chunks.push(current);
  return chunks;
}

// --------------- routes ---------------

const ALL_DIR = path.join(__dirname, "All");

app.get("/api/status", (_req, res) => {
  res.json({
    loaded: ALL_RECORDS.length > 0,
    totalRecords: ALL_RECORDS.length,
    chunks: CHUNKS.length,
    emails: ALL_RECORDS.filter((r) => r.type === "email").length,
    attachments: ALL_RECORDS.filter((r) => r.type === "attachment").length,
  });
});

async function callLLM(messages) {
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) throw new Error("No OPENAI_API_KEY set in .env");

  const resp = await fetch("https://api.openai.com/v1/chat/completions", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`,
    },
    body: JSON.stringify({
      model: "gpt-4o-mini",
      messages,
      temperature: 0.2,
      max_tokens: 2048,
    }),
  });
  if (!resp.ok) {
    const text = await resp.text();
    throw new Error(`LLM error: ${resp.status} ${text}`);
  }
  const data = await resp.json();
  return data.choices?.[0]?.message?.content || "";
}

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

app.post("/api/ask", async (req, res) => {
  try {
    const { question } = req.body || {};
    if (!question) return res.status(400).json({ error: "No question provided." });
    if (ALL_RECORDS.length === 0) return res.status(400).json({ error: "No data loaded." });

    const apiKey = process.env.OPENAI_API_KEY;
    if (!apiKey) {
      return res.json({ answer: "No OPENAI_API_KEY set in .env. Add your key and restart." });
    }

    // MAP: send each chunk to GPT with the question, collect partial answers
    const chunkPrompt =
      "You are analysing a batch of email and attachment records (JSONL). " +
      "Extract ALL information relevant to the user's question from this batch. " +
      "Include specific details: names, dates, numbers, filenames, quotes. " +
      "If nothing relevant is found, reply with just: [NO_RELEVANT_DATA]";

    const partialAnswers = [];
    for (let i = 0; i < CHUNKS.length; i++) {
      const chunkText = CHUNKS[i].join("\n");
      try {
        const partial = await callLLM([
          { role: "system", content: chunkPrompt + `\n\nRECORDS (batch ${i + 1}/${CHUNKS.length}):\n` + chunkText },
          { role: "user", content: question },
        ]);
        if (partial && !partial.includes("[NO_RELEVANT_DATA]")) {
          partialAnswers.push(`[Batch ${i + 1}]\n${partial}`);
        }
      } catch (e) {
        if (e.message.includes("429")) {
          await sleep(5000);
          i--;
          continue;
        }
        partialAnswers.push(`[Batch ${i + 1}] Error: ${e.message}`);
      }
      if (i < CHUNKS.length - 1) await sleep(1000);
    }

    if (partialAnswers.length === 0) {
      return res.json({ answer: "No relevant information found across all records.", chunksProcessed: CHUNKS.length });
    }

    // REDUCE: combine all partial answers into one final answer
    const combined = partialAnswers.join("\n\n");
    const finalAnswer = await callLLM([
      {
        role: "system",
        content:
          "You are an expert analyst. Below are extracted findings from multiple batches of emails, documents, and spreadsheets. " +
          "Combine them into one clear, thorough answer. Remove duplicates. Cite specific filenames, people, dates, and data points. " +
          "Structure your answer with clear sections if appropriate.\n\n" +
          "FINDINGS:\n" + combined,
      },
      { role: "user", content: question },
    ]);

    res.json({ answer: finalAnswer, chunksProcessed: CHUNKS.length });
  } catch (error) {
    res.status(500).json({ error: error.message || "Ask failed." });
  }
});

const port = process.env.PORT || 3000;
app.listen(port, async () => {
  console.log(`Server running at http://localhost:${port}`);
  if (fs.existsSync(ALL_DIR)) {
    console.log("Auto-loading All/ folder...");
    ALL_RECORDS = await loadAllFolder(ALL_DIR);
    CHUNKS = buildChunks(ALL_RECORDS);
    const emails = ALL_RECORDS.filter((r) => r.type === "email").length;
    const attachments = ALL_RECORDS.filter((r) => r.type === "attachment").length;
    console.log(`Loaded ${ALL_RECORDS.length} records (${emails} emails, ${attachments} attachments) in ${CHUNKS.length} chunks`);
  }
});
