const statusEl = document.getElementById("status");
const chunksEl = document.getElementById("chunks");

async function load() {
  try {
    const res = await fetch("/api/chunks");
    const data = await res.json();
    if (!res.ok) throw new Error(data.error);

    statusEl.textContent = `${data.totalRecords} records (${data.emails} emails + ${data.attachments} attachments) split into ${data.chunks.length} chunks. Paste each into ChatGPT in order.`;

    for (let i = 0; i < data.chunks.length; i++) {
      const chunk = data.chunks[i];
      const card = document.createElement("div");
      card.className = "chunk-card";

      const header = document.createElement("div");
      header.className = "chunk-header";

      const title = document.createElement("h2");
      title.textContent = `Chunk ${i + 1} of ${data.chunks.length}`;

      const info = document.createElement("span");
      info.textContent = `${chunk.records} records · ${chunk.chars.toLocaleString()} chars`;

      const btn = document.createElement("button");
      btn.textContent = "Copy";
      btn.addEventListener("click", () => {
        navigator.clipboard.writeText(chunk.text).then(() => {
          btn.textContent = "Copied!";
          btn.classList.add("copied");
          setTimeout(() => {
            btn.textContent = "Copy";
            btn.classList.remove("copied");
          }, 2000);
        });
      });

      header.appendChild(title);
      header.appendChild(info);
      header.appendChild(btn);

      const ta = document.createElement("textarea");
      ta.value = chunk.text;
      ta.readOnly = true;

      card.appendChild(header);
      card.appendChild(ta);
      chunksEl.appendChild(card);
    }
  } catch (e) {
    statusEl.textContent = e.message || "Failed to load.";
    statusEl.style.color = "#b91c1c";
    setTimeout(load, 3000);
  }
}

load();
