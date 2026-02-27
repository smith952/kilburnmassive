const statusEl = document.getElementById("status");
const promptInput = document.getElementById("promptInput");
const askBtn = document.getElementById("askBtn");
const resultCard = document.getElementById("resultCard");
const resultOutput = document.getElementById("resultOutput");

let records = [];

function setStatus(msg, isError = false) {
  statusEl.textContent = msg;
  statusEl.style.color = isError ? "#b91c1c" : "#4b5563";
}

async function loadData() {
  try {
    const res = await fetch("/api/convert-folder", { method: "POST" });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Load failed.");

    records = data.records || [];
    const parts = [];
    if (data.emailCount) parts.push(`${data.emailCount} emails`);
    if (data.attachmentCount) parts.push(`${data.attachmentCount} attachments`);
    setStatus(`${parts.join(" + ") || data.count + " files"} loaded. Ask away.`);
    askBtn.disabled = false;
    promptInput.focus();
  } catch (error) {
    setStatus(error.message || "Failed to load emails.", true);
  }
}

loadData();

async function ask() {
  const question = promptInput.value.trim();
  if (!question) return;
  if (!records.length) {
    setStatus("No data loaded.", true);
    return;
  }

  try {
    askBtn.disabled = true;
    askBtn.textContent = "Thinking...";
    resultCard.style.display = "block";
    resultOutput.textContent = "Thinking...";

    const res = await fetch("/api/ask", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ question, records }),
    });

    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Request failed.");

    resultOutput.textContent = data.answer || "No answer returned.";
  } catch (error) {
    resultOutput.textContent = error.message || "Something went wrong.";
  } finally {
    askBtn.disabled = false;
    askBtn.textContent = "Ask";
  }
}

askBtn.addEventListener("click", ask);

promptInput.addEventListener("keydown", (e) => {
  if (e.key === "Enter" && !e.shiftKey) {
    e.preventDefault();
    ask();
  }
});
