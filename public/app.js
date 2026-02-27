const statusEl = document.getElementById("status");
const promptInput = document.getElementById("promptInput");
const askBtn = document.getElementById("askBtn");
const resultCard = document.getElementById("resultCard");
const resultOutput = document.getElementById("resultOutput");

function setStatus(msg, isError = false) {
  statusEl.textContent = msg;
  statusEl.style.color = isError ? "#b91c1c" : "#4b5563";
}

async function checkReady() {
  try {
    const res = await fetch("/api/status");
    const data = await res.json();
    if (data.loaded) {
      setStatus(`${data.emails} emails + ${data.attachments} attachments loaded. Ask away.`);
      askBtn.disabled = false;
      promptInput.focus();
    } else {
      setStatus("Loading data...");
      setTimeout(checkReady, 2000);
    }
  } catch (_e) {
    setStatus("Server not reachable.", true);
  }
}

checkReady();

async function ask() {
  const question = promptInput.value.trim();
  if (!question) return;

  try {
    askBtn.disabled = true;
    askBtn.textContent = "Thinking...";
    resultCard.style.display = "block";
    resultOutput.textContent = "Searching relevant files and analysing...";

    const res = await fetch("/api/ask", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ question }),
    });

    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Request failed.");

    let output = data.answer || "No answer returned.";
    if (data.filesUsed) {
      output += `\n\n---\nFiles referenced: ${data.filesUsed}`;
    }
    resultOutput.textContent = output;
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
