/* Mail Merge Engine v1.0.0
 * Auth: Office.js SSO -> Microsoft Graph API
 * Batching: 20 requests per Graph $batch call
 * Security: Zero data stored server-side. All processing in browser RAM.
 */

const GRAPH_BATCH_URL = "https://graph.microsoft.com/v1.0/$batch";
const BATCH_SIZE = 20;
const BATCH_DELAY_MS = 1500;

let accessToken = null;
let parsedRecipients = [];
let csvHeaders = [];

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    log("Office.js ready. Host: Outlook.", "info");
    document.getElementById("mergeBtn").addEventListener("click", runMerge);
  }
});

/* ─── LOGGING ─────────────────────────────────────────────────── */

function log(message, type = "info") {
  const logEl = document.getElementById("statusLog");
  const entry = document.createElement("p");
  entry.className = `log-entry log-${type}`;
  const time = new Date().toLocaleTimeString();
  entry.textContent = `[${time}] ${message}`;
  logEl.appendChild(entry);
  logEl.scrollTop = logEl.scrollHeight;
}

function clearLog() {
  document.getElementById("statusLog").innerHTML = "";
  log("Log cleared.", "info");
}

/* ─── TAG INSERTION ────────────────────────────────────────────── */

function insertTag(tag) {
  Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, (result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      log("Could not read email body.", "error");
      return;
    }
    const current = result.value;
    const updated = current.replace("</body>", `${tag}</body>`);
    Office.context.mailbox.item.body.setAsync(updated, { coercionType: Office.CoercionType.Html }, (setResult) => {
      if (setResult.status !== Office.AsyncResultStatus.Succeeded) {
        log(`Failed to insert tag: ${tag}`, "error");
      } else {
        log(`Inserted tag: ${tag}`, "success");
      }
    });
  });
}

function promptCustomTag() {
  const name = prompt("Enter custom field name (e.g. invoice_number):");
  if (name && name.trim()) {
    const tag = `{{${name.trim().toLowerCase().replace(/\s+/g, "_")}}}`;
    const bar = document.getElementById("tagBar");
    const newTag = document.createElement("span");
    newTag.className = "tag";
    newTag.textContent = tag;
    newTag.onclick = () => insertTag(tag);
    bar.insertBefore(newTag, bar.lastElementChild);
    insertTag(tag);
  }
}

/* ─── CSV PARSING ──────────────────────────────────────────────── */

function parseCSV(raw) {
  const lines = raw.trim().split("\n").filter(l => l.trim());
  if (lines.length < 2) return { headers: [], rows: [] };

  const headers = lines[0].split(",").map(h => h.trim().toLowerCase());
  const rows = [];

  for (let i = 1; i < lines.length; i++) {
    const cols = lines[i].split(",").map(c => c.trim());
    if (cols.length < headers.length) continue;
    const row = {};
    headers.forEach((h, idx) => { row[h] = cols[idx] || ""; });
    rows.push(row);
  }

  return { headers, rows };
}

function parseAndPreview() {
  const raw = document.getElementById("csvInput").value;
  const { headers, rows } = parseCSV(raw);

  if (!headers.includes("email")) {
    log("CSV must contain an 'email' column.", "error");
    return;
  }

  csvHeaders = headers;
  parsedRecipients = rows;

  const countEl = document.getElementById("recipientCount");
  countEl.textContent = `${rows.length} recipient${rows.length !== 1 ? "s" : ""} loaded`;

  renderPreviewTable(headers, rows.slice(0, 5));
  log(`Parsed ${rows.length} recipients. Showing first 5 in preview.`, "success");
}

function renderPreviewTable(headers, rows) {
  const container = document.getElementById("previewTable");
  container.classList.remove("hidden");

  let html = `<label class="label">Preview (first 5 rows)</label><table class="preview-table"><thead><tr>`;
  headers.forEach(h => { html += `<th>${h}</th>`; });
  html += `</tr></thead><tbody>`;
  rows.forEach(row => {
    html += "<tr>";
    headers.forEach(h => { html += `<td>${row[h] || ""}</td>`; });
    html += "</tr>";
  });
  html += "</tbody></table>";
  container.innerHTML = html;
}

/* ─── TOKEN REPLACEMENT ────────────────────────────────────────── */

function personalize(template, recipient) {
  let result = template;
  Object.keys(recipient).forEach(key => {
    const regex = new RegExp(`\\{\\{${key}\\}\\}`, "gi");
    result = result.replace(regex, escapeHtml(recipient[key]));
  });
  return result;
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/* ─── AUTH ─────────────────────────────────────────────────────── */

function getAccessToken() {
  return new Promise((resolve, reject) => {
    Office.context.auth.getAccessTokenAsync(
      { allowSignInPrompt: true, allowConsentPrompt: true, forMSGraphAccess: true },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(new Error(`Auth failed: ${result.error.message}`));
        }
      }
    );
  });
}

/* ─── GRAPH BATCH SEND ─────────────────────────────────────────── */

function buildEmailRequest(id, toEmail, subject, htmlBody) {
  return {
    id: String(id),
    method: "POST",
    url: "/me/sendMail",
    headers: { "Content-Type": "application/json" },
    body: {
      message: {
        subject: subject,
        body: { contentType: "HTML", content: htmlBody },
        toRecipients: [{ emailAddress: { address: toEmail } }]
      },
      saveToSentItems: true
    }
  };
}

function chunkArray(arr, size) {
  const chunks = [];
  for (let i = 0; i < arr.length; i += size) {
    chunks.push(arr.slice(i, i + size));
  }
  return chunks;
}

function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

async function sendBatch(requests, token) {
  const response = await fetch(GRAPH_BATCH_URL, {
    method: "POST",
    headers: {
      "Authorization": `Bearer ${token}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ requests })
  });

  if (!response.ok) {
    throw new Error(`Batch HTTP error: ${response.status} ${response.statusText}`);
  }

  return await response.json();
}

/* ─── MAIN MERGE RUNNER ────────────────────────────────────────── */

async function runMerge() {
  const btn = document.getElementById("mergeBtn");

  if (parsedRecipients.length === 0) {
    parseAndPreview();
    if (parsedRecipients.length === 0) {
      log("No valid recipients found. Check your CSV format.", "error");
      return;
    }
  }

  btn.disabled = true;
  btn.textContent = "⏳ Sending...";
  log("Starting mail merge...", "info");

  try {
    log("Requesting Microsoft 365 access token via Office.js SSO...", "info");
    accessToken = await getAccessToken();
    log("Access token obtained. Processing recipients...", "success");
  } catch (err) {
    log(`Authentication error: ${err.message}`, "error");
    btn.disabled = false;
    btn.textContent = "▶ Run mail merge";
    return;
  }

  const subjectTemplate = document.getElementById("subjectInput").value.trim();
  if (!subjectTemplate) {
    log("Subject line is empty.", "error");
    btn.disabled = false;
    btn.textContent = "▶ Run mail merge";
    return;
  }

  let emailBodyTemplate = "";
  try {
    emailBodyTemplate = await new Promise((resolve, reject) => {
      Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(new Error("Could not read email body."));
        }
      });
    });
  } catch (err) {
    log(`Body read error: ${err.message}`, "error");
    btn.disabled = false;
    btn.textContent = "▶ Run mail merge";
    return;
  }

  const requests = parsedRecipients.map((recipient, idx) => {
    const personalizedSubject = personalize(subjectTemplate, recipient);
    const personalizedBody = personalize(emailBodyTemplate, recipient);
    return buildEmailRequest(idx + 1, recipient.email, personalizedSubject, personalizedBody);
  });

  const batches = chunkArray(requests, BATCH_SIZE);
  const totalBatches = batches.length;
  let totalSent = 0;
  let totalFailed = 0;

  log(`${parsedRecipients.length} emails split into ${totalBatches} batch${totalBatches > 1 ? "es" : ""} of ${BATCH_SIZE}.`, "info");

  for (let i = 0; i < batches.length; i++) {
    const batch = batches[i];
    log(`Sending batch ${i + 1}/${totalBatches} (${batch.length} emails)...`, "info");

    try {
      const result = await sendBatch(batch, accessToken);
      const responses = result.responses || [];

      responses.forEach(r => {
        if (r.status >= 200 && r.status < 300) {
          totalSent++;
        } else {
          totalFailed++;
          log(`Failed for request ID ${r.id}: ${r.status} — ${JSON.stringify(r.body)}`, "error");
        }
      });

      log(`Batch ${i + 1}/${totalBatches} complete. Sent: ${totalSent} | Failed: ${totalFailed}`, "success");
    } catch (err) {
      log(`Batch ${i + 1} error: ${err.message}`, "error");
      totalFailed += batch.length;
    }

    if (i < batches.length - 1) {
      log(`Waiting ${BATCH_DELAY_MS / 1000}s before next batch to avoid throttling...`, "info");
      await delay(BATCH_DELAY_MS);
    }
  }

  log(`─── Merge complete. ✅ Sent: ${totalSent} | ❌ Failed: ${totalFailed} ───`, totalFailed > 0 ? "warning" : "success");

  btn.disabled = false;
  btn.textContent = "▶ Run mail merge";
}
