/* Mail Merge Engine v1.0.3
 * Auth: Office.js SSO -> Microsoft Graph API
 * Batching: 20 requests per Graph $batch call
 * Security: Zero data stored server-side. All processing in browser RAM.
 */

const GRAPH_BATCH_URL = "https://graph.microsoft.com/v1.0/$batch";
const GRAPH_MAILBOX_URL = "https://graph.microsoft.com/v1.0/me/mailboxSettings";
const BATCH_SIZE = 20;
const BATCH_DELAY_MS = 1500;
const MAX_RETRIES = 3;
const EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const MAX_RECIPIENTS = 10000;
const MAX_PAYLOAD_BYTES = 3.5 * 1024 * 1024; // warn at 3.5 MB (hard limit is 4 MB)

const LS_KEY_SUBJECT = "mailmerge_subject";
const LS_KEY_CSV = "mailmerge_csv";
const LS_KEY_TAGS = "mailmerge_custom_tags";

// Default tags always shown in the tag bar
const DEFAULT_TAGS = ["{{first_name}}", "{{last_name}}", "{{email}}", "{{company}}", "{{title}}"];

let parsedRecipients = [];
let csvHeaders = [];
let mergeRunning = false;
let cancelRequested = false;
let subjectHasFocus = false;

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    log("Office.js ready. Host: Outlook.", "info");

    document.getElementById("mergeBtn").addEventListener("click", handleMergeClick);
    document.getElementById("stopBtn").addEventListener("click", handleStop);
    document.getElementById("previewBtn").addEventListener("click", parseAndPreview);
    document.getElementById("clearLogBtn").addEventListener("click", clearLog);
    document.getElementById("addCustomTagBtn").addEventListener("click", addCustomTag);
    document.getElementById("confirmSendBtn").addEventListener("click", confirmSend);
    document.getElementById("confirmCancelBtn").addEventListener("click", dismissModal);
    document.getElementById("uploadCsvBtn").addEventListener("click", () => {
      document.getElementById("csvFileInput").click();
    });
    document.getElementById("csvFileInput").addEventListener("change", handleCsvFileUpload);

    // Track subject focus so tag clicks know where to insert
    const subjectInput = document.getElementById("subjectInput");
    subjectInput.addEventListener("focus", () => { subjectHasFocus = true; });
    subjectInput.addEventListener("blur", () => { subjectHasFocus = false; });

    // Tag bar: delegate clicks to tag spans
    document.getElementById("tagBar").addEventListener("click", (e) => {
      const tag = e.target.dataset.tag;
      if (tag) insertTag(tag);
    });

    // Allow Enter key in custom tag input
    document.getElementById("customTagInput").addEventListener("keydown", (e) => {
      if (e.key === "Enter") addCustomTag();
    });

    // Persist subject changes
    document.getElementById("subjectInput").addEventListener("input", () => {
      localStorage.setItem(LS_KEY_SUBJECT, document.getElementById("subjectInput").value);
    });

    // Persist CSV textarea changes
    document.getElementById("csvInput").addEventListener("input", () => {
      localStorage.setItem(LS_KEY_CSV, document.getElementById("csvInput").value);
    });

    restoreLocalState();
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

/* ─── LOCAL STATE PERSISTENCE ─────────────────────────────────── */

function restoreLocalState() {
  const savedSubject = localStorage.getItem(LS_KEY_SUBJECT);
  if (savedSubject) {
    document.getElementById("subjectInput").value = savedSubject;
    log("Restored subject from saved session.", "info");
  }

  const savedCsv = localStorage.getItem(LS_KEY_CSV);
  if (savedCsv) {
    document.getElementById("csvInput").value = savedCsv;
    log("Restored CSV data from saved session.", "info");
    parseAndPreview();
  }

  const savedTags = JSON.parse(localStorage.getItem(LS_KEY_TAGS) || "[]");
  savedTags.forEach(tag => addTagToBar(tag));
}

function saveCustomTagsToStorage() {
  const tagEls = document.querySelectorAll("#tagBar [data-tag]");
  const tags = Array.from(tagEls)
    .map(el => el.dataset.tag)
    .filter(t => !DEFAULT_TAGS.includes(t));
  localStorage.setItem(LS_KEY_TAGS, JSON.stringify(tags));
}

/* ─── TAG INSERTION ────────────────────────────────────────────── */

function insertTag(tag) {
  if (subjectHasFocus) {
    // Insert at cursor position in the subject text input
    const input = document.getElementById("subjectInput");
    const start = input.selectionStart;
    const end = input.selectionEnd;
    const value = input.value;
    input.value = value.slice(0, start) + tag + value.slice(end);
    const newCursor = start + tag.length;
    input.setSelectionRange(newCursor, newCursor);
    input.focus();
    log(`Inserted tag into subject: ${tag}`, "success");
  } else {
    Office.context.mailbox.item.body.setSelectedDataAsync(
      tag,
      { coercionType: Office.CoercionType.Html },
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          log(`Failed to insert tag: ${tag}`, "error");
        } else {
          log(`Inserted tag into body: ${tag}`, "success");
        }
      }
    );
  }
}

function addCustomTag() {
  const input = document.getElementById("customTagInput");
  const name = input.value.trim();
  if (!name) return;

  const normalized = name.toLowerCase().replace(/\s+/g, "_");
  const tag = `{{${normalized}}}`;
  addTagToBar(tag);
  insertTag(tag);
  input.value = "";
}

function addTagToBar(tag) {
  const existing = document.querySelector(`#tagBar [data-tag="${CSS.escape(tag)}"]`);
  if (existing) return;
  const bar = document.getElementById("tagBar");
  const newTag = document.createElement("span");
  newTag.className = "tag";
  newTag.dataset.tag = tag;
  newTag.textContent = tag;
  bar.appendChild(newTag);
  saveCustomTagsToStorage();
}

/* ─── CSV FILE UPLOAD ──────────────────────────────────────────── */

function handleCsvFileUpload(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (event) => {
    document.getElementById("csvInput").value = event.target.result;
    log(`Loaded file: ${file.name} (${(file.size / 1024).toFixed(1)} KB)`, "info");
    parseAndPreview();
  };
  reader.onerror = () => {
    log(`Failed to read file: ${file.name}`, "error");
  };
  reader.readAsText(file);

  // Reset so the same file can be re-loaded if needed
  e.target.value = "";
}

/* ─── CSV PARSING (RFC 4180) ───────────────────────────────────── */

function parseCSVLine(line) {
  const fields = [];
  let i = 0;
  while (i < line.length) {
    if (line[i] === '"') {
      // Quoted field
      let field = "";
      i++; // skip opening quote
      while (i < line.length) {
        if (line[i] === '"') {
          if (line[i + 1] === '"') {
            // Escaped quote
            field += '"';
            i += 2;
          } else {
            i++; // skip closing quote
            break;
          }
        } else {
          field += line[i];
          i++;
        }
      }
      fields.push(field.trim());
      if (line[i] === ",") i++;
    } else {
      // Unquoted field
      const end = line.indexOf(",", i);
      if (end === -1) {
        fields.push(line.slice(i).trim());
        break;
      } else {
        fields.push(line.slice(i, end).trim());
        i = end + 1;
      }
    }
  }
  return fields;
}

function parseCSV(raw) {
  const lines = raw.trim().split("\n").filter(l => l.trim());
  if (lines.length < 2) return { headers: [], rows: [] };

  const headers = parseCSVLine(lines[0]).map(h => h.toLowerCase());
  const rows = [];

  for (let i = 1; i < lines.length; i++) {
    const cols = parseCSVLine(lines[i]);
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

  if (rows.length > MAX_RECIPIENTS) {
    log(`🛑 CSV contains ${rows.length.toLocaleString()} rows — exceeds the Microsoft 365 hard limit of 10,000 outbound recipients per 24 hours. Trim your list before sending to avoid a corporate outbox lockout.`, "error");
    document.getElementById("recipientCount").textContent = `${rows.length.toLocaleString()} recipients — OVER LIMIT`;
    return;
  }

  csvHeaders = headers;
  parsedRecipients = rows;

  // Auto-populate tag bar with any new headers from this CSV
  headers.forEach(h => {
    const tag = `{{${h}}}`;
    if (!DEFAULT_TAGS.includes(tag)) {
      addTagToBar(tag);
    }
  });

  const countEl = document.getElementById("recipientCount");
  countEl.textContent = `${rows.length} recipient${rows.length !== 1 ? "s" : ""} loaded`;

  renderPreviewTable(headers, rows.slice(0, 5));
  log(`Parsed ${rows.length} recipients. Showing first 5 in preview.`, "success");
}

function renderPreviewTable(headers, rows) {
  const container = document.getElementById("previewTable");
  container.classList.remove("hidden");

  let html = `<label class="label">Preview (first 5 rows)</label><table class="preview-table"><thead><tr>`;
  headers.forEach(h => { html += `<th>${escapeHtml(h)}</th>`; });
  html += `</tr></thead><tbody>`;
  rows.forEach(row => {
    html += "<tr>";
    headers.forEach(h => { html += `<td>${escapeHtml(row[h] || "")}</td>`; });
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

function findUnresolvedTokens(text) {
  const matches = text.match(/\{\{[^}]+\}\}/g);
  return matches ? [...new Set(matches)] : [];
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

/* ─── EMAIL VALIDATION ─────────────────────────────────────────── */

function validateRecipients(recipients) {
  const valid = [];
  const invalid = [];
  recipients.forEach((r, idx) => {
    if (EMAIL_REGEX.test(r.email)) {
      valid.push(r);
    } else {
      invalid.push({ row: idx + 2, email: r.email }); // +2: 1-based + header row
    }
  });
  return { valid, invalid };
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

/* ─── PRE-FLIGHT CHECKS ────────────────────────────────────────── */

async function checkMailboxReady(token) {
  const response = await fetch(GRAPH_MAILBOX_URL, {
    headers: { "Authorization": `Bearer ${token}` }
  });
  if (!response.ok) {
    throw new Error(`Mailbox pre-flight failed (${response.status}). The account may be unprovisioned or undergoing migration.`);
  }
}

function checkPayloadSize(bodyTemplate) {
  const bytes = new TextEncoder().encode(bodyTemplate).length;
  if (bytes > MAX_PAYLOAD_BYTES) {
    log(`⚠️ Email body is ~${(bytes / 1024 / 1024).toFixed(2)} MB — dangerously close to the 4 MB Graph API per-message limit. Compress or remove inline images before sending.`, "warning");
  }
}

/* ─── GRAPH BATCH SEND ─────────────────────────────────────────── */

function buildEmailRequest(id, toEmail, subject, htmlBody, saveToSent) {
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
      saveToSentItems: saveToSent
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

async function sendBatchWithRetry(requests, token) {
  for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
    const response = await fetch(GRAPH_BATCH_URL, {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${token}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({ requests })
    });

    // Top-level 429: the entire batch request was throttled
    if (response.status === 429) {
      const retryAfter = parseInt(response.headers.get("Retry-After") || "10", 10);
      if (attempt < MAX_RETRIES) {
        log(`Rate limited (429). Retrying batch in ${retryAfter}s (attempt ${attempt}/${MAX_RETRIES})...`, "warning");
        await delay(retryAfter * 1000);
        continue;
      } else {
        throw new Error(`Rate limited after ${MAX_RETRIES} retries.`);
      }
    }

    if (!response.ok) {
      throw new Error(`Batch HTTP error: ${response.status} ${response.statusText}`);
    }

    const data = await response.json();
    const responses = data.responses || [];

    // Check for per-request 429s and retry those individual requests
    const throttled = responses.filter(r => r.status === 429);
    if (throttled.length > 0 && attempt < MAX_RETRIES) {
      const retryAfterHeader = throttled
        .map(r => parseInt((r.headers || {})["Retry-After"] || "10", 10))
        .reduce((a, b) => Math.max(a, b), 10);
      const throttledIds = new Set(throttled.map(r => r.id));
      const retryRequests = requests.filter(r => throttledIds.has(r.id));
      log(`${throttled.length} request(s) throttled. Retrying in ${retryAfterHeader}s (attempt ${attempt}/${MAX_RETRIES})...`, "warning");
      await delay(retryAfterHeader * 1000);
      // Merge successful responses with retried ones
      const okResponses = responses.filter(r => r.status !== 429);
      const retryResult = await sendBatchWithRetry(retryRequests, token);
      return { responses: [...okResponses, ...(retryResult.responses || [])] };
    }

    return data;
  }
}

/* ─── CONFIRMATION MODAL ───────────────────────────────────────── */

let pendingMergeResolve = null;

function showConfirmModal(count) {
  return new Promise((resolve) => {
    pendingMergeResolve = resolve;
    document.getElementById("confirmText").textContent =
      `This will send ${count} personalized email${count !== 1 ? "s" : ""}. Continue?`;
    document.getElementById("confirmModal").classList.remove("hidden");
  });
}

function confirmSend() {
  dismissModal();
  if (pendingMergeResolve) pendingMergeResolve(true);
}

function dismissModal() {
  document.getElementById("confirmModal").classList.add("hidden");
  if (pendingMergeResolve) pendingMergeResolve(false);
  pendingMergeResolve = null;
}

/* ─── CANCEL ───────────────────────────────────────────────────── */

function handleStop() {
  cancelRequested = true;
  log("Stop requested — will halt after current batch.", "warning");
  document.getElementById("stopBtn").disabled = true;
}

function setMergeRunning(running) {
  mergeRunning = running;
  const mergeBtn = document.getElementById("mergeBtn");
  const stopBtn = document.getElementById("stopBtn");
  mergeBtn.disabled = running;
  mergeBtn.textContent = running ? "⏳ Sending..." : "▶ Run mail merge";
  stopBtn.disabled = false;
  stopBtn.classList.toggle("hidden", !running);
}

/* ─── MAIN MERGE RUNNER ────────────────────────────────────────── */

async function handleMergeClick() {
  // 1. Ensure recipients are parsed
  if (parsedRecipients.length === 0) {
    parseAndPreview();
    if (parsedRecipients.length === 0) {
      log("No valid recipients found. Check your CSV format.", "error");
      return;
    }
  }

  // 2. Validate subject before doing anything expensive
  const subjectTemplate = document.getElementById("subjectInput").value.trim();
  if (!subjectTemplate) {
    log("Subject line is empty.", "error");
    return;
  }

  // 3. Validate email addresses
  const { valid, invalid } = validateRecipients(parsedRecipients);
  if (invalid.length > 0) {
    invalid.forEach(({ row, email }) => {
      log(`Row ${row}: invalid email address "${email}" — skipped.`, "warning");
    });
  }
  if (valid.length === 0) {
    log("No valid email addresses found.", "error");
    return;
  }

  // 4. Show confirmation modal
  const confirmed = await showConfirmModal(valid.length);
  if (!confirmed) {
    log("Merge cancelled.", "info");
    return;
  }

  // 5. Read email body template
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
    return;
  }

  // 6. Payload size pre-flight
  checkPayloadSize(emailBodyTemplate);

  // 7. Mailbox provisioning pre-flight
  let preflightToken;
  try {
    preflightToken = await getAccessToken();
    await checkMailboxReady(preflightToken);
    log("Mailbox pre-flight OK.", "info");
  } catch (err) {
    log(`Pre-flight error: ${err.message}`, "error");
    return;
  }

  // 8. Check for unresolved tokens on the first recipient as a representative sample
  const sampleSubject = personalize(subjectTemplate, valid[0]);
  const sampleBody = personalize(emailBodyTemplate, valid[0]);
  const unresolvedSubject = findUnresolvedTokens(sampleSubject);
  const unresolvedBody = findUnresolvedTokens(sampleBody);
  const allUnresolved = [...new Set([...unresolvedSubject, ...unresolvedBody])];
  if (allUnresolved.length > 0) {
    log(`Warning: unresolved token(s) found — these will be sent literally: ${allUnresolved.join(", ")}`, "warning");
  }

  // 9. Run merge
  const saveToSent = document.getElementById("saveToSentItems").checked;
  setMergeRunning(true);
  cancelRequested = false;
  log("Starting mail merge...", "info");

  const requests = valid.map((recipient, idx) => {
    const personalizedSubject = personalize(subjectTemplate, recipient);
    const personalizedBody = personalize(emailBodyTemplate, recipient);
    return buildEmailRequest(idx + 1, recipient.email, personalizedSubject, personalizedBody, saveToSent);
  });

  const batches = chunkArray(requests, BATCH_SIZE);
  const totalBatches = batches.length;
  let totalSent = 0;
  let totalFailed = 0;

  log(`${valid.length} emails split into ${totalBatches} batch${totalBatches > 1 ? "es" : ""} of up to ${BATCH_SIZE}.`, "info");

  for (let i = 0; i < batches.length; i++) {
    if (cancelRequested) {
      log("Merge stopped by user.", "warning");
      break;
    }

    const batch = batches[i];
    log(`Sending batch ${i + 1}/${totalBatches} (${batch.length} emails)...`, "info");

    // Re-fetch token per batch to avoid expiry on long runs
    let token;
    try {
      token = await getAccessToken();
    } catch (err) {
      log(`Authentication error on batch ${i + 1}: ${err.message}`, "error");
      totalFailed += batch.length;
      break;
    }

    try {
      const result = await sendBatchWithRetry(batch, token);
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

    if (i < batches.length - 1 && !cancelRequested) {
      log(`Waiting ${BATCH_DELAY_MS / 1000}s before next batch to avoid throttling...`, "info");
      await delay(BATCH_DELAY_MS);
    }
  }

  log(`─── Merge complete. ✅ Sent: ${totalSent} | ❌ Failed: ${totalFailed} ───`, totalFailed > 0 ? "warning" : "success");
  setMergeRunning(false);
}
