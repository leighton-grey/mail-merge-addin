/* Mail Merge Engine v2.6.0
 * Auth: Office.js SSO -> Microsoft Graph API
 * Batching: 20 requests per Graph $batch call (reduced dynamically when attachment present)
 * Privacy: Subject + CSV cached in browser localStorage only. Nothing stored server-side.
 *
 * v2.6.0 — 7 UX features for non-technical users:
 *   1. Quick email entry mode — toggle replaces CSV area with a simple one-per-line email textarea;
 *      add-in converts to internal CSV invisibly, shows recipient count in real time
 *   2. Import from Outlook To field — button reads native compose To addresses into the recipient
 *      list and clears Outlook's To field so the merge engine can populate it per-recipient
 *   3. Smart paste detection — paste handler detects tab-separated content (copied from Excel /
 *      Google Sheets) and auto-converts tabs→commas before inserting into CSV textarea
 *   4. Drag-and-drop CSV/Excel files — drop zone overlay on the CSV textarea; dropped files are
 *      processed identically to the file picker (supports .csv, .xlsx, .xls)
 *   5. Prominent Step 1 test row — "Send test to myself" extracted from action buttons into its
 *      own visually distinct row (brand orange pill, "Step 1 — Test first" label) above the footer
 *   6. Plain-English pre-send summary — confirmation modal now shows a human-readable card:
 *      recipient count, estimated time, and first email preview (To address + personalised subject)
 *   7. accentColor updated to brand orange #F58220 in manifest; manifest icons changed from
 *      external URLs to bundled PNGs inside the zip to fix M365 Admin Center upload validation
 *
 * v2.4.0 — 6 bug fixes (merge complete timer guard, cancel vs done, bare localStorage,
 *   failedRecipients cleared on reload, mergeInProgress user message, broadcast re-entrancy);
 *   10 UX/a11y fixes (footer flex-wrap, ARIA progress bar, static tag chip keyboard, modal ARIA
 *   roles+labels, Escape key modal dismiss, send report columns, scheduling active badge, retry
 *   button cleared on reload, template restores expand accordion, CSV paste change event)
 *
 * v2.3.0 — Logical fixes (sendBatchWithRetry non-recursive, daily cap preserves failed list,
 *   broadcast skip_if, record_count mixed runs, inverted window guard, broadcast sendOutcomes,
 *   simulate skip_if parity, CSV row stamps); Performance (simulate async yield, innerHTML loop
 *   batching, TOKEN_REGEX hoisted, _originalIndex stamp, saveRateLimitState debounced);
 *   Security (handleCheckErrors XSS, testRowSelect XSS, unsubscribe_link href escape,
 *   fill-in secondary substitution, DNS disclosure log, CRLF header strip, EMAIL_REGEX
 *   tightened, CDN crossorigin)
 *
 * v2.2.0 — 16 QoL features: test send row selector, broadcast BCC cap warning, sample CSV download,
 *   merge_table/unsubscribe_link in UI, progress bar persist, preview paging, sendOutcomes
 *   accumulation, rate limit persistence, auth expiry guidance, simulate/dry-run mode,
 *   template saves scheduling settings, mixed send breakdown in confirmation, first-run banner,
 *   tag keyboard nav, modal focus trap, ISO date pruning fix
 *
 * v2.1.1 — 14 bug fixes: double-send guard, broadcast opt-out check, merge try/finally,
 *   429 retry depth limit, broadcast empty toRecipients, fill-in $ replacement, retry finally,
 *   contacts pagination, pre-send time estimate, DKIM selector2, optout XSS, group empty email log,
 *   spurious setMergeRunning, sendOutcomes accumulation
 *
 * v2.1.0 — 8 bug fixes (insertTag Mac coercion, fill-in isolation, merge table HTML escaping,
 *   CSV delimiter auto-detect, broadcast plain-text strip, checkErrors false positives,
 *   confirm modal, optional chaining compat); 5 features (unsubscribe_link token,
 *   pre-send confirmation, send jitter, signature preservation warning, CSV auto-detect)
 *
 * v1.2.0 — CC/BCC, Reply-To, per-recipient attachments, {{greeting_line}}, Send As, dedup, schedule
 * v1.2.1 — Importance, read/delivery receipts, plain text mode, custom X- headers, multi-TO per row
 * v1.3.0 — Sensitivity, categories, template save/load, retry failed, test send, log export, suppression
 * v1.4.0 bug fixes:
 *   - Subject personalization no longer HTML-encodes values (only HTML body does — prevents &amp; in subjects)
 *   - CSV rows with fewer columns than headers are padded instead of silently dropped
 *   - parsedRecipients cleared when an over-limit CSV is loaded (prevents stale list being sent)
 *   - Test send button disabled while a merge is running (prevents concurrent Graph calls)
 *   - Scheduled send now tracks failedRecipients for retry
 *   - Removed unused errorCount variable in per-recipient file handler
 * v1.4.0 additions:
 *   - Preview all: step through every recipient's personalised subject + body before sending
 *   - Progress bar: live visual percentage during send
 *   - Configurable batch delay: user-adjustable inter-batch pause (0–10 s)
 *   - Send summary report: downloadable CSV of per-recipient send outcomes
 * v1.4.1 bug fixes:
 *   - personalize(): single-pass replacement prevents double-substitution; regex no longer built from
 *     column names, eliminating regex injection for headers containing . + * ? etc.
 *   - Smart tokens {{today}} and {{now}} added (Classic Outlook Date field equivalent)
 *   - resolveAttachmentForRecipient(): per-missing-filename warning emitted once, not once per row
 *   - handleTestSend(): now applies replyTo, sendAs, receipts, and custom headers — matches real merge
 *   - Confirm modal shows unique address count for multi-TO rows
 *   - Sensitivity setting logged at merge start alongside priority
 * v1.5.0 additions:
 *   1. Multiple shared attachments — sharedAttachments[] array, multi-file picker
 *   2. TO display name — display_name CSV column sets the name on the first TO address
 *   3. BCC self — checkbox appends sender's address to every email's BCC list
 *   4. Follow-up flag — checkbox sets message.flag = { flagStatus: "flagged" } on sent items
 *   5. Message expiry — datetime-local input sets message.expiryDateTime (Exchange deferred delivery)
 *   6. List-Unsubscribe header — RFC 2369/8058 headers added per-recipient with token support
 *   7. Conditional merge — {{if:column=value:true text:false text}} syntax in subject/body
 *   8. Inline CID images — inlineImages Map, CID embedding via isInline + contentId
 *   buildEmailRequest refactored from 17 positional params to opts object
 * v1.5.1 — Tabbed UI redesign: Recipients / Compose / Options tabs, sticky footer with action
 *           buttons always visible, accordion sections for Scheduling and Advanced options,
 *           mini log in footer
 * v1.6.0 — Bug fixes: greeting_line double-escape, processConditionals colon in trueText,
 *           sendScheduledMessages saveToSent warning, parseAddressList multi-delimiter,
 *           validateRecipients logs invalid multi-TO addresses, handleRetryFailed state
 *           corruption, parseCSV RFC 4180 embedded newlines, handleTestSend attachment size guard.
 *           New features: Import from Outlook Contacts, skip_if column, merge field formatting
 *           pipes ({{field|upper}} etc.), recipient list filter/sort UI, templates save full settings.
 * v1.7.0 — Match Fields dialog, body template save/load, configurable greeting line,
 *           per-recipient send_at scheduling, Save as Drafts mode.
 * v1.8.0 — Check for Errors pre-flight, row checkboxes, contact group import, in-compose field picker,
 *           {{record_num}}/{{record_count}} tokens.
 * v1.9.0 — 8 bug fixes (record_num offset, test send tokens, checkbox state, CSV escaping, dead code);
 *           broadcast mode, multi-criteria AND/OR filter, duplicate send guard, GAL directory import.
 * v1.9.1 — 15 bug fixes: initTabs scope, validateRecipients null guard, broadcast validation/suppression/multi-TO/toRecipients, retry restore logic, stop-button race, directory search encoding, duplicate guard multi-TO, contact import header matching, log row numbers, opt-smart class, dead code cleanup
 * v1.9.2 — 11 open bug fixes (broadcast options/images/suppression, retry/draft CSV restore, filter reset, preview-all filtered set, record_num offset, test send guard); Mac compatibility (setSelectedDataAsync text coercion, setAsync image warning, error 13012); Exchange/hybrid detection (restUrl startup check, IMAP messaging)
 * v2.0.0 — 7 new features:
 *   1. Token fallback values: {{field|Fallback Text}} — empty fields use the fallback literal
 *   2. Sending window / business hours: restrict sends to Mon–Fri 09:00–17:00 (configurable)
 *   3. Max emails per hour / daily cap: rate limiting with automatic pause
 *   4. Many-to-one grouped row merge: group rows by email, {{merge_table}} token
 *   5. Managed persistent opt-out / unsubscribe list: localStorage suppression with UI
 *   6. Fill-in prompt at merge start: {{fill_in:Prompt}} tokens show modal before send
 *   7. SPF/DKIM/DMARC DNS pre-flight check: warns for large sends (500+) if DNS records missing
 */

const GRAPH_BATCH_URL     = "https://graph.microsoft.com/v1.0/$batch";
const BATCH_SIZE          = 20;
const BATCH_DELAY_MS      = 1500;   // default — overridden at runtime by batchDelayInput
const MAX_RETRIES         = 3;
const EMAIL_REGEX         = /^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)+$/;
// P3: hoisted to module scope — prevents recompilation on every personalize() call
const TOKEN_REGEX         = /\{\{([^}|]+)(?:\|([^}]*))?\}\}/gi;
const MAX_RECIPIENTS      = 10000;
const MAX_PAYLOAD_BYTES   = 3.5 * 1024 * 1024;
const MAX_ATTACHMENT_BYTES = 3 * 1024 * 1024;

const LS_KEY_SUBJECT      = "mailmerge_subject";
const LS_KEY_CSV          = "mailmerge_csv";
const LS_KEY_TAGS         = "mailmerge_custom_tags";
const LS_KEY_TEMPLATES    = "mailmerge_templates";
const LS_KEY_GREETING     = "mailmerge_greeting";
const LS_KEY_SEND_HISTORY = "mailmerge_send_history"; // { email: isoTimestamp }

const CANONICAL_FIELDS = [
  { key: "email",        label: "Email *",        required: true  },
  { key: "first_name",   label: "First name",      required: false },
  { key: "last_name",    label: "Last name",       required: false },
  { key: "salutation",   label: "Salutation",      required: false },
  { key: "gender",       label: "Gender",          required: false },
  { key: "company",      label: "Company",         required: false },
  { key: "title",        label: "Job title",       required: false },
  { key: "cc",           label: "CC",              required: false },
  { key: "bcc",          label: "BCC",             required: false },
  { key: "reply_to",     label: "Reply-To",        required: false },
  { key: "attachment",   label: "Attachment file", required: false },
  { key: "display_name", label: "Display name",    required: false },
  { key: "skip_if",      label: "Skip if",         required: false },
  { key: "send_at",      label: "Send at",         required: false },
];

const DEFAULT_TAGS = [
  "{{first_name}}", "{{last_name}}", "{{email}}", "{{company}}", "{{title}}",
  "{{greeting_line}}", "{{today}}", "{{now}}", "{{record_num}}", "{{record_count}}",
  "{{merge_table}}", "{{unsubscribe_link}}"
];

let mergeInProgress           = false;  // BUG 1: double-send guard
let broadcastInProgress       = false;  // UX Bug 6: re-entrancy guard for handleBroadcast
let _mergeCompletedSuccessfully = false; // Feature 5: suppress progress hide on success
let previewTablePage          = 0;      // Feature 6: preview table paging
const PREVIEW_PAGE_SIZE       = 10;     // Feature 6: rows per page
let parsedRecipients          = [];
let cancelRequested           = false;
let subjectHasFocus           = false;  // true while #subjectInput has keyboard focus

// _subjectWasFocused: snapshot of subjectHasFocus captured on mousedown of a tag chip.
// This is necessary because blur fires BEFORE click in the browser event order —
// by the time the click handler runs, subjectHasFocus is already false. We snapshot it
// during mousedown (which fires before blur) so insertTag() knows whether to target
// the subject line or the email body.
let _subjectWasFocused        = false;

// _hintResetTimer: debounce handle for resetting the #tagsHint text after
// the subject input loses focus. The 150 ms delay covers the mousedown→blur→click
// sequence, so the hint doesn't flip to "body" before the click completes.
let _hintResetTimer           = null;

// tagTarget: controls where clicking a tag chip inserts the token.
//   "body"    → setSelectedDataAsync (cursor position in Outlook compose body)
//   "subject" → subject.getAsync + setAsync (appends to Outlook native subject)
// Toggled by the Body / Subject segmented control in the tag header.
// Note: when #subjectInput itself is focused, insertTag() always writes
// into that field regardless of tagTarget (subjectHasFocus takes priority).
let tagTarget                 = "body";
let _lastOutlookSubject       = null;   // last value read from / pushed to Outlook's native subject
let _subjectPushTimer         = null;   // debounce timer for pushing to Outlook
let _subjectPollTimer         = null;   // interval for polling Outlook subject
let sharedAttachments         = [];   // array of { name, contentType, contentBytes, sizeBytes }
let perRecipientFiles         = new Map();
let inlineImages              = new Map(); // filename.toLowerCase() → { name, contentType, contentBytes, sizeBytes }
let suppressionSet            = new Set();
let failedRecipients          = [];
let sendOutcomes              = [];   // { email, status, timestamp, error? } — for summary report
let previewBodyTemplate       = "";   // cached body for preview-all mode
const warnedMissingAttachments = new Set(); // deduplicate per-filename warnings across a merge run
let previewIndex        = 0;    // current row index in preview navigator
let previewRecipients   = [];   // filtered recipient set used by preview-all modal (A10)

let fieldMapping = {}; // { canonicalName: csvColumnName }
let greetingConfig = { format: "dear_sal_last", fallback: "Dear Valued Customer" };
let draftsMode = false;

let selectedRowIndices = null; // null = all selected; Set of original parsedRecipients indices when subset
let editTableHeaders = [];   // working copy of column names during an edit session
let editTableRows    = [];   // working copy: array of string arrays (one per recipient)

// Contacts Groups tab state
let groupsData = [];
let selectedGroups = new Set(); // group IDs
let contactsActiveTab = "contacts"; // "contacts" | "groups" | "directory"

// Directory tab state
let directoryData = [];
let selectedDirectory = new Set();

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    log("Office.js ready. Host: Outlook.", "info");

    // C1 / C2: Check for Exchange Online mailbox — restUrl is null for IMAP/Gmail-only accounts
    const restUrl = Office.context.mailbox.restUrl;
    if (!restUrl) {
      const banner = document.createElement("div");
      banner.style.cssText = "background:#eb5757;color:#fff;padding:10px 14px;font-size:12px;line-height:1.5;";
      banner.innerHTML = `<strong>&#9888; No Exchange Online mailbox detected.</strong><br>
        This add-in sends email via Microsoft Graph and requires an Exchange Online mailbox.<br><br>
        <strong>Common causes:</strong><br>
        &bull; Your Microsoft 365 license includes Office apps but not Exchange Online
        (e.g. M365 Apps for Business/Enterprise without an email plan).<br>
        &bull; Outlook is configured with a Gmail or other IMAP account rather than
        a Microsoft 365 Exchange Online account.<br>
        &bull; If you use <strong>JumpCloud</strong> as your identity provider, your user account
        may not be synced to the M365 Cloud Directory Integration — contact your IT administrator
        to ensure your account is provisioned with an Exchange Online mailbox.<br><br>
        <em>Note: having Google Workspace for email does not prevent this add-in from working —
        you just need an M365 account with Exchange Online as well, and Outlook must be signed in
        with that M365 account (not the Google account).</em><br><br>
        <a href="https://learn.microsoft.com/en-us/microsoft-365/admin/misc/why-cant-i-do-mail-merge"
           style="color:#fff;text-decoration:underline;" target="_blank">Learn more about Exchange Online requirements</a>`;
      document.body.insertBefore(banner, document.body.firstChild);
      ["mergeBtn","testSendBtn","previewAllBtn","saveDraftsBtn","broadcastBtn"].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.disabled = true;
      });
    }

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

    document.getElementById("uploadAttachmentBtn").addEventListener("click", () => {
      document.getElementById("attachmentFileInput").click();
    });
    document.getElementById("attachmentFileInput").addEventListener("change", handleAttachmentUpload);
    document.getElementById("clearSharedAttachmentsBtn").addEventListener("click", clearSharedAttachments);

    document.getElementById("uploadPerRecipientBtn").addEventListener("click", () => {
      document.getElementById("perRecipientFilesInput").click();
    });
    document.getElementById("perRecipientFilesInput").addEventListener("change", handlePerRecipientFilesUpload);
    document.getElementById("clearPerRecipientBtn").addEventListener("click", clearPerRecipientFiles);

    document.getElementById("uploadInlineImagesBtn").addEventListener("click", () => {
      document.getElementById("inlineImagesInput").click();
    });
    document.getElementById("inlineImagesInput").addEventListener("change", handleInlineImagesUpload);
    document.getElementById("clearInlineImagesBtn").addEventListener("click", clearInlineImages);

    document.getElementById("scheduleEnabled").addEventListener("change", (e) => {
      document.getElementById("scheduleRow").classList.toggle("hidden", !e.target.checked);
      updateSchedulingBadge();
    });

    document.getElementById("expiryEnabled").addEventListener("change", (e) => {
      document.getElementById("expiryRow").classList.toggle("hidden", !e.target.checked);
    });

    const sendingWindowEnabledEl = document.getElementById("sendingWindowEnabled");
    if (sendingWindowEnabledEl) {
      sendingWindowEnabledEl.addEventListener("change", function(e) {
        const sendingWindowRowEl = document.getElementById("sendingWindowRow");
        if (sendingWindowRowEl) sendingWindowRowEl.classList.toggle("hidden", !e.target.checked);
        updateSchedulingBadge();
      });
    }

    document.getElementById("customHeadersEnabled").addEventListener("change", (e) => {
      document.getElementById("customHeadersRow").classList.toggle("hidden", !e.target.checked);
    });

    document.getElementById("saveTemplateBtn").addEventListener("click", saveTemplate);
    document.getElementById("deleteTemplateBtn").addEventListener("click", deleteTemplate);
    document.getElementById("templateSelect").addEventListener("change", loadTemplate);

    document.getElementById("testSendBtn").addEventListener("click", handleTestSend);
    document.getElementById("previewAllBtn").addEventListener("click", handlePreviewAll);

    document.getElementById("retryFailedBtn").addEventListener("click", handleRetryFailed);
    document.getElementById("downloadLogBtn").addEventListener("click", downloadLog);
    document.getElementById("downloadReportBtn").addEventListener("click", downloadSendReport);

    document.getElementById("uploadSuppressionBtn").addEventListener("click", () => {
      document.getElementById("suppressionFileInput").click();
    });
    document.getElementById("suppressionFileInput").addEventListener("change", handleSuppressionUpload);
    document.getElementById("clearSuppressionBtn").addEventListener("click", clearSuppression);

    // Preview modal navigation
    document.getElementById("previewCloseBtn").addEventListener("click", closePreviewModal);
    document.getElementById("previewPrevBtn").addEventListener("click", () => {
      if (previewIndex > 0) { previewIndex--; renderPreviewEntry(); }
    });
    document.getElementById("previewNextBtn").addEventListener("click", () => {
      if (previewIndex < previewRecipients.length - 1) { previewIndex++; renderPreviewEntry(); }
    });

    // Batch delay label live update
    const batchDelayInput = document.getElementById("batchDelayInput");
    batchDelayInput.addEventListener("input", () => {
      const sec = parseFloat(batchDelayInput.value);
      document.getElementById("batchDelayLabel").textContent =
        sec === 0 ? "(no delay between batches)" : `(${sec} s between batches)`;
    });

    // Helper: set tagTarget and keep the Body/Subject toggle pills in sync.
    // Call this from any focus/blur/click handler so the UI always reflects reality.
    function applyTagTarget(target) {
      tagTarget = target;
      const bodyBtn    = document.getElementById("targetBodyBtn");
      const subjectBtn = document.getElementById("targetSubjectBtn");
      if (target === "subject") {
        subjectBtn.classList.add("active");
        bodyBtn.classList.remove("active");
      } else {
        bodyBtn.classList.add("active");
        subjectBtn.classList.remove("active");
      }
    }

    const subjectInput = document.getElementById("subjectInput");
    // On focus: mark that subject has focus, auto-switch toggle to Subject,
    // and update the tag hint so users know clicking a tag targets the subject.
    subjectInput.addEventListener("focus", () => {
      subjectHasFocus = true;
      applyTagTarget("subject");
      clearTimeout(_hintResetTimer);  // cancel any pending reset from a previous blur
      updateTagHint(true);
    });
    // On blur: mark focus as lost, auto-switch toggle back to Body,
    // then reset the hint after 150 ms.
    // The 150 ms delay covers the full mousedown→blur→click sequence —
    // if the user clicked a tag chip, the click fires within ~50-80 ms
    // of blur, so the hint (and _subjectWasFocused) are still correct
    // when insertTag() runs.
    subjectInput.addEventListener("blur", () => {
      subjectHasFocus = false;
      applyTagTarget("body");
      _hintResetTimer = setTimeout(() => updateTagHint(false), 150);
    });

    // Cross-frame focus fix: when the user clicks in the Outlook compose area
    // (a different iframe or native window), the subjectInput blur event does NOT
    // fire — cross-frame clicks are invisible to the source frame's blur listeners.
    // However, the taskpane's *window* does lose focus reliably.
    // We auto-switch to Body because clicking into Outlook almost always means
    // the user intends to place the cursor in the compose body.
    // (Native subject vs body is indistinguishable from inside the task pane iframe.)
    window.addEventListener("blur", () => {
      if (subjectHasFocus) {
        subjectHasFocus = false;
        clearTimeout(_hintResetTimer);
        updateTagHint(false);
      }
      // Always snap back to Body when task pane loses focus — native Outlook click.
      applyTagTarget("body");
    });

    // mousedown on tagBar: snapshot whether subject had focus BEFORE blur fires.
    // We check e.target to make sure the user isn't clicking the ✕ remove button —
    // that click should never trigger an insert.
    document.getElementById("tagBar").addEventListener("mousedown", (e) => {
      if (!e.target.classList.contains("tag-remove-btn")) {
        _subjectWasFocused = subjectHasFocus;
      }
    });

    // click on tagBar: use closest() so clicking the inner .tag-label span
    // (or anywhere in the chip) still finds the parent [data-tag] element.
    // Skip the click entirely if the remove button was clicked.
    document.getElementById("tagBar").addEventListener("click", (e) => {
      if (e.target.classList.contains("tag-remove-btn")) return; // handled by its own listener
      const chip = e.target.closest("[data-tag]");
      if (chip) insertTag(chip.dataset.tag, _subjectWasFocused);
    });

    // Body / Subject toggle — manual override. applyTagTarget keeps the pills in sync.
    document.getElementById("targetBodyBtn").addEventListener("click", () => {
      applyTagTarget("body");
    });
    document.getElementById("targetSubjectBtn").addEventListener("click", () => {
      applyTagTarget("subject");
    });

    document.getElementById("customTagInput").addEventListener("keydown", (e) => {
      if (e.key === "Enter") addCustomTag();
    });

    document.getElementById("subjectInput").addEventListener("input", () => {
      const val = document.getElementById("subjectInput").value;
      lsSet(LS_KEY_SUBJECT, val);
      // Push to Outlook's native subject field (debounced 400 ms)
      clearTimeout(_subjectPushTimer);
      _subjectPushTimer = setTimeout(() => pushSubjectToOutlook(val), 400);
    });

    // Sync button: pull current Outlook subject into taskpane on demand
    const syncSubjectBtn = document.getElementById("syncSubjectBtn");
    if (syncSubjectBtn) {
      syncSubjectBtn.addEventListener("click", () => syncSubjectFromOutlook(true));
    }

    const _csvInputEl = document.getElementById("csvInput");
    const _saveCsv = () => lsSet(LS_KEY_CSV, _csvInputEl.value);
    _csvInputEl.addEventListener("input", _saveCsv);
    _csvInputEl.addEventListener("change", _saveCsv);

    restoreLocalState();

    // Sync subject from Outlook's native compose field on first load,
    // then poll every 500 ms for live-like bidirectional sync between
    // Outlook's compose field and the taskpane subject input.
    syncSubjectFromOutlook(true);
    _subjectPollTimer = setInterval(() => syncSubjectFromOutlook(false), 500);

    initTabs();
    initAccordions();
    initOptOutUI();

    document.getElementById("toggleLogBtn").addEventListener("click", () => {
      document.getElementById("fullLogArea").classList.toggle("hidden");
    });

    // Import from Contacts (Feature 1)
    document.getElementById("importContactsBtn").addEventListener("click", handleImportContacts);
    document.getElementById("contactsCloseBtn").addEventListener("click", () => {
      document.getElementById("contactsModal").classList.add("hidden");
    });
    document.getElementById("contactsSearch").addEventListener("input", (e) => filterContacts(e.target.value));
    document.getElementById("contactsSelectAllBtn").addEventListener("click", () => {
      const visible = document.getElementById("contactsList").querySelectorAll("input[type=checkbox]");
      const allChecked = [...visible].every(cb => cb.checked);
      visible.forEach(cb => {
        cb.checked = !allChecked;
        if (!allChecked) selectedContacts.add(cb.dataset.email);
        else selectedContacts.delete(cb.dataset.email);
      });
      updateContactsSelectedCount();
    });
    document.getElementById("contactsImportBtn").addEventListener("click", () => {
      if (contactsActiveTab === "directory") importDirectorySelected();
      else importSelectedContacts();
    });

    // Filter / sort bar (Feature 4)
    document.getElementById("applyFilterBtn").addEventListener("click", applyFilterSort);
    document.getElementById("clearFilterBtn").addEventListener("click", clearFilterSort);

    // Match Fields modal (Feature 1)
    document.getElementById("mapFieldsBtn").addEventListener("click", () => {
      const headers = parsedRecipients.length > 0 ? Object.keys(parsedRecipients[0]).filter(k => !k.startsWith("_")) : [];
      openMatchFieldsModal(headers);
    });
    document.getElementById("matchFieldsCloseBtn").addEventListener("click", () => {
      document.getElementById("matchFieldsModal").classList.add("hidden");
    });
    document.getElementById("matchFieldsCancelBtn").addEventListener("click", () => {
      document.getElementById("matchFieldsModal").classList.add("hidden");
    });
    document.getElementById("matchFieldsApplyBtn").addEventListener("click", applyMatchFields);

    // Body template save/load (Feature 2)
    document.getElementById("saveBodyToTemplateBtn").addEventListener("click", saveBodyToTemplate);
    document.getElementById("loadBodyFromTemplateBtn").addEventListener("click", loadBodyFromTemplate);

    // Greeting line config (Feature 3)
    document.getElementById("greetingFormat").addEventListener("change", () => {
      greetingConfig.format = document.getElementById("greetingFormat").value;
      lsSet(LS_KEY_GREETING, JSON.stringify(greetingConfig));
    });
    document.getElementById("greetingFallback").addEventListener("input", () => {
      greetingConfig.fallback = document.getElementById("greetingFallback").value.trim() || "Dear Valued Customer";
      lsSet(LS_KEY_GREETING, JSON.stringify(greetingConfig));
    });

    // Save as Drafts (Feature 5)
    document.getElementById("saveDraftsBtn").addEventListener("click", handleSaveDrafts);

    // Check for Errors (Feature 1 v1.8)
    document.getElementById("checkErrorsBtn").addEventListener("click", handleCheckErrors);
    document.getElementById("checkErrorsCloseBtn").addEventListener("click", () => {
      _closeModalWithTrap("checkErrorsModal"); // Feature 15
    });

    // Edit recipient table modal
    document.getElementById("editTableCloseBtn").addEventListener("click", () => {
      document.getElementById("editTableModal").classList.add("hidden");
    });
    document.getElementById("editTableAddColBtn").addEventListener("click", editTableAddColumn);
    document.getElementById("editTableSaveBtn").addEventListener("click", saveEditTableChanges);
    document.getElementById("newColNameInput").addEventListener("keydown", function(e) {
      if (e.key === "Enter") { e.preventDefault(); editTableAddColumn(); }
    });

    // Insert field at cursor (Feature 4 v1.8)
    document.getElementById("insertFieldBtn").addEventListener("click", handleInsertField);

    // Broadcast mode (v1.9.0)
    document.getElementById("broadcastBtn").addEventListener("click", handleBroadcast);

    // Duplicate send history clear (v1.9.0)
    document.getElementById("clearSendHistoryBtn").addEventListener("click", () => {
      lsRemove(LS_KEY_SEND_HISTORY);
      log("Send history cleared.", "info");
    });

    // Multi-criteria filter (v1.9.0)
    document.getElementById("addFilterCondBtn").addEventListener("click", () => addFilterCondition());

    // Contacts modal tabs (Feature 3 v1.8)
    document.getElementById("contactsTabContacts").addEventListener("click", () => {
      contactsActiveTab = "contacts";
      document.getElementById("contactsTabContacts").classList.add("active");
      document.getElementById("contactsTabGroups").classList.remove("active");
      document.getElementById("contactsTabDirectory").classList.remove("active");
      document.getElementById("contactsList").style.display = "";
      document.getElementById("groupsList").style.display = "none";
      document.getElementById("directoryList").style.display = "none";
      document.getElementById("directorySearchRow").style.display = "none";
      document.getElementById("contactsSearch").style.display = "";
      document.getElementById("contactsSelectAllBtn").style.display = "";
      document.getElementById("contactsImportBtn").style.display = "";
      document.getElementById("contactsImportBtn").textContent = "Import selected";
      document.getElementById("contactsSelectedCount").style.display = "";
    });
    document.getElementById("contactsTabGroups").addEventListener("click", () => {
      contactsActiveTab = "groups";
      document.getElementById("contactsTabGroups").classList.add("active");
      document.getElementById("contactsTabContacts").classList.remove("active");
      document.getElementById("contactsTabDirectory").classList.remove("active");
      document.getElementById("contactsList").style.display = "none";
      document.getElementById("groupsList").style.display = "";
      document.getElementById("directoryList").style.display = "none";
      document.getElementById("directorySearchRow").style.display = "none";
      document.getElementById("contactsSearch").style.display = "none";
      document.getElementById("contactsSelectAllBtn").style.display = "none";
      document.getElementById("contactsImportBtn").style.display = "none";
      document.getElementById("contactsSelectedCount").style.display = "none";
      loadGroups();
    });
    document.getElementById("contactsTabDirectory").addEventListener("click", () => {
      contactsActiveTab = "directory";
      ["contactsTabContacts","contactsTabGroups","contactsTabDirectory"].forEach(id => {
        document.getElementById(id).classList.remove("active");
      });
      document.getElementById("contactsTabDirectory").classList.add("active");
      document.getElementById("contactsList").style.display = "none";
      document.getElementById("groupsList").style.display = "none";
      document.getElementById("directoryList").style.display = "";
      document.getElementById("directorySearchRow").style.display = "";
      document.getElementById("contactsSearch").style.display = "none";
      document.getElementById("contactsSelectAllBtn").style.display = "none";
      document.getElementById("contactsImportBtn").style.display = "";
      document.getElementById("contactsImportBtn").textContent = "Import selected";
      document.getElementById("contactsSelectedCount").style.display = "none";
      directoryData = [];
      selectedDirectory.clear();
      document.getElementById("directoryList").innerHTML =
        '<p style="padding:8px;color:#605e5c;">Type at least 2 characters to search.</p>';
    });

    // Directory search input — debounced
    let directorySearchTimer = null;
    document.getElementById("directorySearch").addEventListener("input", (e) => {
      clearTimeout(directorySearchTimer);
      directorySearchTimer = setTimeout(() => searchDirectory(e.target.value.trim()), 400);
    });

    // UX 3: keyboard activation for static tag chips (Enter / Space)
    document.addEventListener("keydown", (e) => {
      if ((e.key === "Enter" || e.key === " ") && e.target.classList.contains("tag")) {
        e.preventDefault();
        e.target.click();
      }
    });

    // UX 5: Escape key closes the topmost visible modal
    document.addEventListener("keydown", (e) => {
      if (e.key !== "Escape") return;
      const visibleModals = [
        "fillInModal", "preSendModal", "confirmModal", "dupeSendModal",
        "checkErrorsModal", "previewModal", "matchFieldsModal", "contactsModal",
        "editTableModal"
      ];
      for (const id of visibleModals) {
        const el = document.getElementById(id);
        if (el && !el.classList.contains("hidden")) {
          const cancelBtn = el.querySelector(
            "#fillInCancelBtn, #preSendCancelBtn, #confirmCancelBtn, #dupeSendCancelBtn, " +
            "#checkErrorsCloseBtn, #previewCloseBtn, #matchFieldsCancelBtn, #contactsCloseBtn, " +
            "#editTableCloseBtn, " +
            "button[class*='cancel'], button[class*='close']"
          );
          if (cancelBtn) cancelBtn.click();
          else el.classList.add("hidden");
          break;
        }
      }
    });

    // UX 8: hide retry button when CSV textarea is manually cleared
    document.getElementById("csvInput")?.addEventListener("input", () => {
      if (!document.getElementById("csvInput").value.trim()) {
        document.getElementById("retryFailedBtn")?.classList.add("hidden");
        failedRecipients = [];
      }
    });

    // ── Smart paste detection ──────────────────────────────────────
    // When a user pastes tab-separated content (Excel/Sheets copy),
    // automatically convert it to CSV so they don't need to save as .csv first.
    document.getElementById("csvInput")?.addEventListener("paste", (e) => {
      const raw = (e.clipboardData || window.clipboardData)?.getData("text");
      if (!raw) return;
      const lines = raw.split(/\r?\n/);
      // Tab-separated if the majority of lines contain tabs
      const tabLines = lines.filter(l => l.includes("\t")).length;
      if (tabLines < Math.max(1, lines.length * 0.5)) return; // not tab-separated
      e.preventDefault();
      // Convert: tabs → commas, quote fields that contain commas
      const csv = lines.map(line =>
        line.split("\t").map(cell => {
          if (cell.includes(",") || cell.includes('"') || cell.includes("\n")) {
            return '"' + cell.replace(/"/g, '""') + '"';
          }
          return cell;
        }).join(",")
      ).join("\n");
      const ta = document.getElementById("csvInput");
      const start = ta.selectionStart, end = ta.selectionEnd;
      ta.value = ta.value.slice(0, start) + csv + ta.value.slice(end);
      ta.selectionStart = ta.selectionEnd = start + csv.length;
      ta.dispatchEvent(new Event("input"));
      ta.dispatchEvent(new Event("change"));
      log("Detected Excel/Sheets table paste — converted to CSV automatically.", "info");
    });

    // ── Drag-and-drop CSV / Excel files ───────────────────────────
    const csvDropZone = document.getElementById("csvDropZone");
    if (csvDropZone) {
      csvDropZone.addEventListener("dragover", (e) => {
        e.preventDefault();
        csvDropZone.classList.add("drag-over");
      });
      csvDropZone.addEventListener("dragleave", () => {
        csvDropZone.classList.remove("drag-over");
      });
      csvDropZone.addEventListener("drop", (e) => {
        e.preventDefault();
        csvDropZone.classList.remove("drag-over");
        const file = e.dataTransfer?.files?.[0];
        if (!file) return;
        // Reuse the existing file upload handler by injecting the file into csvFileInput
        const dt = new DataTransfer();
        dt.items.add(file);
        const fi = document.getElementById("csvFileInput");
        fi.files = dt.files;
        fi.dispatchEvent(new Event("change"));
      });
    }

    // ── Quick email entry mode toggle ─────────────────────────────
    document.getElementById("quickModeBtn")?.addEventListener("click", toggleQuickMode);
    document.getElementById("quickEmailInput")?.addEventListener("input", () => {
      updateQuickModeCSV();
    });

    // ── Sync recipients from Outlook's native To field ────────────
    document.getElementById("syncRecipientsBtn")?.addEventListener("click", () => {
      syncRecipientsFromOutlook();
    });

    // Feature 8: load persisted rate limit state
    loadRateLimitState();

    // UX 7: wire rate limit badge to maxPerHour and dailyCap inputs
    document.getElementById("maxPerHour")?.addEventListener("input", updateRateLimitBadge);
    document.getElementById("dailyCap")?.addEventListener("input", updateRateLimitBadge);

    // Feature 3: sample CSV download
    document.getElementById("downloadSampleCsv")?.addEventListener("click", (e) => {
      e.preventDefault();
      const sampleCsv = [
        "email,first_name,last_name,company,title",
        "alice@example.com,Alice,Smith,Acme Corp,Director",
        "bob@example.com,Bob,Jones,Globex,Manager",
        "carol@example.com,Carol,Williams,Initech,Analyst"
      ].join("\r\n");
      const blob = new Blob([sampleCsv], { type: "text/csv" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "mail-merge-sample.csv";
      a.click();
      URL.revokeObjectURL(url);
    });

    // Feature 10: simulate button
    document.getElementById("simulateBtn")?.addEventListener("click", handleSimulate);

    // Feature 13: getting started banner
    const hasSeenWelcome = lsGet("mm_welcome_dismissed", null);
    const hasCsv = !!lsGet(LS_KEY_CSV, null);
    if (!hasSeenWelcome && !hasCsv) {
      document.getElementById("gettingStartedBanner")?.classList.remove("hidden");
    }
    document.getElementById("dismissGettingStarted")?.addEventListener("click", (e) => {
      e.preventDefault();
      document.getElementById("gettingStartedBanner")?.classList.add("hidden");
      lsSet("mm_welcome_dismissed", "1");
    });
  }
});

/* ─── SWIPE-BACK / HORIZONTAL SCROLL GUARD ────────────────────
   On macOS + Chrome, a two-finger horizontal swipe on a trackpad
   inside an Office add-in iframe can trigger the browser's
   back/forward navigation, navigating away from the compose window.

   Fix: intercept wheel events where the horizontal delta exceeds
   the vertical delta (i.e. a deliberate horizontal swipe), and
   call preventDefault() so Chrome never sees it as a navigation
   gesture. We only do this when the page itself has no horizontal
   scroll room (scrollWidth === clientWidth), meaning the swipe
   was never meant to scroll content anyway.
   ─────────────────────────────────────────────────────────── */
(function installSwipeGuard() {
  document.addEventListener("wheel", function(e) {
    // Only intercept clearly horizontal swipes
    if (Math.abs(e.deltaX) <= Math.abs(e.deltaY)) return;
    // If the element being scrolled has horizontal scroll room, let it scroll
    let el = e.target;
    while (el && el !== document.documentElement) {
      if (el.scrollWidth > el.clientWidth) return;
      el = el.parentElement;
    }
    // No horizontal scroll room anywhere in the chain — block the swipe
    e.preventDefault();
  }, { passive: false });
})();

/* ─── LOGGING ─────────────────────────────────────────────────── */

/**
 * Append a timestamped message to the full status log (#statusLog) and
 * also update the single-line mini log strip in the sticky footer.
 *
 * Log levels and their colour treatment:
 *   "info"    — muted grey-purple  — routine progress messages
 *   "success" — green              — successful email sends
 *   "warning" — amber              — non-fatal issues (skipped rows, etc.)
 *   "error"   — red                — hard failures requiring attention
 *
 * @param {string} message - Human-readable log message
 * @param {"info"|"success"|"warning"|"error"} [type="info"] - Log level
 */
function log(message, type = "info") {
  const logEl = document.getElementById("statusLog");
  const entry = document.createElement("p");
  entry.className = `log-entry log-${type}`;   // CSS colours defined in taskpane.css
  const time = new Date().toLocaleTimeString();
  entry.textContent = `[${time}] ${message}`;
  logEl.appendChild(entry);
  logEl.scrollTop = logEl.scrollHeight;  // auto-scroll so latest message is always visible

  // Mirror the latest message in the compact footer mini log strip.
  // The colour codes here match the CSS log-* colours but as inline
  // styles so they work on the dark footer background.
  const mini = document.getElementById("footerLogMini");
  if (mini) {
    mini.textContent = `[${time}] ${message}`;
    mini.style.color =
      type === "success" ? "#6fcf97" :
      type === "error"   ? "#eb5757" :
      type === "warning" ? "#f2994a" : "#c8c6c4";
  }
}

/**
 * Clear all entries from the status log and reset to the initial ready state.
 */
function clearLog() {
  document.getElementById("statusLog").innerHTML = "";
  log("Log cleared.", "info");
}

/* ─── OUTLOOK SUBJECT SYNC ──────────────────────────────────────
   Two-way bridge between the taskpane subject field and Outlook's
   native compose subject field, so non-technical users can type in
   either place and have them stay in sync automatically.
   ─────────────────────────────────────────────────────────────── */

/**
 * Read Outlook's native subject → update the taskpane input.
 * @param {boolean} force  true = always overwrite, false = only if Outlook changed since last sync
 */
function syncSubjectFromOutlook(force) {
  // Don't overwrite the taskpane while the user is actively typing in it
  if (!force && subjectHasFocus) return;
  try {
    Office.context.mailbox.item.subject.getAsync(result => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) return;
      const outlookSubject = result.value || "";
      // Skip if nothing changed since we last read/pushed (avoids echo-back loops)
      if (!force && outlookSubject === _lastOutlookSubject) return;
      _lastOutlookSubject = outlookSubject;
      const input = document.getElementById("subjectInput");
      if (!input) return;
      // Don't interrupt typing — recheck focus inside the async callback too
      if (!force && subjectHasFocus) return;
      if (input.value !== outlookSubject) {
        input.value = outlookSubject;
        lsSet(LS_KEY_SUBJECT, outlookSubject);
      }
    });
  } catch (e) { /* Office context unavailable — fail silently */ }
}

/**
 * Push the taskpane subject value to Outlook's native compose subject field.
 * @param {string} value
 */
function pushSubjectToOutlook(value) {
  try {
    _lastOutlookSubject = value; // record what we pushed so the next poll won't echo it back
    Office.context.mailbox.item.subject.setAsync(value, () => {});
  } catch (e) { /* fail silently */ }
}

/* ─── QUICK EMAIL ENTRY MODE ────────────────────────────────────
   A simple "paste or type emails one per line" mode that removes
   the CSV barrier for non-technical users doing small sends.
   ─────────────────────────────────────────────────────────────── */

let _quickModeActive = false;

function toggleQuickMode() {
  _quickModeActive = !_quickModeActive;
  const btn     = document.getElementById("quickModeBtn");
  const csvArea = document.getElementById("csvSection");
  const qArea   = document.getElementById("quickSection");

  if (_quickModeActive) {
    csvArea?.classList.add("hidden");
    qArea?.classList.remove("hidden");
    if (btn) btn.textContent = "📋 Switch to CSV mode";
    document.getElementById("quickEmailInput")?.focus();
  } else {
    csvArea?.classList.remove("hidden");
    qArea?.classList.add("hidden");
    if (btn) btn.textContent = "📧 Quick: type emails";
  }
}

let _quickModeParseTimer = null;

function updateQuickModeCSV() {
  const raw = document.getElementById("quickEmailInput")?.value || "";
  const emails = raw.split(/[\n,;]+/).map(e => e.trim()).filter(e => e && EMAIL_REGEX.test(e));
  const csvInput = document.getElementById("csvInput");
  if (!csvInput) return;
  if (emails.length === 0) {
    csvInput.value = "";
  } else {
    csvInput.value = "email\n" + emails.join("\n");
  }
  csvInput.dispatchEvent(new Event("input"));
  csvInput.dispatchEvent(new Event("change"));
  // Debounce parseAndPreview — don't re-parse on every keypress
  clearTimeout(_quickModeParseTimer);
  if (emails.length > 0) {
    _quickModeParseTimer = setTimeout(() => parseAndPreview(), 350);
  }
}

/* ─── SYNC RECIPIENTS FROM OUTLOOK'S TO FIELD ───────────────────
   Reads the native To field of the compose window and pre-populates
   the CSV with those addresses — great for small ad-hoc sends.
   ─────────────────────────────────────────────────────────────── */

function syncRecipientsFromOutlook() {
  try {
    Office.context.mailbox.item.to.getAsync(result => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        log("Could not read To field from Outlook.", "warning");
        return;
      }
      const recipients = result.value || [];
      if (recipients.length === 0) {
        log("Outlook's To field is empty — add addresses there first.", "warning");
        return;
      }
      const csvLines = ["email,display_name"];
      recipients.forEach(r => {
        const addr = (r.emailAddress || "").trim();
        const name = (r.displayName || "").replace(/,/g, " ");
        if (addr) csvLines.push(`${addr},${name}`);
      });
      const csvInput = document.getElementById("csvInput");
      if (csvInput) {
        csvInput.value = csvLines.join("\n");
        csvInput.dispatchEvent(new Event("input"));
        csvInput.dispatchEvent(new Event("change"));
        parseAndPreview();
        log(`Imported ${recipients.length} address${recipients.length !== 1 ? "es" : ""} from Outlook's To field.`, "success");
      }
      // Clear Outlook's To field so the merge engine populates it per-recipient
      Office.context.mailbox.item.to.setAsync([], () => {});
    });
  } catch (e) {
    log("syncRecipientsFromOutlook: " + e.message, "error");
  }
}

/**
 * Wire up the tab navigation bar.
 *
 * Each .tab-btn has a data-tab attribute matching the id of a .tab-pane
 * (e.g. data-tab="compose" → #tab-compose). On click:
 *   1. Remove .active from all buttons, set aria-selected="false"
 *   2. Hide all .tab-pane elements
 *   3. Activate the clicked button and reveal the matching pane
 *   4. If the user switched to the Compose tab, trigger a subject sync
 *      so the taskpane subject reflects any changes made in Outlook's
 *      native compose area since the last poll.
 */
function initTabs() {
  document.querySelectorAll(".tab-bar .tab-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      const target = btn.dataset.tab;
      // Deactivate all tab buttons
      document.querySelectorAll(".tab-bar .tab-btn").forEach(b => {
        b.classList.remove("active");
        b.setAttribute("aria-selected", "false");
      });
      // Hide all tab panels
      document.querySelectorAll(".tab-pane").forEach(p => p.classList.add("hidden"));
      // Activate the clicked button and show its panel
      btn.classList.add("active");
      btn.setAttribute("aria-selected", "true");
      document.getElementById(`tab-${target}`).classList.remove("hidden");
      // Pull any changes the user typed directly in Outlook's subject field
      if (target === "compose") syncSubjectFromOutlook(false);
    });
  });
}

/**
 * Wire up all .accordion-hdr buttons.
 *
 * Each header has a data-accordion attribute that matches the ID suffix of
 * its body element (e.g. data-accordion="optout" → #accordion-optout).
 * On click, toggle the body's .hidden class and the header's .open class
 * (the .open class rotates the › chevron via CSS).
 */
function initAccordions() {
  document.querySelectorAll(".accordion-hdr").forEach(hdr => {
    hdr.addEventListener("click", () => {
      const body = document.getElementById(`accordion-${hdr.dataset.accordion}`);
      if (!body) return;
      const isOpen = !body.classList.contains("hidden");
      body.classList.toggle("hidden", isOpen);   // close if open, open if closed
      hdr.classList.toggle("open", !isOpen);      // sync the chevron rotation
    });
  });
}

/* ─── LOCAL STATE PERSISTENCE ─────────────────────────────────── */

/**
 * Safely read a value from localStorage, returning defaultVal on any error
 * (quota exceeded, storage restricted by the Office webview, JSON parse fail).
 *
 * If defaultVal is a non-null object, the stored string is JSON-parsed so
 * callers receive an object back rather than a raw string.
 *
 * @param {string} key         - localStorage key
 * @param {*}      [defaultVal] - Value to return when key is absent or errors
 * @returns {*} Stored value (or defaultVal)
 */
function lsGet(key, defaultVal) {
  try {
    const raw = localStorage.getItem(key);
    if (raw === null) return defaultVal !== undefined ? defaultVal : null;
    // Auto-parse JSON when the caller passed a non-null object as the default —
    // this signals that the stored value should be an object, not a string.
    if (defaultVal !== null && typeof defaultVal === "object") {
      try { return JSON.parse(raw); } catch { return defaultVal; }
    }
    return raw;
  } catch { return defaultVal !== undefined ? defaultVal : null; }
}

/**
 * Safely write a value to localStorage, silently swallowing quota or
 * webview-restriction errors so the app continues to work read-only.
 *
 * @param {string} key   - localStorage key
 * @param {string} value - String value to store (stringify objects before calling)
 */
function lsSet(key, value) {
  try { localStorage.setItem(key, value); } catch { /* quota exceeded or restricted webview */ }
}

/**
 * Safely delete a key from localStorage.
 * @param {string} key
 */
function lsRemove(key) {
  try { localStorage.removeItem(key); } catch { /* ignore */ }
}

/**
 * Restore persisted state from the previous session.
 * Called once during Office.onReady() after UI elements are set up.
 *
 * Restores:
 *   - Subject line text (LS_KEY_SUBJECT)
 *   - CSV recipient data (LS_KEY_CSV) — triggers parseAndPreview() to rebuild the table
 *   - Custom tag chips (LS_KEY_TAGS)
 *   - Saved email templates (rendered into the template list)
 *   - Greeting format / fallback settings (LS_KEY_GREETING)
 */
function restoreLocalState() {
  const savedSubject = lsGet(LS_KEY_SUBJECT);
  if (savedSubject) {
    document.getElementById("subjectInput").value = savedSubject;
    log("Restored subject from saved session.", "info");
  }

  const savedCsv = lsGet(LS_KEY_CSV);
  if (savedCsv) {
    document.getElementById("csvInput").value = savedCsv;
    log("Restored CSV data from saved session.", "info");
    parseAndPreview();  // rebuild recipient table from the restored CSV string
  }

  // Restore user-added custom tag chips (non-default, non-smart tags).
  // DEFAULT_TAGS are already rendered in the HTML; duplicates are silently ignored by addTagToBar().
  const savedTags = JSON.parse(lsGet(LS_KEY_TAGS) || "[]");
  savedTags.forEach(tag => addTagToBar(tag));

  renderTemplateList();  // rebuild the saved-template picker in the Compose tab

  // Restore greeting format and fallback text.
  // The stored value may be a JSON string or already-parsed object depending on
  // which version of lsGet was used when it was saved — handle both.
  const savedGreeting = lsGet(LS_KEY_GREETING, null);
  if (savedGreeting) {
    try {
      const g = typeof savedGreeting === "object" ? savedGreeting : JSON.parse(savedGreeting);
      greetingConfig = g;
      document.getElementById("greetingFormat").value   = g.format   || "dear_sal_last";
      document.getElementById("greetingFallback").value = g.fallback || "Dear Valued Customer";
    } catch { /* ignore malformed stored data */ }
  }
}

/**
 * Persist the current set of custom tag chips to localStorage so they
 * survive page reloads. Filters out DEFAULT_TAGS (built-in chips that are
 * already in the HTML and don't need to be stored separately).
 *
 * Called by addTagToBar() (when a new chip is added) and by removeTagFromBar()
 * (when a chip is removed).
 */
function saveCustomTagsToStorage() {
  const tagEls = document.querySelectorAll("#tagBar [data-tag]");
  const tags = Array.from(tagEls)
    .map(el => el.dataset.tag)
    .filter(t => !DEFAULT_TAGS.includes(t));  // exclude built-in default chips
  lsSet(LS_KEY_TAGS, JSON.stringify(tags));
}

/* ─── TEMPLATE SAVE / LOAD ─────────────────────────────────────── */

function getTemplates() {
  return JSON.parse(lsGet(LS_KEY_TEMPLATES) || "{}");
}

function getTemplateState() {
  return {
    subject:         document.getElementById("subjectInput").value,
    replyTo:         document.getElementById("replyToInput").value,
    importance:      document.getElementById("importanceSelect").value,
    sensitivity:     document.getElementById("sensitivitySelect").value,
    saveToSent:      document.getElementById("saveToSentItems").checked,
    dedup:           document.getElementById("deduplicateEnabled").checked,
    readReceipt:     document.getElementById("requestReadReceipt").checked,
    deliveryReceipt: document.getElementById("requestDeliveryReceipt").checked,
    bccSelf:         document.getElementById("bccSelfEnabled").checked,
    flagFollowup:    document.getElementById("flagForFollowup").checked,
    plainText:       document.getElementById("plainTextMode").checked,
    categories:      document.getElementById("categoriesInput").value,
    customHeaders:   document.getElementById("customHeadersInput").value,
    customHeadersEnabled: document.getElementById("customHeadersEnabled").checked,
    listUnsubscribe: document.getElementById("listUnsubscribeInput").value,
    sendAs:          document.getElementById("sendAsInput").value,
    greetingFormat:  greetingConfig.format,
    greetingFallback: greetingConfig.fallback,
    // Feature 11: scheduling and rate limit settings
    scheduleEnabled:      document.getElementById("scheduleEnabled")?.checked || false,
    scheduledTime:        document.getElementById("scheduledTime")?.value || "",
    sendingWindowEnabled: document.getElementById("sendingWindowEnabled")?.checked || false,
    windowStart:          document.getElementById("windowStart")?.value || "09:00",
    windowEnd:            document.getElementById("windowEnd")?.value || "17:00",
    maxPerHour:           document.getElementById("maxPerHour")?.value || "0",
    dailyCap:             document.getElementById("dailyCap")?.value || "0",
    jitterMin:            document.getElementById("jitterMinInput")?.value || "0",
    jitterMax:            document.getElementById("jitterMaxInput")?.value || "0",
  };
}

function applyTemplateState(state) {
  if (state.subject         !== undefined) document.getElementById("subjectInput").value            = state.subject;
  if (state.replyTo         !== undefined) document.getElementById("replyToInput").value            = state.replyTo;
  if (state.importance      !== undefined) document.getElementById("importanceSelect").value        = state.importance;
  if (state.sensitivity     !== undefined) document.getElementById("sensitivitySelect").value       = state.sensitivity;
  if (state.saveToSent      !== undefined) document.getElementById("saveToSentItems").checked       = state.saveToSent;
  if (state.dedup           !== undefined) document.getElementById("deduplicateEnabled").checked    = state.dedup;
  if (state.readReceipt     !== undefined) document.getElementById("requestReadReceipt").checked    = state.readReceipt;
  if (state.deliveryReceipt !== undefined) document.getElementById("requestDeliveryReceipt").checked = state.deliveryReceipt;
  if (state.bccSelf         !== undefined) document.getElementById("bccSelfEnabled").checked        = state.bccSelf;
  if (state.flagFollowup    !== undefined) document.getElementById("flagForFollowup").checked       = state.flagFollowup;
  if (state.plainText       !== undefined) document.getElementById("plainTextMode").checked         = state.plainText;
  if (state.categories      !== undefined) document.getElementById("categoriesInput").value         = state.categories;
  if (state.customHeaders   !== undefined) document.getElementById("customHeadersInput").value      = state.customHeaders;
  if (state.customHeadersEnabled !== undefined) {
    document.getElementById("customHeadersEnabled").checked = state.customHeadersEnabled;
    document.getElementById("customHeadersRow").classList.toggle("hidden", !state.customHeadersEnabled);
  }
  if (state.listUnsubscribe !== undefined) document.getElementById("listUnsubscribeInput").value    = state.listUnsubscribe;
  if (state.sendAs          !== undefined) document.getElementById("sendAsInput").value             = state.sendAs;
  if (state.greetingFormat  !== undefined) {
    greetingConfig.format = state.greetingFormat;
    document.getElementById("greetingFormat").value = state.greetingFormat;
  }
  if (state.greetingFallback !== undefined) {
    greetingConfig.fallback = state.greetingFallback;
    document.getElementById("greetingFallback").value = state.greetingFallback;
  }
  // Feature 11: restore scheduling and rate limit settings
  if (state.scheduleEnabled !== undefined) {
    const el = document.getElementById("scheduleEnabled");
    if (el) { el.checked = state.scheduleEnabled; el.dispatchEvent(new Event("change")); }
  }
  if (state.scheduledTime) { const el = document.getElementById("scheduledTime"); if (el) el.value = state.scheduledTime; }
  if (state.sendingWindowEnabled !== undefined) {
    const el = document.getElementById("sendingWindowEnabled");
    if (el) { el.checked = state.sendingWindowEnabled; el.dispatchEvent(new Event("change")); }
  }
  if (state.windowStart) { const el = document.getElementById("windowStart"); if (el) el.value = state.windowStart; }
  if (state.windowEnd)   { const el = document.getElementById("windowEnd");   if (el) el.value = state.windowEnd; }
  if (state.maxPerHour !== undefined) { const el = document.getElementById("maxPerHour"); if (el) el.value = state.maxPerHour; }
  if (state.dailyCap    !== undefined) { const el = document.getElementById("dailyCap");    if (el) el.value = state.dailyCap; }
  if (state.jitterMin   !== undefined) { const el = document.getElementById("jitterMinInput"); if (el) el.value = state.jitterMin; }
  if (state.jitterMax   !== undefined) { const el = document.getElementById("jitterMaxInput"); if (el) el.value = state.jitterMax; }

  // UX 9: if scheduling settings were restored as active, expand the scheduling accordion
  if (state.scheduleEnabled || state.sendingWindowEnabled) {
    const schBody = document.getElementById("accordion-scheduling");
    if (schBody && schBody.classList.contains("hidden")) {
      schBody.classList.remove("hidden");
      const schHdr = schBody.previousElementSibling;
      if (schHdr) {
        schHdr.classList.add("open");
        const schChevron = schHdr.querySelector(".accordion-chevron");
        if (schChevron) schChevron.textContent = "▾";
      }
    }
  }
  // UX 9: if rate limit settings were restored as active, expand the advanced accordion
  const mphRestored = parseInt(state.maxPerHour || "0", 10);
  const dcRestored  = parseInt(state.dailyCap   || "0", 10);
  if (mphRestored > 0 || dcRestored > 0) {
    const advBody = document.getElementById("accordion-advanced");
    if (advBody && advBody.classList.contains("hidden")) {
      advBody.classList.remove("hidden");
      const advHdr = advBody.previousElementSibling;
      if (advHdr) {
        advHdr.classList.add("open");
        const advChevron = advHdr.querySelector(".accordion-chevron");
        if (advChevron) advChevron.textContent = "▾";
      }
    }
  }

  // UX 7: refresh badges after template restore
  updateSchedulingBadge();
  updateRateLimitBadge();
}

function saveTemplate() {
  const name = document.getElementById("templateNameInput").value.trim();
  if (!name) { log("Enter a template name.", "warning"); return; }
  const templates = lsGet(LS_KEY_TEMPLATES, {});
  templates[name] = getTemplateState();
  lsSet(LS_KEY_TEMPLATES, JSON.stringify(templates));
  populateTemplateSelect();
  document.getElementById("templateSelect").value = name;
  document.getElementById("deleteTemplateBtn").disabled = false;
  log(`Template "${name}" saved.`, "success");
}

function loadTemplate() {
  const name = document.getElementById("templateSelect").value;
  if (!name) { document.getElementById("deleteTemplateBtn").disabled = true; return; }
  const templates = lsGet(LS_KEY_TEMPLATES, {});
  const tpl = templates[name];
  if (!tpl) return;
  applyTemplateState(tpl);
  document.getElementById("deleteTemplateBtn").disabled = false;
  log(`Template "${name}" loaded.`, "info");
}

function deleteTemplate() {
  const name = document.getElementById("templateSelect").value;
  if (!name) { log("Select a template to delete.", "error"); return; }
  const templates = getTemplates();
  delete templates[name];
  lsSet(LS_KEY_TEMPLATES, JSON.stringify(templates));
  renderTemplateList();
  log(`Template "${name}" deleted.`, "info");
}

function renderTemplateList() {
  const sel = document.getElementById("templateSelect");
  const templates = getTemplates();
  const names = Object.keys(templates);
  sel.innerHTML = `<option value="">— load template —</option>` +
    names.map(n => `<option value="${escapeHtml(n)}">${escapeHtml(n)}</option>`).join("");
  document.getElementById("deleteTemplateBtn").disabled = names.length === 0;
}

// Alias used by saveTemplate / loadTemplate (Feature 5)
function populateTemplateSelect() { renderTemplateList(); }

/* ─── MATCH FIELDS DIALOG (Feature 1) ──────────────────────────── */

function openMatchFieldsModal(csvHeaders) {
  const body = document.getElementById("matchFieldsBody");
  body.innerHTML = CANONICAL_FIELDS.map(f => {
    const currentMapping = fieldMapping[f.key] || "";
    const options = ['<option value="">— skip —</option>']
      .concat(csvHeaders.map(h => {
        const autoMatch = autoMatchField(f.key, h) && !currentMapping;
        const selected = currentMapping === h ? " selected" : (autoMatch ? " selected" : "");
        return `<option value="${escapeHtml(h)}"${selected}>${escapeHtml(h)}</option>`;
      }));
    return `<div style="display:flex;align-items:center;gap:8px;margin-bottom:6px;">
      <label style="width:110px;font-size:11px;flex-shrink:0;">${escapeHtml(f.label)}</label>
      <select class="select-inline" data-canonical="${f.key}" style="flex:1;">${options.join("")}</select>
    </div>`;
  }).join("");
  document.getElementById("matchFieldsModal").classList.remove("hidden");
}

function autoMatchField(canonicalKey, csvHeader) {
  const h = csvHeader.toLowerCase().replace(/[\s_-]+/g, "");
  const synonyms = {
    email:        ["email", "emailaddress", "e-mail", "mail"],
    first_name:   ["firstname", "givenname", "forename", "fname"],
    last_name:    ["lastname", "surname", "familyname", "lname"],
    salutation:   ["salutation", "title", "honorific", "prefix"],
    gender:       ["gender", "sex"],
    company:      ["company", "organisation", "organization", "employer", "firm", "companyname"],
    title:        ["jobtitle", "position", "role", "occupation"],
    cc:           ["cc", "carboncopy"],
    bcc:          ["bcc", "blindcc"],
    reply_to:     ["replyto", "reply_to", "replytoemail"],
    attachment:   ["attachment", "file", "filename", "attachfile"],
    display_name: ["displayname", "name", "fullname"],
    skip_if:      ["skipif", "skip", "donotsend", "optout"],
    send_at:      ["sendat", "sendtime", "scheduletime", "senddate"],
  };
  return (synonyms[canonicalKey] || []).includes(h);
}

function applyMatchFields() {
  const selects = document.getElementById("matchFieldsBody").querySelectorAll("select");
  const newMapping = {};
  selects.forEach(sel => {
    if (sel.value) newMapping[sel.dataset.canonical] = sel.value;
  });
  if (!newMapping.email) {
    log("Match Fields: 'Email' column is required — please map it.", "warning");
    return;
  }
  fieldMapping = newMapping;
  document.getElementById("matchFieldsModal").classList.add("hidden");
  log(`Field mapping applied. Email column: "${fieldMapping.email}".`, "success");
  // Re-parse with new mapping
  parseAndPreview();
}

function applyFieldMapping(rawRow) {
  if (!Object.keys(fieldMapping).length) return rawRow;
  const mapped = Object.assign({}, rawRow); // keep original columns too
  CANONICAL_FIELDS.forEach(f => {
    const csvCol = fieldMapping[f.key];
    if (csvCol && rawRow[csvCol] !== undefined) {
      mapped[f.key] = rawRow[csvCol];
    }
  });
  return mapped;
}

/* ─── BODY TEMPLATE SAVE / LOAD (Feature 2) ────────────────────── */

async function saveBodyToTemplate() {
  const name = document.getElementById("templateNameInput").value.trim()
             || document.getElementById("templateSelect").value;
  if (!name) { log("Select or name a template first.", "warning"); return; }
  try {
    const body = await getComposeBodyAsync();
    const templates = lsGet(LS_KEY_TEMPLATES, {});
    if (!templates[name]) templates[name] = {};
    templates[name].body = body;
    lsSet(LS_KEY_TEMPLATES, JSON.stringify(templates));
    log(`Body saved to template "${name}".`, "success");
  } catch (e) {
    log(`Failed to read compose body: ${e.message}`, "error");
  }
}

async function loadBodyFromTemplate() {
  const name = document.getElementById("templateSelect").value;
  if (!name) { log("Select a template first.", "warning"); return; }
  const templates = lsGet(LS_KEY_TEMPLATES, {});
  const tpl = templates[name];
  if (!tpl || !tpl.body) { log(`Template "${name}" has no saved body.`, "warning"); return; }

  // B2: On Mac, setAsync with HTML coercion can corrupt embedded image src attributes.
  const isMac = Office.context.platform === Office.PlatformType.Mac;
  if (isMac && tpl.body.match(/<img/i)) {
    log("Warning: Mac: loading a body template with embedded images may corrupt image links in the sent email (known Office.js limitation on Mac). Proceeding anyway.", "warning");
  }

  // Signature preservation warning — loading a template replaces the entire compose body on all platforms
  log("ℹ Loading template will replace the current compose body (including your signature). Tip: place your signature at the bottom of the template to preserve it.", "info");

  try {
    await setComposeBodyAsync(tpl.body);
    log(`Body loaded from template "${name}".`, "success");
  } catch (e) {
    log(`Failed to set compose body: ${e.message}`, "error");
  }
}

function getComposeBodyAsync() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync(
      Office.CoercionType.Html,
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
        else reject(new Error(result.error.message));
      }
    );
  });
}

function setComposeBodyAsync(html) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.setAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error(result.error.message));
      }
    );
  });
}

/* ─── SAVE AS DRAFTS (Feature 5) ───────────────────────────────── */

async function handleSaveDrafts() {
  const savedCsvText = document.getElementById("csvInput").value;
  const savedRecipients = parsedRecipients.slice();
  draftsMode = true;
  try {
    await handleMergeClick();
  } finally {
    draftsMode = false;
    // A8: Restore if handleMergeClick cleared it on zero-failure completion
    if (!document.getElementById("csvInput").value && savedCsvText) {
      document.getElementById("csvInput").value = savedCsvText;
      lsSet(LS_KEY_CSV, savedCsvText);
      parsedRecipients = savedRecipients;
    }
  }
}

/* ─── TEST SEND ────────────────────────────────────────────────── */

async function handleTestSend() {
  // A1: Re-entrancy guard — disable button for the duration
  document.getElementById("testSendBtn").disabled = true;
  const savedFillIn = window._fillInValues;
  window._fillInValues = null;
  try {
  const subjectTemplate = document.getElementById("subjectInput").value.trim();
  if (!subjectTemplate) { log("Subject line is empty.", "error"); return; }
  if (parsedRecipients.length === 0) {
    log("Load recipients first — test send uses the first row for personalisation.", "error");
    return;
  }

  const selfEmail = Office.context.mailbox.userProfile.emailAddress;
  if (!selfEmail) { log("Could not determine your email address.", "error"); return; }

  // Feature 1: use selected row index
  const rowIdx = parseInt(document.getElementById("testRowSelect")?.value || "0", 10);
  const sample = parsedRecipients[Math.min(rowIdx, parsedRecipients.length - 1)] || parsedRecipients[0];
  const displayRowNum = rowIdx + 1;

  log(`Test send: personalising with row ${displayRowNum} data and sending to ${selfEmail}...`, "info");

  let emailBodyTemplate = "";
  try {
    emailBodyTemplate = await new Promise((resolve, reject) => {
      Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
        else reject(new Error("Could not read email body."));
      });
    });
  } catch (err) { log(`Body read error: ${err.message}`, "error"); return; }

  const sampleWithMeta = Object.assign({}, sample, {
    record_num:   String(rowIdx + 1),
    record_count: String(parsedRecipients.length || 1)
  });
  const subject = personalize(subjectTemplate, sampleWithMeta, false);  // plain text — no HTML escaping
  const body    = personalize(emailBodyTemplate, sampleWithMeta, true); // HTML body — escape merge values

  const importance             = document.getElementById("importanceSelect").value;
  const sensitivity            = document.getElementById("sensitivitySelect").value;
  const plainTextMode          = document.getElementById("plainTextMode").checked;
  const categoriesRaw          = document.getElementById("categoriesInput").value.trim();
  const categories             = categoriesRaw ? categoriesRaw.split(",").map(c => c.trim()).filter(Boolean) : [];
  const replyTo                = document.getElementById("replyToInput").value.trim();
  const sendAs                 = document.getElementById("sendAsInput").value.trim();
  const isReadReceiptRequested     = document.getElementById("requestReadReceipt").checked;
  const isDeliveryReceiptRequested = document.getElementById("requestDeliveryReceipt").checked;
  const customHeadersEnabled   = document.getElementById("customHeadersEnabled").checked;
  const globalCustomHeaders    = customHeadersEnabled
    ? parseCustomHeaders(document.getElementById("customHeadersInput").value)
    : [];
  const listUnsubscribeTemplate = document.getElementById("listUnsubscribeInput").value.trim();
  const testCustomHeaders = [
    ...globalCustomHeaders,
    ...(listUnsubscribeTemplate
      ? buildUnsubHeaders(personalize(listUnsubscribeTemplate, sampleWithMeta, false))
      : [])
  ];
  const flagged    = document.getElementById("flagForFollowup").checked;
  const inlineImagesArr = Array.from(inlineImages.values());
  const recipientAttachments = resolveAttachmentsForRecipient(sampleWithMeta);

  let token;
  try { token = await getAccessToken(); }
  catch (err) { log(`Auth error: ${err.message}`, "error"); return; }

  const bodyContent     = plainTextMode ? stripHtmlToText(body) : body;
  const bodyContentType = plainTextMode ? "Text" : "HTML";

  const message = {
    subject: `[TEST] ${subject}`,
    body: { contentType: bodyContentType, content: bodyContent },
    toRecipients: [{ emailAddress: { address: selfEmail } }]
  };
  if (importance && importance !== "normal")   message.importance = importance;
  if (sensitivity && sensitivity !== "normal") message.sensitivity = sensitivity;
  if (categories.length > 0)                  message.categories = categories;
  if (replyTo)                                message.replyTo = [{ emailAddress: { address: replyTo } }];
  if (isReadReceiptRequested)                 message.isReadReceiptRequested = true;
  if (isDeliveryReceiptRequested)             message.isDeliveryReceiptRequested = true;
  if (testCustomHeaders.length > 0)           message.internetMessageHeaders = testCustomHeaders;
  if (flagged)                                message.flag = { flagStatus: "flagged" };

  const fileAttachments   = recipientAttachments.map(a => ({
    "@odata.type": "#microsoft.graph.fileAttachment",
    name: a.name, contentType: a.contentType, contentBytes: a.contentBytes
  }));
  const inlineAttachments = inlineImagesArr.map(img => ({
    "@odata.type": "#microsoft.graph.fileAttachment",
    name: img.name, contentType: img.contentType, contentBytes: img.contentBytes,
    isInline: true, contentId: img.name
  }));
  const allAttachments = [...fileAttachments, ...inlineAttachments];
  if (allAttachments.length > 0) message.attachments = allAttachments;

  // Size guard
  const totalAttachBytes = sharedAttachments.reduce((s, a) => s + a.sizeBytes, 0)
    + [...perRecipientFiles.values()].reduce((s, a) => s + a.sizeBytes, 0)
    + [...inlineImages.values()].reduce((s, a) => s + a.sizeBytes, 0);
  if (totalAttachBytes > 3.5 * 1024 * 1024) {
    log(`Test send: attachments total ${(totalAttachBytes / 1024 / 1024).toFixed(1)} MB — may exceed the 4 MB Graph sendMail limit. Attempting anyway.`, "warning");
  }

  const payload = { message, saveToSentItems: true };
  if (sendAs) {
    payload.message.from = { emailAddress: { address: sendAs } };
  }

  try {
    const res = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
      method: "POST",
      headers: { "Authorization": `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(`${res.status} — ${JSON.stringify(err)}`);
    }
    log(`✅ Test email sent to ${selfEmail}. Subject: "[TEST] ${subject}"`, "success");
  } catch (err) {
    log(`Test send failed: ${err.message}`, "error");
  }
  } finally {
    // A1: Re-enable button regardless of outcome
    window._fillInValues = savedFillIn;
    document.getElementById("testSendBtn").disabled = false;
  }
}

/* ─── PREVIEW ALL ──────────────────────────────────────────────── */

async function handlePreviewAll() {
  // A10: Use filtered/sorted set so preview matches what will actually be sent
  const toPreview = getFilteredSortedRecipients();
  if (!toPreview.length) {
    log("No recipients to preview.", "warning");
    return;
  }
  const subjectTemplate = document.getElementById("subjectInput").value.trim();
  if (!subjectTemplate) { log("Subject line is empty.", "error"); return; }

  const bodyHtmlForPreview = await getComposeBodyAsync().catch(function() { return ""; });
  if (!window._fillInValues && /\{\{fill_in:/i.test(bodyHtmlForPreview)) {
    log("Note: {{fill_in:…}} tokens found — run a merge first to populate fill-in values. Previewing with placeholder tokens.", "warning");
  }

  log("Reading email body for preview...", "info");
  try {
    previewBodyTemplate = await new Promise((resolve, reject) => {
      Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
        else reject(new Error("Could not read email body."));
      });
    });
  } catch (err) { log(`Body read error: ${err.message}`, "error"); return; }

  previewRecipients = toPreview; // A10: store filtered set for navigation
  previewIndex = 0;
  renderPreviewEntry();
  _openModalWithTrap("previewModal"); // Feature 15: focus trap
}

function renderPreviewEntry() {
  // A10: Use previewRecipients (filtered set) instead of parsedRecipients
  const total = previewRecipients.length;
  if (!total || previewIndex >= total) return;
  const recipient = Object.assign({}, previewRecipients[previewIndex], {
    record_num:   String(previewIndex + 1),
    record_count: String(total)
  });
  const subjectTemplate = document.getElementById("subjectInput").value;

  const subject  = personalize(subjectTemplate, recipient, false);
  const body     = personalize(previewBodyTemplate, recipient, true);
  const plainTextMode = document.getElementById("plainTextMode").checked;
  const displayBody   = plainTextMode
    ? `<pre style="white-space:pre-wrap;font-family:Segoe UI,sans-serif;font-size:13px;padding:8px;">${escapeHtml(stripHtmlToText(body))}</pre>`
    : body;

  document.getElementById("previewCounter").textContent = `Preview ${previewIndex + 1} / ${total}`;
  document.getElementById("previewSubject").textContent = subject;
  document.getElementById("previewBodyFrame").srcdoc    = displayBody;

  document.getElementById("previewPrevBtn").disabled = previewIndex === 0;
  document.getElementById("previewNextBtn").disabled = previewIndex === total - 1;
}

function closePreviewModal() {
  _closeModalWithTrap("previewModal"); // Feature 15
}

/* ─── DOWNLOAD LOG ─────────────────────────────────────────────── */

function downloadLog() {
  const entries = document.querySelectorAll("#statusLog .log-entry");
  if (!entries.length) { log("Log is empty.", "info"); return; }
  const text = Array.from(entries).map(el => el.textContent).join("\n");
  const blob = new Blob([text], { type: "text/plain" });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement("a");
  a.href     = url;
  a.download = `mail-merge-log-${new Date().toISOString().slice(0, 19).replace(/:/g, "-")}.txt`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  log("Log exported.", "info");
}

/* ─── SEND SUMMARY REPORT ──────────────────────────────────────── */

function downloadSendReport() {
  if (!sendOutcomes.length) { log("No send outcomes to report.", "info"); return; }
  const rows = [["row_num","email","display_name","subject_used","status","timestamp","error"]];
  sendOutcomes.forEach(o => rows.push([
    o.rowNum      || "",
    o.email       || "",
    o.displayName || "",
    o.subjectUsed || "",
    o.status      || "",
    o.timestamp   || "",
    o.error       || ""
  ].map(v => csvField(String(v)))));
  const csv  = rows.map(r => r.join(",")).join("\n");
  const blob = new Blob([csv], { type: "text/csv" });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement("a");
  a.href     = url;
  a.download = `mail-merge-report-${new Date().toISOString().slice(0, 10)}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  log("Send report exported.", "info");
}

/* ─── PROGRESS BAR ─────────────────────────────────────────────── */

function setProgress(current, total) {
  const pct = total > 0 ? Math.round((current / total) * 100) : 0;
  document.getElementById("progressContainer").classList.remove("hidden");
  const fill = document.getElementById("progressFill");
  fill.style.width = `${pct}%`;
  fill.setAttribute("aria-valuenow", pct);
  document.getElementById("progressLabel").textContent = `${pct}%  (${current} / ${total})`;
}

function hideProgress() {
  document.getElementById("progressContainer").classList.add("hidden");
  document.getElementById("progressFill").style.width = "0%";
}

/* Feature 5: show completion state and auto-hide after 8 seconds */
function showMergeComplete(sent, total) {
  _mergeCompletedSuccessfully = true;
  const fill  = document.getElementById("progressFill");
  const label = document.getElementById("progressLabel");
  const container = document.getElementById("progressContainer");
  if (container) container.classList.remove("hidden");
  if (fill)  fill.style.width = "100%";
  if (label) label.textContent = "✓ Done — " + sent + " of " + total + " sent";
  setTimeout(function() {
    if (!mergeInProgress) {  // only hide if no merge is running
      const c = document.getElementById("progressContainer");
      if (c) c.classList.add("hidden");
      _mergeCompletedSuccessfully = false;
    }
  }, 8000);
}

/* ─── SUPPRESSION LIST ─────────────────────────────────────────── */

function handleSuppressionUpload(e) {
  const file = e.target.files[0];
  if (!file) return;
  e.target.value = "";
  const reader = new FileReader();
  reader.onload = (event) => {
    const text   = event.target.result;
    const emails = text.split(/[\n,;\r]+/)
      .map(s => s.trim().toLowerCase())
      .filter(s => EMAIL_REGEX.test(s));
    emails.forEach(em => suppressionSet.add(em));
    document.getElementById("suppressionLabel").textContent =
      `${suppressionSet.size} address${suppressionSet.size !== 1 ? "es" : ""} suppressed`;
    document.getElementById("clearSuppressionBtn").classList.remove("hidden");
    log(`Suppression list loaded: ${suppressionSet.size} address${suppressionSet.size !== 1 ? "es" : ""} will be excluded.`, "info");
  };
  reader.onerror = () => log(`Failed to read suppression file.`, "error");
  reader.readAsText(file);
}

function clearSuppression() {
  suppressionSet.clear();
  document.getElementById("suppressionLabel").textContent = "None loaded";
  document.getElementById("clearSuppressionBtn").classList.add("hidden");
  log("Suppression list cleared.", "info");
}

/* ─── RETRY FAILED ─────────────────────────────────────────────── */

async function handleRetryFailed() {
  if (!failedRecipients.length) return;
  // A7 + BUG 7: Save both recipients and CSV text — restore is now in a finally block
  const savedRecipients = parsedRecipients.slice();
  const savedCsvText = document.getElementById("csvInput").value;
  parsedRecipients = failedRecipients.slice();
  failedRecipients = [];
  try {
    await handleMergeClick();
  } finally {
    // Restore regardless of outcome (error or success)
    parsedRecipients = savedRecipients;
    if (!document.getElementById("csvInput").value && savedCsvText) {
      document.getElementById("csvInput").value = savedCsvText;
      lsSet(LS_KEY_CSV, savedCsvText);
    }
  }
}

/* ─── TAG INSERTION ────────────────────────────────────────────── */

/**
 * Insert a tag placeholder into either the subject line or the email body.
 *
 * @param {string} tag          - The placeholder string, e.g. "{{first_name}}"
 * @param {boolean} forceSubject - When true, insert into the subject input even
 *                                 if subjectHasFocus is currently false. This is
 *                                 needed because blur fires before click in the
 *                                 browser event order — by the time the click
 *                                 handler runs, subjectHasFocus is already false.
 *                                 The caller passes _subjectWasFocused (captured
 *                                 on mousedown, before blur) as this argument.
 */
function insertTag(tag, forceSubject = false) {
  if (forceSubject || subjectHasFocus) {
    // ── Path 1: Insert at cursor in the taskpane #subjectInput ────
    // Used when the taskpane subject field has (or just had) focus.
    // The input's existing "input" listener then debounces a push to
    // the native Outlook subject via pushSubjectToOutlook().
    const input = document.getElementById("subjectInput");
    const start = input.selectionStart;
    const end   = input.selectionEnd;
    const value = input.value;
    input.value = value.slice(0, start) + tag + value.slice(end);
    const newCursor = start + tag.length;
    input.setSelectionRange(newCursor, newCursor);

    // Fire "input" so lsSet() persists the value and the 400 ms push
    // timer to Outlook's native subject field starts.
    input.dispatchEvent(new Event("input"));

    // Restore focus so the user can click another tag immediately.
    input.focus();
    log(`Inserted tag into subject (taskpane): ${tag}`, "success");

  } else if (tagTarget === "subject") {
    // ── Path 2: Append to the Outlook native subject via Office.js ─
    // Used when the "Subject" toggle is active and #subjectInput is
    // not focused. We read the current subject, append the tag, write
    // it back, and mirror the result into #subjectInput so the two
    // stay in sync (the poll would catch it anyway but this is instant).
    Office.context.mailbox.item.subject.getAsync((getResult) => {
      if (getResult.status !== Office.AsyncResultStatus.Succeeded) {
        log(`Could not read subject: ${getResult.error.message}`, "error");
        return;
      }
      const newSubject = getResult.value + tag;
      Office.context.mailbox.item.subject.setAsync(newSubject, (setResult) => {
        if (setResult.status !== Office.AsyncResultStatus.Succeeded) {
          log(`Failed to set subject: ${setResult.error.message}`, "error");
        } else {
          // Mirror into taskpane subject input immediately.
          const input = document.getElementById("subjectInput");
          input.value = newSubject;
          // Dispatch "input" to persist to localStorage (skip the
          // push-to-Outlook debounce — we just came FROM Outlook).
          lsSet(LS_KEY_SUBJECT, newSubject);
          log(`Inserted tag into subject (native): ${tag}`, "success");
        }
      });
    });

  } else {
    // ── Path 3: Insert into the email body via Office.js ──────────
    // setSelectedDataAsync inserts at the current cursor position in
    // the Outlook compose body. On macOS, only Text coercion is
    // supported; HTML works on Windows.
    const isMac = Office.context.platform === Office.PlatformType.Mac;
    const coercionType = isMac ? Office.CoercionType.Text : Office.CoercionType.Html;
    Office.context.mailbox.item.body.setSelectedDataAsync(
      tag,
      { coercionType },
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          log(`Failed to insert tag into body: ${tag}`, "error");
        } else {
          log(`Inserted tag into body: ${tag}`, "success");
        }
      }
    );
  }
}

/**
 * Read the value from #customTagInput, normalise it into a {{tag}} format,
 * add it to the tag bar as a custom chip, and immediately insert it into
 * the currently focused field (body or subject line).
 *
 * Normalisation rules: lowercase, spaces replaced with underscores.
 * Example: "First Name" → {{first_name}}
 *
 * Called by the "Add tag" button click handler and the Enter keydown
 * listener on #customTagInput.
 */
function addCustomTag() {
  const input = document.getElementById("customTagInput");
  const name  = input.value.trim();
  if (!name) return;
  // Normalise: lowercase + underscores so the tag matches CSV column name conventions
  const normalized = name.toLowerCase().replace(/\s+/g, "_");
  const tag = `{{${normalized}}}`;
  addTagToBar(tag);      // add the chip to the bar (no-op if already present)
  insertTag(tag);        // insert immediately into the focused field
  input.value = "";      // clear the input for the next custom tag
}

/**
 * Add a tag chip to #tagBar if it isn't already there.
 *
 * Each chip structure:
 *   <span class="tag [tag-smart|tag-custom]" data-tag="{{field}}">
 *     <span class="tag-label">{{field}}</span>
 *     <!-- only for non-default, non-smart chips: -->
 *     <button class="tag-remove-btn" aria-label="Remove {{field}} tag">✕</button>
 *   </span>
 *
 * The .tag-label span separates the display text from the ✕ button so
 * that e.target.closest("[data-tag]") in the click handler still resolves
 * to the chip even when the user clicks the label text directly.
 *
 * Remove buttons are added only to custom (user-added / CSV-column) chips.
 * Default chips (those in DEFAULT_TAGS[]) and smart chips (type === "smart")
 * are permanent and do not get a remove button.
 *
 * @param {string} tag   - The placeholder string, e.g. "{{first_name}}"
 * @param {string} [type] - "smart" for computed tags like {{today}} / {{greeting_line}}
 */
function addTagToBar(tag, type) {
  // Bail out early if this tag chip already exists (prevents duplicates
  // when restoring from localStorage or re-parsing the same CSV).
  const existing = document.querySelector(`#tagBar [data-tag="${CSS.escape(tag)}"]`);
  if (existing) return;

  const bar  = document.getElementById("tagBar");
  const chip = document.createElement("span");
  chip.className   = type === "smart" ? "tag tag-smart" : "tag tag-custom";
  chip.dataset.tag = tag;

  // Label span — holds the visible tag text (e.g. "{{company}}").
  // Using a child span (rather than setting textContent directly on the chip)
  // means clicks on the label still bubble up to the chip's [data-tag] element,
  // which closest() in the click handler can find.
  const labelSpan = document.createElement("span");
  labelSpan.className   = "tag-label";
  labelSpan.textContent = tag;
  chip.appendChild(labelSpan);

  // Remove button — only for non-default, non-smart custom chips.
  // DEFAULT_TAGS is the array of built-in placeholder names defined near
  // the top of this file. Smart chips (type === "smart") are also permanent.
  const isDefault = DEFAULT_TAGS.includes(tag);
  const isSmart   = type === "smart";
  if (!isDefault && !isSmart) {
    const removeBtn = document.createElement("button");
    removeBtn.className = "tag-remove-btn";
    removeBtn.textContent = "✕";
    removeBtn.setAttribute("aria-label", `Remove ${tag} tag`);
    removeBtn.title = `Remove ${tag}`;
    // stopPropagation prevents the remove click from also triggering
    // the chip's insert handler (which is registered on the tagBar div).
    removeBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      removeTagFromBar(tag);
    });
    // Also block mousedown propagation so _subjectWasFocused isn't
    // snapshotted incorrectly when the user clicks the remove button.
    removeBtn.addEventListener("mousedown", (e) => e.stopPropagation());
    chip.appendChild(removeBtn);
  }

  // Keyboard accessibility: Enter / Space triggers a click on the chip.
  chip.setAttribute("tabindex", "0");
  chip.setAttribute("role", "button");
  chip.setAttribute("aria-label", "Insert " + tag);
  chip.addEventListener("keydown", function(e) {
    if (e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      chip.click();
    }
  });

  bar.appendChild(chip);
  // Persist non-smart custom tags to localStorage so they survive page reloads.
  if (type !== "smart") saveCustomTagsToStorage();
}

/**
 * Remove a custom tag chip from #tagBar and update localStorage.
 * Only call this for user-added or CSV-column chips — default and smart
 * chips should never be removed.
 *
 * @param {string} tag - The placeholder string to remove, e.g. "{{city}}"
 */
function removeTagFromBar(tag) {
  const chip = document.querySelector(`#tagBar [data-tag="${CSS.escape(tag)}"]`);
  if (chip) {
    chip.remove();
    saveCustomTagsToStorage();  // update persisted custom tags list
    log(`Removed custom tag: ${tag}`, "info");
  }
}

/**
 * Update the #tagsHint hint text to tell the user whether clicking a
 * tag chip will insert into the subject line or the email body.
 * Called on subjectInput focus (targetIsSubject=true) and on blur
 * (targetIsSubject=false, after a 150 ms delay via _hintResetTimer).
 *
 * @param {boolean} targetIsSubject - true when subject input is focused
 */
function updateTagHint(targetIsSubject) {
  const el = document.getElementById("tagsHint");
  if (!el) return;
  el.innerHTML = targetIsSubject
    ? "👆 Click any tag to insert into the <strong>subject line</strong>"
    : "👆 Click any tag to insert it into your <strong>email body</strong>";
}

/* ─── FILE UPLOAD: CSV + EXCEL ─────────────────────────────────── */

function handleCsvFileUpload(e) {
  const file = e.target.files[0];
  if (!file) return;
  e.target.value = "";
  const ext = file.name.split(".").pop().toLowerCase();
  if (ext === "xlsx" || ext === "xls") {
    loadExcelFile(file);
  } else {
    loadCsvFile(file);
  }
}

function loadCsvFile(file) {
  const reader = new FileReader();
  reader.onload = (event) => {
    document.getElementById("csvInput").value = event.target.result;
    log(`Loaded CSV: ${file.name} (${(file.size / 1024).toFixed(1)} KB)`, "info");
    parseAndPreview();
  };
  reader.onerror = () => log(`Failed to read file: ${file.name}`, "error");
  reader.readAsText(file);
}

function loadExcelFile(file) {
  if (typeof XLSX === "undefined") {
    log("Excel support unavailable — SheetJS failed to load. Try refreshing the task pane.", "error");
    return;
  }
  const reader = new FileReader();
  reader.onload = (event) => {
    try {
      const data      = new Uint8Array(event.target.result);
      const workbook  = XLSX.read(data, { type: "array" });
      const sheetName = workbook.SheetNames[0];
      const csv       = XLSX.utils.sheet_to_csv(workbook.Sheets[sheetName]);
      document.getElementById("csvInput").value = csv;
      log(`Loaded Excel: ${file.name} — sheet "${sheetName}" (${(file.size / 1024).toFixed(1)} KB)`, "info");
      parseAndPreview();
    } catch (err) {
      log(`Failed to parse Excel file: ${err.message}`, "error");
    }
  };
  reader.onerror = () => log(`Failed to read file: ${file.name}`, "error");
  reader.readAsArrayBuffer(file);
}

/* ─── ATTACHMENT UPLOAD: SHARED ────────────────────────────────── */

function handleAttachmentUpload(e) {
  const files = Array.from(e.target.files);
  e.target.value = "";
  if (!files.length) return;

  let pending     = files.length;
  let loadedCount = 0;

  function checkDone() {
    pending--;
    if (pending === 0) {
      updateSharedAttachmentsLabel();
      if (loadedCount > 0) {
        log(`Loaded ${loadedCount} shared attachment${loadedCount !== 1 ? "s" : ""}.`, "success");
      }
    }
  }

  files.forEach(file => {
    if (file.size > MAX_ATTACHMENT_BYTES) {
      log(`Shared attachment too large: ${file.name} ` +
          `(${(file.size / 1024 / 1024).toFixed(1)} MB). Max 3 MB — skipped.`, "error");
      checkDone();
      return;
    }
    const reader = new FileReader();
    reader.onload = (event) => {
      const [meta, contentBytes] = event.target.result.split(",");
      const contentType = meta.split(":")[1].split(";")[0];
      sharedAttachments.push({ name: file.name, contentType, contentBytes, sizeBytes: file.size });
      loadedCount++;
      checkDone();
    };
    reader.onerror = () => {
      log(`Failed to read attachment: ${file.name}`, "error");
      checkDone();
    };
    reader.readAsDataURL(file);
  });
}

function updateSharedAttachmentsLabel() {
  if (sharedAttachments.length === 0) {
    document.getElementById("attachmentLabel").textContent = "None";
    document.getElementById("clearSharedAttachmentsBtn").classList.add("hidden");
  } else {
    const totalBytes = sharedAttachments.reduce((s, a) => s + a.sizeBytes, 0);
    document.getElementById("attachmentLabel").textContent =
      `${sharedAttachments.length} file${sharedAttachments.length !== 1 ? "s" : ""} (${(totalBytes / 1024).toFixed(1)} KB total)`;
    document.getElementById("clearSharedAttachmentsBtn").classList.remove("hidden");
  }
}

function clearSharedAttachments() {
  sharedAttachments = [];
  document.getElementById("attachmentLabel").textContent = "None";
  document.getElementById("clearSharedAttachmentsBtn").classList.add("hidden");
  log("Shared attachments cleared.", "info");
}

/* ─── ATTACHMENT UPLOAD: PER-RECIPIENT ─────────────────────────── */

function handlePerRecipientFilesUpload(e) {
  const files = Array.from(e.target.files);
  e.target.value = "";
  if (!files.length) return;

  let pending     = files.length;
  let loadedCount = 0;

  function checkDone() {
    pending--;
    if (pending === 0) {
      const total = perRecipientFiles.size;
      document.getElementById("perRecipientLabel").textContent =
        `${total} file${total !== 1 ? "s" : ""} loaded`;
      document.getElementById("clearPerRecipientBtn").classList.remove("hidden");
      if (loadedCount > 0) {
        log(`Loaded ${loadedCount} per-recipient file${loadedCount !== 1 ? "s" : ""}.`, "success");
      }
    }
  }

  files.forEach(file => {
    if (file.size > MAX_ATTACHMENT_BYTES) {
      log(`Per-recipient file too large: ${file.name} ` +
          `(${(file.size / 1024 / 1024).toFixed(1)} MB). Max 3 MB.`, "error");
      checkDone();
      return;
    }
    const reader = new FileReader();
    reader.onload = (event) => {
      const [meta, contentBytes] = event.target.result.split(",");
      const contentType = meta.split(":")[1].split(";")[0];
      perRecipientFiles.set(file.name.toLowerCase(), {
        name: file.name,
        contentType,
        contentBytes,
        sizeBytes: file.size
      });
      loadedCount++;
      checkDone();
    };
    reader.onerror = () => {
      log(`Failed to read: ${file.name}`, "error");
      checkDone();
    };
    reader.readAsDataURL(file);
  });
}

function clearPerRecipientFiles() {
  perRecipientFiles.clear();
  document.getElementById("perRecipientLabel").textContent = "None loaded";
  document.getElementById("clearPerRecipientBtn").classList.add("hidden");
  log("Per-recipient attachment files cleared.", "info");
}

/* ─── ATTACHMENT UPLOAD: INLINE IMAGES ─────────────────────────── */

function handleInlineImagesUpload(e) {
  const files = Array.from(e.target.files);
  e.target.value = "";
  if (!files.length) return;

  let pending     = files.length;
  let loadedCount = 0;

  function checkDone() {
    pending--;
    if (pending === 0) {
      const total = inlineImages.size;
      document.getElementById("inlineImagesLabel").textContent =
        `${total} image${total !== 1 ? "s" : ""} loaded`;
      document.getElementById("clearInlineImagesBtn").classList.remove("hidden");
      if (loadedCount > 0) {
        log(`Loaded ${loadedCount} inline image${loadedCount !== 1 ? "s" : ""}.`, "success");
      }
    }
  }

  files.forEach(file => {
    if (file.size > MAX_ATTACHMENT_BYTES) {
      log(`Inline image too large: ${file.name} ` +
          `(${(file.size / 1024 / 1024).toFixed(1)} MB). Max 3 MB — skipped.`, "error");
      checkDone();
      return;
    }
    const reader = new FileReader();
    reader.onload = (event) => {
      const [meta, contentBytes] = event.target.result.split(",");
      const contentType = meta.split(":")[1].split(";")[0];
      inlineImages.set(file.name.toLowerCase(), {
        name: file.name, contentType, contentBytes, sizeBytes: file.size
      });
      loadedCount++;
      checkDone();
    };
    reader.onerror = () => {
      log(`Failed to read inline image: ${file.name}`, "error");
      checkDone();
    };
    reader.readAsDataURL(file);
  });
}

function clearInlineImages() {
  inlineImages.clear();
  document.getElementById("inlineImagesLabel").textContent = "None loaded";
  document.getElementById("clearInlineImagesBtn").classList.add("hidden");
  log("Inline images cleared.", "info");
}

function resolveAttachmentsForRecipient(recipient) {
  const filename = (recipient.attachment || "").trim();
  if (filename) {
    const perFile = perRecipientFiles.get(filename.toLowerCase());
    if (perFile) return [perFile];
    const key = filename.toLowerCase();
    if (!warnedMissingAttachments.has(key)) {
      warnedMissingAttachments.add(key);
      log(`No loaded file matches "${filename}" — falling back to shared attachment(s) for rows that reference it.`, "warning");
    }
    return [...sharedAttachments];
  }
  return [...sharedAttachments];
}

/* ─── CSV PARSING (RFC 4180) ───────────────────────────────────── */

function parseCSV(raw) {
  // RFC 4180-compliant parser that handles quoted fields with embedded newlines

  // Auto-detect delimiter from first non-empty line
  let delimiter = ",";
  const firstLine = raw.split(/\r?\n/).find(function(l) { return l.trim().length > 0; }) || "";
  const tabCount   = (firstLine.match(/\t/g)  || []).length;
  const semiCount  = (firstLine.match(/;/g)   || []).length;
  const commaCount = (firstLine.match(/,/g)   || []).length;
  if (tabCount > commaCount && tabCount > semiCount) delimiter = "\t";
  else if (semiCount > commaCount) delimiter = ";";
  // else keep ","

  const rawRows = [];
  let row = [];
  let field = "";
  let inQuotes = false;
  let i = 0;
  const n = raw.length;

  while (i < n) {
    const ch = raw[i];
    if (inQuotes) {
      if (ch === '"') {
        if (i + 1 < n && raw[i + 1] === '"') {
          // Escaped quote
          field += '"';
          i += 2;
        } else {
          // End of quoted field
          inQuotes = false;
          i++;
        }
      } else {
        field += ch;
        i++;
      }
    } else {
      if (ch === '"') {
        inQuotes = true;
        i++;
      } else if (ch === delimiter) {
        row.push(field);
        field = "";
        i++;
      } else if (ch === '\r') {
        // Handle \r\n and bare \r
        row.push(field);
        field = "";
        rawRows.push(row);
        row = [];
        if (i + 1 < n && raw[i + 1] === '\n') i++;
        i++;
      } else if (ch === '\n') {
        row.push(field);
        field = "";
        rawRows.push(row);
        row = [];
        i++;
      } else {
        field += ch;
        i++;
      }
    }
  }
  // Final field/row
  if (field || row.length > 0) {
    row.push(field);
    rawRows.push(row);
  }
  // Remove trailing empty row
  if (rawRows.length > 0 && rawRows[rawRows.length - 1].every(f => f === "")) {
    rawRows.pop();
  }

  if (rawRows.length < 2) return { headers: [], rows: [] };

  const headers = rawRows[0].map(h => h.trim().toLowerCase());
  const rows    = [];

  for (let r = 1; r < rawRows.length; r++) {
    const cols = rawRows[r].map(c => c.trim());
    // Skip rows that are entirely blank
    if (cols.every(c => c === "")) continue;
    // Pad short rows instead of silently dropping them (handles optional trailing columns)
    while (cols.length < headers.length) cols.push("");
    const rowObj = {};
    headers.forEach((h, idx) => { rowObj[h] = cols[idx] || ""; });
    rowObj._csvRow = r + 1; // L8: +1 for header row, already 1-based from loop start r=1
    rowObj._originalIndex = rows.length; // P4: stamp before push to avoid O(n) indexOf
    rows.push(rowObj);
  }

  return { headers, rows };
}

/* ─── FILTER / SORT STATE ──────────────────────────────────────── */

let activeFilter = null; // { conditions, sortCol, sortDir, logic }
let filterConditionCount = 0;

function populateFilterSortBar(headers) {
  const sortCol = document.getElementById("sortCol");
  // P2: build all options then assign once — avoids repeated innerHTML += reparse
  sortCol.innerHTML = '<option value="">Sort col…</option>' +
    headers.map(h => `<option value="${escapeHtml(h)}">${escapeHtml(h)}</option>`).join("");
  document.getElementById("filterSortBar").classList.remove("hidden");
  // Start with one blank condition row visible
  if (document.getElementById("filterConditions").children.length === 0) {
    addFilterCondition();
  }
}

function addFilterCondition(colVal = "", opVal = "contains", valVal = "") {
  filterConditionCount++;
  const id = filterConditionCount;
  const headers = parsedRecipients.length > 0 ? Object.keys(parsedRecipients[0]).filter(k => !k.startsWith("_")) : [];
  const colOptions = headers.map(h =>
    `<option value="${escapeHtml(h)}" ${h === colVal ? "selected" : ""}>${escapeHtml(h)}</option>`
  ).join("");

  const div = document.createElement("div");
  div.id = `filter-cond-${id}`;
  div.style.cssText = "display:flex;gap:6px;align-items:center;margin-bottom:4px;flex-wrap:wrap;";
  div.innerHTML = `
    <select class="select-inline filter-col-sel" style="min-width:90px;">
      <option value="">col…</option>${colOptions}
    </select>
    <select class="select-inline filter-op-sel" style="min-width:100px;">
      <option value="contains" ${opVal==="contains"?"selected":""}>contains</option>
      <option value="not_contains" ${opVal==="not_contains"?"selected":""}>not contains</option>
      <option value="eq" ${opVal==="eq"?"selected":""}>= equals</option>
      <option value="neq" ${opVal==="neq"?"selected":""}>&ne; not equals</option>
      <option value="gt" ${opVal==="gt"?"selected":""}>&gt; greater than</option>
      <option value="lt" ${opVal==="lt"?"selected":""}>&lt; less than</option>
      <option value="empty" ${opVal==="empty"?"selected":""}>is empty</option>
      <option value="not_empty" ${opVal==="not_empty"?"selected":""}>is not empty</option>
    </select>
    <input type="text" class="input input-inline filter-val-inp" value="${escapeHtml(valVal)}"
           placeholder="value…" style="max-width:100px;" />
    <button class="btn-icon filter-remove-btn" data-cond-id="${id}" title="Remove">&#10005;</button>
  `;
  div.querySelector(".filter-remove-btn").addEventListener("click", () => {
    const condRowEl = document.getElementById("filter-cond-" + id);
    if (condRowEl) condRowEl.remove();
  });
  document.getElementById("filterConditions").appendChild(div);
}

function getFilterConditions() {
  const rows = document.querySelectorAll("#filterConditions > div");
  return [...rows].map(row => ({
    col: row.querySelector(".filter-col-sel").value,
    op:  row.querySelector(".filter-op-sel").value,
    val: row.querySelector(".filter-val-inp").value,
  })).filter(c => c.col); // ignore rows with no column selected
}

function applyCondition(row, cond) {
  const cellVal = String(row[cond.col] !== null && row[cond.col] !== undefined ? row[cond.col] : "").trim();
  const testVal = cond.val.trim();
  const numCell = parseFloat(cellVal);
  const numTest = parseFloat(testVal);
  const bothNum = !isNaN(numCell) && !isNaN(numTest);
  switch (cond.op) {
    case "contains":     return cellVal.toLowerCase().includes(testVal.toLowerCase());
    case "not_contains": return !cellVal.toLowerCase().includes(testVal.toLowerCase());
    case "eq":           return cellVal.toLowerCase() === testVal.toLowerCase();
    case "neq":          return cellVal.toLowerCase() !== testVal.toLowerCase();
    case "gt":           return bothNum ? numCell > numTest : cellVal > testVal;
    case "lt":           return bothNum ? numCell < numTest : cellVal < testVal;
    case "empty":        return cellVal === "";
    case "not_empty":    return cellVal !== "";
    default:             return true;
  }
}

function getFilteredSortedRecipients() {
  let result = parsedRecipients.slice();

  // Row selection
  if (selectedRowIndices !== null) {
    result = result.filter((_, i) => selectedRowIndices.has(i));
  }

  if (activeFilter) {
    const { conditions = [], sortCol, sortDir, logic = "AND" } = activeFilter;
    if (conditions.length) {
      result = result.filter(row => {
        if (logic === "OR") return conditions.some(c => applyCondition(row, c));
        return conditions.every(c => applyCondition(row, c));
      });
    }
    if (sortCol) {
      result.sort((a, b) => {
        const av = String(a[sortCol] !== null && a[sortCol] !== undefined ? a[sortCol] : "").toLowerCase();
        const bv = String(b[sortCol] !== null && b[sortCol] !== undefined ? b[sortCol] : "").toLowerCase();
        const cmp = av < bv ? -1 : av > bv ? 1 : 0;
        return sortDir === "desc" ? -cmp : cmp;
      });
    }
  }
  return result;
}

function applyFilterSort() {
  const conditions = getFilterConditions();
  const sortCol = document.getElementById("sortCol").value;
  const sortDir = document.getElementById("sortDir").value;
  const logic   = document.getElementById("filterLogic").value; // "AND" | "OR"
  activeFilter  = (conditions.length || sortCol) ? { conditions, sortCol, sortDir, logic } : null;
  renderFilteredPreview();
}

function clearFilterSort() {
  activeFilter = null;
  document.getElementById("filterConditions").innerHTML = "";
  filterConditionCount = 0;
  document.getElementById("sortCol").value   = "";
  document.getElementById("sortDir").value   = "asc";
  document.getElementById("filterLogic").value = "AND";
  document.getElementById("filterCountLabel").textContent = "";
  renderFilteredPreview();
}

function renderFilteredPreview() {
  const rows = getFilteredSortedRecipients();
  document.getElementById("filterCountLabel").textContent =
    activeFilter ? `${rows.length} of ${parsedRecipients.length} rows` : "";
  renderPreviewTable(rows);
}

function parseAndPreview() {
  const raw = document.getElementById("csvInput").value;
  const { headers, rows } = parseCSV(raw);

  // Apply field mapping to determine the effective email column
  const emailColMapped = fieldMapping.email || null;
  const hasEmailHeader = headers.includes("email") || (emailColMapped && headers.includes(emailColMapped));

  if (!hasEmailHeader) {
    // Show map fields button and auto-open modal
    document.getElementById("mapFieldsBtn").style.display = "";
    if (!fieldMapping.email) {
      log("No 'email' column detected — opening Match Fields to map columns.", "warning");
      openMatchFieldsModal(headers);
    } else {
      log("CSV must contain an 'email' column (or map a column to email).", "error");
    }
    return;
  }

  // Apply field mapping to all rows
  const mappedRows = rows.map(row => applyFieldMapping(row));

  if (mappedRows.length > MAX_RECIPIENTS) {
    log(`🛑 CSV contains ${mappedRows.length.toLocaleString()} rows — exceeds the Microsoft 365 hard limit of ` +
        `10,000 outbound recipients per 24 hours. Trim your list before sending.`, "error");
    document.getElementById("recipientCount").textContent =
      `${mappedRows.length.toLocaleString()} recipients — OVER LIMIT`;
    parsedRecipients = [];   // clear so old list cannot be sent accidentally
    return;
  }

  parsedRecipients = mappedRows;
  sendOutcomes = [];   // BUG 14: reset outcomes on new CSV load so they accumulate across retries
  activeFilter = null; // reset filter on fresh parse
  previewTablePage = 0; // Feature 6: reset paging on new CSV load
  // A9: Clear stale condition rows DOM and counter so old rows don't persist across CSV reloads
  document.getElementById("filterConditions").innerHTML = "";
  filterConditionCount = 0;
  selectedRowIndices = null; // reset row selection on fresh parse
  // Bug 4: clear failed recipients list and hide retry button on new CSV load
  failedRecipients = [];
  const _parseRetryBtn = document.getElementById("retryFailedBtn");
  if (_parseRetryBtn) { _parseRetryBtn.classList.add("hidden"); _parseRetryBtn.textContent = ""; }

  // Feature 13: hide getting-started banner once CSV is loaded
  const bannerEl = document.getElementById("gettingStartedBanner");
  if (bannerEl) bannerEl.classList.add("hidden");

  // Feature 1: rebuild test-row selector
  const testRowSel = document.getElementById("testRowSelect");
  if (testRowSel) {
    testRowSel.innerHTML = mappedRows.map(function(r, i) {
      const label = escapeHtml((r.email || r.first_name || "").slice(0, 28)); // S2: escape CSV data
      return '<option value="' + i + '">Row ' + (i + 1) + ': ' + label + '</option>';
    }).join("");
  }

  // Show map fields button now that we have headers
  document.getElementById("mapFieldsBtn").style.display = "";

  const emailCounts = {};
  mappedRows.forEach(r => {
    const e = (r.email || "").toLowerCase();
    if (e) emailCounts[e] = (emailCounts[e] || 0) + 1;
  });
  const dupes = Object.entries(emailCounts).filter(([, count]) => count > 1);
  if (dupes.length > 0) {
    const dupeList = dupes.map(([email, count]) => `${email} (×${count})`).join(", ");
    log(`Warning: ${dupes.length} duplicate email address${dupes.length !== 1 ? "es" : ""} found — ` +
        `${dupeList}. Enable "Remove duplicates" in send options or edit the file.`, "warning");
  }

  headers.forEach(h => {
    const tag = `{{${h}}}`;
    if (!DEFAULT_TAGS.includes(tag)) addTagToBar(tag);
  });

  // Feature 4: always expose merge_table and unsubscribe_link as smart tags
  addTagToBar("{{merge_table}}", "smart");
  addTagToBar("{{unsubscribe_link}}", "smart");

  document.getElementById("recipientCount").textContent =
    `${mappedRows.length} recipient${mappedRows.length !== 1 ? "s" : ""} loaded`;

  populateFilterSortBar(Object.keys(mappedRows[0] || {}).filter(k => !k.startsWith("_")));
  renderPreviewTable(parsedRecipients);
  populateInsertFieldSelect(headers);
  log(`Parsed ${mappedRows.length} recipients. Showing preview.`, "success");
}

function renderPreviewTable(rows) {
  const container = document.getElementById("previewTable");
  container.classList.remove("hidden");

  if (!rows || rows.length === 0) {
    container.innerHTML = '<p class="hint" style="padding:4px;">No rows to display.</p>';
    return;
  }

  // Feature 6: paging
  const totalPages = Math.ceil(rows.length / PREVIEW_PAGE_SIZE);
  previewTablePage = Math.min(previewTablePage, Math.max(0, totalPages - 1));
  const pageRows = rows.slice(previewTablePage * PREVIEW_PAGE_SIZE, (previewTablePage + 1) * PREVIEW_PAGE_SIZE);

  const headers = Object.keys(rows[0]).filter(h => !h.startsWith("_"));
  let html = '<div style="display:flex;align-items:center;gap:6px;margin-bottom:4px;">' +
    '<label class="label" style="margin:0;">Preview (rows ' + (previewTablePage * PREVIEW_PAGE_SIZE + 1) + '–' + Math.min((previewTablePage + 1) * PREVIEW_PAGE_SIZE, rows.length) + ' of ' + rows.length + ')</label>' +
    '<button class="btn-secondary btn-edit-table" id="openEditTableBtn">✏ Edit table</button>' +
    '</div>';
  html += '<table class="preview-table"><thead><tr>';
  html += '<th><input type="checkbox" id="selectAllRowsChk" checked title="Select/deselect all" /></th>';
  headers.forEach(h => { html += '<th>' + escapeHtml(h) + '</th>'; });
  html += '</tr></thead><tbody>';
  pageRows.forEach(function(row) {
    const originalIdx = row._originalIndex !== undefined ? row._originalIndex : parsedRecipients.indexOf(row); // P4
    const isChecked = (selectedRowIndices === null || selectedRowIndices.has(originalIdx)) ? "checked" : "";
    html += '<tr>';
    html += '<td><input type="checkbox" class="row-select-chk" data-idx="' + originalIdx + '" ' + isChecked + ' /></td>';
    headers.forEach(h => { html += '<td>' + escapeHtml(row[h] || "") + '</td>'; });
    html += '</tr>';
  });
  html += '</tbody></table>';

  // Paging controls
  if (totalPages > 1) {
    html += '<div class="preview-paging">' +
      '<button id="prevPageBtn" class="btn-sm"' + (previewTablePage === 0 ? ' disabled' : '') + '>‹ Prev</button>' +
      '<span>Page ' + (previewTablePage + 1) + ' of ' + totalPages + ' (' + rows.length + ' total)</span>' +
      '<button id="nextPageBtn" class="btn-sm"' + (previewTablePage >= totalPages - 1 ? ' disabled' : '') + '>Next ›</button>' +
      '</div>';
  }

  container.innerHTML = html;

  // Wire paging buttons
  var prevBtn = document.getElementById("prevPageBtn");
  var nextBtn = document.getElementById("nextPageBtn");
  if (prevBtn) {
    prevBtn.addEventListener("click", function() {
      previewTablePage = Math.max(0, previewTablePage - 1);
      renderPreviewTable(rows);
    });
  }
  if (nextBtn) {
    nextBtn.addEventListener("click", function() {
      previewTablePage = Math.min(totalPages - 1, previewTablePage + 1);
      renderPreviewTable(rows);
    });
  }

  // Wire "Edit table" button — opens the in-memory edit modal
  const openEditBtn = document.getElementById("openEditTableBtn");
  if (openEditBtn) openEditBtn.addEventListener("click", openEditTableModal);

  // Wire up header checkbox
  document.getElementById("selectAllRowsChk").addEventListener("change", function(e) {
    if (e.target.checked) {
      // Select all — null means "all included"
      selectedRowIndices = null;
      document.querySelectorAll(".row-select-chk").forEach(function(chk) { chk.checked = true; });
    } else {
      // Deselect all — empty Set means "none included"
      selectedRowIndices = new Set();
      document.querySelectorAll(".row-select-chk").forEach(function(chk) { chk.checked = false; });
    }
    updateSelectionCount();
  });

  // Wire up row checkboxes
  document.querySelectorAll(".row-select-chk").forEach(function(chk) {
    chk.addEventListener("change", function() {
      const idx = parseInt(chk.dataset.idx, 10);
      if (chk.checked) {
        if (selectedRowIndices !== null) {
          selectedRowIndices.add(idx);
          // If now all are selected, reset to null
          if (selectedRowIndices.size === parsedRecipients.length) {
            selectedRowIndices = null;
          }
        }
        // If null, already all selected — nothing to do
      } else {
        if (selectedRowIndices === null) {
          // Was all-selected; build full Set and remove this one
          selectedRowIndices = new Set(parsedRecipients.map(function(_, i) { return i; }));
          selectedRowIndices.delete(idx);
        } else {
          selectedRowIndices.delete(idx);
        }
      }
      const allChks = document.querySelectorAll(".row-select-chk");
      const allChecked = [...allChks].every(function(c) { return c.checked; });
      const selectAllChk = document.getElementById("selectAllRowsChk");
      if (selectAllChk) selectAllChk.checked = allChecked;
      updateSelectionCount();
    });
  });
}

function updateSelectionCount() {
  const total = parsedRecipients.length;
  const selected = selectedRowIndices === null ? total : selectedRowIndices.size;
  const countEl = document.getElementById("recipientCount");
  if (countEl) {
    countEl.textContent = selected === total
      ? `${total} recipients`
      : `${selected} of ${total} selected`;
  }
}

/* ─── TOKEN REPLACEMENT ────────────────────────────────────────── */

function resolveGreetingLine(recipient) {
  const sal   = String(recipient.salutation  || "").trim();
  const first = String(recipient.first_name  || "").trim();
  const last  = String(recipient.last_name   || "").trim();
  const cfg   = greetingConfig;
  const fallback = cfg.fallback || "Dear Valued Customer";

  let parts = [];
  switch (cfg.format) {
    case "dear_sal_last":
      parts = ["Dear", sal, last].filter(Boolean);
      break;
    case "dear_first":
      parts = ["Dear", first].filter(Boolean);
      break;
    case "dear_first_last":
      parts = ["Dear", first, last].filter(Boolean);
      break;
    case "hi_first":
      parts = ["Hi", first].filter(Boolean);
      break;
    case "to_sal_last":
      parts = ["To", sal, last].filter(Boolean);
      break;
    default:
      parts = ["Dear", sal, last].filter(Boolean);
  }
  // Need at least one name component
  const hasName = (sal || first || last);
  return hasName && parts.length > 1 ? parts.join(" ") : fallback;
}

/* ─── CONDITIONAL MERGE ────────────────────────────────────────── */

/**
 * Processes {{if:column=value:true text:false text}} conditionals before token replacement.
 * Supports operators: =, !=, >, <, >=, <=
 * Also supports truthy check: {{if:column:true text:false text}}
 */
function processConditionals(template, recipient) {
  return template.replace(/\{\{if:([^}]+)\}\}/gi, (match, inner) => {
    // Find first two colons to separate: condition : trueText : falseText
    const firstColon  = inner.indexOf(":");
    if (firstColon === -1) return match;
    const secondColon = inner.indexOf(":", firstColon + 1);
    if (secondColon === -1) return match;

    const condPart  = inner.slice(0, firstColon).trim();
    const trueText  = inner.slice(firstColon + 1, secondColon);
    const falseText = inner.slice(secondColon + 1);

    const opMatch = condPart.match(/^(.+?)(>=|<=|!=|=|>|<)(.+)$/);
    let result;
    if (opMatch) {
      const col     = opMatch[1].trim().toLowerCase();
      const op      = opMatch[2];
      const val     = opMatch[3].trim();
      const cellVal = String(recipient[col] !== null && recipient[col] !== undefined ? recipient[col] : "").trim();
      const numCell = parseFloat(cellVal);
      const numVal  = parseFloat(val);
      const bothNum = !isNaN(numCell) && !isNaN(numVal);
      if      (op === "=")  result = cellVal.toLowerCase() === val.toLowerCase();
      else if (op === "!=") result = cellVal.toLowerCase() !== val.toLowerCase();
      else if (op === ">")  result = bothNum ? numCell > numVal  : cellVal > val;
      else if (op === "<")  result = bothNum ? numCell < numVal  : cellVal < val;
      else if (op === ">=") result = bothNum ? numCell >= numVal : cellVal >= val;
      else if (op === "<=") result = bothNum ? numCell <= numVal : cellVal <= val;
      else result = false;
    } else {
      const _cv = recipient[condPart.toLowerCase()]; result = !!String(_cv !== null && _cv !== undefined ? _cv : "").trim();
    }
    return result ? trueText : falseText;
  });
}

/* ─── MERGE FIELD FILTERS ──────────────────────────────────────── */

function applyFilter(value, filter) {
  if (!filter) return value;
  const f = filter.trim().toLowerCase();
  if (f === "upper")    return value.toUpperCase();
  if (f === "lower")    return value.toLowerCase();
  if (f === "title")    return value.replace(/\b\w/g, c => c.toUpperCase());
  if (f === "trim")     return value.trim();
  if (f === "currency") {
    const n = parseFloat(value);
    return isNaN(n) ? value : n.toLocaleString(undefined, { style: "currency", currency: "USD", minimumFractionDigits: 2 });
  }
  if (f === "number") {
    const n = parseFloat(value);
    return isNaN(n) ? value : n.toLocaleString();
  }
  if (f === "date") {
    const d = new Date(value);
    return isNaN(d.getTime()) ? value : d.toLocaleDateString(undefined, { year: "numeric", month: "long", day: "numeric" });
  }
  // Unknown filter name = treat as fallback value for empty fields
  return value.trim() ? value : filter;
}

/**
 * @param {boolean} escapeValues — true for HTML body (prevents CSV data from injecting HTML tags),
 *                                 false for plain-text fields like subject line
 *
 * Single-pass replacement: one regex scans the template once, looking up each matched token in a Map.
 * This avoids two bugs in the previous per-key approach:
 *   1. Regex injection — column names like "order.id" would create `.` wildcards in the RegExp.
 *   2. Double-replacement — a value containing "{{another_token}}" would be substituted again.
 *
 * Two-pass escape fix (Bug 1): raw CSV values are escaped first, then greeting_line is built from
 * already-escaped values so it is never double-escaped.
 *
 * Pipe filter support (Feature 3): {{field|upper}}, {{amount|currency}}, etc.
 */
function personalize(template, recipient, escapeValues = true) {
  // Apply fill-in values (Feature 6) before any other processing
  if (window._fillInValues) {
    template = applyFillInValues(template, window._fillInValues);
  }
  template = processConditionals(template, recipient);
  const now = new Date();

  // First pass: escape (or not) all raw CSV values
  const escapedRecipient = {};
  Object.keys(recipient).forEach(key => {
    escapedRecipient[key.toLowerCase()] = escapeValues
      ? escapeHtml(String(recipient[key] !== null && recipient[key] !== undefined ? recipient[key] : ""))
      : String(recipient[key] !== null && recipient[key] !== undefined ? recipient[key] : "");
  });

  // Build greeting_line from already-escaped values — don't escape again
  const greetingRecipient = escapeValues ? escapedRecipient : recipient;
  const greetingLine = resolveGreetingLine({
    salutation: (greetingRecipient.salutation !== null && greetingRecipient.salutation !== undefined ? greetingRecipient.salutation : (recipient.salutation !== null && recipient.salutation !== undefined ? recipient.salutation : "")),
    first_name: (greetingRecipient.first_name !== null && greetingRecipient.first_name !== undefined ? greetingRecipient.first_name : (recipient.first_name !== null && recipient.first_name !== undefined ? recipient.first_name : "")),
    last_name:  (greetingRecipient.last_name  !== null && greetingRecipient.last_name  !== undefined ? greetingRecipient.last_name  : (recipient.last_name  !== null && recipient.last_name  !== undefined ? recipient.last_name  : "")),
    gender:     (greetingRecipient.gender     !== null && greetingRecipient.gender     !== undefined ? greetingRecipient.gender     : (recipient.gender     !== null && recipient.gender     !== undefined ? recipient.gender     : ""))
  });

  const lookup = new Map(Object.entries(escapedRecipient));
  lookup.set("greeting_line", greetingLine);
  lookup.set("today", now.toLocaleDateString(undefined, { year: "numeric", month: "long", day: "numeric" }));
  lookup.set("now",   now.toLocaleTimeString(undefined, { hour: "2-digit", minute: "2-digit" }));

  // Single pass: match {{key}} or {{key|filter}} and replace if the key is known; leave unknown tokens intact.
  TOKEN_REGEX.lastIndex = 0; // P3: reset global regex state before reuse
  let result = template.replace(TOKEN_REGEX, (match, key, filter) => {
    const val = lookup.get(key.trim().toLowerCase());
    if (val === undefined) return match;
    return applyFilter(val, filter ? filter.trim() : null);
  });

  // Replace {{merge_table}} with grouped rows table if applicable
  if (recipient._groupedRows && recipient._groupedRows.length > 1) {
    const headers = Object.keys(recipient._groupedRows[0]).filter(k => !k.startsWith("_"));
    result = result.replace(/\{\{merge_table\}\}/gi, buildMergeTable(recipient._groupedRows, headers));
  } else {
    result = result.replace(/\{\{merge_table\}\}/gi, "");
  }

  // {{unsubscribe_link}} → mailto: opt-out link
  if (result.indexOf("{{unsubscribe_link}}") !== -1 || /\{\{unsubscribe_link\}\}/i.test(result)) {
    const recipEmail = encodeURIComponent(recipient.email || "");
    const senderAddr = (Office.context.mailbox && Office.context.mailbox.userProfile && Office.context.mailbox.userProfile.emailAddress) || "";
    const unsubLink = '<a href="mailto:' + escapeHtml(senderAddr) + '?subject=UNSUBSCRIBE%20' + recipEmail + '&body=Please%20remove%20me%20from%20your%20mailing%20list." style="color:#888;font-size:11px;">Unsubscribe</a>'; // S3
    result = result.replace(/\{\{unsubscribe_link\}\}/gi, unsubLink);
  }

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

function stripHtmlToText(html) {
  return html
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n\n")
    .replace(/<\/div>/gi, "\n")
    .replace(/<\/li>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#039;/g, "'")
    .replace(/&nbsp;/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

/* ─── ADDRESS LIST PARSER ──────────────────────────────────────── */

function parseAddressList(raw) {
  if (!raw || !raw.trim()) return [];
  return raw.split(/[;,\s]+/)
    .map(a => a.trim())
    .filter(a => a.length > 0 && EMAIL_REGEX.test(a))
    .map(address => ({ emailAddress: { address } }));
}

function parseCustomHeaders(str) {
  if (!str || !str.trim()) return [];
  return str.split("\n")
    .map(line => line.trim())
    .filter(line => line.includes(":"))
    .map(line => {
      const colonIdx = line.indexOf(":");
      return {
        name:  line.slice(0, colonIdx).trim().replace(/[\r\n]/g, ""),   // S6: strip CRLF injection
        value: line.slice(colonIdx + 1).trim().replace(/[\r\n]/g, "")   // S6: strip CRLF injection
      };
    })
    .filter(h => h.name && h.value);
}

/* ─── EMAIL VALIDATION ─────────────────────────────────────────── */

function validateRecipients(recipients) {
  const valid   = [];
  const invalid = [];
  for (let i = 0; i < recipients.length; i++) {
    const r = recipients[i];

    // skip_if column: skip row if value is truthy (non-empty, non-zero, non-false)
    const skipVal = String(r.skip_if !== null && r.skip_if !== undefined ? r.skip_if : "").trim().toLowerCase();
    if (skipVal && skipVal !== "0" && skipVal !== "false" && skipVal !== "no") {
      log(`Row ${r._csvRow || (i + 2)} (${r.email || "?"}): skipped — skip_if = "${r.skip_if}"`, "info");
      continue;
    }

    if (!r.email) {
      log(`Row ${r._csvRow || (i + 2)}: missing email address — row skipped.`, "warning");
      continue;
    }
    const addresses = r.email.split(/[;,\s]+/).map(s => s.trim()).filter(Boolean);
    // Log a warning for each invalid address in a multi-address row
    addresses.forEach(addr => {
      if (!EMAIL_REGEX.test(addr)) {
        log(`Row ${r._csvRow || (i + 2)}: invalid address skipped — "${addr}"`, "warning");
      }
    });
    const hasValid  = addresses.some(a => EMAIL_REGEX.test(a));
    if (hasValid) { valid.push(r); }
    else          { invalid.push({ row: i + 2, email: r.email }); }
  }
  return { valid, invalid };
}

/* ─── AUTH ─────────────────────────────────────────────────────── */

function ssoErrorMessage(code) {
  switch (code) {
    case 13001:
      return "Your session has expired or requires re-authorization. " +
             "Please close and reopen the Mail Merge task pane, then sign in again when prompted. " +
             "Your failed recipients will be preserved for retry. " +
             "(Tip: sign in with your M365 email address, not your JumpCloud credentials.)";
    case 13002:
      return "Authentication was cancelled or access was denied. Please try again.";
    case 13003:
      return "Personal Microsoft accounts (Outlook.com / Hotmail) are not supported — " +
             "this add-in requires a Microsoft 365 work or school account with Exchange Online. " +
             "If you are using a JumpCloud-managed account, ensure your user is provisioned " +
             "and synced to Microsoft 365 in the JumpCloud Cloud Directory Integration.";
    case 13004:
      return "Add-in configuration error — the Entra ID app registration may be misconfigured. " +
             "Contact your IT administrator and reference SSO error 13004. " +
             "If your organisation uses JumpCloud as its IdP, ensure federation with " +
             "Microsoft 365/Entra ID is enabled and the add-in has been admin-consented.";
    case 13005:
      return "Your account is not configured for this add-in. " +
             "If you use JumpCloud SSO, ensure your user account is bound to the " +
             "M365 Cloud Directory Integration in JumpCloud before signing in. " +
             "Contact your IT administrator and reference SSO error 13005.";
    case 13006:
      return "An error occurred in Office. Please save your work, restart Outlook, and try again.";
    case 13007:
      return "Your session has expired or requires re-authorization. " +
             "Please close and reopen the Mail Merge task pane, then sign in again when prompted. " +
             "Your failed recipients will be preserved for retry.";
    case 13008:
      return "Your organisation's policies prevented the consent prompt from appearing. " +
             "An administrator must grant permissions for this add-in in Entra ID.";
    case 13009:
      return "Admin consent is required before this add-in can send email. " +
             "Your Microsoft 365 administrator must grant permissions. Ask them to visit: " +
             "https://login.microsoftonline.com/common/adminconsent" +
             "?client_id=a3a648da-9dc0-48ce-948a-ba2434afcadd";
    case 13010:
      return "A navigation error occurred in the sign-in flow. " +
             "Please try again. If the problem persists, update Edge.";
    case 13012:
      // B3: Add-in not supported in current Outlook host
      return "This add-in is not supported in your version of Outlook. Please update Outlook or use Outlook on the web.";
    case 13013:
      return "Too many authentication requests. Please wait a moment and try again.";
    default:
      return `Authentication failed (code ${code}). ` +
             "Please restart Outlook and try again. " +
             "If the problem persists, contact your IT administrator.";
  }
}

function getAccessToken() {
  return new Promise((resolve, reject) => {
    Office.context.auth.getAccessTokenAsync(
      { allowSignInPrompt: true, allowConsentPrompt: true, forMSGraphAccess: true },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          const code = result.error.code;
          // SSO errors that indicate config/federation issues — fall back to MSAL dialog
          // (common with JumpCloud IdP, misconfigured Entra, or tenant SSO restrictions)
          const fallbackCodes = [13001, 13002, 13003, 13004, 13005, 13006, 13007, 13008, 13009, 13010, 13012, 13013];
          if (fallbackCodes.includes(code)) {
            log("SSO unavailable (code " + code + ") — opening sign-in dialog…", "warning");
            getTokenViaDialog().then(resolve).catch(reject);
          } else {
            reject(new Error(ssoErrorMessage(code)));
          }
        }
      }
    );
  });
}

/* ─── MSAL DIALOG FALLBACK (for JumpCloud/federation SSO failures) ─ */

function getTokenViaDialog() {
  return new Promise((resolve, reject) => {
    const dialogUrl = "https://leighton-grey.github.io/mail-merge-addin/auth-dialog.html?v=5";
    Office.context.ui.displayDialogAsync(
      dialogUrl,
      { height: 60, width: 35, promptBeforeOpen: false },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          reject(new Error("Could not open sign-in dialog: " + asyncResult.error.message));
          return;
        }
        const dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          dialog.close();
          try {
            const msg = JSON.parse(arg.message);
            if (msg.type === "token") {
              resolve(msg.token);
            } else {
              reject(new Error(msg.message || "Authentication failed"));
            }
          } catch {
            reject(new Error("Invalid response from auth dialog"));
          }
        });
        dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
          // 12006 = user closed the dialog manually
          reject(new Error(arg.error === 12006 ? "Sign-in cancelled" : "Dialog closed unexpectedly (error " + arg.error + ")"));
        });
      }
    );
  });
}

/* ─── PRE-FLIGHT CHECKS ────────────────────────────────────────── */

function checkPayloadSize(bodyTemplate) {
  const bytes = new TextEncoder().encode(bodyTemplate).length;
  if (bytes > MAX_PAYLOAD_BYTES) {
    log(`⚠️ Email body is ~${(bytes / 1024 / 1024).toFixed(2)} MB — dangerously close to the 4 MB ` +
        `Graph API per-message limit. Compress or remove inline images before sending.`, "warning");
  }
}

/* ─── SENDING WINDOW (Feature 2) ──────────────────────────────── */

function msUntilWindowOpens() {
  const sendingWindowEl = document.getElementById("sendingWindowEnabled");
  const enabled = sendingWindowEl && sendingWindowEl.checked;
  if (!enabled) return 0;
  const now = new Date();
  const day = now.getDay(); // 0=Sun, 6=Sat
  const startParts = (document.getElementById("windowStart").value || "09:00").split(":").map(Number);
  const endParts = (document.getElementById("windowEnd").value || "17:00").split(":").map(Number);
  const startMins = startParts[0] * 60 + startParts[1];
  const endMins = endParts[0] * 60 + endParts[1];
  // L5: guard against inverted window (end <= start) which causes an infinite loop
  if (startMins >= endMins) {
    log("⚠ Sending window: end time must be after start time. Window check skipped.", "warning");
    return 0; // don't block sending
  }
  const nowMins = now.getHours() * 60 + now.getMinutes();

  const isWeekday = day >= 1 && day <= 5;
  const inWindow = isWeekday && nowMins >= startMins && nowMins < endMins;
  if (inWindow) return 0;

  // Calculate ms until next window open
  let next = new Date(now);
  next.setSeconds(0, 0);
  // If today is a weekday and before window start, wait until today's start
  if (isWeekday && nowMins < startMins) {
    next.setHours(startParts[0], startParts[1], 0, 0);
    return next - now;
  }
  // Otherwise: next Monday (or next weekday) at window start
  let daysAhead = 1;
  while (true) {
    const nextDay = (day + daysAhead) % 7;
    if (nextDay >= 1 && nextDay <= 5) break;
    daysAhead++;
  }
  next.setDate(next.getDate() + daysAhead);
  next.setHours(startParts[0], startParts[1], 0, 0);
  return next - now;
}

/* ─── RATE LIMITING (Feature 3) ───────────────────────────────── */

let emailsSentThisHour = 0;
let emailsSentToday = 0;
let hourWindowStart = Date.now();
let dayWindowStart = new Date().setHours(0, 0, 0, 0);

/* ─── RATE LIMIT STATE PERSISTENCE (Feature 8) ─────────────────── */

function loadRateLimitState() {
  const saved = lsGet("mm_rate_state", null);
  if (saved) {
    const now = Date.now();
    if (saved.hourWindowStart && (now - saved.hourWindowStart) < 3600000) {
      emailsSentThisHour = saved.emailsSentThisHour || 0;
      hourWindowStart = saved.hourWindowStart;
    }
    const todayStart = new Date().setHours(0, 0, 0, 0);
    if (saved.dayWindowStart && saved.dayWindowStart >= todayStart) {
      emailsSentToday = saved.emailsSentToday || 0;
      dayWindowStart = saved.dayWindowStart;
    }
  }
}

function saveRateLimitState() {
  lsSet("mm_rate_state", JSON.stringify({ emailsSentThisHour, emailsSentToday, hourWindowStart, dayWindowStart }));
}

async function checkRateLimits() {
  // Refresh hour window if needed
  if (Date.now() - hourWindowStart >= 3600000) {
    hourWindowStart = Date.now();
    emailsSentThisHour = 0;
  }
  // Refresh day window if needed
  if (Date.now() - dayWindowStart >= 86400000) {
    dayWindowStart = new Date().setHours(0, 0, 0, 0);
    emailsSentToday = 0;
  }

  const maxPerHourEl2 = document.getElementById("maxPerHour");
  const dailyCapEl2   = document.getElementById("dailyCap");
  const maxPerHour = parseInt(maxPerHourEl2 ? maxPerHourEl2.value : "0", 10);
  const dailyCap   = parseInt(dailyCapEl2   ? dailyCapEl2.value   : "0", 10);

  if (dailyCap > 0 && emailsSentToday >= dailyCap) {
    log(`Daily cap of ${dailyCap} reached. Stopping merge.`, "warning");
    return false;
  }

  if (maxPerHour > 0 && emailsSentThisHour >= maxPerHour) {
    const msUntilReset = hourWindowStart + 3600000 - Date.now();
    const waitMins = Math.ceil(msUntilReset / 60000);
    log(`⏸ Hourly limit of ${maxPerHour} reached. Pausing ${waitMins} min until hour resets…`, "warning");
    await new Promise(r => setTimeout(r, msUntilReset));
    hourWindowStart = Date.now();
    emailsSentThisHour = 0;
    log("▶ Resuming sends (hour window reset).", "info");
  }

  return true;
}

/* ─── GRAPH BATCH SEND ─────────────────────────────────────────── */

function buildEmailRequest(id, toEmail, opts) {
  const {
    subject, htmlBody, saveToSent,
    attachments = [], inlineImages: inlineImagesOpt = [],
    replyTo, sendAs,
    cc = [], bcc = [],
    importance, isReadReceiptRequested = false, isDeliveryReceiptRequested = false,
    customHeaders = [], plainText = false,
    sensitivity, categories = [],
    displayName = "", flagged = false, expiryISO = null
  } = opts;

  const toAddresses     = toEmail.split(";").map(s => s.trim()).filter(s => EMAIL_REGEX.test(s));
  const bodyContent     = plainText ? stripHtmlToText(htmlBody) : htmlBody;
  const bodyContentType = plainText ? "Text" : "HTML";

  const toRecipients = toAddresses.map((address, i) => {
    if (i === 0 && displayName) {
      return { emailAddress: { address, name: displayName } };
    }
    return { emailAddress: { address } };
  });

  const message = {
    subject,
    body: { contentType: bodyContentType, content: bodyContent },
    toRecipients
  };

  if (importance && importance !== "normal") message.importance = importance;
  if (sensitivity && sensitivity !== "normal") message.sensitivity = sensitivity;
  if (categories && categories.length > 0) message.categories = categories;
  if (isReadReceiptRequested)     message.isReadReceiptRequested   = true;
  if (isDeliveryReceiptRequested) message.isDeliveryReceiptRequested = true;
  if (replyTo) message.replyTo = [{ emailAddress: { address: replyTo } }];
  if (sendAs)  message.from    = { emailAddress: { address: sendAs } };
  if (cc  && cc.length  > 0) message.ccRecipients  = cc;
  if (bcc && bcc.length > 0) message.bccRecipients = bcc;
  if (customHeaders && customHeaders.length > 0) message.internetMessageHeaders = customHeaders;
  if (flagged) message.flag = { flagStatus: "flagged" };
  if (expiryISO) message.expiryDateTime = { dateTime: expiryISO, timeZone: "UTC" };

  const fileAttachments = attachments.map(a => ({
    "@odata.type": "#microsoft.graph.fileAttachment",
    name: a.name, contentType: a.contentType, contentBytes: a.contentBytes
  }));
  const inlineAttachments = inlineImagesOpt.map(img => ({
    "@odata.type": "#microsoft.graph.fileAttachment",
    name: img.name, contentType: img.contentType, contentBytes: img.contentBytes,
    isInline: true, contentId: img.name
  }));
  const allAttachments = [...fileAttachments, ...inlineAttachments];
  if (allAttachments.length > 0) message.attachments = allAttachments;

  const batchUrl  = draftsMode ? "/me/messages" : "/me/sendMail";
  const batchBody = draftsMode
    ? message
    : { message, saveToSentItems: saveToSent };

  return {
    id: String(id),
    method: "POST",
    url: batchUrl,
    headers: { "Content-Type": "application/json" },
    body: batchBody
  };
}

function effectiveBatchSize(attachmentSizeBytes) {
  if (!attachmentSizeBytes) return BATCH_SIZE;
  const base64Bytes   = Math.ceil(attachmentSizeBytes * 1.37);
  const perEmailBytes = base64Bytes + 4096;
  return Math.max(1, Math.min(BATCH_SIZE, Math.floor(3_500_000 / perEmailBytes)));
}

function buildUnsubHeaders(url) {
  return [
    { name: "List-Unsubscribe",      value: `<${url}>` },
    { name: "List-Unsubscribe-Post", value: "List-Unsubscribe=One-Click" }
  ];
}

function chunkArray(arr, size) {
  const chunks = [];
  for (let i = 0; i < arr.length; i += size) chunks.push(arr.slice(i, i + size));
  return chunks;
}

function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function getJitterMs() {
  const minEl = document.getElementById("jitterMinInput");
  const maxEl = document.getElementById("jitterMaxInput");
  const min = parseInt(minEl ? minEl.value : "0", 10);
  const max = parseInt(maxEl ? maxEl.value : "0", 10);
  if (!min && !max) return 0;
  const lo = Math.min(min, max);
  const hi = Math.max(min, max);
  return (lo + Math.random() * (hi - lo)) * 1000;
}

/**
 * Send a Microsoft Graph $batch request with automatic retry on rate-limit (429) responses.
 *
 * Why non-recursive: earlier versions used tail-call recursion which caused
 * "maximum call stack exceeded" errors when a persistent 429 produced many
 * retry rounds. This iterative version uses a for-loop with a pending queue
 * to avoid that, while still retrying only the throttled sub-set.
 *
 * Retry logic:
 *   - If the top-level batch response is 429, wait Retry-After seconds and
 *     retry all pending requests (the whole batch was rejected by the gateway).
 *   - If individual requests inside the batch are 429, collect them and retry
 *     only those on the next loop iteration.
 *   - After MAX_RETRIES exhausted attempts, mark remaining items as failed
 *     rather than looping forever.
 *
 * @param {Object[]} requests - Graph batch request objects (id, method, url, headers, body)
 * @param {string}   token    - OAuth2 Bearer token for the Authorization header
 * @returns {Promise<{responses: Object[]}>} - Combined responses for all request IDs
 */
async function sendBatchWithRetry(requests, token) {
  let pending = requests.slice();
  const results = [];

  for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
    if (!pending.length) break;

    const res = await fetch(GRAPH_BATCH_URL, {
      method: "POST",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify({ requests: pending })
    });

    if (!res.ok) {
      if (res.status === 429) {
        const retryAfter = parseInt(res.headers.get("Retry-After") || "10", 10);
        if (attempt < MAX_RETRIES) {
          log(`Rate limited (429). Retrying entire batch in ${retryAfter}s (attempt ${attempt}/${MAX_RETRIES})...`, "warning");
          await delay(retryAfter * 1000);
          continue; // retry all pending
        }
        // Final attempt still throttled — fail all pending
        log(`Rate limited after ${MAX_RETRIES} attempts — marking ${pending.length} request(s) as failed.`, "error");
        pending.forEach(req => results.push({ id: req.id, status: 429, body: { error: { message: "Rate limited" } } }));
        break;
      }
      throw new Error(`Batch HTTP error: ${res.status} ${res.statusText}`);
    }

    const data = await res.json();
    const throttled = [];

    for (const resp of (data.responses || [])) {
      if (resp.status === 429) {
        const original = pending.find(r => r.id === resp.id);
        if (original) throttled.push(original);
      } else {
        results.push(resp);
      }
    }

    if (throttled.length && attempt < MAX_RETRIES) {
      const retryAfter = parseInt(
        ((data.responses || []).find(r => r.status === 429)?.headers?.["Retry-After"]) || "10", 10
      );
      log(`${throttled.length} request(s) throttled. Retrying in ${retryAfter}s (attempt ${attempt}/${MAX_RETRIES})...`, "warning");
      await delay(retryAfter * 1000);
      pending = throttled; // narrow to only throttled items
      continue;
    } else if (throttled.length) {
      // Exhausted retries — record as failed
      log(`${throttled.length} request(s) still throttled after ${MAX_RETRIES} attempts — marking as failed.`, "error");
      throttled.forEach(req => results.push({ id: req.id, status: 429, body: { error: { message: "Rate limited after retries" } } }));
    }
    break; // all items resolved
  }

  return { responses: results };
}

/* ─── SCHEDULED SEND ───────────────────────────────────────────── */

async function sendScheduledMessages(
  recipients, subjectTemplate, bodyTemplate, saveToSent,
  replyTo, sendAs, scheduledTimeISO,
  importance, isReadReceiptRequested, isDeliveryReceiptRequested,
  customHeaders, plainText,
  sensitivity, categories,
  bccSelf, selfEmail, flagged, expiryISO, listUnsubscribeTemplate, inlineImagesArr,
  globalOffset = 0,       // A6: offset into the full valid array so record_num is correct in mixed runs
  totalRecipientCount = 0 // L4: total valid count for correct record_count in mixed runs
) {
  if (!saveToSent) {
    log("⚠ Scheduled send: 'Save to Sent' cannot be suppressed for scheduled messages — Exchange always saves deferred sends. Continuing.", "warning");
  }
  log(`Scheduling ${recipients.length} emails for ${new Date(scheduledTimeISO).toLocaleString()}...`, "info");
  log("Exchange deferred delivery active — emails will send even if Outlook is closed.", "info");

  let totalSent   = 0;
  let totalFailed = 0;
  const totalCount = totalRecipientCount || recipients.length; // L4

  for (let i = 0; i < recipients.length; i++) {
    if (cancelRequested) {
      log("Merge stopped by user.", "warning");
      break;
    }

    const recipient = Object.assign({}, recipients[i], {
      record_num:   String(globalOffset + i + 1),  // A6: use global offset for correct record_num in mixed runs
      record_count: String(totalCount)
    });
    const personalizedSubject = personalize(subjectTemplate, recipient, false);  // subject = plain text
    const personalizedBody    = personalize(bodyTemplate, recipient, true);      // body = HTML
    const cc         = parseAddressList(recipient.cc  || "");
    const bcc        = parseAddressList(recipient.bcc || "");
    if (bccSelf && selfEmail) bcc.push({ emailAddress: { address: selfEmail } });
    const recipientAttachments = resolveAttachmentsForRecipient(recipient);
    const timestamp  = new Date().toISOString();

    // Build per-recipient custom headers (global + list-unsubscribe)
    const perRecipientHeaders = [
      ...customHeaders,
      ...(listUnsubscribeTemplate
        ? buildUnsubHeaders(personalize(listUnsubscribeTemplate, recipient, false))
        : [])
    ];

    let token;
    try {
      token = await getAccessToken();
    } catch (err) {
      log(`Auth error at message ${i + 1}: ${err.message}`, "error");
      for (let j = i; j < recipients.length; j++) {
        failedRecipients.push(recipients[j]);
        sendOutcomes.push({ rowNum: recipients[j]._csvRow || (globalOffset + j + 1), email: recipients[j].email, displayName: recipients[j].display_name || recipients[j].first_name || "", subjectUsed: "", status: "failed", timestamp, error: err.message });
      }
      totalFailed += recipients.length - i;
      break;
    }

    try {
      const toAddresses     = recipient.email.split(";").map(s => s.trim()).filter(s => EMAIL_REGEX.test(s));
      const displayName     = (recipient.display_name || [recipient.first_name, recipient.last_name].filter(Boolean).join(" ")).trim();
      const bodyContent     = plainText ? stripHtmlToText(personalizedBody) : personalizedBody;
      const bodyContentType = plainText ? "Text" : "HTML";

      const toRecipients = toAddresses.map((address, idx) => {
        if (idx === 0 && displayName) return { emailAddress: { address, name: displayName } };
        return { emailAddress: { address } };
      });

      // Per-recipient send_at override
      let deliveryTime = scheduledTimeISO;
      if (recipient.send_at) {
        const parsed = new Date(recipient.send_at);
        if (!isNaN(parsed.getTime())) {
          deliveryTime = parsed.toISOString();
        } else {
          log(`Row for ${recipient.email}: invalid send_at value "${recipient.send_at}" — using global schedule time.`, "warning");
        }
      }

      const message = {
        subject: personalizedSubject,
        body: { contentType: bodyContentType, content: bodyContent },
        toRecipients,
        singleValueExtendedProperties: [{
          id: "SystemTime 0x000F",
          value: deliveryTime
        }]
      };

      if (importance && importance !== "normal") message.importance = importance;
      if (sensitivity && sensitivity !== "normal") message.sensitivity = sensitivity;
      if (categories && categories.length > 0) message.categories = categories;
      if (isReadReceiptRequested)     message.isReadReceiptRequested     = true;
      if (isDeliveryReceiptRequested) message.isDeliveryReceiptRequested = true;
      const effectiveReplyTo = (recipient.reply_to || "").trim() || replyTo;
      if (effectiveReplyTo) message.replyTo = [{ emailAddress: { address: effectiveReplyTo } }];
      if (sendAs)  message.from    = { emailAddress: { address: sendAs } };
      if (cc.length  > 0) message.ccRecipients  = cc;
      if (bcc.length > 0) message.bccRecipients = bcc;
      if (perRecipientHeaders.length > 0) message.internetMessageHeaders = perRecipientHeaders;
      if (flagged) message.flag = { flagStatus: "flagged" };
      if (expiryISO) message.expiryDateTime = { dateTime: expiryISO, timeZone: "UTC" };

      const fileAttachments = recipientAttachments.map(a => ({
        "@odata.type": "#microsoft.graph.fileAttachment",
        name: a.name, contentType: a.contentType, contentBytes: a.contentBytes
      }));
      const inlineAttachments = inlineImagesArr.map(img => ({
        "@odata.type": "#microsoft.graph.fileAttachment",
        name: img.name, contentType: img.contentType, contentBytes: img.contentBytes,
        isInline: true, contentId: img.name
      }));
      const allAttachments = [...fileAttachments, ...inlineAttachments];
      if (allAttachments.length > 0) message.attachments = allAttachments;

      const createRes = await fetch("https://graph.microsoft.com/v1.0/me/messages", {
        method: "POST",
        headers: { "Authorization": `Bearer ${token}`, "Content-Type": "application/json" },
        body: JSON.stringify(message)
      });

      if (!createRes.ok) {
        if (createRes.status === 403) {
          const errBody = await createRes.json().catch(() => ({}));
          throw new Error(
            `403 Forbidden — Mail.ReadWrite permission required for scheduled send. ` +
            `Ask your IT admin to add Mail.ReadWrite to the Entra app registration. ` +
            `Graph error: ${JSON.stringify(errBody)}`
          );
        }
        const errBody = await createRes.json().catch(() => ({}));
        throw new Error(`Create draft failed: ${createRes.status} — ${JSON.stringify(errBody)}`);
      }

      const created = await createRes.json();

      const sendRes = await fetch(
        `https://graph.microsoft.com/v1.0/me/messages/${created.id}/send`,
        { method: "POST", headers: { "Authorization": `Bearer ${token}` } }
      );

      if (!sendRes.ok) {
        const errText = await sendRes.text();
        throw new Error(`Send draft failed: ${sendRes.status} — ${errText}`);
      }

      totalSent++;
      sendOutcomes.push({ rowNum: recipient._csvRow || (globalOffset + i + 1), email: recipient.email, displayName: (recipient.display_name || recipient.first_name || ""), subjectUsed: personalizedSubject.slice(0, 100), status: "scheduled", timestamp, error: "" });

      if (totalSent % 10 === 0 || i === recipients.length - 1) {
        log(`Scheduled ${totalSent}/${recipients.length}...`, "info");
      }
    } catch (err) {
      totalFailed++;
      failedRecipients.push(recipient);
      sendOutcomes.push({ rowNum: recipient._csvRow || (globalOffset + i + 1), email: recipient.email, displayName: (recipient.display_name || recipient.first_name || ""), subjectUsed: personalizedSubject.slice(0, 100), status: "failed", timestamp, error: err.message });
      log(`Failed for ${recipient.email}: ${err.message}`, "error");
      if (err.message.includes("403")) {
        for (let j = i + 1; j < recipients.length; j++) {
          failedRecipients.push(recipients[j]);
          sendOutcomes.push({ rowNum: recipients[j]._csvRow || (globalOffset + j + 1), email: recipients[j].email, displayName: (recipients[j].display_name || recipients[j].first_name || ""), subjectUsed: "", status: "failed", timestamp, error: "aborted after 403" });
        }
        totalFailed += recipients.length - i - 1;
        break;
      }
    }

    setProgress(i + 1, recipients.length);

    if (i < recipients.length - 1 && !cancelRequested) {
      await delay(200);
    }
  }

  return { totalSent, totalFailed };
}

/* ─── IMMEDIATE BATCH SEND HELPER ──────────────────────────────── */

async function sendImmediateBatch(
  recipients, subjectTemplate, emailBodyTemplate, saveToSent,
  replyTo, sendAs, importance, isReadReceiptRequested, isDeliveryReceiptRequested,
  globalCustomHeaders, plainTextMode, sensitivity, categories,
  bccSelf, selfEmail, flagged, expiryISO, listUnsubscribeTemplate, inlineImagesArr,
  allValid
) {
  let totalSent   = 0;
  let totalFailed = 0;

  const sharedBytes   = sharedAttachments.reduce((s, a) => s + a.sizeBytes, 0);
  const perRecipMax   = Math.max(...Array.from(perRecipientFiles.values()).map(f => f.sizeBytes), 0);
  const inlineBytes   = inlineImagesArr.reduce((s, img) => s + img.sizeBytes, 0);
  const maxAttachBytes = sharedBytes + perRecipMax + inlineBytes;
  const batchSize = effectiveBatchSize(maxAttachBytes);
  if (maxAttachBytes > 0 && batchSize < BATCH_SIZE) {
    log(`Attachment detected — batch size reduced to ${batchSize} emails/batch to stay within Graph API limits.`, "info");
  }

  const batchDelayMs = parseFloat(document.getElementById("batchDelayInput").value) * 1000;

  // Compute the global starting offset for record_num
  const indexOffset = allValid ? allValid.indexOf(recipients[0]) : 0;
  const totalCount = allValid ? allValid.length : recipients.length;

  const _batchSubjects = []; // parallel array for UX 6 report columns
  const requests = recipients.map((recipient, idx) => {
    const recipientWithMeta = Object.assign({}, recipient, {
      record_num:   String(indexOffset + idx + 1),
      record_count: String(totalCount)
    });
    const personalizedSubject = personalize(subjectTemplate, recipientWithMeta, false);
    _batchSubjects[idx] = personalizedSubject;
    const personalizedBody    = personalize(emailBodyTemplate, recipientWithMeta, true);
    const cc         = parseAddressList(recipient.cc  || "");
    const bcc        = parseAddressList(recipient.bcc || "");
    if (bccSelf && selfEmail) bcc.push({ emailAddress: { address: selfEmail } });
    const recipientAttachments = resolveAttachmentsForRecipient(recipient);
    const displayName = (recipient.display_name || [recipient.first_name, recipient.last_name].filter(Boolean).join(" ")).trim();
    const effectiveReplyTo = (recipient.reply_to || "").trim() || replyTo;
    const perRecipientHeaders = [
      ...globalCustomHeaders,
      ...(listUnsubscribeTemplate
        ? buildUnsubHeaders(personalize(listUnsubscribeTemplate, recipientWithMeta, false))
        : [])
    ];
    return buildEmailRequest(idx + 1, recipient.email, {
      subject: personalizedSubject, htmlBody: personalizedBody,
      saveToSent, attachments: recipientAttachments, inlineImages: inlineImagesArr,
      replyTo: effectiveReplyTo, sendAs, cc, bcc,
      importance, isReadReceiptRequested, isDeliveryReceiptRequested,
      customHeaders: perRecipientHeaders, plainText: plainTextMode,
      sensitivity, categories,
      displayName, flagged, expiryISO
    });
  });

  const batches      = chunkArray(requests, batchSize);
  const totalBatches = batches.length;
  const action       = draftsMode ? "drafts" : "emails";

  log(`${recipients.length} ${action} split into ${totalBatches} batch${totalBatches > 1 ? "es" : ""} of up to ${batchSize}.`, "info");

  let mergeStop = false;
  for (let i = 0; i < batches.length; i++) {
    if (cancelRequested || mergeStop) {
      log("Merge stopped by user.", "warning");
      break;
    }

    // Feature 2: Sending window check (per batch)
    const waitMs = msUntilWindowOpens();
    if (waitMs > 0) {
      const waitMins = Math.ceil(waitMs / 60000);
      log(`⏸ Outside sending window. Pausing ${waitMins} min until window opens…`, "warning");
      await new Promise(r => setTimeout(r, waitMs));
      log("▶ Resuming sends.", "info");
    }

    // Feature 3: Rate limit check (per batch)
    const canContinue = await checkRateLimits();
    if (!canContinue) {
      // L2: push all remaining unprocessed recipients (current batch onwards) to failedRecipients
      const ts = new Date().toISOString();
      for (let bi = i; bi < batches.length; bi++) {
        for (const req of batches[bi]) {
          const recipIdx = parseInt(req.id, 10) - 1;
          if (recipients[recipIdx]) {
            failedRecipients.push(recipients[recipIdx]);
            sendOutcomes.push({
              rowNum: recipients[recipIdx]._csvRow || (indexOffset + recipIdx + 1),
              email: recipients[recipIdx].email || "",
              displayName: recipients[recipIdx].display_name || recipients[recipIdx].first_name || "",
              subjectUsed: (_batchSubjects[recipIdx] || "").slice(0, 100),
              status: "failed",
              timestamp: ts,
              error: "Daily cap reached — not sent"
            });
          }
        }
      }
      mergeStop = true;
      break;
    }

    const batch = batches[i];
    log(`${draftsMode ? "Creating drafts" : "Sending"} batch ${i + 1}/${totalBatches} (${batch.length} ${action})...`, "info");

    let token;
    try {
      token = await getAccessToken();
    } catch (err) {
      log(`Authentication error on batch ${i + 1}: ${err.message}`, "error");
      const ts = new Date().toISOString();
      batch.forEach(req => {
        const idx = parseInt(req.id, 10) - 1;
        if (recipients[idx]) {
          failedRecipients.push(recipients[idx]);
          sendOutcomes.push({ rowNum: recipients[idx]._csvRow || (indexOffset + idx + 1), email: recipients[idx].email, displayName: recipients[idx].display_name || recipients[idx].first_name || "", subjectUsed: (_batchSubjects[idx] || "").slice(0, 100), status: "failed", timestamp: ts, error: err.message });
        }
      });
      totalFailed += batch.length;
      break;
    }

    try {
      const result    = await sendBatchWithRetry(batch, token);
      const responses = result.responses || [];
      const ts = new Date().toISOString();
      const batchSentEmails = [];
      let itemStatus401 = false;
      responses.forEach(r => {
        const recipIdx = parseInt(r.id, 10) - 1;
        if (r.status === 401) {
          // Feature 9: auth token expired mid-batch
          itemStatus401 = true;
          totalFailed++;
          if (recipients[recipIdx]) {
            failedRecipients.push(recipients[recipIdx]);
            sendOutcomes.push({ rowNum: recipients[recipIdx]._csvRow || (indexOffset + recipIdx + 1), email: recipients[recipIdx].email, displayName: recipients[recipIdx].display_name || recipients[recipIdx].first_name || "", subjectUsed: (_batchSubjects[recipIdx] || "").slice(0, 100), status: "failed", timestamp: ts, error: "401 Unauthorized" });
          }
        } else if (r.status >= 200 && r.status < 300) {
          totalSent++;
          emailsSentThisHour++;
          emailsSentToday++;
          if (recipients[recipIdx]) {
            sendOutcomes.push({ rowNum: recipients[recipIdx]._csvRow || (indexOffset + recipIdx + 1), email: recipients[recipIdx].email, displayName: recipients[recipIdx].display_name || recipients[recipIdx].first_name || "", subjectUsed: (_batchSubjects[recipIdx] || "").slice(0, 100), status: draftsMode ? "draft" : "sent", timestamp: ts, error: "" });
            if (!draftsMode) batchSentEmails.push(recipients[recipIdx].email);
          }
        } else {
          totalFailed++;
          if (recipients[recipIdx]) {
            failedRecipients.push(recipients[recipIdx]);
            sendOutcomes.push({ rowNum: recipients[recipIdx]._csvRow || (indexOffset + recipIdx + 1), email: recipients[recipIdx].email, displayName: recipients[recipIdx].display_name || recipients[recipIdx].first_name || "", subjectUsed: (_batchSubjects[recipIdx] || "").slice(0, 100), status: "failed", timestamp: ts, error: JSON.stringify(r.body) });
          }
          log(`Failed for request ID ${r.id}: ${r.status} — ${JSON.stringify(r.body)}`, "error");
        }
      });
      saveRateLimitState(); // P5: call once per batch rather than once per email
      if (itemStatus401) {
        log("⚠ Auth token expired during send. Remaining recipients saved for retry. Reopen the task pane to re-authenticate.", "warning");
        mergeStop = true;
      }
      if (batchSentEmails.length > 0) recordSentEmails(batchSentEmails);
      setProgress(totalSent + totalFailed, recipients.length);
      log(`Batch ${i + 1}/${totalBatches} complete. ${draftsMode ? "Drafted" : "Sent"}: ${totalSent} | Failed: ${totalFailed}`, "success");
    } catch (err) {
      log(`Batch ${i + 1} error: ${err.message}`, "error");
      const ts = new Date().toISOString();
      batch.forEach(req => {
        const failedIdx = parseInt(req.id, 10) - 1;
        if (recipients[failedIdx]) {
          failedRecipients.push(recipients[failedIdx]);
          sendOutcomes.push({ rowNum: recipients[failedIdx]._csvRow || (indexOffset + failedIdx + 1), email: recipients[failedIdx].email, displayName: recipients[failedIdx].display_name || recipients[failedIdx].first_name || "", subjectUsed: (_batchSubjects[failedIdx] || "").slice(0, 100), status: "failed", timestamp: ts, error: err.message });
        }
      });
      totalFailed += batch.length;
      setProgress(totalSent + totalFailed, recipients.length);
    }

    if (i < batches.length - 1 && !cancelRequested) {
      if (batchDelayMs > 0) {
        log(`Waiting ${batchDelayMs / 1000}s before next batch to avoid throttling...`, "info");
        await delay(batchDelayMs);
      }
      const jitter = getJitterMs();
      if (jitter > 0) await delay(jitter);
    }
  }

  return { totalSent, totalFailed };
}

/* ─── PRE-SEND CONFIRMATION MODAL (Feature C) ─────────────────── */

async function showPreSendConfirmation(recipients, batchDelayMs, scheduledCount) {
  const count = recipients.length;
  if (count <= 1) return true; // skip modal for single sends
  scheduledCount = scheduledCount || 0;
  const immediateCount = count - scheduledCount;
  // BUG 9: account for both per-email delay and per-batch API round-trip time
  const BATCH_SIZE_ESTIMATE = 20; // graph batch sends 20 at a time
  const batchCount = Math.ceil(count / BATCH_SIZE_ESTIMATE);
  const batchRoundTripMs = 2000; // ~2s per batch API call
  const totalMs = (count * batchDelayMs) + (batchCount * batchRoundTripMs);
  const estSeconds = totalMs / 1000;
  const estStr = estSeconds < 10 ? "< 10s"
    : estSeconds < 60 ? `~${Math.ceil(estSeconds)}s`
    : estSeconds < 3600 ? `~${Math.ceil(estSeconds / 60)} min`
    : `~${(estSeconds / 3600).toFixed(1)} hr`;
  const first = recipients[0];
  const firstName = (first.display_name || first.first_name || first.email || "").trim();

  // Feature 12: show breakdown of immediate vs scheduled
  let sendBreakdown = "";
  if (scheduledCount > 0 && immediateCount > 0) {
    sendBreakdown = '<br><span style="color:var(--text-muted,#605e5c);font-size:12px;">' + immediateCount + ' immediate &middot; ' + scheduledCount + ' scheduled (per send_at column)</span>';
  } else if (scheduledCount > 0) {
    sendBreakdown = '<br><span style="color:var(--text-muted,#605e5c);font-size:12px;">All ' + scheduledCount + ' scheduled per send_at column</span>';
  }

  // Build a personalized subject preview for the first recipient
  const subjectTemplate = document.getElementById("subjectInput")?.value || "";
  let firstSubjectPreview = "";
  if (subjectTemplate && typeof personalize === "function") {
    try { firstSubjectPreview = personalize(subjectTemplate, first); } catch(e) { firstSubjectPreview = subjectTemplate; }
  }

  document.getElementById("preSendSummary").innerHTML =
    `<div style="background:var(--surface,#f3f2f1);border-radius:10px;padding:10px 12px;margin-bottom:10px;font-size:13px;line-height:1.7;">
      🚀 You're about to send <strong>${count}</strong> personalised email${count !== 1 ? "s" : ""}.${sendBreakdown ? "<br>" + sendBreakdown.replace(/<br>/, "") : ""}
    </div>
    <div style="font-size:12px;color:var(--text-muted,#605e5c);margin-bottom:6px;">First email preview:</div>
    <div style="background:var(--surface,#f3f2f1);border-radius:10px;padding:10px 12px;font-size:12px;line-height:1.6;">
      <strong>To:</strong> ${escapeHtml(firstName)} &lt;${escapeHtml(first.email || "")}&gt;<br>
      ${firstSubjectPreview ? "<strong>Subject:</strong> " + escapeHtml(firstSubjectPreview) + "<br>" : ""}
      <strong>Est. time:</strong> ${estStr}
    </div>`;
  return new Promise(function(resolve) {
    const prevFocus = document.activeElement;
    const modal = document.getElementById("preSendModal");
    modal.classList.remove("hidden");
    const releaseTrap = trapFocus(modal);
    document.getElementById("preSendConfirmBtn").onclick = function() {
      releaseTrap();
      if (prevFocus && typeof prevFocus.focus === "function") prevFocus.focus();
      modal.classList.add("hidden");
      resolve(true);
    };
    document.getElementById("preSendCancelBtn").onclick = function() {
      releaseTrap();
      if (prevFocus && typeof prevFocus.focus === "function") prevFocus.focus();
      modal.classList.add("hidden");
      resolve(false);
    };
  });
}

/* ─── MODAL FOCUS TRAP (Feature 15) ────────────────────────────── */

function trapFocus(modal) {
  const focusable = modal.querySelectorAll(
    'button:not([disabled]), input:not([disabled]), select:not([disabled]), textarea:not([disabled]), [tabindex="0"]'
  );
  if (!focusable.length) return function() {};
  const first = focusable[0];
  const last  = focusable[focusable.length - 1];
  first.focus();

  function handler(e) {
    if (e.key !== "Tab") return;
    if (e.shiftKey) {
      if (document.activeElement === first) { e.preventDefault(); last.focus(); }
    } else {
      if (document.activeElement === last) { e.preventDefault(); first.focus(); }
    }
  }
  modal.addEventListener("keydown", handler);
  return function() { modal.removeEventListener("keydown", handler); };
}

// Map of modalId → { prevFocus, releaseTrap }
const _modalTrapState = new Map();

function _openModalWithTrap(modalId) {
  const modal = document.getElementById(modalId);
  if (!modal) return;
  const prevFocus = document.activeElement;
  modal.classList.remove("hidden");
  const releaseTrap = trapFocus(modal);
  _modalTrapState.set(modalId, { prevFocus, releaseTrap });
}

function _closeModalWithTrap(modalId) {
  const modal = document.getElementById(modalId);
  if (modal) modal.classList.add("hidden");
  const state = _modalTrapState.get(modalId);
  if (state) {
    state.releaseTrap();
    if (state.prevFocus && typeof state.prevFocus.focus === "function") state.prevFocus.focus();
    _modalTrapState.delete(modalId);
  }
}

/* ─── SIMPLE CONFIRM MODAL HELPER ─────────────────────────────── */

function showSimpleConfirm(message) {
  document.getElementById("confirmText").textContent = message;
  const sendBtn = document.getElementById("confirmSendBtn");
  const savedText = sendBtn.textContent;
  sendBtn.textContent = "Yes";
  _openModalWithTrap("confirmModal");
  return new Promise(function(resolve) {
    _confirmModalPrevFocusSaved = true;
    pendingMergeResolve = function(result) {
      sendBtn.textContent = savedText;
      resolve(result);
    };
  });
}
let _confirmModalPrevFocusSaved = false;

/* ─── CONFIRMATION MODAL ───────────────────────────────────────── */

let pendingMergeResolve = null;

function showConfirmModal(rowCount, addressCount, scheduledTimeISO, broadcastMode) {
  return new Promise((resolve) => {
    pendingMergeResolve = resolve;

    const action = draftsMode ? "Save as drafts" : "Send";

    let text;
    if (broadcastMode) {
      text = `Send 1 broadcast email to ${rowCount} recipient${rowCount !== 1 ? "s" : ""} via BCC?`;
    } else {
      // When multi-TO rows are present, addressCount > rowCount — show both figures.
      const emailWord = rowCount !== 1 ? "emails" : "email";
      const countStr  = addressCount > rowCount
        ? `${rowCount} ${emailWord} (reaching ${addressCount} addresses)`
        : `${rowCount} ${emailWord}`;

      if (draftsMode) {
        text = `This will save ${countStr} as drafts in your Drafts folder.`;
      } else if (scheduledTimeISO) {
        text = `This will schedule ${countStr} to be sent at ${new Date(scheduledTimeISO).toLocaleString()}.`;
      } else {
        text = `This will send ${countStr}.`;
      }

      if (sharedAttachments.length > 0) {
        text += ` Each email will include ${sharedAttachments.length} attachment(s).`;
      }
    }
    text += " Continue?";

    document.getElementById("confirmText").textContent = text;
    document.getElementById("confirmSendBtn").textContent = broadcastMode ? "Broadcast" : action;
    _openModalWithTrap("confirmModal"); // Feature 15: focus trap
  });
}

function confirmSend() {
  const resolve = pendingMergeResolve;
  pendingMergeResolve = null;
  _closeModalWithTrap("confirmModal"); // Feature 15: release trap + restore focus
  if (resolve) resolve(true);
}

function dismissModal() {
  const resolve = pendingMergeResolve;
  pendingMergeResolve = null;
  _closeModalWithTrap("confirmModal"); // Feature 15: release trap + restore focus
  if (resolve) resolve(false);
}

/* ─── CANCEL ───────────────────────────────────────────────────── */

function handleStop() {
  cancelRequested = true;
  log("Stop requested — will halt after current batch.", "warning");
  document.getElementById("stopBtn").disabled = true;
}

function setMergeRunning(running) {
  const mergeBtn     = document.getElementById("mergeBtn");
  const stopBtn      = document.getElementById("stopBtn");
  const testBtn      = document.getElementById("testSendBtn");
  const draftsBtn    = document.getElementById("saveDraftsBtn");
  const broadcastBtn = document.getElementById("broadcastBtn");
  mergeBtn.disabled     = running;
  testBtn.disabled      = running;   // prevent concurrent test sends during active merge
  draftsBtn.disabled    = running;   // prevent concurrent drafts run
  broadcastBtn.disabled = running;   // prevent concurrent broadcast
  mergeBtn.textContent = running ? "⏳ Sending..." : "▶ Run mail merge";
  stopBtn.disabled = !running;
  stopBtn.classList.toggle("hidden", !running);
  if (running) {
    // Feature 5: reset completion flag and progress bar at the start of a new merge
    _mergeCompletedSuccessfully = false;
    const fill = document.getElementById("progressFill");
    if (fill) fill.style.width = "0%";
  } else if (!_mergeCompletedSuccessfully) {
    // Feature 5: only hide progress if we did NOT just call showMergeComplete
    hideProgress();
  }
}

/* ─── MAIN MERGE RUNNER ────────────────────────────────────────── */

/**
 * Orchestrate the full mail merge send run.
 *
 * High-level flow:
 *   1. Guard against re-entrant clicks (mergeInProgress flag)
 *   2. Parse and validate the recipient list
 *   3. Acquire an OAuth2 token via Office.js SSO
 *   4. Show a pre-send confirmation modal with a human-readable summary
 *   5. Dispatch emails in Graph $batch chunks of up to BATCH_SIZE (20)
 *   6. After each batch, apply the user-configured batch delay
 *   7. Respect the hourly rate limit and daily cap if set
 *   8. Show the retry button + download report on completion
 *
 * All async waits use await so the UI stays responsive throughout.
 * The cancelRequested flag (set by handleStop()) is checked between
 * batches so the user can abort without killing the current batch.
 *
 * @async
 * @returns {Promise<void>}
 */
async function handleMergeClick() {
  // BUG 1: double-send guard — prevents concurrent merges from multiple clicks during async modals
  if (mergeInProgress) {
    log("A merge is already in progress. Please wait or stop the current merge.", "warning");
    return;
  }
  mergeInProgress = true;
  document.getElementById("mergeBtn").disabled = true;
  try {
  // 1. Ensure recipients are parsed
  if (parsedRecipients.length === 0) {
    parseAndPreview();
    if (parsedRecipients.length === 0) {
      log("No valid recipients found. Check your CSV format.", "error");
      return;
    }
  }

  // 2. Validate subject
  const subjectTemplate = document.getElementById("subjectInput").value.trim();
  if (!subjectTemplate) {
    log("Subject line is empty.", "error");
    return;
  }

  // 3. Read optional send-option inputs
  const replyTo               = document.getElementById("replyToInput").value.trim();
  const sendAs                = document.getElementById("sendAsInput").value.trim();
  const deduplicateEnabled    = document.getElementById("deduplicateEnabled").checked;
  const scheduleEnabled       = document.getElementById("scheduleEnabled").checked;
  const importance            = document.getElementById("importanceSelect").value;
  const sensitivity           = document.getElementById("sensitivitySelect").value;
  const categoriesRaw         = document.getElementById("categoriesInput").value.trim();
  const categories            = categoriesRaw ? categoriesRaw.split(",").map(c => c.trim()).filter(Boolean) : [];
  const isReadReceiptRequested     = document.getElementById("requestReadReceipt").checked;
  const isDeliveryReceiptRequested = document.getElementById("requestDeliveryReceipt").checked;
  const plainTextMode         = document.getElementById("plainTextMode").checked;
  const customHeadersEnabled  = document.getElementById("customHeadersEnabled").checked;
  const globalCustomHeaders   = customHeadersEnabled
    ? parseCustomHeaders(document.getElementById("customHeadersInput").value)
    : [];
  const batchDelayMs = parseFloat(document.getElementById("batchDelayInput").value) * 1000;
  const bccSelf      = document.getElementById("bccSelfEnabled").checked;
  const flagged      = document.getElementById("flagForFollowup").checked;
  const listUnsubscribeTemplate = document.getElementById("listUnsubscribeInput").value.trim();
  let scheduledTimeISO = null;
  let expiryISO        = null;

  if (scheduleEnabled) {
    const scheduledTimeValue = document.getElementById("scheduledTime").value;
    if (!scheduledTimeValue) {
      log("Schedule send is enabled but no date/time is set.", "error");
      return;
    }
    scheduledTimeISO = new Date(scheduledTimeValue).toISOString();
    if (new Date(scheduledTimeISO) <= new Date()) {
      log("Scheduled time must be in the future.", "error");
      return;
    }
  }

  if (document.getElementById("expiryEnabled").checked) {
    const expiryValue = document.getElementById("expiryDateTimeInput").value;
    if (!expiryValue) {
      log("Message expiry is enabled but no date/time is set.", "error");
      return;
    }
    expiryISO = new Date(expiryValue).toISOString();
    if (new Date(expiryISO) <= new Date()) {
      log("Expiry date must be in the future.", "error");
      return;
    }
  }

  let selfEmail = null;
  if (bccSelf) {
    selfEmail = Office.context.mailbox.userProfile.emailAddress;
  }

  // 4. Validate email addresses (use filtered/sorted list if a filter is active)
  const { valid: validRaw, invalid } = validateRecipients(getFilteredSortedRecipients());
  if (invalid.length > 0) {
    invalid.forEach(({ row, email }) => {
      log(`Row ${row}: invalid email address "${email}" — skipped.`, "warning");
    });
  }
  if (validRaw.length === 0) {
    log("No valid email addresses found.", "error");
    return;
  }

  // 5. Deduplicate if requested
  let valid = validRaw;
  if (deduplicateEnabled) {
    const seen   = new Set();
    const before = valid.length;
    valid = valid.filter(r => {
      const primaryEmail = r.email.split(";")[0].trim().toLowerCase();
      if (seen.has(primaryEmail)) return false;
      seen.add(primaryEmail);
      return true;
    });
    const removed = before - valid.length;
    if (removed > 0) {
      log(`Deduplication: removed ${removed} duplicate row${removed !== 1 ? "s" : ""}. ` +
          `Sending to ${valid.length} unique address${valid.length !== 1 ? "es" : ""}.`, "info");
    }
  }

  // 6. Apply suppression list
  if (suppressionSet.size > 0) {
    const before = valid.length;
    valid = valid.filter(r => {
      const primaryEmail = r.email.split(";")[0].trim().toLowerCase();
      return !suppressionSet.has(primaryEmail);
    });
    const suppressed = before - valid.length;
    if (suppressed > 0) {
      log(`Suppression: blocked ${suppressed} address${suppressed !== 1 ? "es" : ""}. ` +
          `Sending to ${valid.length} remaining.`, "info");
    }
    if (valid.length === 0) {
      log("All recipients were suppressed. Nothing to send.", "error");
      return;
    }
  }

  // 6b. Group rows by email if requested (many-to-one merge)
  const groupByEmailEl = document.getElementById("groupByEmail");
  if (groupByEmailEl && groupByEmailEl.checked) {
    valid = groupRecipientsByEmail(valid);
    log(`Grouped by email: ${valid.length} unique recipient(s).`, "info");
  }

  // 6c. Apply persistent opt-out list
  const optOutSet = getOptOutList();
  if (optOutSet.size > 0) {
    const beforeOptOut = valid.length;
    valid = valid.filter(r => {
      const addrs = parseAddressList(r.email || "").map(a => a.emailAddress.address.toLowerCase());
      return !addrs.some(a => optOutSet.has(a));
    });
    if (valid.length < beforeOptOut) {
      log(`Skipped ${beforeOptOut - valid.length} opt-out recipient(s).`, "info");
    }
    if (valid.length === 0) {
      log("All recipients are on the opt-out list. Nothing to send.", "error");
      // BUG 13: do NOT call setMergeRunning(false) here — it was never set to true at this point.
      // The mergeInProgress finally block (BUG 1) handles re-enabling the button.
      return;
    }
  }

  // 7. Duplicate send guard check
  const recentDupes = checkDuplicateSendHistory(valid);
  if (recentDupes.length) {
    const proceed = await showDuplicateWarningModal(recentDupes);
    if (!proceed) return;
  }

  // 8. Show confirmation modal — for multi-TO rows count unique addresses reached, not just row count
  const totalAddresses = valid.reduce((sum, r) => {
    return sum + r.email.split(";").map(s => s.trim()).filter(s => EMAIL_REGEX.test(s)).length;
  }, 0);
  const confirmed = await showConfirmModal(valid.length, totalAddresses, scheduledTimeISO);
  if (!confirmed) {
    log("Merge cancelled.", "info");
    return;
  }

  // 9. Read email body template from compose window
  let emailBodyTemplate = "";
  try {
    emailBodyTemplate = await new Promise((resolve, reject) => {
      Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
        else reject(new Error("Could not read email body."));
      });
    });
  } catch (err) {
    log(`Body read error: ${err.message}`, "error");
    return;
  }

  // 9b. Check for fill-in tokens and prompt user
  const combinedTemplate = subjectTemplate + " " + emailBodyTemplate;
  if (/\{\{fill_in:/i.test(combinedTemplate)) {
    const fillInValues = await collectFillInValues(combinedTemplate);
    if (fillInValues === null) {
      log("Merge cancelled (fill-in prompt dismissed).", "warning");
      return;
    }
    emailBodyTemplate = applyFillInValues(emailBodyTemplate, fillInValues);
    window._fillInValues = fillInValues;
  } else {
    window._fillInValues = null;
  }

  // 10. Payload size pre-flight
  checkPayloadSize(emailBodyTemplate);

  // 10b. DNS pre-flight check for large sends (Feature 7)
  await runDnsPreflightCheck(valid.length);

  // 10c. Pre-send summary confirmation (Feature C) — only for sends > 1 recipient
  const batchDelayForConfirm = parseFloat(document.getElementById("batchDelayInput").value) * 1000;
  const scheduledCount = valid.filter(r => r.send_at && r.send_at.trim()).length; // Feature 12
  const preSendOk = await showPreSendConfirmation(valid, batchDelayForConfirm, scheduledCount);
  if (!preSendOk) {
    log("Merge cancelled.", "info");
    return;
  }

  // 11. Check for unresolved tokens using first recipient as sample
  const sampleSubject = personalize(subjectTemplate, valid[0], false);
  const sampleBody    = personalize(emailBodyTemplate, valid[0], true);
  const allUnresolved = [...new Set([
    ...findUnresolvedTokens(sampleSubject),
    ...findUnresolvedTokens(sampleBody)
  ])];
  if (allUnresolved.length > 0) {
    log(`Warning: unresolved token(s) — these will be sent literally: ${allUnresolved.join(", ")}`, "warning");
  }

  failedRecipients = [];
  // BUG 14: sendOutcomes is NOT reset here — it accumulates across retry runs for the same CSV.
  // It is reset in parseAndPreview() when a new CSV is loaded.
  warnedMissingAttachments.clear();
  document.getElementById("retryFailedBtn").classList.add("hidden");
  document.getElementById("downloadReportBtn").classList.add("hidden");
  cancelRequested = false;   // reset BEFORE enabling Stop
  // BUG 3: wrap send block so setMergeRunning(false) is called even on unexpected throws
  setMergeRunning(true);
  try {
  if (draftsMode) {
    log("Starting mail merge in DRAFTS mode — emails will be saved to Drafts folder, not sent.", "info");
  } else {
    log("Starting mail merge...", "info");
  }
  if (sendAs)                         log(`Send As: ${sendAs}`, "info");
  if (replyTo)                        log(`Reply-To: ${replyTo}`, "info");
  if (importance !== "normal")        log(`Priority: ${importance}`, "info");
  if (sensitivity !== "normal")       log(`Sensitivity: ${sensitivity}`, "info");
  if (isReadReceiptRequested)         log("Read receipt requested.", "info");
  if (isDeliveryReceiptRequested)     log("Delivery receipt requested.", "info");
  if (plainTextMode)                  log("Plain text mode — HTML will be stripped from body.", "info");
  if (globalCustomHeaders.length > 0) log(`Custom headers: ${globalCustomHeaders.map(h => h.name).join(", ")}`, "info");
  if (perRecipientFiles.size > 0)     log(`Per-recipient attachments: ${perRecipientFiles.size} file(s) loaded.`, "info");
  if (bccSelf)                        log(`BCC self enabled: ${selfEmail}`, "info");
  if (flagged)                        log("Follow-up flag: sent items will be flagged.", "info");
  if (expiryISO)                      log(`Message expiry set: ${new Date(expiryISO).toLocaleString()}`, "info");
  if (listUnsubscribeTemplate)        log(`List-Unsubscribe template: ${listUnsubscribeTemplate}`, "info");
  if (sharedAttachments.length > 0)   log(`Shared attachments: ${sharedAttachments.length} file(s) loaded.`, "info");
  if (inlineImages.size > 0)          log(`Inline images: ${inlineImages.size} image(s) loaded.`, "info");

  const saveToSent = document.getElementById("saveToSentItems").checked;

  let totalSent   = 0;
  let totalFailed = 0;

  const inlineImagesArr = Array.from(inlineImages.values());

  // Count recipients with per-row send_at
  const recipientsWithSendAt = valid.filter(r => r.send_at && r.send_at.trim());
  if (recipientsWithSendAt.length > 0) {
    log(`${recipientsWithSendAt.length} recipient(s) have a per-row send_at value — those rows will use individual scheduled delivery.`, "info");
  }

  if (scheduledTimeISO || recipientsWithSendAt.length > 0) {
    // ── SCHEDULED SEND ────────────────────────────────────────────────────
    // Split: rows with send_at go through scheduled path; if global schedule is on,
    // rows without send_at also go through scheduled path using global time.
    const scheduledRecipients = scheduledTimeISO
      ? valid  // all recipients use scheduled path
      : recipientsWithSendAt; // only rows with send_at
    const immediateRecipients = scheduledTimeISO
      ? []
      : valid.filter(r => !(r.send_at && r.send_at.trim()));

    if (scheduledRecipients.length > 0) {
      // A6: compute offset of first scheduled recipient within the full valid array
      const firstScheduledIdx = valid.indexOf(scheduledRecipients[0]);
      ({ totalSent, totalFailed } = await sendScheduledMessages(
        scheduledRecipients, subjectTemplate, emailBodyTemplate, saveToSent,
        replyTo, sendAs, scheduledTimeISO,
        importance, isReadReceiptRequested, isDeliveryReceiptRequested,
        globalCustomHeaders, plainTextMode,
        sensitivity, categories,
        bccSelf, selfEmail, flagged, expiryISO, listUnsubscribeTemplate, inlineImagesArr,
        firstScheduledIdx, valid.length  // L4: pass total valid count for correct record_count
      ));
    }

    // Process immediate recipients (those without send_at and no global schedule)
    if (immediateRecipients.length > 0) {
      log(`${immediateRecipients.length} recipient(s) will be sent immediately (no send_at).`, "info");
      // Fall through to batch send for these recipients — handled below via a temporary override
      const immResults = await sendImmediateBatch(
        immediateRecipients, subjectTemplate, emailBodyTemplate, saveToSent,
        replyTo, sendAs, importance, isReadReceiptRequested, isDeliveryReceiptRequested,
        globalCustomHeaders, plainTextMode, sensitivity, categories,
        bccSelf, selfEmail, flagged, expiryISO, listUnsubscribeTemplate, inlineImagesArr,
        valid
      );
      totalSent   += immResults.totalSent;
      totalFailed += immResults.totalFailed;
    }

  } else if (draftsMode) {
    // ── DRAFTS MODE ($batch with /me/messages) ─────────────────────────────
    const immResults = await sendImmediateBatch(
      valid, subjectTemplate, emailBodyTemplate, saveToSent,
      replyTo, sendAs, importance, isReadReceiptRequested, isDeliveryReceiptRequested,
      globalCustomHeaders, plainTextMode, sensitivity, categories,
      bccSelf, selfEmail, flagged, expiryISO, listUnsubscribeTemplate, inlineImagesArr,
      valid
    );
    totalSent   = immResults.totalSent;
    totalFailed = immResults.totalFailed;
    if (!cancelRequested) {
      log(`${totalSent} draft(s) created in your Drafts folder. Open Outlook to review before sending.`, "success");
    }

  } else {
    // ── IMMEDIATE SEND ($batch) ────────────────────────────────────────────
    const immResults = await sendImmediateBatch(
      valid, subjectTemplate, emailBodyTemplate, saveToSent,
      replyTo, sendAs, importance, isReadReceiptRequested, isDeliveryReceiptRequested,
      globalCustomHeaders, plainTextMode, sensitivity, categories,
      bccSelf, selfEmail, flagged, expiryISO, listUnsubscribeTemplate, inlineImagesArr,
      valid
    );
    totalSent   = immResults.totalSent;
    totalFailed = immResults.totalFailed;
  }

  log(`─── Merge complete. ✅ Sent: ${totalSent} | ❌ Failed: ${totalFailed} ───`,
      totalFailed > 0 ? "warning" : "success");
  // Bug 2: only show "Done" banner if the merge completed without a stop request
  if (!cancelRequested) {
    showMergeComplete(totalSent, totalSent + totalFailed);
  } else {
    // Stopped early — just hide progress bar normally
    const _stopContainer = document.getElementById("progressContainer");
    if (_stopContainer) _stopContainer.classList.add("hidden");
  }
  if (/\{\{unsubscribe_link\}\}/i.test(emailBodyTemplate)) {
    log("ℹ {{unsubscribe_link}} was used. Monitor your inbox for 'UNSUBSCRIBE' replies and use the Opt-Out List to record them.", "info");
  }

  if (failedRecipients.length > 0) {
    const retryBtn = document.getElementById("retryFailedBtn");
    retryBtn.textContent = `↺ Retry failed (${failedRecipients.length})`;
    retryBtn.classList.remove("hidden");
  }

  if (sendOutcomes.length > 0) {
    document.getElementById("downloadReportBtn").classList.remove("hidden");
  }

  if (!cancelRequested && totalFailed === 0) {
    lsRemove(LS_KEY_CSV);
    document.getElementById("csvInput").value = "";
    parsedRecipients = [];
    document.getElementById("recipientCount").textContent = "";
    document.getElementById("previewTable").classList.add("hidden");
    log("CSV cleared from local storage after successful send.", "info");
  }

    setMergeRunning(false); // normal-flow explicit call (BUG 3)
  } catch (err) {
    log(`Unexpected error during merge: ${err.message}`, "error");
  } finally {
    setMergeRunning(false); // BUG 3 safety net — no-op if already called above
  }
  } finally {
    // BUG 1: always release the double-send guard and re-enable the button
    mergeInProgress = false;
    document.getElementById("mergeBtn").disabled = false;
  }
}

/* ─── CSV FIELD ESCAPING HELPER ────────────────────────────────── */

function csvField(val) {
  const s = String(val !== null && val !== undefined ? val : "");
  if (s.includes(",") || s.includes('"') || s.includes("\n")) {
    return '"' + s.replace(/"/g, '""') + '"';
  }
  return s;
}

/* ─── GROUPED ROW MERGE (Feature 4) ───────────────────────────── */

function groupRecipientsByEmail(recipients) {
  const map = new Map();
  for (const r of recipients) {
    const key = (r.email || "").toLowerCase().trim();
    if (!key) {
      // BUG 12: log skipped rows with whitespace-only email instead of silently dropping
      log(`Row skipped in group merge: empty email address (row data: ${JSON.stringify(r).slice(0, 80)}).`, "warning");
      continue;
    }
    if (!map.has(key)) map.set(key, { primary: r, rows: [r] });
    else map.get(key).rows.push(r);
  }
  return Array.from(map.values()).map(({ primary, rows }) => ({ ...primary, _groupedRows: rows }));
}

function buildMergeTable(rows, headers) {
  const cols = headers.filter(h => !["email","cc","bcc","attachment","skip_if","send_at","reply_to","display_name"].includes(h.toLowerCase()));
  let html = '<table border="1" cellpadding="4" cellspacing="0" style="border-collapse:collapse;width:100%"><thead><tr>';
  for (const col of cols) html += `<th style="background:#f0f0f0;padding:4px 8px">${escapeHtml(col)}</th>`;
  html += '</tr></thead><tbody>';
  for (const row of rows) {
    html += '<tr>';
    for (const col of cols) html += `<td style="padding:4px 8px">${escapeHtml(row[col] || '')}</td>`;
    html += '</tr>';
  }
  html += '</tbody></table>';
  return html;
}

/* ─── OPT-OUT / UNSUBSCRIBE LIST (Feature 5) ──────────────────── */

const LS_KEY_OPTOUT = "mm_optout_list";

function getOptOutList() {
  return new Set((lsGet(LS_KEY_OPTOUT, []) || []).map(e => e.toLowerCase().trim()));
}

function saveOptOutList(set) {
  lsSet(LS_KEY_OPTOUT, JSON.stringify(Array.from(set)));
  renderOptOutList();
}

function renderOptOutList() {
  // BUG 11: use data-optout-idx and event delegation instead of inline onclick to prevent XSS
  const list = lsGet(LS_KEY_OPTOUT, []) || [];
  const el = document.getElementById("optoutList");
  const countEl = document.getElementById("optoutCount");
  if (!el) return;
  el.innerHTML = list.map((email, idx) =>
    `<div style="display:flex;justify-content:space-between;align-items:center;padding:2px 4px;">
      <span>${escapeHtml(email)}</span>
      <button data-optout-idx="${idx}" style="font-size:10px;padding:1px 6px;cursor:pointer;border:1px solid var(--border);border-radius:3px;background:var(--bg-secondary)">&#10005;</button>
    </div>`
  ).join("") || '<div style="color:var(--text-muted);padding:4px;">No addresses</div>';

  // Attach listeners after innerHTML set
  el.querySelectorAll("[data-optout-idx]").forEach(btn => {
    btn.addEventListener("click", () => {
      const idx = parseInt(btn.dataset.optoutIdx, 10);
      const currentList = lsGet(LS_KEY_OPTOUT, []) || [];
      currentList.splice(idx, 1);
      lsSet(LS_KEY_OPTOUT, JSON.stringify(currentList));
      renderOptOutList();
    });
  });

  if (countEl) countEl.textContent = `${list.length} address${list.length === 1 ? "" : "es"}`;
}

function removeOptOut(email) {
  const list = new Set((lsGet(LS_KEY_OPTOUT, []) || []).map(e => e.toLowerCase().trim()));
  list.delete(email.toLowerCase().trim());
  saveOptOutList(list);
}

function initOptOutUI() {
  const optoutAddBtn = document.getElementById("optoutAddBtn");
  if (optoutAddBtn) {
    optoutAddBtn.addEventListener("click", function() {
      const input = document.getElementById("optoutAddInput");
      const email = (input.value || "").trim().toLowerCase();
      if (!email || !EMAIL_REGEX.test(email)) { log("Enter a valid email to add to opt-out list.", "warning"); return; }
      const list = getOptOutList();
      list.add(email);
      saveOptOutList(list);
      input.value = "";
      log(`Added ${email} to opt-out list.`, "info");
    });
  }

  const optoutClearBtn = document.getElementById("optoutClearBtn");
  if (optoutClearBtn) {
    optoutClearBtn.addEventListener("click", async function() {
      const confirmed = await showSimpleConfirm("Clear the entire opt-out list? This cannot be undone.");
      if (!confirmed) return;
      saveOptOutList(new Set());
      log("Opt-out list cleared.", "info");
    });
  }

  const optoutImportBtn = document.getElementById("optoutImportBtn");
  if (optoutImportBtn) {
    optoutImportBtn.addEventListener("click", function() {
      if (!parsedRecipients.length) { log("Load a CSV first.", "warning"); return; }
      const list = getOptOutList();
      let added = 0;
      for (const r of parsedRecipients) {
        const addrs = parseAddressList(r.email || "").map(function(a) { return a.emailAddress.address.toLowerCase(); });
        for (const a of addrs) {
          if (!list.has(a)) { list.add(a); added++; }
        }
      }
      saveOptOutList(list);
      log(`Imported ${added} address${added === 1 ? "" : "es"} to opt-out list.`, "success");
    });
  }

  renderOptOutList();
}

/* ─── FILL-IN PROMPT (Feature 6) ──────────────────────────────── */

async function collectFillInValues(template) {
  const regex = /\{\{fill_in:([^}]+)\}\}/gi;
  const prompts = new Map();
  let m;
  while ((m = regex.exec(template)) !== null) {
    prompts.set(m[1].trim(), "");
  }
  if (prompts.size === 0) return null;

  const fieldsEl = document.getElementById("fillInFields");
  fieldsEl.innerHTML = "";
  prompts.forEach((_, prompt) => {
    const id = "fillin_" + prompt.replace(/\W+/g, "_");
    const div = document.createElement("div");
    div.style.cssText = "margin-bottom:10px;";
    div.innerHTML = `<label style="font-size:12px;font-weight:600;display:block;margin-bottom:3px;">${escapeHtml(prompt)}</label>
      <input type="text" id="${id}" placeholder="Enter ${escapeHtml(prompt)}…" style="width:100%;box-sizing:border-box;padding:6px 8px;border:1px solid var(--border);border-radius:4px;font-size:13px;">`;
    fieldsEl.appendChild(div);
  });

  return new Promise((resolve) => {
    _openModalWithTrap("fillInModal"); // Feature 15: focus trap

    document.getElementById("fillInOkBtn").onclick = () => {
      const values = new Map();
      prompts.forEach((_, prompt) => {
        const id = "fillin_" + prompt.replace(/\W+/g, "_");
        const fillInEl = document.getElementById(id); values.set(prompt, fillInEl ? fillInEl.value : "");
      });
      _closeModalWithTrap("fillInModal"); // Feature 15
      resolve(values);
    };

    document.getElementById("fillInCancelBtn").onclick = () => {
      _closeModalWithTrap("fillInModal"); // Feature 15
      resolve(null);
    };
  });
}

function applyFillInValues(template, fillInValues) {
  if (!fillInValues) return template;
  fillInValues.forEach((value, prompt) => {
    // S4: neutralise token delimiters in user input to prevent secondary substitution
    const safeValue = value.replace(/\{\{/g, "{ {").replace(/\}\}/g, "} }");
    const escapedPrompt = prompt.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    // BUG 6: use function replacer to prevent JS interpreting $&, $', $1 etc. in value
    template = template.replace(
      new RegExp(`\\{\\{fill_in:${escapedPrompt}\\}\\}`, "gi"),
      () => safeValue
    );
  });
  return template;
}

/* ─── DNS PRE-FLIGHT CHECK (Feature 7) ────────────────────────── */

async function checkDomainDns(domain) {
  const results = { spf: false, dmarc: false, dkim: false };
  try {
    const spfResp = await fetch(`https://cloudflare-dns.com/dns-query?name=${encodeURIComponent(domain)}&type=TXT`, {
      headers: { "Accept": "application/dns-json" }
    });
    if (spfResp.ok) {
      const data = await spfResp.json();
      results.spf = (data.Answer || []).some(r => (r.data || "").includes("v=spf1"));
    }
  } catch(e) { /* network error, skip */ }

  try {
    const dmarcResp = await fetch(`https://cloudflare-dns.com/dns-query?name=${encodeURIComponent("_dmarc." + domain)}&type=TXT`, {
      headers: { "Accept": "application/dns-json" }
    });
    if (dmarcResp.ok) {
      const data = await dmarcResp.json();
      results.dmarc = (data.Answer || []).some(r => (r.data || "").includes("v=DMARC1"));
    }
  } catch(e) { /* network error, skip */ }

  // BUG 10: check both selector1 and selector2 — non-M365 domains may use selector2 only
  async function checkDkimSelector(selectorDomain) {
    try {
      const res = await fetch(`https://cloudflare-dns.com/dns-query?name=${encodeURIComponent(selectorDomain)}&type=CNAME`, {
        headers: { Accept: "application/dns-json" }
      });
      if (!res.ok) return false;
      const data = await res.json();
      return (data.Answer || []).length > 0;
    } catch { return false; }
  }
  const dkim1 = await checkDkimSelector("selector1._domainkey." + domain);
  const dkim2 = await checkDkimSelector("selector2._domainkey." + domain);
  results.dkim = dkim1 || dkim2;

  return results;
}

async function runDnsPreflightCheck(recipientCount) {
  if (recipientCount < 500) return;
  const senderEmail = (Office.context.mailbox && Office.context.mailbox.userProfile && Office.context.mailbox.userProfile.emailAddress) || "";
  const domain = senderEmail.split("@")[1];
  if (!domain) return;

  // S5: disclose the third-party DNS query so users are aware of the data flow
  log(`ℹ DNS pre-flight: querying Cloudflare DNS (cloudflare-dns.com) for ${domain} SPF/DKIM/DMARC records.`, "info");
  log(`Checking DNS records for ${domain}…`, "info");
  const dns = await checkDomainDns(domain);
  const missing = [];
  if (!dns.spf)  missing.push("SPF");
  if (!dns.dmarc) missing.push("DMARC");
  if (!dns.dkim)  missing.push("DKIM (selector1/selector2)");

  if (missing.length > 0) {
    log(`DNS pre-flight warning for ${domain}: Missing ${missing.join(", ")} records. Large sends without these may be marked as spam or rejected. Contact your IT admin to configure email authentication.`, "warning");
  } else {
    log(`DNS check passed for ${domain}: SPF, DMARC, and DKIM are configured.`, "success");
  }
}

/* ─── IMPORT FROM OUTLOOK CONTACTS (Feature 1) ─────────────────── */

// Contacts picker state
let contactsData = [];       // raw Graph contacts
let selectedContacts = new Set();

async function handleImportContacts() {
  document.getElementById("contactsModal").classList.remove("hidden");
  document.getElementById("contactsList").innerHTML = '<p style="padding:8px;color:#605e5c;">Loading…</p>';
  document.getElementById("contactsSearch").value = "";
  selectedContacts.clear();
  updateContactsSelectedCount();

  try {
    const token = await getAccessToken();
    // BUG 8: paginate through all contacts instead of capping at 500
    let url = "https://graph.microsoft.com/v1.0/me/contacts?$select=displayName,givenName,surname,emailAddresses&$top=500&$orderby=displayName";
    let allContacts = [];
    while (url) {
      const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
      if (!res.ok) throw new Error(`Graph ${res.status}`);
      const data = await res.json();
      allContacts = allContacts.concat((data.value || []).filter(c => c.emailAddresses && c.emailAddresses.length > 0));
      url = data["@odata.nextLink"] || null;
    }
    contactsData = allContacts;
    renderContactsList(contactsData);
  } catch (err) {
    document.getElementById("contactsList").innerHTML =
      `<p style="padding:8px;color:#eb5757;">Failed to load contacts: ${err.message}</p>`;
  }
}

function renderContactsList(contacts) {
  const el = document.getElementById("contactsList");
  if (!contacts.length) {
    el.innerHTML = '<p style="padding:8px;color:#605e5c;">No contacts found.</p>';
    return;
  }
  el.innerHTML = contacts.map((c) => {
    const email = c.emailAddresses[0].address;
    const name  = c.displayName || `${c.givenName || ""} ${c.surname || ""}`.trim();
    const checked = selectedContacts.has(email) ? "checked" : "";
    return `<label style="display:flex;align-items:center;gap:8px;padding:5px 8px;cursor:pointer;border-bottom:1px solid #f3f2f1;">
      <input type="checkbox" data-email="${escapeHtml(email)}" data-name="${escapeHtml(name)}"
             data-first="${escapeHtml(c.givenName||"")}" data-last="${escapeHtml(c.surname||"")}" ${checked} />
      <span><strong style="font-size:12px;">${escapeHtml(name)}</strong><br>
        <span style="color:#605e5c;font-size:11px;">${escapeHtml(email)}</span>
      </span>
    </label>`;
  }).join("");
  el.querySelectorAll("input[type=checkbox]").forEach(cb => {
    cb.addEventListener("change", () => {
      if (cb.checked) selectedContacts.add(cb.dataset.email);
      else selectedContacts.delete(cb.dataset.email);
      updateContactsSelectedCount();
    });
  });
}

function updateContactsSelectedCount() {
  const el = document.getElementById("contactsSelectedCount");
  if (el) el.textContent = selectedContacts.size
    ? `${selectedContacts.size} selected`
    : "";
}

function importSelectedContacts() {
  const selected = contactsData.filter(c =>
    selectedContacts.has(c.emailAddresses[0].address)
  );
  if (!selected.length) { log("No contacts selected.", "warning"); return; }

  // Build CSV rows
  const header = "email,first_name,last_name";
  const rows = selected.map(c => {
    const email = c.emailAddresses[0].address;
    const first = c.givenName  || "";
    const last  = c.surname    || "";
    return `${csvField(email)},${csvField(first)},${csvField(last)}`;
  });

  const csvInput = document.getElementById("csvInput");
  const existing = csvInput.value.trim();
  const emailHeader = "email,first_name,last_name";
  const existingLines = existing.split("\n").map(l => l.trim()).filter(Boolean);
  const existingHeader = existingLines[0] || "";

  if (!existing || existingHeader === emailHeader) {
    // Empty or already has matching header — append (or start fresh)
    csvInput.value = (existing ? existing + "\n" : emailHeader + "\n") + rows.join("\n");
  } else {
    // Different header — start fresh, warn
    log("Contact import: existing CSV has different columns — starting fresh with imported contacts.", "warning");
    csvInput.value = emailHeader + "\n" + rows.join("\n");
  }
  csvInput.dispatchEvent(new Event("input")); // trigger localStorage save

  document.getElementById("contactsModal").classList.add("hidden");
  log(`Imported ${selected.length} contact(s) from address book.`, "success");
  parseAndPreview();
}

// Contacts search filter
function filterContacts(query) {
  if (!query.trim()) { renderContactsList(contactsData); return; }
  const q = query.toLowerCase();
  renderContactsList(contactsData.filter(c => {
    const name  = (c.displayName || "").toLowerCase();
    const email = ((c.emailAddresses[0] && c.emailAddresses[0].address) || "").toLowerCase();
    return name.includes(q) || email.includes(q);
  }));
}

/* ─── CONTACT GROUPS IMPORT (Feature 3 v1.8) ───────────────────── */

async function loadGroups() {
  document.getElementById("groupsList").innerHTML = '<p style="padding:8px;color:#605e5c;">Loading…</p>';
  try {
    const token = await getAccessToken();
    // Fetch contact folders (personal contact groups)
    const res = await fetch(
      "https://graph.microsoft.com/v1.0/me/contactFolders?$top=50",
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!res.ok) throw new Error(`Graph ${res.status}`);
    const data = await res.json();
    groupsData = data.value || [];
    renderGroupsList();
  } catch(err) {
    document.getElementById("groupsList").innerHTML =
      `<p style="padding:8px;color:#eb5757;">Failed to load groups: ${err.message}</p>`;
  }
}

function renderGroupsList() {
  const el = document.getElementById("groupsList");
  if (!groupsData.length) {
    el.innerHTML = '<p style="padding:8px;color:#605e5c;">No contact folders found.</p>';
    return;
  }
  el.innerHTML = groupsData.map(g =>
    `<div style="padding:6px 8px;border-bottom:1px solid #f3f2f1;display:flex;align-items:center;gap:8px;cursor:pointer;"
          class="group-item" data-id="${escapeHtml(g.id)}" data-name="${escapeHtml(g.displayName)}">
      <span style="font-size:16px;">📁</span>
      <span style="font-size:12px;">${escapeHtml(g.displayName)}</span>
      <button class="btn-secondary" style="margin-left:auto;font-size:11px;padding:2px 8px;"
              data-group-id="${escapeHtml(g.id)}" data-group-name="${escapeHtml(g.displayName)}">
        Import
      </button>
    </div>`
  ).join("");

  el.querySelectorAll("button[data-group-id]").forEach(btn => {
    btn.addEventListener("click", () => importContactFolder(btn.dataset.groupId, btn.dataset.groupName));
  });
}

async function importContactFolder(folderId, folderName) {
  log(`Loading contacts from folder "${folderName}"…`, "info");
  try {
    const token = await getAccessToken();
    // BUG 8: paginate through all folder contacts instead of capping at 500
    let url = `https://graph.microsoft.com/v1.0/me/contactFolders/${folderId}/contacts?$select=displayName,givenName,surname,emailAddresses&$top=500`;
    let allContacts = [];
    while (url) {
      const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
      if (!res.ok) throw new Error(`Graph ${res.status}`);
      const data = await res.json();
      allContacts = allContacts.concat((data.value || []).filter(c => c.emailAddresses && c.emailAddresses.length > 0));
      url = data["@odata.nextLink"] || null;
    }
    const contacts = allContacts;
    if (!contacts.length) { log(`Folder "${folderName}" has no contacts with email addresses.`, "warning"); return; }

    const header = "email,first_name,last_name";
    const rows = contacts.map(c => {
      const email = c.emailAddresses[0].address;
      const first = c.givenName  || "";
      const last  = c.surname    || "";
      return `${csvField(email)},${csvField(first)},${csvField(last)}`;
    });

    const csvInput = document.getElementById("csvInput");
    const existing = csvInput.value.trim();
    const emailHeader = "email,first_name,last_name";
    const existingLines = existing.split("\n").map(l => l.trim()).filter(Boolean);
    const existingHeader = existingLines[0] || "";

    if (!existing || existingHeader === emailHeader) {
      // Empty or already has matching header — append (or start fresh)
      csvInput.value = (existing ? existing + "\n" : emailHeader + "\n") + rows.join("\n");
    } else {
      // Different header — start fresh, warn
      log("Contact import: existing CSV has different columns — starting fresh with imported contacts.", "warning");
      csvInput.value = emailHeader + "\n" + rows.join("\n");
    }
    csvInput.dispatchEvent(new Event("input"));
    document.getElementById("contactsModal").classList.add("hidden");
    log(`Imported ${contacts.length} contact(s) from folder "${folderName}".`, "success");
    parseAndPreview();
  } catch(err) {
    log(`Failed to import folder: ${err.message}`, "error");
  }
}

/* ─── DUPLICATE SEND GUARD (Feature 3 v1.9) ────────────────────── */

function recordSentEmails(emails) {
  const history = lsGet(LS_KEY_SEND_HISTORY, {});
  const now = new Date().toISOString();
  emails.forEach(rawEmail => {
    parseAddressList(rawEmail).forEach(entry => {
      history[entry.emailAddress.address.toLowerCase()] = now;
    });
  });
  // Prune entries older than 30 days (Feature 16: explicit numeric comparison avoids ISO string bugs)
  const cutoffMs = Date.now() - (30 * 24 * 60 * 60 * 1000);
  Object.keys(history).forEach(k => {
    if (new Date(history[k]).getTime() < cutoffMs) delete history[k];
  });
  lsSet(LS_KEY_SEND_HISTORY, JSON.stringify(history));
}

function checkDuplicateSendHistory(recipients) {
  const history = lsGet(LS_KEY_SEND_HISTORY, {});
  return recipients.filter(r => {
    const addrs = parseAddressList(r.email || "");
    return addrs.some(a => history[a.emailAddress.address.toLowerCase()]);
  });
}

function showDuplicateWarningModal(dupes) {
  return new Promise(resolve => {
    const body  = document.getElementById("dupeSendBody");
    body.innerHTML = `
      <p>${dupes.length} recipient(s) in this list were sent to in the past 30 days:</p>
      <ul style="max-height:120px;overflow-y:auto;font-size:11px;margin:6px 0 6px 16px;">
        ${dupes.slice(0, 10).map(r => `<li>${escapeHtml(r.email)}</li>`).join("")}
        ${dupes.length > 10 ? `<li>… and ${dupes.length - 10} more</li>` : ""}
      </ul>
      <p>Send anyway?</p>
    `;
    _openModalWithTrap("dupeSendModal"); // Feature 15: focus trap
    document.getElementById("dupeSendProceedBtn").onclick = () => {
      _closeModalWithTrap("dupeSendModal");
      resolve(true);
    };
    document.getElementById("dupeSendCancelBtn").onclick = () => {
      _closeModalWithTrap("dupeSendModal");
      resolve(false);
    };
  });
}

/* ─── BROADCAST MODE (Feature 1 v1.9) ──────────────────────────── */

async function handleBroadcast() {
  if (broadcastInProgress) {
    log("A broadcast is already in progress.", "warning");
    return;
  }
  broadcastInProgress = true;
  try {
  const rawAll = getFilteredSortedRecipients();
  if (!rawAll.length) { log("No recipients loaded.", "warning"); return; }

  // L3: filter skip_if rows (same logic as validateRecipients)
  const rawValid = rawAll.filter(r => {
    const s = (r.skip_if || "").trim().toLowerCase();
    return !s || s === "0" || s === "false" || s === "no";
  });
  if (rawValid.length < rawAll.length) {
    log(`Broadcast: ${rawAll.length - rawValid.length} row(s) skipped via skip_if column.`, "info");
  }

  const subject = document.getElementById("subjectInput").value.trim();
  if (!subject) { log("Enter a subject before broadcasting.", "warning"); return; }

  // A3: Apply suppression list — use parseAddressList so multi-TO rows match individual suppressed addresses
  const afterSuppression = rawValid.filter(r => {
    const addrs = parseAddressList(r.email || "");
    return !addrs.some(a => suppressionSet.has(a.emailAddress.address.toLowerCase()));
  });
  const suppressedCount = rawValid.length - afterSuppression.length;
  if (suppressedCount > 0) {
    log(`Broadcast: ${suppressedCount} suppressed address(es) excluded.`, "info");
  }

  // BUG 2: Apply persistent opt-out list
  const optOutSet = getOptOutList();
  const afterOptOut = afterSuppression.filter(r => {
    const addrs = parseAddressList(r.email || "");
    return !addrs.some(a => optOutSet.has(a.emailAddress.address.toLowerCase()));
  });
  if (afterOptOut.length < afterSuppression.length) {
    log(`Broadcast: ${afterSuppression.length - afterOptOut.length} opt-out recipient(s) excluded.`, "info");
  }

  // Validate emails
  const validEmails = afterOptOut.filter(r => {
    const addrs = parseAddressList(r.email || "");
    return addrs.length > 0;
  });
  const skipped = afterOptOut.length - validEmails.length;
  if (skipped > 0) log(`Broadcast: ${skipped} row(s) with invalid email skipped.`, "warning");

  if (!validEmails.length) { log("No valid recipients for broadcast.", "warning"); return; }

  // Duplicate guard
  const recentDupes = checkDuplicateSendHistory(validEmails);
  if (recentDupes.length) {
    const proceed = await showDuplicateWarningModal(recentDupes);
    if (!proceed) return; // A2: just return — do NOT call setMergeRunning(false) here; it was never set to true
  }

  // Confirm
  const confirmed = await showConfirmModal(validEmails.length, validEmails.length,
    null, /* broadcastMode = */ true);
  if (!confirmed) return;

  setMergeRunning(true);
  try {
    const token = await getAccessToken();
    let bodyHtml = await getComposeBodyAsync();
    const saveToSent  = document.getElementById("saveToSentItems").checked;
    const replyTo     = document.getElementById("replyToInput").value.trim();
    const importance  = document.getElementById("importanceSelect").value;
    // A4: Read all Options tab settings
    const plainTextMode          = document.getElementById("plainTextMode").checked;
    const sensitivity            = document.getElementById("sensitivitySelect").value;
    const categoriesRaw          = document.getElementById("categoriesInput").value.trim();
    const categories             = categoriesRaw ? categoriesRaw.split(",").map(c => c.trim()).filter(Boolean) : [];
    const isReadReceiptRequested     = document.getElementById("requestReadReceipt").checked;
    const isDeliveryReceiptRequested = document.getElementById("requestDeliveryReceipt").checked;
    const customHeadersEnabled   = document.getElementById("customHeadersEnabled").checked;
    const globalCustomHeaders    = customHeadersEnabled
      ? parseCustomHeaders(document.getElementById("customHeadersInput").value)
      : [];
    const flagged     = document.getElementById("flagForFollowup").checked;
    const expiryEnabled = document.getElementById("expiryEnabled").checked;
    let expiryISO = null;
    if (expiryEnabled) {
      const expiryValue = document.getElementById("expiryDateTimeInput").value;
      if (expiryValue) expiryISO = new Date(expiryValue).toISOString();
    }
    const sendAs      = document.getElementById("sendAsInput").value.trim();
    const bccSelf     = document.getElementById("bccSelfEnabled").checked;
    let selfEmail = "";
    try { selfEmail = Office.context.mailbox.userProfile.emailAddress || ""; } catch(e) { /* ignore */ }

    const bccAddresses = validEmails.flatMap(r =>
      parseAddressList(r.email || "").map(entry => ({
        emailAddress: { address: entry.emailAddress.address, name: r.display_name || "" }
      }))
    );
    // A4: BCC self
    if (bccSelf && selfEmail) {
      bccAddresses.push({ emailAddress: { address: selfEmail } });
    }

    // Feature 2: BCC recipient cap warnings
    if (bccAddresses.length > 999) {
      log("✗ Broadcast cancelled: over 999 addresses exceeds the Graph API hard limit.", "error");
      setMergeRunning(false);
      return;
    }
    if (bccAddresses.length > 500) {
      log(`⚠ Broadcast: ${bccAddresses.length} BCC recipients exceeds Exchange Online's 500-address limit. The broadcast will likely fail. Consider splitting into smaller groups.`, "warning");
    }

    // A4: plain text mode
    const bodyContent     = plainTextMode ? stripHtmlToText(bodyHtml) : bodyHtml;
    const bodyContentType = plainTextMode ? "Text" : "HTML";

    // BUG 5: toRecipients cannot be empty — fall back to sendAs, then a safe placeholder
    if (!selfEmail && !sendAs) {
      log("Broadcast: sender email unavailable — using fallback undisclosed-recipients address in To field. Some servers may reject this.", "warning");
    }
    const message = {
      subject,
      body: { contentType: bodyContentType, content: bodyContent },
      toRecipients: (selfEmail || sendAs)
        ? [{ emailAddress: { address: selfEmail || sendAs, name: "Undisclosed Recipients" } }]
        : [{ emailAddress: { address: "undisclosed-recipients@noreply.invalid", name: "Undisclosed Recipients" } }],
      bccRecipients: bccAddresses,
      // A11: only set importance if not "normal"
      ...(importance !== "normal" && { importance }),
    };
    if (replyTo) message.replyTo = [{ emailAddress: { address: replyTo } }];
    // A4: optional fields
    if (sensitivity && sensitivity !== "normal") message.sensitivity = sensitivity;
    if (categories.length > 0) message.categories = categories;
    if (isReadReceiptRequested) message.isReadReceiptRequested = true;
    if (isDeliveryReceiptRequested) message.isDeliveryReceiptRequested = true;
    if (globalCustomHeaders.length > 0) message.internetMessageHeaders = globalCustomHeaders;
    if (flagged) message.flag = { flagStatus: "flagged" };
    if (expiryISO) message.expiryDateTime = { dateTime: expiryISO, timeZone: "UTC" };
    if (sendAs) message.from = { emailAddress: { address: sendAs } };

    // A5 + Build attachments: shared attachments first, then inline images
    const attachmentsList = sharedAttachments.map(a => ({
      "@odata.type": "#microsoft.graph.fileAttachment",
      name: a.name,
      contentType: a.contentType,
      contentBytes: a.contentBytes,
    }));
    // A5: Append inline CID images
    inlineImages.forEach((img) => {
      attachmentsList.push({
        "@odata.type": "#microsoft.graph.fileAttachment",
        name: img.name,
        contentType: img.contentType,
        contentBytes: img.contentBytes,
        isInline: true,
        contentId: img.name,
      });
    });
    if (attachmentsList.length > 0) message.attachments = attachmentsList;

    const res = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ message, saveToSentItems: saveToSent }),
    });

    if (res.ok) {
      log(`Broadcast sent to ${validEmails.length} recipients (BCC).`, "success");
      recordSentEmails(validEmails.map(r => r.email));
      // L6: record broadcast send outcomes so the report includes broadcast recipients
      const ts = new Date().toISOString();
      bccAddresses.forEach(addr => {
        sendOutcomes.push({
          rowNum: "",
          email: typeof addr === "string" ? addr : (addr.emailAddress && addr.emailAddress.address) || "",
          displayName: (typeof addr === "object" && addr.emailAddress && addr.emailAddress.name) || "",
          subjectUsed: subject.slice(0, 100),
          status: "broadcast",
          timestamp: ts,
          error: ""
        });
      });
      if (sendOutcomes.length > 0) document.getElementById("downloadReportBtn").classList.remove("hidden");
    } else {
      const err = await res.json().catch(() => ({}));
      log(`Broadcast failed: ${(err.error && err.error.message) || res.status}`, "error");
    }
  } catch(e) {
    log(`Broadcast error: ${e.message}`, "error");
  } finally {
    setMergeRunning(false);
  }
  } finally {
    broadcastInProgress = false;
  }
}

/* ─── SCHEDULING / RATE LIMIT ACTIVE BADGES (UX 7) ─────────────── */

function updateSchedulingBadge() {
  const active = document.getElementById("scheduleEnabled")?.checked ||
                 document.getElementById("sendingWindowEnabled")?.checked;
  const badge = document.getElementById("schedulingActiveBadge");
  if (badge) badge.classList.toggle("hidden", !active);
}

function updateRateLimitBadge() {
  const mphVal = parseInt(document.getElementById("maxPerHour")?.value || "0", 10);
  const dcVal  = parseInt(document.getElementById("dailyCap")?.value   || "0", 10);
  const active = mphVal > 0 || dcVal > 0;
  const badge  = document.getElementById("rateLimitActiveBadge");
  if (badge) badge.classList.toggle("hidden", !active);
}

/* ─── SIMULATE / DRY-RUN MODE (Feature 10) ─────────────────────── */

async function handleSimulate() {
  if (!parsedRecipients.length) { log("Load a CSV first.", "warning"); return; }
  const valid = getFilteredSortedRecipients();
  if (!valid.length) { log("No recipients match current filter.", "warning"); return; }

  const simBtn = document.getElementById("simulateBtn");
  if (simBtn) simBtn.disabled = true;
  try {
    const subject = document.getElementById("subjectInput")?.value || "";
    let bodyHtml = "";
    try { bodyHtml = await getComposeBodyAsync(); } catch(e) { bodyHtml = ""; }

    const rows = [["record_num", "email", "display_name", "subject_preview", "body_tokens_missing", "skip_if_result", "attachment_found"]];

    for (let i = 0; i < valid.length; i++) {
      if (i > 0 && i % 50 === 0) await new Promise(r => setTimeout(r, 0)); // P1: yield to UI thread
      const r = valid[i];
      const augmented = Object.assign({}, r, {
        record_num:   String(i + 1),
        record_count: String(valid.length)
      });

      const personalizedSubject = personalize(subject, augmented, false);
      const personalizedBody    = personalize(bodyHtml,  augmented, true);

      // Find unresolved tokens (still contain {{ }})
      const unresolvedMatches = personalizedBody.match(/\{\{[^}]+\}\}/g) || [];
      const unresolved = unresolvedMatches.join("; ") || "none";

      // skip_if result — L7: logic now matches validateRecipients exactly
      const skipVal = (r.skip_if || "").trim().toLowerCase();
      const skipResult = (skipVal && skipVal !== "0" && skipVal !== "false" && skipVal !== "no") ? "SKIP" : "send";

      rows.push([
        String(i + 1),
        r.email || "",
        r.display_name || r.first_name || "",
        personalizedSubject.slice(0, 80),
        unresolved,
        skipResult,
        r.attachment ? "check manually" : "n/a"
      ]);
    }

    const csv = rows.map(row => row.map(v => csvField(String(v))).join(",")).join("\r\n");
    const blob = new Blob([csv], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "merge-simulation-" + new Date().toISOString().slice(0, 10) + ".csv";
    a.click();
    URL.revokeObjectURL(url);
    log("Simulation complete: " + valid.length + " rows. Check the downloaded CSV for token issues.", "success");
  } catch(e) {
    log("Simulation failed: " + e.message, "error");
  } finally {
    if (simBtn) simBtn.disabled = false;
  }
}

/* ─── GAL / COMPANY DIRECTORY IMPORT (Feature 4 v1.9) ──────────── */

async function searchDirectory(query) {
  const el = document.getElementById("directoryList");
  if (!query || query.length < 2) {
    el.innerHTML = '<p style="padding:8px;color:#605e5c;">Type at least 2 characters to search the company directory.</p>';
    return;
  }
  el.innerHTML = '<p style="padding:8px;color:#605e5c;">Searching…</p>';
  try {
    const token = await getAccessToken();
    const url = `https://graph.microsoft.com/v1.0/me/people?$search=${encodeURIComponent(query)}&$select=displayName,givenName,surname,scoredEmailAddresses&$top=25`;
    const res = await fetch(url, {
      headers: { Authorization: `Bearer ${token}`, ConsistencyLevel: "eventual" }
    });
    if (res.status === 403) {
      el.innerHTML = '<p style="padding:8px;color:#eb5757;">Directory search requires the People.Read permission on the Entra app registration.</p>';
      return;
    }
    if (!res.ok) throw new Error(`Graph ${res.status}`);
    const data = await res.json();
    directoryData = (data.value || []).filter(p =>
      p.scoredEmailAddresses && p.scoredEmailAddresses.length > 0
    );
    renderDirectoryList();
  } catch(err) {
    el.innerHTML = `<p style="padding:8px;color:#eb5757;">Search failed: ${err.message}</p>`;
  }
}

function renderDirectoryList() {
  const el = document.getElementById("directoryList");
  if (!directoryData.length) {
    el.innerHTML = '<p style="padding:8px;color:#605e5c;">No results.</p>';
    return;
  }
  el.innerHTML = directoryData.map((p) => {
    const email = p.scoredEmailAddresses[0].address;
    const name  = p.displayName || "";
    const checked = selectedDirectory.has(email) ? "checked" : "";
    return `<label style="display:flex;align-items:center;gap:8px;padding:5px 8px;cursor:pointer;border-bottom:1px solid #f3f2f1;">
      <input type="checkbox" data-email="${escapeHtml(email)}" data-name="${escapeHtml(name)}"
             data-first="${escapeHtml(p.givenName||"")}" data-last="${escapeHtml(p.surname||"")}" ${checked} />
      <span><strong style="font-size:12px;">${escapeHtml(name)}</strong><br>
        <span style="color:#605e5c;font-size:11px;">${escapeHtml(email)}</span>
      </span>
    </label>`;
  }).join("");
  el.querySelectorAll("input[type=checkbox]").forEach(cb => {
    cb.addEventListener("change", () => {
      if (cb.checked) selectedDirectory.add(cb.dataset.email);
      else selectedDirectory.delete(cb.dataset.email);
    });
  });
}

function importDirectorySelected() {
  const selected = directoryData.filter(p =>
    selectedDirectory.has(p.scoredEmailAddresses[0].address)
  );
  if (!selected.length) { log("No directory contacts selected.", "warning"); return; }
  const header = "email,first_name,last_name";
  const rows = selected.map(p => {
    const email = p.scoredEmailAddresses[0].address;
    return `${csvField(email)},${csvField(p.givenName||"")},${csvField(p.surname||"")}`;
  });
  const csvInput = document.getElementById("csvInput");
  const existing = csvInput.value.trim();
  const emailHeader = "email,first_name,last_name";
  const existingLines = existing.split("\n").map(l => l.trim()).filter(Boolean);
  const existingHeader = existingLines[0] || "";

  if (!existing || existingHeader === emailHeader) {
    // Empty or already has matching header — append (or start fresh)
    csvInput.value = (existing ? existing + "\n" : emailHeader + "\n") + rows.join("\n");
  } else {
    // Different header — start fresh, warn
    log("Contact import: existing CSV has different columns — starting fresh with imported contacts.", "warning");
    csvInput.value = emailHeader + "\n" + rows.join("\n");
  }
  csvInput.dispatchEvent(new Event("input"));
  document.getElementById("contactsModal").classList.add("hidden");
  log(`Imported ${selected.length} person(s) from company directory.`, "success");
  parseAndPreview();
}

/* ─── IN-COMPOSE FIELD PICKER (Feature 4 v1.8) ─────────────────── */

function populateInsertFieldSelect(headers) {
  const sel = document.getElementById("insertFieldSelect");
  const smart = ["greeting_line", "today", "now", "record_num", "record_count", "merge_table", "fill_in:Your prompt", "unsubscribe_link"];
  // P2: build all options then assign once — avoids repeated innerHTML += reparse
  const smartOpts = smart.map(t => `<option value="{{${t}}}">✨ {{${t}}}</option>`).join("");
  const colOpts   = headers.map(h => `<option value="{{${escapeHtml(h)}}}">{{${escapeHtml(h)}}}</option>`).join("");
  sel.innerHTML = '<option value="">— pick field —</option>' + smartOpts + colOpts;
}

async function handleInsertField() {
  const sel = document.getElementById("insertFieldSelect");
  const token = sel.value;
  if (!token) { log("Pick a field to insert.", "warning"); return; }
  // B1: On Mac, CoercionType.Html causes font-size bugs and fails after multiple calls.
  // Tokens like {{first_name}} are plain text so Text coercion is safe.
  const isMac = Office.context.platform === Office.PlatformType.Mac;
  const coercionType = isMac ? Office.CoercionType.Text : Office.CoercionType.Html;
  try {
    await new Promise((resolve, reject) => {
      Office.context.mailbox.item.body.setSelectedDataAsync(
        token,
        { coercionType },
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
          else reject(new Error(result.error.message));
        }
      );
    });
    log(`Inserted ${token} into body.`, "info");
  } catch(e) {
    log(`Could not insert into body: ${e.message}`, "error");
  }
}

/* ─── CHECK FOR ERRORS PRE-FLIGHT (Feature 1 v1.8) ─────────────── */

async function handleCheckErrors() {
  if (!parsedRecipients.length) {
    log("Load a CSV first before checking for errors.", "warning");
    return;
  }

  // Collect known tokens: CSV headers + smart tokens
  const csvHeaders = new Set(Object.keys(parsedRecipients[0]).map(k => k.toLowerCase()));
  const smartTokens = new Set(["greeting_line", "today", "now", "record_num", "record_count", "merge_table", "unsubscribe_link"]);
  const allKnown = new Set([...csvHeaders, ...smartTokens]);

  // Get subject
  const subject = document.getElementById("subjectInput").value || "";

  // Get body HTML
  let bodyHtml = "";
  try {
    bodyHtml = await getComposeBodyAsync();
  } catch(e) {
    bodyHtml = "";
  }

  const combined = subject + " " + bodyHtml;

  // Find all {{token}} and {{token|filter}} references
  const tokenRegex = /\{\{([^}|]+)(?:\|[^}]*)?\}\}/gi;
  const usedTokens = new Set();
  let m;
  while ((m = tokenRegex.exec(combined)) !== null) {
    // Skip {{if:...}} conditionals — handled separately; skip built-in runtime tokens
    const key = m[1].trim().toLowerCase();
    if (key.startsWith("if:") || key.startsWith("fill_in:") || key === "merge_table") continue;
    usedTokens.add(key);
  }

  // Also extract tokens used inside {{if:...}} conditions
  const ifRegex = /\{\{if:([^}]+)\}\}/gi;
  while ((m = ifRegex.exec(combined)) !== null) {
    const inner = m[1];
    // condition part: before first colon
    const condPart = inner.slice(0, inner.indexOf(":")).trim().toLowerCase();
    const opMatch = condPart.match(/^(.+?)(>=|<=|!=|=|>|<)(.+)$/);
    if (opMatch) usedTokens.add(opMatch[1].trim());
    else usedTokens.add(condPart);
  }

  const unknownTokens = [...usedTokens].filter(t => !allKnown.has(t));

  // For each used token that IS known, find rows where it's empty
  const blanksByToken = {};
  [...usedTokens].filter(t => allKnown.has(t) && !smartTokens.has(t)).forEach(t => {
    const emptyRows = parsedRecipients
      .map((r, i) => ({ idx: i + 1, email: r.email || "?", val: r[t] }))
      .filter(r => !r.val || !String(r.val).trim());
    if (emptyRows.length) blanksByToken[t] = emptyRows;
  });

  // Build report HTML
  const lines = [];

  if (!usedTokens.size) {
    lines.push('<p style="color:#605e5c;">No merge tokens found in subject or body.</p>');
  } else {
    lines.push(`<p style="margin-bottom:8px;"><strong>${usedTokens.size}</strong> token(s) used across subject and body.</p>`);
  }

  if (unknownTokens.length) {
    lines.push(`<p style="color:#eb5757;margin-bottom:4px;"><strong>⚠ ${unknownTokens.length} unknown token(s)</strong> — not in CSV headers or smart tokens:</p>`);
    lines.push('<ul style="margin:0 0 8px 16px;">');
    unknownTokens.forEach(t => lines.push(`<li><code>{{${escapeHtml(t)}}}</code> — will not be replaced</li>`)); // S1
    lines.push('</ul>');
  } else if (usedTokens.size) {
    lines.push('<p style="color:#107c10;margin-bottom:8px;">✔ All tokens match CSV columns or smart tokens.</p>');
  }

  if (Object.keys(blanksByToken).length) {
    lines.push(`<p style="color:#f2994a;margin-bottom:4px;"><strong>⚠ Blank values found:</strong></p>`);
    Object.entries(blanksByToken).forEach(([token, rows]) => {
      lines.push(`<p style="margin:4px 0 2px;"><code>{{${escapeHtml(token)}}}</code> is blank in ${rows.length} row(s):</p>`); // S1
      lines.push('<ul style="margin:0 0 6px 16px;color:#605e5c;">');
      rows.slice(0, 5).forEach(r => lines.push(`<li>Row ${r.idx}: ${escapeHtml(r.email)}</li>`));
      if (rows.length > 5) lines.push(`<li>… and ${rows.length - 5} more</li>`);
      lines.push('</ul>');
    });
  } else if (usedTokens.size && !unknownTokens.length) {
    lines.push('<p style="color:#107c10;">✔ No blank values found for any used token.</p>');
  }

  // Overall verdict
  const hasIssues = unknownTokens.length || Object.keys(blanksByToken).length;
  lines.push(`<hr style="border:none;border-top:1px solid #edebe9;margin:8px 0;">`);
  lines.push(hasIssues
    ? `<p style="color:#f2994a;font-weight:600;">Ready to send with warnings. Review issues above.</p>`
    : `<p style="color:#107c10;font-weight:600;">✔ No errors found. Ready to send.</p>`);

  document.getElementById("checkErrorsBody").innerHTML = lines.join("");
  _openModalWithTrap("checkErrorsModal"); // Feature 15: focus trap
}

/* ═══════════════════════════════════════════════════════════════════
   EDIT RECIPIENT TABLE — in-memory CSV editor
   Lets users add/rename/delete columns and edit cell values without
   leaving the add-in or modifying the original file.
   ═══════════════════════════════════════════════════════════════════ */

function openEditTableModal() {
  if (!parsedRecipients.length) {
    log("Load a recipient list first before editing.", "error");
    return;
  }
  editTableHeaders = Object.keys(parsedRecipients[0]).filter(k => !k.startsWith("_"));
  editTableRows = parsedRecipients.map(row =>
    editTableHeaders.map(h => row[h] !== undefined ? String(row[h]) : "")
  );
  const newColInput = document.getElementById("newColNameInput");
  if (newColInput) newColInput.value = "";
  renderEditTable();
  document.getElementById("editTableModal").classList.remove("hidden");
  const firstCell = document.querySelector("#editTableGrid .edit-cell-input");
  if (firstCell) firstCell.focus();
}

function renderEditTable() {
  const grid = document.getElementById("editTableGrid");
  let html = '<table><thead><tr>';
  html += '<th style="width:28px;min-width:28px;"></th>';
  editTableHeaders.forEach(function(h, colIdx) {
    html +=
      '<th>' +
        '<div class="edit-col-header-wrap">' +
          '<input class="edit-col-header-input" data-col="' + colIdx + '" value="' + escapeHtml(h) + '" title="Rename column" />' +
          '<button class="edit-col-del" data-col="' + colIdx + '" title="Remove column">✕</button>' +
        '</div>' +
      '</th>';
  });
  html += '</tr></thead><tbody>';
  editTableRows.forEach(function(row, rowIdx) {
    html += '<tr>';
    html += '<td class="edit-row-num">' + (rowIdx + 1) + '</td>';
    row.forEach(function(val, colIdx) {
      html +=
        '<td><input class="edit-cell-input"' +
          ' data-row="' + rowIdx + '"' +
          ' data-col="' + colIdx + '"' +
          ' value="' + escapeHtml(val) + '"' +
          ' /></td>';
    });
    html += '</tr>';
  });
  html += '</tbody></table>';
  grid.innerHTML = html;

  grid.querySelectorAll(".edit-col-header-input").forEach(function(input) {
    input.addEventListener("change", function() {
      const col = parseInt(this.dataset.col, 10);
      const trimmed = this.value.trim();
      editTableHeaders[col] = trimmed || editTableHeaders[col];
      this.value = editTableHeaders[col];
    });
  });

  grid.querySelectorAll(".edit-col-del").forEach(function(btn) {
    btn.addEventListener("click", function() {
      const col = parseInt(this.dataset.col, 10);
      if (editTableHeaders.length <= 1) { log("Cannot remove the last column.", "error"); return; }
      flushEditTableInputs();
      editTableHeaders.splice(col, 1);
      editTableRows.forEach(function(row) { row.splice(col, 1); });
      renderEditTable();
    });
  });

  grid.querySelectorAll(".edit-cell-input").forEach(function(input) {
    input.addEventListener("input", function() {
      editTableRows[parseInt(this.dataset.row, 10)][parseInt(this.dataset.col, 10)] = this.value;
    });
  });
}

function flushEditTableInputs() {
  const grid = document.getElementById("editTableGrid");
  grid.querySelectorAll(".edit-col-header-input").forEach(function(input) {
    const col = parseInt(input.dataset.col, 10);
    const trimmed = input.value.trim();
    if (trimmed) editTableHeaders[col] = trimmed;
  });
  grid.querySelectorAll(".edit-cell-input").forEach(function(input) {
    editTableRows[parseInt(input.dataset.row, 10)][parseInt(input.dataset.col, 10)] = input.value;
  });
}

function editTableAddColumn() {
  const input = document.getElementById("newColNameInput");
  const name  = (input ? input.value.trim() : "");
  if (!name) {
    log("Enter a column name in the box next to '+ Add column'.", "warning");
    if (input) input.focus();
    return;
  }
  if (editTableHeaders.map(h => h.toLowerCase()).includes(name.toLowerCase())) {
    log('A column named "' + escapeHtml(name) + '" already exists.', "warning");
    if (input) input.focus();
    return;
  }
  flushEditTableInputs();
  editTableHeaders.push(name);
  editTableRows.forEach(function(row) { row.push(""); });
  if (input) input.value = "";
  renderEditTable();
  const grid = document.getElementById("editTableGrid");
  grid.scrollLeft = grid.scrollWidth;
  const newCells = grid.querySelectorAll(".edit-cell-input[data-col='" + (editTableHeaders.length - 1) + "']");
  if (newCells.length) newCells[0].focus();
}

function saveEditTableChanges() {
  flushEditTableInputs();
  const oldRecipients = parsedRecipients.slice();
  parsedRecipients = editTableRows.map(function(row, i) {
    const obj = { _originalIndex: i };
    // Preserve internal tracking fields from original row so send logic still works
    if (oldRecipients[i] && oldRecipients[i]._csvRow !== undefined) {
      obj._csvRow = oldRecipients[i]._csvRow;
    }
    editTableHeaders.forEach(function(h, col) {
      obj[h] = row[col] !== undefined ? row[col] : "";
    });
    return obj;
  });
  document.getElementById("editTableModal").classList.add("hidden");
  refreshAfterTableEdit();
}

function refreshAfterTableEdit() {
  const headers = editTableHeaders.slice();
  document.getElementById("recipientCount").textContent =
    parsedRecipients.length + ' recipient' + (parsedRecipients.length !== 1 ? 's' : '') + ' loaded';
  const testRowSel = document.getElementById("testRowSelect");
  if (testRowSel) {
    testRowSel.innerHTML = parsedRecipients.map(function(r, i) {
      const label = escapeHtml((r.email || r.first_name || "").slice(0, 28));
      return '<option value="' + i + '">Row ' + (i + 1) + ': ' + label + '</option>';
    }).join("");
  }
  headers.forEach(function(h) {
    const tag = '{{' + h + '}}';
    if (!DEFAULT_TAGS.includes(tag)) addTagToBar(tag);
  });
  populateFilterSortBar(headers);
  populateInsertFieldSelect(headers);
  activeFilter       = null;
  previewTablePage   = 0;
  selectedRowIndices = null;
  renderPreviewTable(parsedRecipients);
  log(
    'Table saved — ' + parsedRecipients.length + ' row' + (parsedRecipients.length !== 1 ? 's' : '') +
    ', ' + headers.length + ' column' + (headers.length !== 1 ? 's' : '') + '.',
    'success'
  );
}
