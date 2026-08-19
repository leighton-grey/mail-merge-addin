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

// ── Microsoft Graph API constants ─────────────────────────────────────────
// Graph's $batch endpoint lets us send up to 20 API calls in a single HTTP
// request. This massively reduces round-trips and is critical for large merges.
// Without batching, sending 200 emails would require 200 separate HTTP requests.
const GRAPH_BATCH_URL     = "https://graph.microsoft.com/v1.0/$batch";

// How many sendMail calls to pack into each $batch request. Graph's hard limit
// is 20 — exceeding it returns a 400 error. We reduce this dynamically when
// per-recipient attachments are present (larger payloads = smaller batches).
const BATCH_SIZE          = 20;

// Milliseconds to wait between consecutive batch requests. Without a delay,
// back-to-back batches can trigger Graph's throttle (429 Too Many Requests).
// 1500 ms ≈ 13 emails/sec, well within Microsoft's per-user send limits.
// Overridden at runtime by the batchDelayInput UI control.
const BATCH_DELAY_MS      = 1500;

// How many times to retry a failed batch request before giving up. Graph
// sends 429 responses with a Retry-After header when rate-limited; the
// retry loop reads that header and waits the specified time before retrying.
const MAX_RETRIES         = 3;

// ── Licensing ──────────────────────────────────────────────────────────────
// LICENSE_ENFORCEMENT = false means the license check runs silently in the
// background — it logs the result but never blocks features. Flip to true
// when you want to gate the add-in behind a paid subscription. This way you
// can ship the check code to production first, verify it works correctly in
// the logs, and then flip the flag when ready — with zero code changes.
// IMPORTANT: never change this without explicit instruction from Leighton.
const LICENSE_ENFORCEMENT      = false;

// The Azure Function URL that validates a user's subscription. Replace
// YOUR_FUNCTION_APP with the actual Function App name before going live.
// The function receives userId + tenantId as query params and returns
// { licensed: bool, plan: string, token: string, expiresAt: ISO date }.
const LICENSE_API_URL          = "https://YOUR_FUNCTION_APP.azurewebsites.net/api/license";

// localStorage keys for the license cache. We cache the result for 24 hours
// so the add-in doesn't hit the license API on every startup — that would be
// slow and hammers your Azure Function bill unnecessarily.
const LICENSE_CACHE_KEY        = "mailmerge_license_token";
const LICENSE_CACHE_EXPIRY_KEY = "mailmerge_license_expiry";
// ──────────────────────────────────────────────────────────────────────────

// RFC 5322-compliant email address regex. More permissive than a simple
// "contains @" check but strict enough to catch typos like "bob@" or "@company.com".
// We use this everywhere we need to validate an address before sending to Graph.
// Note: this regex does NOT allow IP-address domains (e.g. user@[192.168.1.1])
// which is intentional — Exchange Online doesn't accept them anyway.
const EMAIL_REGEX         = /^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)+$/;

// Regex that matches mail merge tokens like {{first_name}} or {{field|Fallback}}.
// The two capture groups are: (1) the field key, (2) the optional fallback value
// after the pipe. Declared at module scope (hoisted) so JavaScript compiles the
// regex once when the script loads rather than recompiling it inside every call
// to personalize(). The 'g' flag means it finds ALL tokens in a template.
// IMPORTANT: always reset TOKEN_REGEX.lastIndex = 0 before calling .replace()
// or .exec() in a loop — the 'g' flag maintains state between calls, and
// forgetting to reset it causes intermittent "skips every other token" bugs.
const TOKEN_REGEX         = /\{\{([^}|]+)(?:\|([^}]*))?\}\}/gi;

// Hard cap on recipient list size — prevents UI freezes on huge CSV files.
// Graph can handle much larger sends; this limit is about browser performance.
const MAX_RECIPIENTS      = 10000;

// Graph's documented limit per sendMail request is 4 MB. We warn at 3.5 MB
// to leave headroom for base64-encoded attachments and JSON overhead.
const MAX_PAYLOAD_BYTES   = 3.5 * 1024 * 1024;

// Graph's limit for a single attachment is 3 MB (via the non-resumable path).
// For larger files you'd need the resumable upload session API — we don't support
// that here, so anything above this limit is rejected with a warning.
const MAX_ATTACHMENT_BYTES = 3 * 1024 * 1024;

// ── localStorage keys ─────────────────────────────────────────────────────
// Centralising key strings as constants prevents typos from silently creating
// orphaned localStorage entries that are never read back. If you ever need to
// rename a key (e.g. after a breaking schema change), change it in one place
// here rather than hunting through the codebase for every "mailmerge_xxx" string.

// Persists the subject-line template across page refreshes. Stored as a plain
// string — the same text that populates #subjectInput on load.
const LS_KEY_SUBJECT      = "mailmerge_subject";

// Persists the raw CSV text across refreshes. Allows the user to close the
// add-in task pane and reopen it without losing their recipient list.
const LS_KEY_CSV          = "mailmerge_csv";

// Persists user-defined custom tags (beyond the built-in DEFAULT_TAGS list).
// Stored as a JSON-serialised array of strings like ["{{promo_code}}", ...].
const LS_KEY_TAGS         = "mailmerge_custom_tags";

// Persists saved named templates — subject, body, and all settings.
// Stored as a JSON array of template objects; see saveTemplate() for the schema.
const LS_KEY_TEMPLATES    = "mailmerge_templates";

// Persists the greeting-line configuration (format + fallback).
// Stored as a JSON object matching greetingConfig's shape.
const LS_KEY_GREETING     = "mailmerge_greeting";

// Persists a map of { email → ISO timestamp } for the last successful send to
// each address. Used by the scheduling / sending-window logic to enforce
// per-recipient cooldown periods and the daily-cap check.
const LS_KEY_SEND_HISTORY = "mailmerge_send_history"; // { email: isoTimestamp }

// ── Canonical field definitions ───────────────────────────────────────────
// CANONICAL_FIELDS is the authoritative list of column names the add-in
// understands natively. When the user uploads a CSV, the Match Fields dialog
// maps arbitrary CSV column headers onto these canonical keys.
//
// The `required` flag drives validation: if no column is mapped to "email",
// parseCSV/validateRecipients will refuse to proceed. All other fields are
// optional — missing values are treated as empty strings, which is fine for
// personalisation tokens (they render as blank rather than "{{first_name}}").
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

// ── Default tag chips ─────────────────────────────────────────────────────
// These are the built-in tokens shown as clickable chips in the tag toolbar.
// They map to either: (a) a canonical CSV column (first_name, email, etc.),
// or (b) a computed smart token resolved at send time (today, now, record_num).
// Users can add additional chips via addCustomTag(); those are persisted under
// LS_KEY_TAGS. DEFAULT_TAGS themselves are hard-coded and never stored —
// they're always rendered fresh from this array.
const DEFAULT_TAGS = [
  "{{first_name}}", "{{last_name}}", "{{email}}", "{{company}}", "{{title}}",
  "{{greeting_line}}", "{{today}}", "{{now}}", "{{record_num}}", "{{record_count}}",
  "{{merge_table}}", "{{unsubscribe_link}}"
];

// ── Module-level state variables ──────────────────────────────────────────
// These are the "global" state variables for the add-in. Keeping them at module
// scope (rather than inside functions) lets every function share them without
// threading values through call chains. The trade-off is that you must be careful
// about resetting them between runs — stale state from a previous merge is a
// common source of bugs (several fixes in the changelog address exactly this).

// mergeInProgress: the double-send guard. Before starting any merge operation,
// check this flag; set it to true at the start and false in the finally block.
// Without this, a user clicking "Send" twice quickly (or the scheduler firing
// while a manual merge runs) would launch two concurrent Graph API sessions,
// causing duplicate emails and confusing log output.
let mergeInProgress           = false;  // BUG 1: double-send guard

// broadcastInProgress: same concept for the broadcast (one-to-many BCC) path,
// which has its own entry point (handleBroadcast) and therefore its own flag.
// Broadcast and normal merge can't safely run simultaneously — both write to
// sendOutcomes, failedRecipients, and the progress bar.
let broadcastInProgress       = false;  // UX Bug 6: re-entrancy guard for handleBroadcast

// _mergeCompletedSuccessfully: set to true in the merge finally block when
// everything sent without cancellation. The progress-hide animation checks this
// flag — if the merge succeeded, we leave the progress bar at 100% for a moment
// rather than hiding it immediately, giving the user visual confirmation.
let _mergeCompletedSuccessfully = false; // Feature 5: suppress progress hide on success

// previewTablePage / PREVIEW_PAGE_SIZE: the recipient preview table is paginated
// to avoid rendering thousands of DOM rows, which freezes the browser. Page 0
// is the first page. PREVIEW_PAGE_SIZE controls rows per page.
let previewTablePage          = 0;      // Feature 6: preview table paging
const PREVIEW_PAGE_SIZE       = 10;     // Feature 6: rows per page

// parsedRecipients: the working recipient list produced by parseCSV() or the
// quick-entry parser. Each element is a plain object whose keys are the CSV
// column names (mapped to canonical names by fieldMapping). This array drives
// everything — preview, validation, the merge loop, retry, and the summary report.
// Cleared on each CSV reload to prevent the v1.4.0 "stale list" bug.
let parsedRecipients          = [];

// cancelRequested: set to true by handleStop(). The merge loop checks this flag
// at the start of each batch iteration and exits cleanly if it's set.
// Using a flag (rather than throwing an exception) gives the loop a chance to
// finish the current in-flight batch before stopping, avoiding partial sends.
let cancelRequested           = false;

// subjectHasFocus: true while the #subjectInput text field has keyboard focus.
// Used by insertTag() to decide whether to insert the token into the subject
// field or the Outlook compose body. Must be tracked here because focus state
// is lost by the time a chip's click handler fires (blur fires before click).
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
// tagTarget: controls where clicking a tag chip inserts the token.
//   "body"    → Office.js setSelectedDataAsync (cursor position in Outlook compose body)
//   "subject" → Office.js subject.getAsync + setAsync (appends to Outlook native subject)
// Toggled by the Body / Subject segmented control in the tag header.
// Note: when #subjectInput itself is focused, insertTag() always writes
// into that field regardless of tagTarget (subjectHasFocus takes priority).
let tagTarget                 = "body";

// _lastOutlookSubject: caches the most recent subject string read from or pushed
// to Outlook's native compose subject field via subject.getAsync/setAsync.
// Lets the poll/push debounce skip redundant API calls when nothing has changed.
let _lastOutlookSubject       = null;   // last value read from / pushed to Outlook's native subject

// _subjectPushTimer: debounce timer ID for the push direction (local → Outlook).
// We don't call subject.setAsync on every keystroke — that would spam the
// Office.js API and cause flickering. Instead we wait 400 ms after the last
// keystroke before pushing. clearTimeout/setTimeout pattern resets the timer
// on each new keystroke.
let _subjectPushTimer         = null;   // debounce timer for pushing to Outlook

// _subjectPollTimer: handle for the setInterval that polls Outlook's native
// subject field. Because Outlook can modify the subject independently (e.g. when
// the user types directly in the subject box), we poll every 2 s and sync
// back to #subjectInput if the value changed externally.
let _subjectPollTimer         = null;   // interval for polling Outlook subject

// sharedAttachments: files the user has attached via the "Add shared attachment"
// picker. These are sent identically to EVERY recipient. Each entry is:
//   { name: string, contentType: string, contentBytes: base64, sizeBytes: number }
// Kept in memory only — large files here inflate every email's payload.
let sharedAttachments         = [];   // array of { name, contentType, contentBytes, sizeBytes }

// perRecipientFiles: Map from filename (lowercased) → attachment object for
// per-recipient attachments. The CSV "attachment" column contains filenames;
// this Map is populated when the user drops a folder of files. At send time,
// resolveAttachmentForRecipient() looks up the recipient's filename here.
let perRecipientFiles         = new Map();

// inlineImages: Map from filename (lowercased) → attachment object for CID-embedded
// inline images. When the body HTML contains <img src="cid:filename.png">, we look
// the file up here and attach it with isInline:true and the matching contentId.
let inlineImages              = new Map(); // filename.toLowerCase() → { name, contentType, contentBytes, sizeBytes }

// suppressionSet: Set of email addresses that have opted out / unsubscribed.
// Loaded from localStorage on startup. Recipients whose normalised address
// appears here are silently skipped during the merge loop (never sent to Graph).
// Managed via the Suppression UI — users can add/remove/export addresses.
let suppressionSet            = new Set();

// failedRecipients: array of recipient objects that failed to send during the
// most recent merge run. Populated inside the batch loop on Graph errors.
// Drives the "Retry failed" button — handleRetryFailed() reads this array
// and re-sends only these rows.
let failedRecipients          = [];

// sendOutcomes: accumulates { email, status, timestamp, error? } for every send
// attempt in the current merge run, including retries. Used to generate the
// downloadable CSV summary report at the end of a merge. Note: this array
// accumulates across retries within a single run (not across separate runs).
let sendOutcomes              = [];   // { email, status, timestamp, error? } — for summary report

// previewBodyTemplate: caches the raw HTML body string at the moment preview-all
// is launched. We snapshot it here so that edits made in Outlook's compose window
// during a long preview session don't affect the previews already generated.
let previewBodyTemplate       = "";   // cached body for preview-all mode

// warnedMissingAttachments: tracks which per-recipient attachment filenames have
// already generated a "file not found" warning in the current run. Without this,
// a 500-row merge where 10 rows reference a missing file would emit 10 identical
// warning lines in the log — confusing and hard to scan.
const warnedMissingAttachments = new Set(); // deduplicate per-filename warnings across a merge run

// previewIndex / previewRecipients: state for the preview-all modal navigator.
// previewIndex is the zero-based position in previewRecipients (the filtered
// working set, not all of parsedRecipients). Kept here so the prev/next buttons
// can update the modal without re-reading parsedRecipients.
let previewIndex        = 0;    // current row index in preview navigator
let previewRecipients   = [];   // filtered recipient set used by preview-all modal (A10)

// fieldMapping: maps canonical field names to the actual CSV column headers the
// user has in their file. E.g. { email: "Email Address", first_name: "Given Name" }.
// Populated by the Match Fields dialog. When empty (no mapping set), parseCSV()
// assumes the CSV headers exactly match the canonical keys.
let fieldMapping = {}; // { canonicalName: csvColumnName }

// greetingConfig: drives the {{greeting_line}} smart token. `format` is one of
// the predefined greeting patterns (e.g. "dear_sal_last" → "Dear Mr Smith"),
// and `fallback` is used when the required name fields are missing.
// Persisted to localStorage under LS_KEY_GREETING.
let greetingConfig = { format: "dear_sal_last", fallback: "Dear Valued Customer", customTemplate: "" };

// draftsMode: when true, buildEmailRequest() changes the Graph endpoint from
// /sendMail to /messages (creates a draft) instead of sending immediately.
// Useful for large sends that need manual review before delivery.
let draftsMode = false;

// selectedRowIndices: controls which rows participate in the merge.
// null means "all rows" (the common case). A Set of parsedRecipients array
// indices means "only these rows" — set when the user checks specific rows in
// the recipient table, or when using preview-filtered send.
let selectedRowIndices = null; // null = all selected; Set of original parsedRecipients indices when subset

// editTableHeaders / editTableRows: working copies used while the inline
// recipient-table editor is open. We edit copies rather than parsedRecipients
// directly so the user can cancel without corrupting the working list.
let editTableHeaders = [];   // working copy of column names during an edit session
let editTableRows    = [];   // working copy: array of string arrays (one per recipient)

// ── Contacts/Groups/Directory tab state ───────────────────────────────────
// These three variables hold the data and selection state for the import-from-
// contacts panel. Kept at module scope because the tab can be refreshed without
// losing the user's checkbox selections.

// groupsData: array of contact group objects returned by the Graph contacts API.
let groupsData = [];

// selectedGroups: Set of group IDs (Graph object IDs) the user has ticked.
// Used by importSelectedGroups() to decide which groups to expand into recipients.
let selectedGroups = new Set(); // group IDs

// contactsActiveTab: which sub-tab is currently visible in the contacts panel.
// "contacts" | "groups" | "directory" — drives CSS visibility and button state.
let contactsActiveTab = "contacts"; // "contacts" | "groups" | "directory"

// ── Directory tab state ────────────────────────────────────────────────────
// directoryData: results of the most recent Graph People / GAL search.
let directoryData = [];

// selectedDirectory: Set of email addresses the user has ticked in the directory
// results list. Used by importSelectedDirectory() to append rows to the CSV.
let selectedDirectory = new Set();

// ── License check ─────────────────────────────────────────────────────────
/**
 * Checks whether the signed-in user's tenant has an active subscription.
 * Result is cached in localStorage for 24 hours (matching the JWT expiry on the server).
 *
 * When LICENSE_ENFORCEMENT is false this function runs and logs silently — it does not
 * block any features. Flip LICENSE_ENFORCEMENT = true when ready to gate the product.
 */
async function checkLicense() {
  // ── Step 1: check the local cache ────────────────────────────────────────
  // Before making any network request, check whether we already have a valid
  // license result cached in localStorage. The cache key holds the license token
  // and the expiry key holds a Unix timestamp (ms). If both exist and the
  // timestamp hasn't expired, we trust the cached result and skip the API call.
  // parseInt with radix 10 is explicit best-practice — always pass the radix
  // to avoid surprising behaviour if the string starts with "0x" or "0".
  const cachedToken  = localStorage.getItem(LICENSE_CACHE_KEY);
  const cachedExpiry = localStorage.getItem(LICENSE_CACHE_EXPIRY_KEY);
  if (cachedToken && cachedExpiry && Date.now() < parseInt(cachedExpiry, 10)) {
    log("License: valid (cached)", "info");
    return; // Exit early — no need to hit the API
  }

  // ── Step 2: read user profile from the Office mailbox object ─────────────
  // Office.context.mailbox.userProfile gives us the signed-in user's email
  // address without needing a Graph call. This is available as soon as
  // Office.onReady fires, which has already happened before checkLicense runs.
  let userId   = null;
  let tenantId = null;
  let email    = null;
  try {
    const account = Office.context.mailbox.userProfile;
    // Guard against null — on some Outlook versions userProfile may not be populated
    email = account && account.emailAddress ? account.emailAddress : null;
    // userId and tenantId come from the JWT token claims below, not from this object
    const idToken = Office.context.auth && Office.context.auth.getAccessTokenAsync
      ? null  // Will be populated below via getAccessTokenAsync
      : null;
  } catch (e) {
    // Non-fatal: we still attempt the SSO path below, which gives us the claims
    log("License: could not read Office profile — " + e.message, "warn");
  }

  // ── Step 3: get Office SSO token and extract user identity claims ─────────
  // Office.auth.getAccessToken returns a JWT signed by Microsoft. We decode
  // the middle (payload) segment to extract:
  //   tid = tenant ID (identifies the organisation's Microsoft 365 tenant)
  //   oid = object ID (uniquely identifies the user within Entra ID)
  // These two values together uniquely identify the paying customer.
  // We use allowSignInPrompt: false so this runs silently in the background —
  // we never want the license check to interrupt the user with a sign-in popup.
  // If SSO isn't available (code 13000, etc.) we just skip the check rather
  // than blocking the user — the add-in fails open, not closed.
  let ssoToken = null;
  try {
    await new Promise((resolve) => {
      Office.auth.getAccessToken({ allowSignInPrompt: false, allowConsentPrompt: false }, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          ssoToken = result.value;
          // A JWT has three dot-separated Base64url-encoded segments:
          //   header.payload.signature
          // We only need the payload (index [1]). Base64url uses - and _ instead
          // of + and /, so we swap those before calling atob() which expects
          // standard Base64. JSON.parse then turns the decoded string into an object.
          try {
            const payload = JSON.parse(atob(ssoToken.split(".")[1].replace(/-/g, "+").replace(/_/g, "/")));
            tenantId = payload.tid || null; // Tenant ID — identifies the organisation
            userId   = payload.oid || null; // Object ID — identifies the individual user
          } catch { /* non-fatal — payload may not decode on some token issuers */ }
        }
        resolve(); // Always resolve so the outer await continues regardless of success/failure
      });
    });
  } catch (e) {
    // SSO can fail in many environments (legacy Outlook, IMAP-only, etc.) — non-fatal
    log("License: SSO token unavailable — " + e.message, "warn");
  }

  // If we couldn't identify the user, skip the check entirely.
  // The add-in continues to work — we never block on a failed identity lookup.
  if (!userId || !tenantId) {
    log("License: could not determine user identity, skipping check.", "warn");
    return;
  }

  // ── Step 4: call the remote license API ──────────────────────────────────
  // URLSearchParams builds a properly encoded query string from key-value pairs.
  // Using it instead of string concatenation prevents parameter injection if
  // any of the values contain special characters like &, =, or #.
  try {
    const params = new URLSearchParams({ userId, tenantId });
    if (email) params.set("email", email); // Optional — helps with support lookup

    // Await the fetch. If the server is down, the catch block below handles it.
    const res = await fetch(`${LICENSE_API_URL}?${params.toString()}`);

    if (!res.ok) {
      // res.ok is true for 2xx status codes. Anything else (4xx, 5xx) is a server error.
      // We warn but never block — a flaky API server should never prevent email sends.
      log(`License API returned ${res.status} — treating as unlicensed.`, "warn");
      if (LICENSE_ENFORCEMENT) showLicenseGate("server_error");
      return;
    }

    // Parse the JSON response body. Expected shape:
    //   { licensed: bool, plan: string, token: string, expiresAt: ISO8601, reason?: string }
    const data = await res.json();

    if (data.licensed) {
      log(`License: active — plan: ${data.plan}`, "info");
      // Cache until the server-provided expiry so we don't hit the API on every startup.
      // new Date(data.expiresAt).getTime() converts the ISO string to a Unix timestamp in ms.
      localStorage.setItem(LICENSE_CACHE_KEY, data.token || "valid");
      localStorage.setItem(LICENSE_CACHE_EXPIRY_KEY, String(new Date(data.expiresAt).getTime()));
    } else {
      // Licensed = false — the server explicitly says this user/tenant isn't subscribed.
      log(`License: not active — reason: ${data.reason}`, "warn");
      if (LICENSE_ENFORCEMENT) showLicenseGate(data.reason);
    }
  } catch (err) {
    log("License check network error (non-fatal): " + err.message, "warn");
    // Never block the add-in on a network error — fail open
  }
}

/**
 * Shows a non-dismissible overlay when LICENSE_ENFORCEMENT is true and the user is unlicensed.
 * @param {string} reason - "no_subscription" | "suspended" | "server_error"
 */
function showLicenseGate(reason) {
  const messages = {
    no_subscription: {
      title: "Subscription required",
      body: "Mail Merge requires an active subscription. Purchase one through Microsoft AppSource to continue.",
      link: "https://appsource.microsoft.com"
    },
    suspended: {
      title: "Subscription suspended",
      body: "Your subscription is currently suspended, likely due to a payment issue. Please update your payment method in Microsoft 365 Admin Center.",
      link: "https://admin.microsoft.com"
    },
    server_error: {
      title: "License check unavailable",
      body: "We could not verify your subscription right now. Please try again later or contact support.",
      link: null
    }
  };

  const m = messages[reason] || messages["no_subscription"];

  const gate = document.createElement("div");
  gate.style.cssText = [
    "position:fixed", "inset:0", "z-index:99999",
    "background:rgba(243,242,251,0.97)", "backdrop-filter:blur(8px)",
    "display:flex", "align-items:center", "justify-content:center",
    "padding:24px"
  ].join(";");

  const card = document.createElement("div");
  card.style.cssText = [
    "background:#fff", "border-radius:20px", "padding:36px 32px",
    "max-width:360px", "width:100%", "text-align:center",
    "box-shadow:0 8px 40px rgba(108,98,212,0.16)"
  ].join(";");

  const icon = document.createElement("div");
  icon.style.cssText = "font-size:44px;margin-bottom:14px";
  icon.textContent = reason === "suspended" ? "⚠️" : "🔒";

  const h2 = document.createElement("h2");
  h2.style.cssText = "font-size:18px;font-weight:700;color:#1C1C1E;margin-bottom:10px";
  h2.textContent = m.title;

  const p = document.createElement("p");
  p.style.cssText = "font-size:14px;color:#6B6B6B;line-height:1.6;margin-bottom:20px";
  p.textContent = m.body;

  card.appendChild(icon);
  card.appendChild(h2);
  card.appendChild(p);

  if (m.link) {
    const btn = document.createElement("a");
    btn.href = m.link;
    btn.target = "_blank";
    btn.rel = "noopener noreferrer";
    btn.style.cssText = [
      "display:inline-block",
      "background:linear-gradient(135deg,#6C62D4,#534AB7)",
      "color:#fff", "border-radius:99px", "padding:10px 28px",
      "font-size:14px", "font-weight:600", "text-decoration:none"
    ].join(";");
    btn.textContent = reason === "suspended" ? "Manage subscription" : "Get a subscription";
    card.appendChild(btn);
  }

  gate.appendChild(card);
  document.body.appendChild(gate);
}
// ──────────────────────────────────────────────────────────────────────────

// ── Office.onReady ────────────────────────────────────────────────────────
// Office.onReady is the entry point for every Office Add-in. It fires once the
// Office.js library has finished loading and established a connection to the host
// application (Outlook). All DOM manipulation and event-listener registration
// must happen inside this callback — manipulating the DOM before this fires can
// produce "API not initialised" errors in some Outlook versions.
//
// The `info` object tells us which host and platform we're on. We guard on
// info.host === Office.HostType.Outlook so the script doesn't run if it's
// accidentally loaded in Excel or Word.
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    log("Office.js ready. Host: Outlook.", "info");

    // License check — fire-and-forget. The .catch() here ensures a network
    // failure during license validation never crashes the Office.onReady handler
    // and never blocks the rest of the add-in from initialising.
    // (runs silently; enforcement controlled by LICENSE_ENFORCEMENT flag)
    checkLicense().catch(err => log("License check failed (non-fatal): " + err.message, "warn"));

    // C1 / C2: Check for Exchange Online mailbox — restUrl is null for IMAP/Gmail-only accounts.
    // Office.context.mailbox.restUrl is only populated when the account is hosted in
    // Exchange Online. If it's null, this is an IMAP/Google account and Graph API calls
    // will all fail with 401 — show a clear banner rather than letting the user click
    // "Send" and get a cryptic error message.
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
           style="color:#fff;text-decoration:underline;" target="_blank" rel="noopener noreferrer">Learn more about Exchange Online requirements</a>`;
      // Insert at the very top of the page so it's the first thing the user sees.
      document.body.insertBefore(banner, document.body.firstChild);
      // Disable all send/draft buttons — they would just fail with 401 anyway.
      ["mergeBtn","testSendBtn","previewAllBtn","saveDraftsBtn","broadcastBtn"].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.disabled = true;
      });
    }

    // ── Core action buttons ────────────────────────────────────────────────
    document.getElementById("mergeBtn").addEventListener("click", handleMergeClick);
    document.getElementById("stopBtn").addEventListener("click", handleStop);
    // previewBtn may not exist in all versions of the HTML, so guard with if.
    const previewBtn = document.getElementById("previewBtn");
    if (previewBtn) previewBtn.addEventListener("click", parseAndPreview);

    // ── Optional columns disclosure triangle ──────────────────────────────
    // The optional-columns list (CC, BCC, reply_to, etc.) is hidden by default
    // to keep the CSV help section concise. The toggle button reveals it.
    // We handle both click and keyboard (Enter/Space) for accessibility.
    const csvOptionalToggle = document.getElementById("csvOptionalToggle");
    const csvOptionalList   = document.getElementById("csvOptionalList");
    if (csvOptionalToggle && csvOptionalList) {
      function toggleCsvOptional() {
        // Read the CURRENT state from the DOM, not a variable — the DOM is
        // the source of truth for toggle state so it stays in sync even if
        // something else changes the class externally.
        const open = !csvOptionalList.classList.contains("hidden");
        csvOptionalList.classList.toggle("hidden", open);
        csvOptionalToggle.textContent = open ? "Optional columns ▾" : "Optional columns ▴";
        // ARIA: aria-expanded must track the open state for screen readers.
        csvOptionalToggle.setAttribute("aria-expanded", String(!open));
      }
      csvOptionalToggle.addEventListener("click", toggleCsvOptional);
      csvOptionalToggle.addEventListener("keydown", e => {
        // Space and Enter are the ARIA-compliant activation keys for a button.
        if (e.key === "Enter" || e.key === " ") { e.preventDefault(); toggleCsvOptional(); }
      });
    }

    // ── Log controls ──────────────────────────────────────────────────────
    document.getElementById("clearLogBtn").addEventListener("click", clearLog);

    // ── Confirmation modal ─────────────────────────────────────────────────
    // The modal has two exit paths: confirm (proceed with send) and cancel.
    // confirmSend() is the real merge trigger — handleMergeClick() just opens
    // the modal. This two-step prevents accidental large sends.
    document.getElementById("confirmSendBtn").addEventListener("click", confirmSend);
    document.getElementById("confirmCancelBtn").addEventListener("click", dismissModal);

    // ── CSV upload ─────────────────────────────────────────────────────────
    // The actual <input type="file"> is hidden; the button triggers a click on
    // it programmatically. This pattern gives us full control over the button's
    // appearance without fighting browser default file-input styling.
    document.getElementById("uploadCsvBtn").addEventListener("click", () => {
      document.getElementById("csvFileInput").click();
    });
    document.getElementById("csvFileInput").addEventListener("change", handleCsvFileUpload);

    // ── Shared attachment upload ───────────────────────────────────────────
    document.getElementById("uploadAttachmentBtn").addEventListener("click", () => {
      document.getElementById("attachmentFileInput").click();
    });
    document.getElementById("attachmentFileInput").addEventListener("change", handleAttachmentUpload);
    document.getElementById("clearSharedAttachmentsBtn").addEventListener("click", clearSharedAttachments);

    // ── Per-recipient attachment upload ────────────────────────────────────
    // perRecipientFilesInput accepts a folder (or multiple files via multi-select).
    // The "change" handler populates the perRecipientFiles Map keyed by filename.
    document.getElementById("uploadPerRecipientBtn").addEventListener("click", () => {
      document.getElementById("perRecipientFilesInput").click();
    });
    document.getElementById("perRecipientFilesInput").addEventListener("change", handlePerRecipientFilesUpload);
    document.getElementById("clearPerRecipientBtn").addEventListener("click", clearPerRecipientFiles);

    // ── Inline images upload ────────────────────────────────────────────────
    document.getElementById("uploadInlineImagesBtn").addEventListener("click", () => {
      document.getElementById("inlineImagesInput").click();
    });
    document.getElementById("inlineImagesInput").addEventListener("change", handleInlineImagesUpload);
    document.getElementById("clearInlineImagesBtn").addEventListener("click", clearInlineImages);

    // ── Scheduling accordion toggle ────────────────────────────────────────
    // The "schedule" row is hidden until the checkbox is checked. We also
    // update the scheduling badge (a visual indicator in the tab) so the user
    // can see at a glance that scheduling is active even when the accordion is
    // collapsed.
    document.getElementById("scheduleEnabled").addEventListener("change", (e) => {
      document.getElementById("scheduleRow").classList.toggle("hidden", !e.target.checked);
      updateSchedulingBadge();
    });

    // ── Message expiry toggle ──────────────────────────────────────────────
    document.getElementById("expiryEnabled").addEventListener("change", (e) => {
      document.getElementById("expiryRow").classList.toggle("hidden", !e.target.checked);
    });

    // ── Sending window toggle ──────────────────────────────────────────────
    // The sending-window feature restricts sends to business hours. The row
    // with the time/day controls is hidden until the feature is enabled.
    const sendingWindowEnabledEl = document.getElementById("sendingWindowEnabled");
    if (sendingWindowEnabledEl) {
      sendingWindowEnabledEl.addEventListener("change", function(e) {
        const sendingWindowRowEl = document.getElementById("sendingWindowRow");
        if (sendingWindowRowEl) sendingWindowRowEl.classList.toggle("hidden", !e.target.checked);
        updateSchedulingBadge();
      });
    }

    // ── Custom X- headers toggle ───────────────────────────────────────────
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

    // ── Batch delay range input ────────────────────────────────────────────
    // The batchDelayInput slider controls how long to pause between batch
    // requests. We update a human-readable label on every "input" event
    // (not just "change") so the user sees feedback as they drag the slider.
    const batchDelayInput = document.getElementById("batchDelayInput");
    batchDelayInput.addEventListener("input", () => {
      const sec = parseFloat(batchDelayInput.value);
      document.getElementById("batchDelayLabel").textContent =
        sec === 0 ? "(no delay between batches)" : `(${sec} s between batches)`;
    });

    // ── Subject / Body tag-target toggle ──────────────────────────────────
    // Helper: set tagTarget and keep the Body/Subject toggle pills in sync.
    // This is extracted into a named function rather than inlined because
    // multiple event handlers (focus, blur, window blur, button click) all
    // need to change the target — DRY principle.
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

    // ── New line button in tag bar ─────────────────────────────────────────
    // Inserts a line break into the Outlook body from the taskpane, so the
    // user doesn't have to click back into the body just to press Enter between
    // consecutive tag insertions.
    const newlineBtn = document.getElementById("insertNewlineBtn");
    if (newlineBtn) newlineBtn.addEventListener("click", insertNewline);

    // ── Subject input — localStorage sync + Outlook push ──────────────────
    // Every keystroke: save to localStorage immediately (so it survives
    // a task pane reload) AND debounce a push to Outlook's native subject
    // field. We debounce at 400 ms because subject.setAsync is an async
    // Office.js API call — calling it on every keystroke would queue up
    // hundreds of overlapping async operations.
    document.getElementById("subjectInput").addEventListener("input", () => {
      const val = document.getElementById("subjectInput").value;
      lsSet(LS_KEY_SUBJECT, val);
      // Push to Outlook's native subject field (debounced 400 ms)
      clearTimeout(_subjectPushTimer);
      _subjectPushTimer = setTimeout(() => pushSubjectToOutlook(val), 400);
    });

    // ── Sync subject button ────────────────────────────────────────────────
    // Manual "pull from Outlook" button. The boolean arg (true) means "show
    // a toast if the subject was updated" — useful for on-demand syncs where
    // the user wants visible confirmation.
    const syncSubjectBtn = document.getElementById("syncSubjectBtn");
    if (syncSubjectBtn) {
      syncSubjectBtn.addEventListener("click", () => syncSubjectFromOutlook(true));
    }

    // ── CSV textarea — localStorage autosave ──────────────────────────────
    // We listen to both "input" (handles typing) and "change" (handles
    // paste-via-context-menu and programmatic value changes that don't
    // fire "input"). Both are needed for full coverage.
    const _csvInputEl = document.getElementById("csvInput");
    const _saveCsv = () => lsSet(LS_KEY_CSV, _csvInputEl.value);
    _csvInputEl.addEventListener("input", _saveCsv);
    _csvInputEl.addEventListener("change", _saveCsv);

    // ── Restore previous session state ────────────────────────────────────
    // restoreLocalState() reads localStorage and populates the subject input,
    // CSV textarea, custom tags, and greeting config from the previous session.
    // This must run BEFORE syncSubjectFromOutlook so the UI reflects the saved
    // subject before we potentially overwrite it with Outlook's value.
    restoreLocalState();

    // ── Subject bidirectional sync ────────────────────────────────────────
    // Initial sync: pull the native Outlook subject into the taskpane on load.
    // Then poll every 500 ms so edits in Outlook's native subject box appear
    // in the taskpane in near-real-time. 500 ms is a compromise — fast enough
    // to feel responsive, slow enough not to overwhelm the Office.js bridge.
    syncSubjectFromOutlook(true);
    _subjectPollTimer = setInterval(() => syncSubjectFromOutlook(false), 500);

    // ── Initialise UI subsystems ───────────────────────────────────────────
    initTabs();       // Set up the tab-switching logic (Recipients / Compose / Options)
    initAccordions(); // Set up the expand/collapse accordion sections
    initOptOutUI();   // Load suppressionSet from localStorage and render the opt-out list

    // ── Footer log toggle ──────────────────────────────────────────────────
    // The footer shows a one-line "mini log" by default; clicking the toggle
    // expands the full scrollable log panel. Both panels hide when the other shows.
    document.getElementById("toggleLogBtn").addEventListener("click", () => {
      const mini = document.getElementById("footerLogMini");
      const full = document.getElementById("fullLogArea");
      const open = !full.classList.contains("hidden");
      full.classList.toggle("hidden", open);
      mini.classList.toggle("hidden", open);
    });

    // ── Contacts import modal ──────────────────────────────────────────────
    // handleImportContacts() fetches contacts from Graph and renders a
    // checkbox list in the modal. The search input filters the rendered list
    // client-side (no re-fetch) using filterContacts().
    document.getElementById("importContactsBtn").addEventListener("click", handleImportContacts);
    document.getElementById("contactsCloseBtn").addEventListener("click", () => {
      document.getElementById("contactsModal").classList.add("hidden");
    });
    document.getElementById("contactsSearch").addEventListener("input", (e) => filterContacts(e.target.value));
    // Select-all button: if all visible checkboxes are already checked, uncheck
    // all (toggle-off); otherwise check all. Uses spread to convert NodeList.
    document.getElementById("contactsSelectAllBtn").addEventListener("click", () => {
      const visible = document.getElementById("contactsList").querySelectorAll("input[type=checkbox]");
      const allChecked = [...visible].every(cb => cb.checked);
      visible.forEach(cb => {
        cb.checked = !allChecked;
        // Maintain the selectedContacts Set alongside checkbox state.
        if (!allChecked) selectedContacts.add(cb.dataset.email);
        else selectedContacts.delete(cb.dataset.email);
      });
      updateContactsSelectedCount();
    });
    // The import button serves both contacts and directory — delegate based on
    // which sub-tab is currently active.
    document.getElementById("contactsImportBtn").addEventListener("click", () => {
      if (contactsActiveTab === "directory") importDirectorySelected();
      else importSelectedContacts();
    });

    // ── Recipient filter popup ─────────────────────────────────────────────
    // Apply and Clear buttons live inside the popup; both close it when done.
    document.getElementById("applyFilterBtn").addEventListener("click", applyFilterSort);
    document.getElementById("clearFilterBtn").addEventListener("click", clearFilterSort);
    // Close button (✕) in the popup header — just hides the popup, no filter change.
    document.getElementById("closeFilterPopupBtn").addEventListener("click", closeFilterPopup);
    // Clicking the backdrop (outside the panel) also closes the popup.
    document.getElementById("filterPopupBackdrop").addEventListener("click", closeFilterPopup);

    // ── Match Fields modal ─────────────────────────────────────────────────
    // Opens a dialog that maps the user's CSV column names to the canonical
    // field names (email, first_name, etc.) that the add-in expects.
    // We filter out internal "_"-prefixed keys (like _originalIndex) from
    // the header list — those are implementation details, not real columns.
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

    // ── Body template save/load ────────────────────────────────────────────
    document.getElementById("saveBodyToTemplateBtn").addEventListener("click", saveBodyToTemplate);
    document.getElementById("loadBodyFromTemplateBtn").addEventListener("click", loadBodyFromTemplate);

    // ── Greeting-line configuration ────────────────────────────────────────
    // Persist the greeting config to localStorage on every change so it
    // survives page refreshes. The "change" event is used for the select
    // (fires when the user picks an option) and "input" for the text fields.
    document.getElementById("greetingFormat").addEventListener("change", () => {
      greetingConfig.format = document.getElementById("greetingFormat").value;
      // Show the custom template input only when "Custom template…" is selected.
      // Hidden for all preset formats since they use fixed column mappings.
      const isCustom = greetingConfig.format === "custom";
      document.getElementById("greetingCustomWrap").classList.toggle("hidden", !isCustom);
      lsSet(LS_KEY_GREETING, JSON.stringify(greetingConfig));
    });
    document.getElementById("greetingFallback").addEventListener("input", () => {
      // Fall back to the default string rather than saving an empty fallback,
      // which would result in blank greeting lines for recipients missing names.
      greetingConfig.fallback = document.getElementById("greetingFallback").value.trim() || "Dear Valued Customer";
      lsSet(LS_KEY_GREETING, JSON.stringify(greetingConfig));
    });
    // Custom greeting template — free-text field shown when format === "custom".
    // Supports any {{column}} token from the loaded CSV, e.g. "Dear {{title}} {{last_name}},"
    document.getElementById("greetingCustomTemplate").addEventListener("input", () => {
      greetingConfig.customTemplate = document.getElementById("greetingCustomTemplate").value;
      lsSet(LS_KEY_GREETING, JSON.stringify(greetingConfig));
    });

    // ── Save as Drafts mode ────────────────────────────────────────────────
    document.getElementById("saveDraftsBtn").addEventListener("click", handleSaveDrafts);

    // ── Check for Errors pre-flight ────────────────────────────────────────
    // Opens a modal that runs pre-send validation checks without sending
    // anything: required fields, duplicate emails, attachment sizes, etc.
    document.getElementById("checkErrorsBtn").addEventListener("click", handleCheckErrors);
    document.getElementById("checkErrorsCloseBtn").addEventListener("click", () => {
      // _closeModalWithTrap restores focus to the previously focused element
      // (accessibility requirement — modals must return focus on close).
      _closeModalWithTrap("checkErrorsModal"); // Feature 15
    });

    // ── Edit recipient table modal ─────────────────────────────────────────
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

    // ── Contacts modal sub-tabs ────────────────────────────────────────────
    // The contacts modal has three sub-tabs: Contacts (personal address book),
    // Groups (contact groups / distribution lists), and Directory (GAL / People API).
    // Switching tabs updates contactsActiveTab and shows/hides the appropriate
    // panels. We manipulate style.display directly (rather than class toggling)
    // because we need precise control over which elements are visible in each tab.
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

    // ── Directory search — debounced ───────────────────────────────────────
    // We debounce the search at 400 ms to avoid firing a Graph API request on
    // every keystroke. The local timer variable is scoped inside onReady so
    // it doesn't pollute the module scope.
    let directorySearchTimer = null;
    document.getElementById("directorySearch").addEventListener("input", (e) => {
      clearTimeout(directorySearchTimer);
      directorySearchTimer = setTimeout(() => searchDirectory(e.target.value.trim()), 400);
    });

    // ── Keyboard: Enter/Space activates tag chips ──────────────────────────
    // Tag chips have tabIndex=0 so they're focusable, but they're <span>
    // elements, not <button>s — browsers don't fire click events for Enter/Space
    // on non-interactive elements. We polyfill that behaviour here globally.
    // The check for e.target.classList.contains("tag") is intentionally broad —
    // it catches both DEFAULT_TAGS chips and user-defined custom tag chips.
    document.addEventListener("keydown", (e) => {
      if ((e.key === "Enter" || e.key === " ") && e.target.classList.contains("tag")) {
        e.preventDefault();
        e.target.click();
      }
    });

    // ── Keyboard: Escape closes topmost visible modal ──────────────────────
    // This single listener handles ALL modals rather than each modal adding its
    // own Escape handler, which would create dozens of competing listeners.
    // We iterate the list in display-priority order and close only the first
    // visible one — prevents accidentally closing two modals at once.
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
          // Prefer clicking the cancel/close button so it runs any cleanup logic
          // (e.g. _closeModalWithTrap restoring focus). Fall back to just hiding
          // the element if no cancel button is found.
          const cancelBtn = el.querySelector(
            "#fillInCancelBtn, #preSendCancelBtn, #confirmCancelBtn, #dupeSendCancelBtn, " +
            "#checkErrorsCloseBtn, #previewCloseBtn, #matchFieldsCancelBtn, #contactsCloseBtn, " +
            "#editTableCloseBtn, " +
            "button[class*='cancel'], button[class*='close']"
          );
          if (cancelBtn) cancelBtn.click();
          else el.classList.add("hidden");
          break; // Only close the topmost modal
        }
      }
    });

    // ── CSV clear → reset retry state ─────────────────────────────────────
    // When the user manually clears the CSV textarea, the failedRecipients list
    // is no longer valid (it references rows from the previous CSV). Hide the
    // retry button and reset the list to prevent a stale retry from sending to
    // the wrong people.
    document.getElementById("csvInput")?.addEventListener("input", () => {
      if (!document.getElementById("csvInput").value.trim()) {
        document.getElementById("retryFailedBtn")?.classList.add("hidden");
        failedRecipients = [];
      }
    });

    // ── Smart paste detection ──────────────────────────────────────────────
    // When a user pastes tab-separated content (Excel/Sheets copy), automatically
    // convert it to CSV so they don't need to export from Excel first.
    // We guard on "majority of lines contain tabs" (≥50%) to avoid converting
    // normal text that happens to contain a single tab.
    document.getElementById("csvInput")?.addEventListener("paste", (e) => {
      const raw = (e.clipboardData || window.clipboardData)?.getData("text");
      if (!raw) return;
      const lines = raw.split(/\r?\n/);
      // Tab-separated if the majority of lines contain tabs.
      const tabLines = lines.filter(l => l.includes("\t")).length;
      if (tabLines < Math.max(1, lines.length * 0.5)) return; // not tab-separated — leave paste unmodified
      e.preventDefault(); // We'll insert the converted CSV ourselves
      // RFC 4180 quoting: fields containing commas, quotes, or newlines must be
      // double-quoted, with internal quotes escaped by doubling them ("").
      const csv = lines.map(line =>
        line.split("\t").map(cell => {
          if (cell.includes(",") || cell.includes('"') || cell.includes("\n")) {
            return '"' + cell.replace(/"/g, '""') + '"';
          }
          return cell;
        }).join(",")
      ).join("\n");
      // Insert at cursor position (replacing any selection) rather than
      // appending to the end, so paste works naturally at any cursor location.
      const ta = document.getElementById("csvInput");
      const start = ta.selectionStart, end = ta.selectionEnd;
      ta.value = ta.value.slice(0, start) + csv + ta.value.slice(end);
      ta.selectionStart = ta.selectionEnd = start + csv.length;
      // Dispatch both "input" and "change" to ensure the autosave handler
      // and any other listeners pick up the new value.
      ta.dispatchEvent(new Event("input"));
      ta.dispatchEvent(new Event("change"));
      log("Detected Excel/Sheets table paste — converted to CSV automatically.", "info");
    });

    // ── Drag-and-drop CSV / Excel onto the drop zone ───────────────────────
    // The drop zone is an overlay on the CSV textarea. We intercept dragover to
    // show visual feedback, and on drop we inject the file into the hidden
    // csvFileInput and trigger its "change" handler — reusing all the existing
    // file-type detection and parsing code without duplication.
    const csvDropZone = document.getElementById("csvDropZone");
    if (csvDropZone) {
      csvDropZone.addEventListener("dragover", (e) => {
        e.preventDefault(); // Must preventDefault to allow drop
        csvDropZone.classList.add("drag-over"); // CSS highlights the drop target
      });
      csvDropZone.addEventListener("dragleave", () => {
        csvDropZone.classList.remove("drag-over");
      });
      csvDropZone.addEventListener("drop", (e) => {
        e.preventDefault();
        csvDropZone.classList.remove("drag-over");
        const file = e.dataTransfer?.files?.[0];
        if (!file) return;
        // DataTransfer API: the only cross-browser way to programmatically set
        // an <input type="file">'s file list. We can't assign .files directly.
        const dt = new DataTransfer();
        dt.items.add(file);
        const fi = document.getElementById("csvFileInput");
        fi.files = dt.files;
        fi.dispatchEvent(new Event("change")); // Trigger the upload handler
      });
    }

    // ── Quick email entry mode ─────────────────────────────────────────────
    // Quick mode replaces the CSV textarea with a simple one-email-per-line input.
    // toggleQuickMode() switches the visible panel and updateQuickModeCSV()
    // converts the plain text into a single-column CSV in the background.
    document.getElementById("quickModeBtn")?.addEventListener("click", toggleQuickMode);
    document.getElementById("quickEmailInput")?.addEventListener("input", () => {
      updateQuickModeCSV(); // Convert current lines to CSV and update parsedRecipients
    });

    // ── Import recipients from Outlook's native To field ──────────────────
    // syncRecipientsFromOutlook() reads the compose To recipients via
    // Office.js toRecipients.getAsync, converts them to CSV rows, and clears
    // Outlook's To field so the merge engine can repopulate it per-recipient.
    document.getElementById("syncRecipientsBtn")?.addEventListener("click", () => {
      syncRecipientsFromOutlook();
    });

    // ── Rate limiting — restore persisted state ────────────────────────────
    // loadRateLimitState() reads maxPerHour and dailyCap from localStorage
    // so the user's limits survive page reloads.
    loadRateLimitState();

    // Whenever the user changes the rate limit inputs, update the badge
    // in the tab that shows the active limit.
    document.getElementById("maxPerHour")?.addEventListener("input", updateRateLimitBadge);
    document.getElementById("dailyCap")?.addEventListener("input", updateRateLimitBadge);

    // ── Sample CSV download ────────────────────────────────────────────────
    // Provides a minimal working example CSV so first-time users can see the
    // expected format without reading documentation. The URL.createObjectURL /
    // revokeObjectURL pattern is the browser-standard way to trigger a download
    // from in-memory data without hitting a server.
    document.getElementById("downloadSampleCsv")?.addEventListener("click", (e) => {
      e.preventDefault();
      const sampleCsv = [
        "email,first_name,last_name,company,title",
        "alice@example.com,Alice,Smith,Acme Corp,Director",
        "bob@example.com,Bob,Jones,Globex,Manager",
        "carol@example.com,Carol,Williams,Initech,Analyst"
      ].join("\r\n"); // CRLF line endings per RFC 4180
      const blob = new Blob([sampleCsv], { type: "text/csv" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "mail-merge-sample.csv";
      a.click(); // Programmatic click triggers the browser's Save dialog
      URL.revokeObjectURL(url); // Free the blob URL — important for memory hygiene
    });

    // ── Simulate / dry-run mode ────────────────────────────────────────────
    // handleSimulate() runs the full merge logic (personalisation, conditionals,
    // dedup, suppression check) but skips the Graph API call — useful for
    // validating a campaign before committing to a real send.
    document.getElementById("simulateBtn")?.addEventListener("click", handleSimulate);

    // ── Getting-started banner ─────────────────────────────────────────────
    // Only shown on first run (no welcome dismissal flag AND no saved CSV).
    // The flag is set when the user dismisses it, so it only shows once.
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
 * Show a brief toast notification for user-facing guidance messages.
 * Toasts slide in from the top and auto-dismiss after `duration` ms.
 * They are in addition to the detailed log (for support use).
 *
 * @param {string} message - Human-readable message to display
 * @param {"info"|"success"|"warning"|"error"} [type="info"] - Severity
 * @param {number} [duration=4000] - Auto-dismiss delay in ms (0 = never)
 */
function showToast(message, type, duration) {
  if (!type) type = "info";
  if (duration === undefined) duration = 4000;
  var icons = { info: "💡", success: "✓", warning: "⚠️", error: "✕" };
  var container = document.getElementById("toastContainer");
  if (!container) return;
  var toast = document.createElement("div");
  toast.className = "toast toast-" + type;
  toast.setAttribute("role", "alert");
  var iconSpan = document.createElement("span");
  iconSpan.className = "toast-icon";
  iconSpan.textContent = icons[type] || "💡";
  var msgSpan = document.createElement("span");
  msgSpan.className = "toast-msg";
  msgSpan.textContent = message;
  var closeBtn = document.createElement("button");
  closeBtn.className = "toast-close";
  closeBtn.setAttribute("aria-label", "Dismiss");
  closeBtn.textContent = "×";
  toast.appendChild(iconSpan);
  toast.appendChild(msgSpan);
  toast.appendChild(closeBtn);
  container.appendChild(toast);
  var dismiss = function() {
    if (toast.parentNode) {
      toast.classList.add("toast-hide");
      setTimeout(function() { if (toast.parentNode) toast.parentNode.removeChild(toast); }, 220);
    }
  };
  closeBtn.addEventListener("click", dismiss);
  if (duration > 0) setTimeout(dismiss, duration);
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

// _quickModeActive: tracks whether the add-in is currently in quick-entry mode.
// Module-scoped so toggleQuickMode() and updateQuickModeCSV() share the same flag.
let _quickModeActive = false;

/**
 * toggleQuickMode() — switch between CSV mode and quick email-entry mode.
 *
 * Quick mode shows a simple textarea where users type/paste one email per line.
 * Under the hood it still generates a CSV and calls parseAndPreview(), so the
 * rest of the merge pipeline is unchanged — this is purely a UX simplification.
 * The CSV area and quick area are mutually exclusive panels toggled here.
 */
function toggleQuickMode() {
  _quickModeActive = !_quickModeActive;
  const btn     = document.getElementById("quickModeBtn");
  const csvArea = document.getElementById("csvSection");
  const qArea   = document.getElementById("quickSection");

  if (_quickModeActive) {
    csvArea?.classList.add("hidden");    // hide the full CSV panel
    qArea?.classList.remove("hidden");  // show the simple email-entry panel
    if (btn) btn.textContent = "📋 Switch to CSV mode";
    document.getElementById("quickEmailInput")?.focus(); // put cursor in the input immediately
  } else {
    csvArea?.classList.remove("hidden");
    qArea?.classList.add("hidden");
    if (btn) btn.textContent = "📧 Quick: type emails";
  }
}

// _quickModeParseTimer: debounce handle for updateQuickModeCSV → parseAndPreview.
// Without debouncing, every keystroke in the quick-entry textarea would
// trigger a full CSV parse + DOM render, which is wasteful and can cause
// visible flicker on long lists.
let _quickModeParseTimer = null;

/**
 * updateQuickModeCSV() — convert quick-entry textarea content into a CSV and
 * trigger parseAndPreview().
 *
 * The function reads the raw quick-entry textarea, splits on newlines/commas/
 * semicolons, filters out anything that isn't a valid email address, then
 * synthesises a minimal single-column CSV ("email\n...") in the hidden csvInput
 * textarea. This keeps all downstream logic (parseCSV, validateRecipients, etc.)
 * working exactly as if the user had typed the CSV manually.
 */
function updateQuickModeCSV() {
  const raw = document.getElementById("quickEmailInput")?.value || "";
  // Accept newline-, comma-, or semicolon-separated addresses.
  // EMAIL_REGEX.test() filters out blank lines and any non-email entries.
  const emails = raw.split(/[\n,;]+/).map(e => e.trim()).filter(e => e && EMAIL_REGEX.test(e));
  const csvInput = document.getElementById("csvInput");
  if (!csvInput) return;

  if (emails.length === 0) {
    // Textarea cleared — wipe the recipient list and preview immediately.
    // We do NOT call parseAndPreview() here because an empty CSV would trigger
    // the Match Fields dialog, which is wrong in this context.
    csvInput.value = "";
    parsedRecipients = [];
    selectedRowIndices = null;
    const countEl = document.getElementById("recipientCount");
    if (countEl) countEl.textContent = "0 recipients";
    renderPreviewTable([]);
    return;
  }

  // Build a minimal single-column CSV with just the email header and addresses.
  csvInput.value = "email\n" + emails.join("\n");
  // Dispatch both events so the autosave handler and any other listeners fire.
  csvInput.dispatchEvent(new Event("input"));
  csvInput.dispatchEvent(new Event("change"));
  // Debounce the expensive parse — only run 350 ms after the user pauses typing.
  clearTimeout(_quickModeParseTimer);
  _quickModeParseTimer = setTimeout(() => parseAndPreview(), 350);
}

/* ─── SYNC RECIPIENTS FROM OUTLOOK'S TO FIELD ───────────────────
   Reads the native To field of the compose window and pre-populates
   the CSV with those addresses — great for small ad-hoc sends.
   ─────────────────────────────────────────────────────────────── */

/**
 * syncRecipientsFromOutlook() — import the native Outlook To-field addresses into
 * the recipient CSV and then clear Outlook's To field.
 *
 * Why clear the To field? Because the merge engine sends each email individually
 * via Graph — if the To field still contains all recipients, they'd also receive
 * the email as CC/TO on every other recipient's personalised copy. Clearing it
 * gives the merge engine a clean slate to set the single per-recipient To address.
 *
 * This function uses Office.js toRecipients.getAsync which is available in all
 * Outlook compose-mode contexts (desktop, web, mobile). The result is an array
 * of EmailAddressDetails objects: { emailAddress, displayName }.
 */
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
      // Build a two-column CSV with email and display_name columns.
      // We strip commas from display names because commas are the CSV delimiter —
      // they'd break parsing if left in an unquoted field.
      const csvLines = ["email,display_name"];
      recipients.forEach(r => {
        const addr = (r.emailAddress || "").trim();
        const name = (r.displayName || "").replace(/,/g, " ");
        if (addr) csvLines.push(`${addr},${name}`);
      });
      const csvInput = document.getElementById("csvInput");
      if (csvInput) {
        csvInput.value = csvLines.join("\n");
        // Fire both events to trigger autosave and any other change listeners.
        csvInput.dispatchEvent(new Event("input"));
        csvInput.dispatchEvent(new Event("change"));
        parseAndPreview(); // Parse the newly set CSV and update the recipient table
        log(`Imported ${recipients.length} address${recipients.length !== 1 ? "es" : ""} from Outlook's To field.`, "success");
      }
      // Clear Outlook's To field — setAsync([]) removes all recipients.
      // The empty callback is intentional: we don't care about the result.
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

  // Restore manually-added custom tag chips (non-default, non-smart tags).
  // On each startup, clear any tags that were persisted by older versions —
  // CSV column tags should never survive a taskpane reload because the CSV
  // data itself is in-memory only. Only tags the user explicitly types via
  // the custom tag input are saved going forward (persist=false for CSV tags).
  lsSet(LS_KEY_TAGS, JSON.stringify([]));
  // (savedTags intentionally empty after the clear — users re-add manual tags
  //  as needed; CSV column tags reappear automatically when CSV is reloaded.)

  renderTemplateList();  // rebuild the saved-template picker in the Compose tab

  // Restore greeting format and fallback text.
  // The stored value may be a JSON string or already-parsed object depending on
  // which version of lsGet was used when it was saved — handle both.
  const savedGreeting = lsGet(LS_KEY_GREETING, null);
  if (savedGreeting) {
    try {
      const g = typeof savedGreeting === "object" ? savedGreeting : JSON.parse(savedGreeting);
      greetingConfig = { customTemplate: "", ...g }; // ensure customTemplate always exists
      document.getElementById("greetingFormat").value          = g.format         || "dear_sal_last";
      document.getElementById("greetingFallback").value        = g.fallback       || "Dear Valued Customer";
      document.getElementById("greetingCustomTemplate").value  = g.customTemplate || "";
      // Show custom template input if the saved format was "custom".
      document.getElementById("greetingCustomWrap").classList.toggle("hidden", g.format !== "custom");
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

/* ─── TEMPLATE SAVE / LOAD ─────────────────────────────────────────────────
   Templates let users save their full compose configuration (subject, options,
   scheduling settings) under a name and restore it later. Templates are stored
   in localStorage as a plain object: { "Template Name": stateObject }.
   ─────────────────────────────────────────────────────────────────────────── */

/**
 * getTemplates() — load the templates object from localStorage.
 *
 * Returns a plain object where keys are template names and values are the
 * state snapshots produced by getTemplateState(). Returns an empty object
 * if nothing has been saved yet.
 *
 * @returns {Object} { [name: string]: templateState }
 */
function getTemplates() {
  return JSON.parse(lsGet(LS_KEY_TEMPLATES) || "{}");
}

/**
 * getTemplateState() — snapshot all saveable UI settings into a plain object.
 *
 * This is called when the user saves a template. Every property here corresponds
 * to a UI element that applyTemplateState() must restore. Keep these two functions
 * in sync — if you add a new option to the add-in, add it here AND in
 * applyTemplateState() or it won't be saved/restored.
 *
 * @returns {Object} A complete snapshot of the current UI state
 */
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
    // Scheduling and rate-limit settings — use ?. and || "" fallbacks because
    // these elements may not exist in older HTML versions.
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

  // Auto-expand the scheduling accordion if any scheduling setting was saved as active.
  // It's confusing to restore a checked "Enable scheduling" checkbox inside a
  // collapsed accordion — the user can't see that scheduling is now active.
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
  // Similarly expand the advanced accordion if a rate limit was restored.
  // parseInt with radix 10: a stored "0" should evaluate to 0, not NaN.
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

  // Refresh the scheduling and rate-limit badges to reflect the restored values.
  updateSchedulingBadge();
  updateRateLimitBadge();
}

/**
 * saveTemplate() — save the current UI state under the name in #templateNameInput.
 * Uses getTemplateState() to capture all relevant settings, then serialises
 * the entire templates object to localStorage.
 */
function saveTemplate() {
  const name = document.getElementById("templateNameInput").value.trim();
  if (!name) { log("Enter a template name.", "warning"); return; }
  const templates = lsGet(LS_KEY_TEMPLATES, {});
  templates[name] = getTemplateState();
  lsSet(LS_KEY_TEMPLATES, JSON.stringify(templates));
  populateTemplateSelect();
  // Pre-select the newly saved name in the dropdown so the user can immediately
  // load or delete it without scrolling.
  document.getElementById("templateSelect").value = name;
  document.getElementById("deleteTemplateBtn").disabled = false;
  log(`Template "${name}" saved.`, "success");
}

/**
 * loadTemplate() — restore the template selected in #templateSelect.
 * Calls applyTemplateState() to push all saved values into the UI.
 */
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

/**
 * deleteTemplate() — remove the selected template from localStorage.
 * Re-renders the template list so the deleted name disappears from the dropdown.
 */
function deleteTemplate() {
  const name = document.getElementById("templateSelect").value;
  if (!name) { log("Select a template to delete.", "error"); return; }
  const templates = getTemplates();
  delete templates[name]; // Remove the key from the object before re-serialising
  lsSet(LS_KEY_TEMPLATES, JSON.stringify(templates));
  renderTemplateList();
  log(`Template "${name}" deleted.`, "info");
}

/**
 * renderTemplateList() — rebuild the #templateSelect dropdown from localStorage.
 *
 * Escapes template names before inserting into innerHTML to prevent XSS via
 * a maliciously crafted template name.
 */
function renderTemplateList() {
  const sel = document.getElementById("templateSelect");
  const templates = getTemplates();
  const names = Object.keys(templates);
  sel.innerHTML = `<option value="">— load template —</option>` +
    names.map(n => `<option value="${escapeHtml(n)}">${escapeHtml(n)}</option>`).join("");
  // Disable delete button when no templates exist — nothing to delete.
  document.getElementById("deleteTemplateBtn").disabled = names.length === 0;
}

// populateTemplateSelect: alias kept for backwards compatibility.
// Earlier versions called this name; newer code uses renderTemplateList() directly.
function populateTemplateSelect() { renderTemplateList(); }

/* ─── MATCH FIELDS DIALOG ───────────────────────────────────────────────────
   The Match Fields dialog lets users map their CSV column names to the canonical
   field names the add-in understands. For example: a CSV with "Email Address"
   maps to the canonical "email" key. Without this, users with non-standard CSV
   headers would get no personalisation — every {{first_name}} would be blank.
   ─────────────────────────────────────────────────────────────────────────── */

/**
 * openMatchFieldsModal(csvHeaders) — render a row of label + <select> for each
 * canonical field, pre-selecting the best auto-match candidate.
 *
 * Builds the modal body purely in memory as an HTML string, then sets
 * innerHTML in one operation — much faster than appending DOM nodes one by one
 * when there are 14 rows. escapeHtml() on every user-supplied value prevents
 * XSS via malicious CSV header names.
 *
 * @param {string[]} csvHeaders - Column headers from the parsed CSV
 */
function openMatchFieldsModal(csvHeaders) {
  const body = document.getElementById("matchFieldsBody");
  body.innerHTML = CANONICAL_FIELDS.map(f => {
    const currentMapping = fieldMapping[f.key] || "";
    const options = ['<option value="">— skip —</option>']
      .concat(csvHeaders.map(h => {
        // Auto-select if: (1) the user hasn't manually mapped this field yet,
        // AND (2) this header matches one of the known synonyms.
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

/**
 * autoMatchField(canonicalKey, csvHeader) — fuzzy-match a CSV column header to
 * a canonical field name by checking against a synonym list.
 *
 * Normalises both keys by lowercasing and stripping spaces, hyphens, and
 * underscores before comparing — so "First Name", "first_name", and "firstname"
 * all match the canonical "first_name" field.
 *
 * @param {string} canonicalKey - A CANONICAL_FIELDS key (e.g. "email", "first_name")
 * @param {string} csvHeader    - A raw column header from the user's CSV
 * @returns {boolean} true if the header is a recognised synonym for this field
 */
function autoMatchField(canonicalKey, csvHeader) {
  // Normalise: lowercase + strip all non-alpha characters to produce a bare key.
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

/**
 * applyMatchFields() — read the user's mapping selections from the modal and
 * store them in the fieldMapping object, then re-parse the CSV.
 *
 * Validates that the "email" field has been mapped before closing the modal —
 * without it, the merge can't address any emails.
 */
function applyMatchFields() {
  const selects = document.getElementById("matchFieldsBody").querySelectorAll("select");
  const newMapping = {};
  selects.forEach(sel => {
    // data-canonical holds the canonical key; sel.value holds the CSV column name.
    if (sel.value) newMapping[sel.dataset.canonical] = sel.value;
  });
  if (!newMapping.email) {
    log("Match Fields: 'Email' column is required — please map it.", "warning");
    showToast("Please map a column to Email before applying.", "warning");
    return;
  }
  fieldMapping = newMapping;
  document.getElementById("matchFieldsModal").classList.add("hidden");
  log(`Field mapping applied. Email column: "${fieldMapping.email}".`, "success");
  // Re-parse so the recipient table and validation immediately reflect the mapping.
  parseAndPreview();
}

/**
 * applyFieldMapping(rawRow) — remap a single parsed-CSV row's keys from the
 * user's column names to the canonical field names.
 *
 * When fieldMapping is empty (the user hasn't opened Match Fields), the raw
 * row is returned unchanged — this is the common case when the CSV already
 * uses canonical column names.
 *
 * We use Object.assign to keep the original columns alongside the mapped ones —
 * custom columns not in CANONICAL_FIELDS still need to be accessible as
 * personalisation tokens (e.g. {{promo_code}} from a non-standard column).
 *
 * @param {Object} rawRow - A single row object from parseCSV()
 * @returns {Object} Row with canonical keys added/overwritten
 */
function applyFieldMapping(rawRow) {
  if (!Object.keys(fieldMapping).length) return rawRow; // No mapping — pass through unchanged
  const mapped = Object.assign({}, rawRow); // Shallow copy — don't mutate the original
  CANONICAL_FIELDS.forEach(f => {
    const csvCol = fieldMapping[f.key];
    // Only remap if there's a mapping for this field AND the CSV has that column.
    if (csvCol && rawRow[csvCol] !== undefined) {
      mapped[f.key] = rawRow[csvCol];
    }
  });
  return mapped;
}

/* ─── BODY TEMPLATE SAVE / LOAD ────────────────────────────────────────────
   These functions read the Outlook compose body (via Office.js body.getAsync)
   and store/restore it as part of a named template in localStorage.
   The body is stored as raw HTML — Outlook's compose HTML, which includes
   inline styles and potentially base64-encoded images.
   ─────────────────────────────────────────────────────────────────────────── */

/**
 * saveBodyToTemplate() — read the current compose body HTML and store it
 * in the named template in localStorage.
 *
 * Uses the template name from #templateNameInput (preferred) or the currently
 * selected template in #templateSelect. Creates the template entry if it
 * doesn't already exist so the body can be saved independently of the settings.
 */
async function saveBodyToTemplate() {
  // Accept name from the new-template input OR from the existing-template dropdown.
  const name = document.getElementById("templateNameInput").value.trim()
             || document.getElementById("templateSelect").value;
  if (!name) { log("Select or name a template first.", "warning"); return; }
  try {
    const body = await getComposeBodyAsync(); // Async Office.js call — must await
    const templates = lsGet(LS_KEY_TEMPLATES, {});
    if (!templates[name]) templates[name] = {}; // Create entry if it doesn't exist
    templates[name].body = body;
    lsSet(LS_KEY_TEMPLATES, JSON.stringify(templates));
    log(`Body saved to template "${name}".`, "success");
  } catch (e) {
    log(`Failed to read compose body: ${e.message}`, "error");
  }
}

/**
 * loadBodyFromTemplate() — restore a previously saved body HTML into the
 * Outlook compose window.
 *
 * Warning: setAsync replaces the ENTIRE compose body, including the user's
 * Outlook signature. Users should include their signature in the template.
 */
async function loadBodyFromTemplate() {
  const name = document.getElementById("templateSelect").value;
  if (!name) { log("Select a template first.", "warning"); return; }
  const templates = lsGet(LS_KEY_TEMPLATES, {});
  const tpl = templates[name];
  if (!tpl || !tpl.body) { log(`Template "${name}" has no saved body.`, "warning"); return; }

  // Mac bug (B2): Office.js body.setAsync with HTML coercion can corrupt base64
  // src attributes on <img> tags on Mac. Warn the user but proceed anyway —
  // there's no workaround, and the text content will still be correct.
  const isMac = Office.context.platform === Office.PlatformType.Mac;
  if (isMac && tpl.body.match(/<img/i)) {
    log("Warning: Mac: loading a body template with embedded images may corrupt image links in the sent email (known Office.js limitation on Mac). Proceeding anyway.", "warning");
  }

  // Signature preservation warning — loading a template replaces the entire compose body on all platforms.
  log("ℹ Loading template will replace the current compose body (including your signature). Tip: place your signature at the bottom of the template to preserve it.", "info");

  try {
    await setComposeBodyAsync(tpl.body);
    log(`Body loaded from template "${name}".`, "success");
  } catch (e) {
    log(`Failed to set compose body: ${e.message}`, "error");
  }
}

/**
 * getComposeBodyAsync() — Promise wrapper around Office.js body.getAsync.
 *
 * We request HTML coercion so we get the full rich-text HTML including
 * inline styles, images, and Outlook-added wrapper elements. Wrapping
 * the callback API in a Promise lets callers use async/await rather than
 * nested callbacks.
 *
 * @returns {Promise<string>} Resolves with the compose body as HTML
 */
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

/**
 * setComposeBodyAsync(html) — Promise wrapper around Office.js body.setAsync.
 *
 * Sets the ENTIRE compose body (replacing any existing content) to the
 * provided HTML. This is a destructive operation — the user's current
 * compose content and signature are lost.
 *
 * @param {string} html - HTML string to set as the compose body
 * @returns {Promise<void>}
 */
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

/* ─── SAVE AS DRAFTS ────────────────────────────────────────────────────────
   Drafts mode runs the normal merge pipeline but saves each email as a draft
   (POST /messages) instead of sending it (POST /sendMail). This is useful for
   large campaigns that need a human review before delivery.
   ─────────────────────────────────────────────────────────────────────────── */

/**
 * handleSaveDrafts() — run the merge in drafts mode.
 *
 * Sets draftsMode = true so buildEmailRequest() uses the /messages endpoint,
 * then delegates to handleMergeClick() for the full merge flow. Restores
 * draftsMode = false and the CSV state in the finally block — always — so
 * a mid-drafts error or cancellation doesn't leave the add-in stuck in
 * drafts mode or with a blank CSV.
 */
async function handleSaveDrafts() {
  // Snapshot CSV state before starting — handleMergeClick may clear the CSV
  // on successful completion (zero failures). We need to restore it because
  // creating drafts isn't the same as "sending" from the user's perspective.
  const savedCsvText = document.getElementById("csvInput").value;
  const savedRecipients = parsedRecipients.slice();
  draftsMode = true;
  try {
    await handleMergeClick();
  } finally {
    draftsMode = false;
    // A8: If handleMergeClick cleared the CSV (its zero-failure cleanup), restore it.
    // Drafts don't constitute a successful "send" — the user may want to send for real
    // after reviewing the drafts.
    if (!document.getElementById("csvInput").value && savedCsvText) {
      document.getElementById("csvInput").value = savedCsvText;
      lsSet(LS_KEY_CSV, savedCsvText);
      parsedRecipients = savedRecipients;
    }
  }
}

/* ─── TEST SEND ─────────────────────────────────────────────────────────────
   Sends a single personalised email to the sender themselves using data from
   a selected recipient row. This lets users verify personalisation, formatting,
   and attachments before committing to a full merge send.
   ─────────────────────────────────────────────────────────────────────────── */

/**
 * handleTestSend() — send a single test email to the logged-in user.
 *
 * Uses the data from the row selected in #testRowSelect (default: row 0) to
 * personalise the subject and body. Sends to the sender's own email address.
 * Applies all the same options as a real merge: custom headers, reply-to,
 * send-as, attachments, inline images, receipts, flags — so the test is a
 * faithful preview of what real recipients will receive.
 *
 * Subject is prefixed with "[TEST]" so the sender can easily identify test
 * emails in their inbox.
 */
async function handleTestSend() {
  // Re-entrancy guard: disable the button for the duration of the async operation.
  // Without this, a slow network call could let the user click twice and send
  // two test emails.
  document.getElementById("testSendBtn").disabled = true;
  // Snapshot and clear _fillInValues — the test send shouldn't use fill-in
  // values from a previous interactive fill-in session.
  const savedFillIn = window._fillInValues;
  window._fillInValues = null;
  try {
  const subjectTemplate = document.getElementById("subjectInput").value.trim();
  if (!subjectTemplate) { log("Subject line is empty.", "error"); return; }
  if (parsedRecipients.length === 0) {
    log("Load recipients first — test send uses the first row for personalisation.", "error");
    return;
  }

  // Get the sender's own email from the Office.js user profile — this is the
  // address the test email will be sent TO.
  const selfEmail = Office.context.mailbox.userProfile.emailAddress;
  if (!selfEmail) { log("Could not determine your email address.", "error"); return; }

  // Use the row index from the test-row selector dropdown.
  // parseInt with radix 10 is defensive against "09" being parsed as octal.
  // Math.min guards against an out-of-bounds index if the CSV shrank since the
  // dropdown was populated.
  const rowIdx = parseInt(document.getElementById("testRowSelect")?.value || "0", 10);
  const sample = parsedRecipients[Math.min(rowIdx, parsedRecipients.length - 1)] || parsedRecipients[0];
  const displayRowNum = rowIdx + 1;

  log(`Test send: personalising with row ${displayRowNum} data and sending to ${selfEmail}...`, "info");

  // Read the Outlook compose body as HTML — this is the template the user has
  // typed (or loaded from a saved template) in the compose window.
  let emailBodyTemplate = "";
  try {
    emailBodyTemplate = await new Promise((resolve, reject) => {
      Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
        else reject(new Error("Could not read email body."));
      });
    });
  } catch (err) { log(`Body read error: ${err.message}`, "error"); return; }

  // Inject smart tokens (record_num, record_count) into the sample row object
  // so they resolve correctly during personalisation.
  const sampleWithMeta = Object.assign({}, sample, {
    record_num:   String(rowIdx + 1),
    record_count: String(parsedRecipients.length || 1)
  });
  // personalize() second arg = htmlEscape flag:
  //   false for subject (plain text — &amp; in subjects looks wrong)
  //   true  for body (HTML — prevents XSS via malicious CSV data)
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

/* ─── PREVIEW ALL ───────────────────────────────────────────────────────────
   Lets the user step through every recipient's personalised subject and body
   before sending. Uses the same personalisation pipeline as the real merge so
   what they see is exactly what recipients will receive.
   ─────────────────────────────────────────────────────────────────────────── */

/**
 * handlePreviewAll() — open the preview modal and snapshot the body template.
 *
 * Reads the current compose body from Outlook (snapshot to previewBodyTemplate)
 * so that edits made during a long preview session don't affect the displayed
 * previews mid-way through. Uses the filtered/sorted recipient set so the
 * preview matches what will actually be sent (respecting any active filters).
 */
async function handlePreviewAll() {
  // Use the filtered/sorted set so the preview matches what will actually be sent.
  const toPreview = getFilteredSortedRecipients();
  if (!toPreview.length) {
    log("No recipients to preview.", "warning");
    showToast("No recipients to preview. Load a CSV or add recipients first.", "warning");
    return;
  }
  const subjectTemplate = document.getElementById("subjectInput").value.trim();
  if (!subjectTemplate) { log("Subject line is empty.", "error"); return; }

  // Warn if fill_in tokens are present — they won't resolve without running a merge first.
  const bodyHtmlForPreview = await getComposeBodyAsync().catch(function() { return ""; });
  if (!window._fillInValues && /\{\{fill_in:/i.test(bodyHtmlForPreview)) {
    log("Note: {{fill_in:…}} tokens found — run a merge first to populate fill-in values. Previewing with placeholder tokens.", "warning");
  }

  log("Reading email body for preview...", "info");
  try {
    // Snapshot the body NOW so the preview is deterministic even if the user
    // edits the compose window while stepping through previews.
    previewBodyTemplate = await new Promise((resolve, reject) => {
      Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
        else reject(new Error("Could not read email body."));
      });
    });
  } catch (err) { log(`Body read error: ${err.message}`, "error"); return; }

  previewRecipients = toPreview; // Store the filtered set for Prev/Next navigation
  previewIndex = 0;              // Always start at the first recipient
  renderPreviewEntry();
  // _openModalWithTrap sets focus to the first focusable element and traps
  // keyboard focus inside the modal (accessibility requirement).
  _openModalWithTrap("previewModal");
}

/**
 * renderPreviewEntry() — personalise and display the current preview recipient.
 *
 * Called on modal open and after each Prev/Next button click.
 * Uses previewBodyTemplate (snapshotted at modal-open time) rather than
 * re-reading the compose body, which ensures consistency across navigation.
 */
function renderPreviewEntry() {
  const total = previewRecipients.length;
  if (!total || previewIndex >= total) return;
  // Inject record_num and record_count as if this were a real merge run.
  const recipient = Object.assign({}, previewRecipients[previewIndex], {
    record_num:   String(previewIndex + 1),
    record_count: String(total)
  });
  const subjectTemplate = document.getElementById("subjectInput").value;

  // personalize() third argument: false = no HTML escaping (subject is plain text),
  // true = HTML-escape merge values (body is rendered in an iframe).
  const subject  = personalize(subjectTemplate, recipient, false);
  const body     = personalize(previewBodyTemplate, recipient, true);
  const plainTextMode = document.getElementById("plainTextMode").checked;
  // In plain-text mode, strip all HTML tags and show the result in a <pre> block
  // so whitespace is preserved. escapeHtml prevents any remaining HTML entities
  // from being interpreted by the browser.
  const displayBody   = plainTextMode
    ? `<pre style="white-space:pre-wrap;font-family:Segoe UI,sans-serif;font-size:13px;padding:8px;">${escapeHtml(stripHtmlToText(body))}</pre>`
    : body;

  document.getElementById("previewCounter").textContent = `Preview ${previewIndex + 1} / ${total}`;
  document.getElementById("previewSubject").textContent = subject;
  // srcdoc renders the HTML in an isolated sandbox without a network request.
  // Safer than innerHTML because the iframe has a different origin.
  document.getElementById("previewBodyFrame").srcdoc    = displayBody;

  // Disable Prev/Next at the boundaries so the user can't navigate out of range.
  document.getElementById("previewPrevBtn").disabled = previewIndex === 0;
  document.getElementById("previewNextBtn").disabled = previewIndex === total - 1;
}

/**
 * closePreviewModal() — close the preview modal and restore focus.
 */
function closePreviewModal() {
  // _closeModalWithTrap restores focus to the element that had focus when the
  // modal was opened (accessibility requirement).
  _closeModalWithTrap("previewModal");
}

/* ─── DOWNLOAD LOG ──────────────────────────────────────────────────────────
   Exports the full text content of the status log panel to a .txt file.
   Useful for support requests — users can attach the exported log to show
   exactly what happened during a merge run.
   ─────────────────────────────────────────────────────────────────────────── */

/**
 * downloadLog() — export all log entries as a plain-text file.
 *
 * Collects the textContent of every .log-entry paragraph, joins with newlines,
 * and triggers a browser download. Timestamp is embedded in the filename for
 * uniqueness. Colons in the ISO timestamp are replaced with hyphens because
 * colons are illegal in Windows filenames.
 */
function downloadLog() {
  const entries = document.querySelectorAll("#statusLog .log-entry");
  if (!entries.length) { log("Log is empty.", "info"); return; }
  const text = Array.from(entries).map(el => el.textContent).join("\n");
  const blob = new Blob([text], { type: "text/plain" });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement("a");
  a.href     = url;
  // Slice(0,19) gives "2024-01-15T14:30:00" — drop the sub-second and timezone parts.
  a.download = `mail-merge-log-${new Date().toISOString().slice(0, 19).replace(/:/g, "-")}.txt`;
  // Must append to DOM before clicking — some browsers require the element to be in the document.
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url); // Release the blob URL to free memory
  log("Log exported.", "info");
}

/* ─── SEND SUMMARY REPORT ───────────────────────────────────────────────────
   After a merge run, downloadSendReport() generates a CSV with one row per
   send attempt (including retries) from the sendOutcomes array. The report is
   useful for auditing: who was contacted, when, and whether it succeeded.
   ─────────────────────────────────────────────────────────────────────────── */

/**
 * downloadSendReport() — export the sendOutcomes array as a downloadable CSV.
 *
 * csvField() is applied to every cell to ensure any values containing commas,
 * quotes, or newlines are properly quoted per RFC 4180. Without this, an error
 * message containing a comma would break the CSV structure.
 */
function downloadSendReport() {
  if (!sendOutcomes.length) { log("No send outcomes to report.", "info"); return; }
  // Build header row first, then one data row per outcome.
  const rows = [["row_num","email","display_name","subject_used","status","timestamp","error"]];
  sendOutcomes.forEach(o => rows.push([
    o.rowNum      || "",
    o.email       || "",
    o.displayName || "",
    o.subjectUsed || "",
    o.status      || "",
    o.timestamp   || "",
    o.error       || ""
  ].map(v => csvField(String(v))))); // csvField wraps values in quotes if needed
  const csv  = rows.map(r => r.join(",")).join("\n");
  const blob = new Blob([csv], { type: "text/csv" });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement("a");
  a.href     = url;
  a.download = `mail-merge-report-${new Date().toISOString().slice(0, 10)}.csv`; // date-only suffix
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  log("Send report exported.", "info");
}

/* ─── PROGRESS BAR ──────────────────────────────────────────────────────────
   Three functions manage the progress bar: setProgress (live update during
   send), hideProgress (reset after cancel), showMergeComplete (success state).
   ─────────────────────────────────────────────────────────────────────────── */

/**
 * setProgress(current, total) — update the progress bar during a merge run.
 *
 * Also updates the ARIA aria-valuenow attribute so screen readers announce
 * progress without the user needing to navigate to the progress bar element.
 *
 * @param {number} current - Number of emails sent so far
 * @param {number} total   - Total number of emails to send
 */
function setProgress(current, total) {
  // Guard against division by zero — happens if this is called before total is known.
  const pct = total > 0 ? Math.round((current / total) * 100) : 0;
  document.getElementById("progressContainer").classList.remove("hidden");
  const fill = document.getElementById("progressFill");
  fill.style.width = `${pct}%`;
  fill.setAttribute("aria-valuenow", pct); // ARIA live region for screen readers
  document.getElementById("progressLabel").textContent = `${pct}%  (${current} / ${total})`;
}

/**
 * hideProgress() — reset and hide the progress bar.
 *
 * Called when a merge is cancelled or an error aborts the run.
 * Does NOT call this on success — showMergeComplete() handles that path
 * and auto-hides after 8 seconds.
 */
function hideProgress() {
  document.getElementById("progressContainer").classList.add("hidden");
  document.getElementById("progressFill").style.width = "0%";
}

/**
 * showMergeComplete(sent, total) — show a "Done" completion state in the
 * progress bar and auto-hide it after 8 seconds.
 *
 * Sets _mergeCompletedSuccessfully to true so the auto-hide timer doesn't
 * conflict with a new merge that starts before the 8 s elapses.
 *
 * @param {number} sent  - Number of emails successfully sent
 * @param {number} total - Total number attempted
 */
function showMergeComplete(sent, total) {
  _mergeCompletedSuccessfully = true;
  const fill  = document.getElementById("progressFill");
  const label = document.getElementById("progressLabel");
  const container = document.getElementById("progressContainer");
  if (container) container.classList.remove("hidden");
  if (fill)  fill.style.width = "100%"; // Always show 100% on completion
  if (label) label.textContent = "✓ Done — " + sent + " of " + total + " sent";
  // Auto-hide after 8 s. Guard with !mergeInProgress in case a new merge
  // started while this timer was running — we don't want to hide a live bar.
  setTimeout(function() {
    if (!mergeInProgress) {
      const c = document.getElementById("progressContainer");
      if (c) c.classList.add("hidden");
      _mergeCompletedSuccessfully = false;
    }
  }, 8000);
}

/* ─── SUPPRESSION LIST ──────────────────────────────────────────────────────
   The suppression list (also called opt-out or unsubscribe list) is a Set of
   email addresses that should NEVER receive emails from this add-in. Recipients
   whose (lowercased) address appears in suppressionSet are silently skipped
   during the merge loop — they receive no email and no error is logged.
   ─────────────────────────────────────────────────────────────────────────── */

/**
 * handleSuppressionUpload(e) — parse a text file of suppressed email addresses
 * and add them to suppressionSet.
 *
 * The file can be newline-separated, comma-separated, or semicolon-separated.
 * We lowercase all addresses before adding them because suppressionSet lookups
 * are case-insensitive — if we stored "Bob@EXAMPLE.COM" and the CSV has
 * "bob@example.com", the lookup would miss without normalisation.
 *
 * Resets e.target.value after reading so the user can re-upload the same file.
 *
 * @param {Event} e - The "change" event from the file input
 */
function handleSuppressionUpload(e) {
  const file = e.target.files[0];
  if (!file) return;
  e.target.value = ""; // Allow re-selecting the same file later
  const reader = new FileReader();
  reader.onload = (event) => {
    const text   = event.target.result;
    // Split on newlines, commas, semicolons, or carriage returns — handles
    // most export formats from email marketing platforms.
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

/**
 * clearSuppression() — empty suppressionSet and reset the UI.
 *
 * After clearing, all previously suppressed addresses will receive emails
 * in the next merge run. This is intentional — the user explicitly clicked
 * "Clear suppression list".
 */
function clearSuppression() {
  suppressionSet.clear();
  document.getElementById("suppressionLabel").textContent = "None loaded";
  document.getElementById("clearSuppressionBtn").classList.add("hidden");
  log("Suppression list cleared.", "info");
}

/* ─── RETRY FAILED ──────────────────────────────────────────────────────────
   After a merge run with partial failures, the "Retry failed" button calls
   handleRetryFailed(). It temporarily swaps parsedRecipients with the failed
   list and runs the full merge pipeline again. The finally block restores the
   original list whether or not the retry succeeded.
   ─────────────────────────────────────────────────────────────────────────── */

/**
 * handleRetryFailed() — retry sending to all recipients that failed in the
 * most recent merge run.
 *
 * The key pattern here is "swap, run, restore in finally":
 *   1. Snapshot the current parsedRecipients and CSV text
 *   2. Replace parsedRecipients with failedRecipients
 *   3. Clear failedRecipients (so the retry itself can populate it again)
 *   4. Run handleMergeClick() which reads parsedRecipients
 *   5. In the finally block, restore parsedRecipients and the CSV textarea
 *      unconditionally — even if the retry throws or is cancelled
 *
 * The CSV restoration in the finally block addresses a bug where handleMergeClick
 * could clear the CSV textarea on zero-failure completion, leaving the user
 * with a blank CSV after a successful retry.
 */
async function handleRetryFailed() {
  if (!failedRecipients.length) return;
  const savedRecipients = parsedRecipients.slice(); // shallow copy is sufficient
  const savedCsvText = document.getElementById("csvInput").value;
  // Temporarily replace the working list with the failed subset
  parsedRecipients = failedRecipients.slice();
  failedRecipients = []; // Clear so the retry can repopulate on new failures
  try {
    await handleMergeClick();
  } finally {
    // Always restore — regardless of success, error, or cancellation.
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
    showToast(`${tag} added to subject.`, "success", 2000);

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
          showToast(`${tag} added to subject.`, "success", 2000);
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
          showToast(`${tag} inserted into body.`, "success", 2000);
        }
      }
    );
  }
}

/**
 * insertNewline() — insert a line break into the Outlook compose body from the taskpane.
 *
 * Clicking a tag chip moves focus to the taskpane iframe, which means the user
 * can't press Enter to create a new line in the body without clicking back there.
 * This function provides a "↵ New line" button in the tag bar as a workaround:
 * it calls setSelectedDataAsync with a newline character so the user can build
 * multi-line templates entirely from the taskpane without switching focus.
 *
 * Mac (Text coercion): "\n" creates a line break in plain-text mode.
 * Windows (Html coercion): "<br>" creates a line break in HTML mode.
 */
function insertNewline() {
  // Only makes sense when target is the body — ignore if subject is active.
  if (tagTarget === "subject" || subjectHasFocus) {
    showToast("Switch to Body mode to insert a new line.", "warning", 2000);
    return;
  }
  const isMac = Office.context.platform === Office.PlatformType.Mac;
  const newlineChar = isMac ? "\n" : "<br>";
  const coercionType = isMac ? Office.CoercionType.Text : Office.CoercionType.Html;
  Office.context.mailbox.item.body.setSelectedDataAsync(
    newlineChar,
    { coercionType },
    (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        log(`Failed to insert newline into body: ${result.error.message}`, "error");
      }
      // No toast for newline — it's a silent action; the visual result in the body is feedback enough.
    }
  );
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
function addTagToBar(tag, type, persist) {
  // persist defaults to true for manually-added tags, false for CSV-column tags.
  // CSV column tags are transient (the CSV data is in-memory only) so they
  // should never be saved to localStorage.
  if (persist === undefined) persist = true;
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
  // Only persist to localStorage for manually-added tags, not CSV column tags.
  if (type !== "smart" && persist) saveCustomTagsToStorage();
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
 * Remove all CSV-derived custom tag chips (tag-custom class) from the bar
 * and clear them from localStorage. Called when a new CSV is loaded so stale
 * column tags from the previous file don't linger.
 */
function clearCsvTags() {
  document.querySelectorAll("#tagBar .tag-custom").forEach(chip => chip.remove());
  lsSet(LS_KEY_TAGS, JSON.stringify([]));
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

/* ─── FILE UPLOAD: CSV + EXCEL ──────────────────────────────────────────────
   These functions handle the file-picker upload path. The drag-and-drop path
   in Office.onReady() reuses handleCsvFileUpload() by injecting the dropped
   file into csvFileInput and triggering a "change" event.
   ─────────────────────────────────────────────────────────────────────────── */

/**
 * handleCsvFileUpload(e) — dispatch a file to either the CSV or Excel loader
 * based on its extension. Called by the csvFileInput "change" event.
 *
 * @param {Event} e - The file input "change" event
 */
function handleCsvFileUpload(e) {
  const file = e.target.files[0];
  if (!file) return;
  e.target.value = ""; // Reset input so the same file can be re-uploaded
  const ext = file.name.split(".").pop().toLowerCase();
  if (ext === "xlsx" || ext === "xls") {
    loadExcelFile(file); // Excel path — uses SheetJS to convert to CSV
  } else {
    loadCsvFile(file);   // Plain CSV or TSV path
  }
}

/**
 * loadCsvFile(file) — read a plain-text CSV file into the csvInput textarea
 * and trigger parseAndPreview().
 *
 * FileReader.readAsText() returns the file as a UTF-8 string. The browser
 * handles BOM (byte-order mark) stripping automatically on most platforms.
 *
 * @param {File} file - The File object to read
 */
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

/**
 * loadExcelFile(file) — use the SheetJS (XLSX) library to convert an Excel
 * workbook's first sheet to CSV, then load it exactly like a .csv file.
 *
 * We read the file as ArrayBuffer (not text) because the XLSX library needs
 * the raw binary data to parse the workbook format. After conversion to CSV
 * string, the rest of the pipeline (parseCSV, validation, etc.) is identical
 * to the plain-text CSV path.
 *
 * The SheetJS CDN library is loaded in taskpane.html. If it failed to load
 * (network error, CSP block), XLSX will be undefined — guard and warn.
 *
 * @param {File} file - The .xlsx or .xls File object to read
 */
function loadExcelFile(file) {
  if (typeof XLSX === "undefined") {
    log("Excel support unavailable — SheetJS failed to load. Try refreshing the task pane.", "error");
    return;
  }
  const reader = new FileReader();
  reader.onload = (event) => {
    try {
      // Wrap the ArrayBuffer in Uint8Array — XLSX.read expects a typed array.
      const data      = new Uint8Array(event.target.result);
      const workbook  = XLSX.read(data, { type: "array" });
      // Always use the first sheet — we don't support multi-sheet selection.
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
  reader.readAsArrayBuffer(file); // Must be ArrayBuffer for SheetJS, not text
}

/* ─── ATTACHMENT UPLOAD: SHARED ─────────────────────────────────────────────
   Shared attachments are the same files attached to every email in the merge.
   They're read as Data URLs (base64) via FileReader and stored in the
   sharedAttachments array. Graph sendMail expects attachments as base64
   contentBytes in the request body.
   ─────────────────────────────────────────────────────────────────────────── */

/**
 * handleAttachmentUpload(e) — read one or more shared attachment files and
 * add them to the sharedAttachments array.
 *
 * The "pending" counter pattern: since FileReader.onload fires asynchronously
 * for each file, we can't know when ALL files have loaded without tracking
 * how many are still in-flight. We decrement pending on each completion (success
 * or error) and call checkDone() which fires the final UI update only once.
 *
 * @param {Event} e - The "change" event from the shared attachment file input
 */
function handleAttachmentUpload(e) {
  const files = Array.from(e.target.files);
  e.target.value = ""; // Reset so the same file can be re-added
  if (!files.length) return;

  let pending     = files.length; // tracks in-flight FileReader callbacks
  let loadedCount = 0;

  function checkDone() {
    pending--;
    if (pending === 0) { // All files processed (success or error)
      updateSharedAttachmentsLabel();
      if (loadedCount > 0) {
        log(`Loaded ${loadedCount} shared attachment${loadedCount !== 1 ? "s" : ""}.`, "success");
      }
    }
  }

  files.forEach(file => {
    // Enforce Graph's per-attachment size limit before reading.
    if (file.size > MAX_ATTACHMENT_BYTES) {
      log(`Shared attachment too large: ${file.name} ` +
          `(${(file.size / 1024 / 1024).toFixed(1)} MB). Max 3 MB — skipped.`, "error");
      checkDone(); // Still must decrement pending even on skip
      return;
    }
    const reader = new FileReader();
    reader.onload = (event) => {
      // readAsDataURL returns "data:image/png;base64,iVBORw0..."
      // We need just the base64 part (after the comma) for Graph,
      // and just the MIME type (between ":" and ";") for contentType.
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

/**
 * updateSharedAttachmentsLabel() — refresh the UI label showing how many
 * shared attachments are loaded and their total size.
 */
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

/**
 * clearSharedAttachments() — empty the sharedAttachments array and reset the UI.
 */
function clearSharedAttachments() {
  sharedAttachments = [];
  document.getElementById("attachmentLabel").textContent = "None";
  document.getElementById("clearSharedAttachmentsBtn").classList.add("hidden");
  log("Shared attachments cleared.", "info");
}

/* ─── ATTACHMENT UPLOAD: PER-RECIPIENT ──────────────────────────────────────
   Per-recipient attachments are matched to recipients by filename. The CSV
   "attachment" column holds the filename; the uploaded files are stored in the
   perRecipientFiles Map keyed by lowercased filename. At send time,
   resolveAttachmentsForRecipient() looks up the filename from the row's
   attachment column in this Map.
   ─────────────────────────────────────────────────────────────────────────── */

/**
 * handlePerRecipientFilesUpload(e) — read multiple files for per-recipient
 * attachments and store them in perRecipientFiles Map (key = lowercased filename).
 *
 * Lowercasing the key ensures filename lookups are case-insensitive —
 * the CSV might say "Invoice_A.pdf" while the file is "invoice_a.pdf".
 *
 * @param {Event} e - The "change" event from the per-recipient file input
 */
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
      // Key by lowercased filename — matching is case-insensitive at resolve time.
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

/**
 * clearPerRecipientFiles() — empty the perRecipientFiles Map and reset the UI.
 */
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

/**
 * clearInlineImages() — empty the inlineImages Map and reset the UI.
 */
function clearInlineImages() {
  inlineImages.clear();
  document.getElementById("inlineImagesLabel").textContent = "None loaded";
  document.getElementById("clearInlineImagesBtn").classList.add("hidden");
  log("Inline images cleared.", "info");
}

/**
 * resolveAttachmentsForRecipient(recipient) — determine which attachment files
 * to include for a specific recipient.
 *
 * Logic:
 *   1. If the recipient's "attachment" field is non-empty, look it up in
 *      perRecipientFiles by lowercased filename.
 *   2. If found, return just that one per-recipient file.
 *   3. If not found (filename in CSV but no matching uploaded file), log a
 *      warning ONCE per unique filename (warnedMissingAttachments deduplicates),
 *      then fall back to the shared attachments.
 *   4. If "attachment" is empty, return the shared attachments (sent to everyone).
 *
 * The fallback behaviour (missing per-recipient file → shared attachments) is
 * intentional — better to send something than nothing. The warning ensures the
 * user knows about the mismatch.
 *
 * @param {Object} recipient - A row object from parsedRecipients
 * @returns {Array} Array of attachment objects { name, contentType, contentBytes, sizeBytes }
 */
function resolveAttachmentsForRecipient(recipient) {
  const filename = (recipient.attachment || "").trim();
  if (filename) {
    // Look up by lowercased filename for case-insensitive matching.
    const perFile = perRecipientFiles.get(filename.toLowerCase());
    if (perFile) return [perFile]; // Found — return this recipient's specific file
    // File named in CSV but not in the uploaded set — warn once per unique filename.
    const key = filename.toLowerCase();
    if (!warnedMissingAttachments.has(key)) {
      warnedMissingAttachments.add(key); // Mark as warned so we don't repeat this
      log(`No loaded file matches "${filename}" — falling back to shared attachment(s) for rows that reference it.`, "warning");
    }
    return [...sharedAttachments]; // Fallback: use shared attachments
  }
  // No per-recipient attachment column value — use shared attachments only.
  return [...sharedAttachments];
}

/* ─── CSV PARSING (RFC 4180) ────────────────────────────────────────────────
   A hand-rolled RFC 4180 parser rather than a split-based approach. Why?
   Because split-based parsers can't handle the three hard cases:
     1. Quoted fields containing commas: "Smith, John",alice@example.com
     2. Quoted fields containing newlines (a single "cell" spanning two lines)
     3. Escaped double-quotes inside a quoted field: "Say ""hello"""
   All three are valid RFC 4180 and occur in real-world CSV exports from
   Excel, Google Sheets, and CRM tools.
   ─────────────────────────────────────────────────────────────────────────── */

/**
 * parseCSV(raw) — RFC 4180-compliant CSV parser with auto-detected delimiter.
 *
 * Returns an object { headers, rows } where:
 *   headers: string[] — the lowercased, trimmed column headers from row 1
 *   rows: Object[]    — each row as { columnName: value, _csvRow: number,
 *                        _originalIndex: number }
 *
 * Internal properties prefixed with "_" are excluded from field-mapping
 * and merge token resolution so they don't accidentally appear in emails.
 *
 * @param {string} raw - The raw CSV text to parse
 * @returns {{ headers: string[], rows: Object[] }}
 */
function parseCSV(raw) {
  // ── Auto-detect delimiter ─────────────────────────────────────────────
  // Count occurrences of each candidate delimiter in the first non-empty line.
  // Whichever appears most is assumed to be the delimiter. Default is comma.
  // This handles TSV files (tabs) and European locale CSV (semicolons) without
  // requiring the user to specify the delimiter manually.
  let delimiter = ",";
  const firstLine = raw.split(/\r?\n/).find(function(l) { return l.trim().length > 0; }) || "";
  const tabCount   = (firstLine.match(/\t/g)  || []).length;
  const semiCount  = (firstLine.match(/;/g)   || []).length;
  const commaCount = (firstLine.match(/,/g)   || []).length;
  if (tabCount > commaCount && tabCount > semiCount) delimiter = "\t";
  else if (semiCount > commaCount) delimiter = ";";
  // else keep ","

  // ── Character-by-character state machine ─────────────────────────────
  // State variables: rawRows accumulates parsed rows, row is the current
  // row's fields, field is the current field being built character by character,
  // inQuotes tracks whether we're inside a double-quoted field.
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
          // RFC 4180 §2.7: two consecutive double-quotes inside a quoted field
          // represent a single literal double-quote. Consume both.
          field += '"';
          i += 2;
        } else {
          // Closing quote: exit quoted mode (the next char should be delimiter or newline).
          inQuotes = false;
          i++;
        }
      } else {
        // Any character inside quotes is literal — including newlines (embedded newlines
        // in quoted fields are valid RFC 4180 and must not end the row).
        field += ch;
        i++;
      }
    } else {
      if (ch === '"') {
        inQuotes = true; // Enter quoted field mode
        i++;
      } else if (ch === delimiter) {
        row.push(field); // End of field
        field = "";
        i++;
      } else if (ch === '\r') {
        // Handle both \r\n (Windows) and bare \r (legacy Mac) line endings.
        row.push(field);
        field = "";
        rawRows.push(row);
        row = [];
        if (i + 1 < n && raw[i + 1] === '\n') i++; // Consume the \n in \r\n
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
  // Final field/row: the last row may not end with a newline.
  if (field || row.length > 0) {
    row.push(field);
    rawRows.push(row);
  }
  // Remove trailing empty row — many CSV exports end with a final newline which
  // produces a spurious empty row. If every field in the last row is "", drop it.
  if (rawRows.length > 0 && rawRows[rawRows.length - 1].every(f => f === "")) {
    rawRows.pop();
  }

  // Need at least a header row + one data row.
  if (rawRows.length < 2) return { headers: [], rows: [] };

  // ── Build header and row objects ──────────────────────────────────────
  // Lowercase headers so column name matching is case-insensitive.
  // Trimming handles leading/trailing spaces from some CSV exports.
  const headers = rawRows[0].map(h => h.trim().toLowerCase());
  const rows    = [];

  for (let r = 1; r < rawRows.length; r++) {
    const cols = rawRows[r].map(c => c.trim());
    // Skip entirely blank rows (e.g. blank lines in the middle of the file).
    if (cols.every(c => c === "")) continue;
    // Pad short rows rather than dropping them — a row with fewer columns than
    // headers just has empty values for the missing fields. This handles the
    // common case where the last few optional columns are omitted.
    while (cols.length < headers.length) cols.push("");
    const rowObj = {};
    headers.forEach((h, idx) => { rowObj[h] = cols[idx] || ""; });
    // _csvRow: the 1-based line number in the original CSV (header = line 1, first
    // data row = line 2). Stored for error messages — "row 5 has an invalid email".
    rowObj._csvRow = r + 1; // +1 because r starts at 1 (skipping header), but header is line 1
    // _originalIndex: position in the rows array BEFORE any filter/sort.
    // Stamped here rather than computed with indexOf() later — O(1) vs O(n).
    rowObj._originalIndex = rows.length;
    rows.push(rowObj);
  }

  return { headers, rows };
}

/* ─── FILTER / SORT STATE ───────────────────────────────────────────────────
   The filter/sort bar lets users narrow the recipient list before sending.
   Filters are applied client-side on parsedRecipients using getFilteredSortedRecipients().
   Multiple filter conditions can be combined with AND or OR logic.
   ─────────────────────────────────────────────────────────────────────────── */

// activeFilter: the currently applied filter/sort specification, or null if
// no filter is active. Shape: { conditions, sortCol, sortDir, logic }
// Used by getFilteredSortedRecipients() to produce the working set.
let activeFilter = null;

// filterConditionCount: monotonically increasing counter used to generate unique
// IDs for condition rows. Never resets — ensures IDs stay unique even after rows
// are removed and new ones added.
let filterConditionCount = 0;

/**
 * populateFilterSortBar(headers) — populate the sort-column dropdown with the
 * actual column names from the loaded CSV and make the filter bar visible.
 *
 * Called every time parseAndPreview() rebuilds the recipient table.
 * Builds the options string in memory and assigns in one operation
 * (avoiding repeated DOM reflows from innerHTML += in a loop).
 *
 * @param {string[]} headers - Column names from the parsed CSV
 */
function populateFilterSortBar(headers) {
  const sortCol = document.getElementById("sortCol");
  // Build all options then assign once — avoids repeated innerHTML += reparse on each iteration.
  sortCol.innerHTML = '<option value="">col…</option>' +
    headers.map(h => `<option value="${escapeHtml(h)}">${escapeHtml(h)}</option>`).join("");
  // NOTE: The filter controls now live inside #filterPopup (a popup overlay) rather
  // than the old #filterSortBar section. The popup is accessible via the Filter
  // button inside the preview table header — no need to show a separate bar here.
  // Ensure at least one blank condition row is ready when the popup first opens.
  if (document.getElementById("filterConditions").children.length === 0) {
    addFilterCondition();
  }
}

/**
 * addFilterCondition(colVal, opVal, valVal) — append a new filter condition row
 * to the #filterConditions container.
 *
 * Each row has a column selector, operator selector, value input, and a remove
 * button. Pre-fills defaults when called from applyFilterSort() after restoring
 * a previous filter state.
 *
 * @param {string} colVal - Column to pre-select (default: "")
 * @param {string} opVal  - Operator to pre-select (default: "contains")
 * @param {string} valVal - Value to pre-fill (default: "")
 */
function addFilterCondition(colVal = "", opVal = "contains", valVal = "") {
  filterConditionCount++;
  const id = filterConditionCount; // Unique ID for this condition row's DOM element
  // Build column options from current parsedRecipients keys, excluding internals.
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
  // Remove button: remove this specific condition row from the DOM.
  div.querySelector(".filter-remove-btn").addEventListener("click", () => {
    const condRowEl = document.getElementById("filter-cond-" + id);
    if (condRowEl) condRowEl.remove();
  });
  document.getElementById("filterConditions").appendChild(div);
}

/**
 * getFilterConditions() — read the current state of all filter condition rows
 * and return them as an array of condition objects.
 *
 * Filters out rows with no column selected (blank col) — they're ignored.
 *
 * @returns {Array<{col: string, op: string, val: string}>}
 */
function getFilterConditions() {
  const rows = document.querySelectorAll("#filterConditions > div");
  return [...rows].map(row => ({
    col: row.querySelector(".filter-col-sel").value,
    op:  row.querySelector(".filter-op-sel").value,
    val: row.querySelector(".filter-val-inp").value,
  })).filter(c => c.col); // ignore rows with no column selected
}

/**
 * applyCondition(row, cond) — test a single recipient row against a single
 * filter condition. Returns true if the row passes (should be included).
 *
 * Numeric comparisons (gt/lt) use parseFloat so "100" > "9" works correctly
 * as a number. String comparisons (contains/eq) are case-insensitive.
 *
 * @param {Object} row  - A recipient row object from parsedRecipients
 * @param {Object} cond - A filter condition { col, op, val }
 * @returns {boolean}
 */
function applyCondition(row, cond) {
  const cellVal = String(row[cond.col] !== null && row[cond.col] !== undefined ? row[cond.col] : "").trim();
  const testVal = cond.val.trim();
  const numCell = parseFloat(cellVal);
  const numTest = parseFloat(testVal);
  // bothNum: true when both values are numeric — enables numeric comparison.
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

/**
 * getFilteredSortedRecipients() — return the subset of parsedRecipients that
 * passes the current row selection and activeFilter criteria, in sort order.
 *
 * This is the single source of truth for "what will actually be sent" — the
 * merge loop, preview-all, and simulate all call this function to get their
 * working set. Modifying activeFilter or selectedRowIndices and calling this
 * again immediately reflects the new state.
 *
 * @returns {Object[]} Filtered and sorted copy of parsedRecipients
 */
function getFilteredSortedRecipients() {
  // Shallow copy so we can filter/sort without mutating parsedRecipients.
  let result = parsedRecipients.slice();

  // Apply row checkbox selection first — this is the most restrictive filter
  // (selectedRowIndices = null means "all", a Set means "only these indices").
  if (selectedRowIndices !== null) {
    // Filter by the original index (assigned at parse time as _originalIndex),
    // not the current array index, so filtering survives sort order changes.
    result = result.filter((_, i) => selectedRowIndices.has(i));
  }

  if (activeFilter) {
    const { conditions = [], sortCol, sortDir, logic = "AND" } = activeFilter;
    if (conditions.length) {
      result = result.filter(row => {
        // AND logic: every condition must pass.
        // OR logic: at least one condition must pass.
        if (logic === "OR") return conditions.some(c => applyCondition(row, c));
        return conditions.every(c => applyCondition(row, c));
      });
    }
    if (sortCol) {
      // Case-insensitive string sort. Null-safe with the ternary fallback to "".
      result.sort((a, b) => {
        const av = String(a[sortCol] !== null && a[sortCol] !== undefined ? a[sortCol] : "").toLowerCase();
        const bv = String(b[sortCol] !== null && b[sortCol] !== undefined ? b[sortCol] : "").toLowerCase();
        const cmp = av < bv ? -1 : av > bv ? 1 : 0;
        return sortDir === "desc" ? -cmp : cmp; // Negate for descending
      });
    }
  }
  return result;
}

/**
 * applyFilterSort() — read the current filter conditions and sort settings from
 * the UI, store them in activeFilter, and re-render the preview table.
 */
function applyFilterSort() {
  const conditions = getFilterConditions();
  const sortCol = document.getElementById("sortCol").value;
  const sortDir = document.getElementById("sortDir").value;
  const logic   = document.getElementById("filterLogic").value; // "AND" | "OR"
  // If no conditions and no sort, clear the filter entirely rather than storing
  // an empty filter object (getFilteredSortedRecipients() checks for null).
  activeFilter  = (conditions.length || sortCol) ? { conditions, sortCol, sortDir, logic } : null;
  renderFilteredPreview();
  // Close the popup once the filter is applied — the preview table re-renders
  // with the dot indicator showing that a filter is active.
  closeFilterPopup();
}

/**
 * clearFilterSort() — reset all filter/sort state and re-render.
 *
 * Clears both the in-memory activeFilter and all the UI controls so the user
 * starts fresh. Resets filterConditionCount so new condition IDs don't clash
 * with stale DOM elements.
 */
function clearFilterSort() {
  activeFilter = null;
  document.getElementById("filterConditions").innerHTML = "";
  filterConditionCount = 0;
  document.getElementById("sortCol").value   = "";
  document.getElementById("sortDir").value   = "asc";
  document.getElementById("filterLogic").value = "AND";
  document.getElementById("filterCountLabel").textContent = "";
  renderFilteredPreview();
  // Close the popup after clearing — the table re-renders without the active dot.
  closeFilterPopup();
}

/**
 * openFilterPopup() — show the #filterPopup overlay.
 * Removes the .hidden class and adds a keydown listener for Escape to close.
 */
function openFilterPopup() {
  const popup = document.getElementById("filterPopup");
  if (!popup) return;
  popup.classList.remove("hidden");
  // Trap Escape key to close the popup — accessibility best practice for dialogs.
  document._filterPopupEscHandler = function(e) {
    if (e.key === "Escape") closeFilterPopup();
  };
  document.addEventListener("keydown", document._filterPopupEscHandler);
}

/**
 * closeFilterPopup() — hide the #filterPopup overlay.
 * Adds the .hidden class back and cleans up the Escape listener.
 */
function closeFilterPopup() {
  const popup = document.getElementById("filterPopup");
  if (!popup) return;
  popup.classList.add("hidden");
  if (document._filterPopupEscHandler) {
    document.removeEventListener("keydown", document._filterPopupEscHandler);
    document._filterPopupEscHandler = null;
  }
}

/**
 * renderFilteredPreview() — apply the current filter and re-render the preview table.
 *
 * Also updates the #filterCountLabel to show how many rows pass the current filter.
 */
function renderFilteredPreview() {
  const rows = getFilteredSortedRecipients();
  document.getElementById("filterCountLabel").textContent =
    activeFilter ? `${rows.length} of ${parsedRecipients.length} rows` : "";
  renderPreviewTable(rows);
}

/**
 * parseAndPreview() — parse the CSV textarea, validate it, and rebuild all
 * dependent UI elements.
 *
 * This is the central "reload" function that runs whenever the recipient list
 * changes: on file upload, on paste, on quick-mode input, on Match Fields apply,
 * and on session restore. It orchestrates:
 *   1. Parse the raw CSV text via parseCSV()
 *   2. Check for an email column (or prompt Match Fields if missing)
 *   3. Apply the field mapping via applyFieldMapping()
 *   4. Enforce the MAX_RECIPIENTS cap
 *   5. Reset all stateful UI (filter, selection, retry button, tag chips)
 *   6. Rebuild the tag bar, test-row selector, and preview table
 *   7. Warn about duplicate email addresses
 */
function parseAndPreview() {
  const raw = document.getElementById("csvInput").value;
  const { headers, rows } = parseCSV(raw);

  // Determine whether the CSV has an email column, accounting for the
  // field mapping (which may remap a differently-named column to "email").
  const emailColMapped = fieldMapping.email || null;
  const hasEmailHeader = headers.includes("email") || (emailColMapped && headers.includes(emailColMapped));

  if (!hasEmailHeader) {
    // No email column found — show the Match Fields modal so the user can
    // map one of their existing columns to the "email" canonical field.
    document.getElementById("mapFieldsBtn").style.display = "";
    if (!fieldMapping.email) {
      // No mapping at all — auto-open the modal as a guided on-boarding step.
      log("No 'email' column detected — opening Match Fields to map columns.", "warning");
      showToast("No 'email' column found. Use Match Fields to map your email column.", "warning");
      openMatchFieldsModal(headers);
    } else {
      // A mapping was set but the mapped column isn't in this CSV — user error.
      log("CSV must contain an 'email' column (or map a column to email).", "error");
      showToast("Your CSV needs an 'email' column, or use Match Fields to map one.", "error");
    }
    return;
  }

  // Apply field mapping: remap user's column names to canonical names.
  const mappedRows = rows.map(row => applyFieldMapping(row));

  // Enforce the per-day recipient cap. Microsoft 365 limits outbound mail to
  // 10,000 recipients per 24 hours per account. Exceeding it results in the
  // account being temporarily blocked from sending — a serious operational risk.
  if (mappedRows.length > MAX_RECIPIENTS) {
    log(`🛑 CSV contains ${mappedRows.length.toLocaleString()} rows — exceeds the Microsoft 365 hard limit of ` +
        `10,000 outbound recipients per 24 hours. Trim your list before sending.`, "error");
    document.getElementById("recipientCount").textContent =
      `${mappedRows.length.toLocaleString()} recipients — OVER LIMIT`;
    parsedRecipients = [];   // Clear so the old (possibly smaller) list can't be sent accidentally
    return;
  }

  // ── Commit the new recipient list ──────────────────────────────────────
  parsedRecipients = mappedRows;
  // Reset sendOutcomes so they don't accumulate across unrelated CSV loads.
  sendOutcomes = [];
  // Clear the active filter — it may reference column names from the old CSV.
  activeFilter = null;
  // Reset to page 0 so the preview table starts at the beginning.
  previewTablePage = 0;
  // Clear stale condition rows from the previous CSV — column names have changed.
  document.getElementById("filterConditions").innerHTML = "";
  filterConditionCount = 0;
  // Reset row selection — the previous selection's indices are meaningless with a new CSV.
  selectedRowIndices = null;
  // Clear the failed list and hide the retry button — they reference old rows.
  failedRecipients = [];
  const _parseRetryBtn = document.getElementById("retryFailedBtn");
  if (_parseRetryBtn) { _parseRetryBtn.classList.add("hidden"); _parseRetryBtn.textContent = ""; }

  // Hide the getting-started banner once the user has loaded a real CSV.
  const bannerEl = document.getElementById("gettingStartedBanner");
  if (bannerEl) bannerEl.classList.add("hidden");

  // Rebuild the test-row selector with the new recipient list.
  // Truncate labels to 28 chars to keep the dropdown manageable.
  // escapeHtml() on CSV data prevents XSS via malicious column values.
  const testRowSel = document.getElementById("testRowSelect");
  if (testRowSel) {
    testRowSel.innerHTML = mappedRows.map(function(r, i) {
      const label = escapeHtml((r.email || r.first_name || "").slice(0, 28));
      return '<option value="' + i + '">Row ' + (i + 1) + ': ' + label + '</option>';
    }).join("");
  }

  // Reveal the "Match Fields" button now that we have actual column headers.
  document.getElementById("mapFieldsBtn").style.display = "";

  // Warn about duplicate email addresses (send to same person twice).
  // This is advisory — duplicates can be removed with the dedup option.
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

  // Rebuild tag chips — remove old CSV column chips and add the new file's columns.
  clearCsvTags();
  headers.forEach(h => {
    const tag = `{{${h}}}`;
    // Don't add a chip for columns already covered by DEFAULT_TAGS (e.g. "email", "first_name").
    if (!DEFAULT_TAGS.includes(tag)) addTagToBar(tag, undefined, false);
  });
  // Always ensure merge_table and unsubscribe_link are available as smart chips,
  // even if those aren't actual CSV column names.
  addTagToBar("{{merge_table}}", "smart");
  addTagToBar("{{unsubscribe_link}}", "smart");

  document.getElementById("recipientCount").textContent =
    `${mappedRows.length} recipient${mappedRows.length !== 1 ? "s" : ""}`;

  populateFilterSortBar(Object.keys(mappedRows[0] || {}).filter(k => !k.startsWith("_")));
  renderPreviewTable(parsedRecipients);
  populateInsertFieldSelect(headers); // Refresh the "Insert field" dropdown in the compose toolbar
  log(`Parsed ${mappedRows.length} recipients. Showing preview.`, "success");
}

/**
 * renderPreviewTable(rows) — render the recipient preview table with paging
 * and row-selection checkboxes.
 *
 * Builds the entire table as an HTML string and assigns it in one innerHTML
 * operation — much faster than appending DOM nodes one by one, especially for
 * large recipient lists. We escapeHtml() every cell value to prevent XSS via
 * malicious CSV data appearing in the table.
 *
 * Paging: PREVIEW_PAGE_SIZE rows per page. The current page number is stored
 * in the module-level previewTablePage variable so it persists across re-renders
 * (e.g. when a filter is applied without changing the page).
 *
 * Row selection: each row has a checkbox with data-idx = the row's _originalIndex.
 * The "select all" header checkbox and individual row checkboxes both update
 * selectedRowIndices. selectedRowIndices === null means "all selected" (the fast
 * path — no Set membership checks needed in the merge loop).
 *
 * @param {Object[]} rows - The rows to render (may be a filtered subset of parsedRecipients)
 */
function renderPreviewTable(rows) {
  const container = document.getElementById("previewTable");
  container.classList.remove("hidden");

  if (!rows || rows.length === 0) {
    container.innerHTML = '<p class="hint" style="padding:4px;">No rows to display.</p>';
    return;
  }

  // Clamp previewTablePage to a valid index after filtering may have shrunk the set.
  const totalPages = Math.ceil(rows.length / PREVIEW_PAGE_SIZE);
  previewTablePage = Math.min(previewTablePage, Math.max(0, totalPages - 1));
  const pageRows = rows.slice(previewTablePage * PREVIEW_PAGE_SIZE, (previewTablePage + 1) * PREVIEW_PAGE_SIZE);

  // Filter out internal "_"-prefixed keys so they don't appear as table columns.
  const headers = Object.keys(rows[0]).filter(h => !h.startsWith("_"));
  // Build the preview header: title label + Edit table button + Filter button.
  // The filter button shows an orange dot (●) when a filter is currently active
  // so users can see at a glance that the table is filtered.
  const filterActiveIndicator = activeFilter ? ' <span class="filter-active-dot" title="Filter active">●</span>' : '';
  // Row 1: preview label + Edit table button (inline, same line as the title).
  // Row 2: Filter button on its own line below, left-aligned under "Preview".
  let html = '<div style="margin-bottom:4px;">' +
    '<div style="display:flex;align-items:center;gap:6px;">' +
    '<label class="label" style="margin:0;">Preview (rows ' + (previewTablePage * PREVIEW_PAGE_SIZE + 1) + '–' + Math.min((previewTablePage + 1) * PREVIEW_PAGE_SIZE, rows.length) + ' of ' + rows.length + ')</label>' +
    '<button class="btn-secondary btn-edit-table" id="openEditTableBtn">✏ Edit table</button>' +
    '</div>' +
    '<div style="margin-top:4px;">' +
    '<button class="btn-filter-popup" id="openFilterBtn" title="Filter and sort recipients">⊞ Filter' + filterActiveIndicator + '</button>' +
    '</div>' +
    '</div>';
  html += '<table class="preview-table"><thead><tr>';
  // "Select all" header checkbox — toggles all visible row checkboxes.
  html += '<th><input type="checkbox" id="selectAllRowsChk" checked title="Select/deselect all" /></th>';
  headers.forEach(h => { html += '<th>' + escapeHtml(h) + '</th>'; });
  html += '</tr></thead><tbody>';
  pageRows.forEach(function(row) {
    // Use _originalIndex (set at parse time) rather than the current iteration
    // index — the current index changes with filtering/sorting, but _originalIndex
    // always refers to the same row in parsedRecipients.
    const originalIdx = row._originalIndex !== undefined ? row._originalIndex : parsedRecipients.indexOf(row);
    const isChecked = (selectedRowIndices === null || selectedRowIndices.has(originalIdx)) ? "checked" : "";
    html += '<tr>';
    html += '<td><input type="checkbox" class="row-select-chk" data-idx="' + originalIdx + '" ' + isChecked + ' /></td>';
    headers.forEach(h => { html += '<td>' + escapeHtml(row[h] || "") + '</td>'; });
    html += '</tr>';
  });
  html += '</tbody></table>';

  // Paging controls — only shown when there's more than one page.
  if (totalPages > 1) {
    html += '<div class="preview-paging">' +
      '<button id="prevPageBtn" class="btn-sm"' + (previewTablePage === 0 ? ' disabled' : '') + '>‹ Prev</button>' +
      '<span>Page ' + (previewTablePage + 1) + ' of ' + totalPages + ' (' + rows.length + ' total)</span>' +
      '<button id="nextPageBtn" class="btn-sm"' + (previewTablePage >= totalPages - 1 ? ' disabled' : '') + '>Next ›</button>' +
      '</div>';
  }

  // Single innerHTML assignment — avoids repeated DOM mutations and reflows.
  container.innerHTML = html;

  // ── Wire paging buttons ────────────────────────────────────────────────
  var prevBtn = document.getElementById("prevPageBtn");
  var nextBtn = document.getElementById("nextPageBtn");
  if (prevBtn) {
    prevBtn.addEventListener("click", function() {
      previewTablePage = Math.max(0, previewTablePage - 1);
      renderPreviewTable(rows); // Re-render the same rows array, new page
    });
  }
  if (nextBtn) {
    nextBtn.addEventListener("click", function() {
      previewTablePage = Math.min(totalPages - 1, previewTablePage + 1);
      renderPreviewTable(rows);
    });
  }

  // ── Wire "Edit table" button ───────────────────────────────────────────
  const openEditBtn = document.getElementById("openEditTableBtn");
  if (openEditBtn) openEditBtn.addEventListener("click", openEditTableModal);

  // ── Wire "Filter" button ───────────────────────────────────────────────
  // Opens the #filterPopup overlay. The popup itself is always in the DOM;
  // we just toggle .hidden to show/hide it.
  const openFilterBtn = document.getElementById("openFilterBtn");
  if (openFilterBtn) openFilterBtn.addEventListener("click", openFilterPopup);

  // ── Wire "Select all" header checkbox ─────────────────────────────────
  document.getElementById("selectAllRowsChk").addEventListener("change", function(e) {
    if (e.target.checked) {
      // null = "all rows selected" (the default state, most efficient).
      selectedRowIndices = null;
      document.querySelectorAll(".row-select-chk").forEach(function(chk) { chk.checked = true; });
    } else {
      // empty Set = "no rows selected".
      selectedRowIndices = new Set();
      document.querySelectorAll(".row-select-chk").forEach(function(chk) { chk.checked = false; });
    }
    updateSelectionCount();
  });

  // ── Wire individual row checkboxes ────────────────────────────────────
  document.querySelectorAll(".row-select-chk").forEach(function(chk) {
    chk.addEventListener("change", function() {
      const idx = parseInt(chk.dataset.idx, 10);
      if (chk.checked) {
        if (selectedRowIndices !== null) {
          selectedRowIndices.add(idx);
          // Optimisation: if now all rows are selected, revert to null
          // (the "all selected" sentinel) to avoid Set membership checks in the loop.
          if (selectedRowIndices.size === parsedRecipients.length) {
            selectedRowIndices = null;
          }
        }
        // If already null (all selected), checking a box is a no-op.
      } else {
        if (selectedRowIndices === null) {
          // Was "all selected" — can't just delete from null.
          // Build a full Set with every index, then remove this one.
          selectedRowIndices = new Set(parsedRecipients.map(function(_, i) { return i; }));
          selectedRowIndices.delete(idx);
        } else {
          selectedRowIndices.delete(idx);
        }
      }
      // Keep the "select all" header checkbox in sync with row state.
      const allChks = document.querySelectorAll(".row-select-chk");
      const allChecked = [...allChks].every(function(c) { return c.checked; });
      const selectAllChk = document.getElementById("selectAllRowsChk");
      if (selectAllChk) selectAllChk.checked = allChecked;
      updateSelectionCount();
    });
  });
}

/**
 * updateSelectionCount() — update the #recipientCount label to reflect the
 * current row-selection state.
 *
 * Shows "N recipients" when all are selected, "M of N selected" when a subset
 * is selected. Called after every checkbox change.
 */
function updateSelectionCount() {
  const total = parsedRecipients.length;
  // selectedRowIndices === null means all selected — no need to count a Set.
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

  // Custom template: user writes their own greeting using {{column}} tokens.
  // Any column from the recipient row is available — e.g. "Dear {{title}} {{last_name}},"
  // for a CSV with a title column containing Mr./Mrs./Dr. etc.
  if (cfg.format === "custom") {
    const tmpl = (cfg.customTemplate || "").trim();
    if (!tmpl) return fallback;
    // Replace {{col}} tokens with the matching field from recipient.
    // recipient already has escaped values (or raw, depending on call site),
    // so we just do a straightforward lookup — no additional escaping here.
    const resolved = tmpl.replace(/\{\{(\w+)\}\}/gi, (_m, key) => {
      const val = String(recipient[key.toLowerCase()] !== undefined && recipient[key.toLowerCase()] !== null
        ? recipient[key.toLowerCase()] : "").trim();
      return val;
    }).trim();
    return resolved || fallback;
  }

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

/**
 * applyFilter(value, filter) — apply a named transformation to a resolved token value.
 *
 * Called by personalize() for tokens written as {{field|filter}}, e.g. {{first_name|upper}}.
 * Built-in filter names:
 *   upper    — ALL CAPS
 *   lower    — all lowercase
 *   title    — Title Case (capitalises the first letter of every word)
 *   trim     — strip leading/trailing whitespace
 *   currency — format as USD currency string using the user's locale (e.g. $1,234.56)
 *   number   — format as a locale number string with thousand separators (e.g. 1,234)
 *   date     — parse value as a date and reformat as long-form (e.g. January 1, 2026)
 *
 * Fallback behaviour: if the filter name doesn't match any of the above AND the value
 * is empty/whitespace, the filter string itself becomes the fallback value. This lets
 * you write {{nickname|Friend}} to output "Friend" when the nickname column is blank.
 * If the value is non-empty, the unknown filter is silently ignored and value is returned.
 *
 * @param {string} value  - the resolved (and already HTML-escaped, if needed) token value
 * @param {string|null} filter - the filter name from the pipe, or null if no pipe was used
 * @returns {string} the transformed value
 */
function applyFilter(value, filter) {
  if (!filter) return value; // No pipe → return value unchanged
  const f = filter.trim().toLowerCase();
  if (f === "upper")    return value.toUpperCase();
  if (f === "lower")    return value.toLowerCase();
  if (f === "title")    return value.replace(/\b\w/g, c => c.toUpperCase()); // \b\w = first char of each word
  if (f === "trim")     return value.trim();
  if (f === "currency") {
    const n = parseFloat(value);
    // parseFloat returns NaN for non-numeric strings — fall back to the raw value if it's not a number
    return isNaN(n) ? value : n.toLocaleString(undefined, { style: "currency", currency: "USD", minimumFractionDigits: 2 });
  }
  if (f === "number") {
    const n = parseFloat(value);
    return isNaN(n) ? value : n.toLocaleString(); // toLocaleString adds thousand separators in the user's locale
  }
  if (f === "date") {
    const d = new Date(value);
    // new Date("invalid") produces an invalid date — getTime() returns NaN, so we can detect it
    return isNaN(d.getTime()) ? value : d.toLocaleDateString(undefined, { year: "numeric", month: "long", day: "numeric" });
  }
  // Unknown filter name = treat as fallback value for empty fields.
  // Allows {{field|Default Text}} to output "Default Text" when the field is blank.
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

/**
 * findUnresolvedTokens(text) — scan text for any {{...}} placeholders that were NOT replaced.
 *
 * Called after personalize() to detect tokens that had no matching CSV column — these would
 * appear literally in the sent email, which is almost always a user error. The Set deduplicates
 * so each unique unresolved token only appears once in the warning log.
 *
 * @param {string} text - the personalized subject or body after all replacements have been applied
 * @returns {string[]} deduplicated array of unresolved token strings, e.g. ["{{middle_name}}"]
 */
function findUnresolvedTokens(text) {
  const matches = text.match(/\{\{[^}]+\}\}/g); // find all remaining {{...}} patterns
  return matches ? [...new Set(matches)] : [];    // deduplicate with Set, then spread back to array
}

/**
 * escapeHtml(str) — the critical XSS defence for the HTML email body.
 *
 * Every value from the CSV that gets inserted into an HTML email body MUST go
 * through this function first. Without it, a CSV cell containing:
 *   <script>alert("hacked")</script>
 * would be injected verbatim into the recipient's email and execute in their browser.
 *
 * The five replacements cover all the characters that have special meaning in HTML:
 *   & → &amp;   (must be first — otherwise subsequent replacements would double-escape)
 *   < → &lt;    (closes open tags, starts new ones)
 *   > → &gt;    (closes open tags)
 *   " → &quot;  (breaks out of attribute values like <img src="...">)
 *   ' → &#039;  (breaks out of single-quoted attributes; &apos; isn't universally supported)
 *
 * Note: String(str) coerces null/undefined/numbers to strings before replacing.
 * Calling .replace() on null would throw a TypeError without this coercion.
 *
 * Subject line values are NOT escaped through this function — the subject is
 * plain text, not HTML, so & in a subject should appear literally, not as &amp;.
 */
function escapeHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")   // Always first — prevents double-escaping of subsequent replacements
    .replace(/</g, "&lt;")    // Prevents opening HTML tags in injected content
    .replace(/>/g, "&gt;")    // Prevents closing HTML tags in injected content
    .replace(/"/g, "&quot;")  // Prevents breaking out of double-quoted HTML attributes
    .replace(/'/g, "&#039;"); // Prevents breaking out of single-quoted HTML attributes
}

/**
 * stripHtmlToText(html) — convert an HTML email body to plain text.
 *
 * Used when plainTextMode is enabled: the user writes their email in HTML compose
 * (so they get formatting tools), but the personalised body is then stripped to plain
 * text before being sent via Graph. This is necessary because Graph's "Text" contentType
 * sends raw text — if you sent the HTML markup literally it would appear as raw tags in
 * the recipient's inbox.
 *
 * The replacements are ordered so that:
 *   1. Block-level elements (<br>, </p>, </div>, </li>) are converted to newlines BEFORE
 *      all tags are stripped — otherwise we'd lose all line breaks entirely.
 *   2. All remaining HTML tags are stripped with a generic <[^>]+> pattern.
 *   3. HTML entities are decoded back to their plain-text characters.
 *   4. Runs of 3+ consecutive newlines are collapsed to 2 (a paragraph break) to avoid
 *      excessive blank lines in the output.
 *
 * @param {string} html - the HTML body string from getComposeBodyAsync
 * @returns {string} plain-text equivalent, suitable for Graph's Text contentType
 */
function stripHtmlToText(html) {
  return html
    .replace(/<br\s*\/?>/gi, "\n")      // <br> / <br /> → newline
    .replace(/<\/p>/gi, "\n\n")         // closing </p> → paragraph break
    .replace(/<\/div>/gi, "\n")         // closing </div> → newline
    .replace(/<\/li>/gi, "\n")          // closing </li> → newline
    .replace(/<[^>]+>/g, "")            // strip all remaining HTML tags
    .replace(/&amp;/g, "&")             // decode HTML entities
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#039;/g, "'")
    .replace(/&nbsp;/g, " ")            // non-breaking space → regular space
    .replace(/\n{3,}/g, "\n\n")         // collapse 3+ blank lines into at most 2
    .trim();                             // remove leading/trailing whitespace from the whole string
}

/* ─── ADDRESS LIST PARSER ──────────────────────────────────────── */

/**
 * parseAddressList(raw) — parse a semicolon/comma/space-separated email address string
 * into the Graph API emailAddress object format.
 *
 * Used in two places:
 *   1. CC/BCC fields from the CSV — each row can have "cc" or "bcc" columns with
 *      multiple addresses separated by any combination of semicolons, commas, or spaces.
 *   2. Suppression/opt-out matching — to extract individual addresses from a multi-TO row.
 *
 * Returns Graph-compatible objects: [{ emailAddress: { address: "foo@bar.com" } }, ...]
 * Only addresses that pass EMAIL_REGEX are included — malformed strings are silently dropped.
 *
 * @param {string} raw - raw address string from a CSV cell or input field
 * @returns {{ emailAddress: { address: string } }[]}
 */
function parseAddressList(raw) {
  if (!raw || !raw.trim()) return [];
  return raw.split(/[;,\s]+/)          // split on any mix of semicolons, commas, or whitespace
    .map(a => a.trim())
    .filter(a => a.length > 0 && EMAIL_REGEX.test(a)) // keep only syntactically valid addresses
    .map(address => ({ emailAddress: { address } }));  // wrap in Graph object format
}

/**
 * parseCustomHeaders(str) — parses the "Custom headers" textarea into Graph-compatible objects.
 *
 * Input format (one header per line):
 *   X-Campaign-ID: summer-2026
 *   X-Mailer: MailMergeAddin/2.6
 *
 * The function returns an array of { name, value } objects that get appended to the
 * Graph sendMail request's internetMessageHeaders array.
 *
 * SECURITY — CRLF injection defence:
 * HTTP headers are separated by \r\n (CRLF). If a user entered a header value like:
 *   X-Foo: bar\r\nBcc: attacker@evil.com
 * without sanitisation, the injected \r\n would create a rogue Bcc header in the raw
 * MIME message and silently copy every email to an attacker. The .replace(/[\r\n]/g, "")
 * calls strip ALL carriage returns and newlines from both the header name and value.
 */
function parseCustomHeaders(str) {
  if (!str || !str.trim()) return []; // Nothing entered — return empty array, not null

  return str.split("\n")           // Split textarea content into individual lines
    .map(line => line.trim())       // Remove leading/trailing whitespace from each line
    .filter(line => line.includes(":")) // Skip blank lines and lines without a colon separator
    .map(line => {
      // Find the first colon — this is the header name/value separator.
      // We use indexOf (not split(":")) so a value like "URL: https://x.com:443"
      // doesn't get split at the port number colon.
      const colonIdx = line.indexOf(":");
      return {
        name:  line.slice(0, colonIdx).trim().replace(/[\r\n]/g, ""),   // S6: strip CRLF injection from name
        value: line.slice(colonIdx + 1).trim().replace(/[\r\n]/g, "")   // S6: strip CRLF injection from value
      };
    })
    .filter(h => h.name && h.value); // Drop any entries where name or value is empty after trimming
}

/* ─── EMAIL VALIDATION ─────────────────────────────────────────── */

/**
 * validateRecipients(recipients) — filters the parsed recipient list before sending.
 *
 * Does three things for each row:
 *   1. skip_if check: if the skip_if CSV column is truthy (non-empty, non-zero,
 *      non-"false", non-"no"), the row is silently excluded from the send.
 *      This lets users mark specific rows to skip without deleting them from the CSV.
 *   2. Email presence check: rows with no email address are logged and skipped.
 *   3. Email format validation: the email field can hold multiple addresses
 *      (semicolon/comma/space-separated). Each address is individually validated
 *      against EMAIL_REGEX. The row is kept if at least ONE address is valid.
 *
 * @param {Array<Object>} recipients - the parsed CSV rows from parsedRecipients
 * @returns {{ valid: Array<Object>, invalid: Array<{row, email}> }}
 */
function validateRecipients(recipients) {
  const valid   = []; // Rows that passed all checks — these get emails sent
  const invalid = []; // Rows that failed email validation — logged in the UI

  for (let i = 0; i < recipients.length; i++) {
    const r = recipients[i];

    // skip_if column: any truthy value (except "0", "false", "no") means skip.
    // We lowercase for case-insensitive comparison — "False", "NO", "FALSE" all skip.
    // String() coerces null/undefined to "".
    const skipVal = String(r.skip_if !== null && r.skip_if !== undefined ? r.skip_if : "").trim().toLowerCase();
    if (skipVal && skipVal !== "0" && skipVal !== "false" && skipVal !== "no") {
      log(`Row ${r._csvRow || (i + 2)} (${r.email || "?"}): skipped — skip_if = "${r.skip_if}"`, "info");
      continue; // Don't add to valid or invalid — it's an intentional skip
    }

    // r._csvRow is stamped by parseCSV to preserve the original 1-based row number
    // (1 = header, so data rows start at 2). We use it in log messages so the user
    // can find the problematic row in their spreadsheet.
    if (!r.email) {
      log(`Row ${r._csvRow || (i + 2)}: missing email address — row skipped.`, "warning");
      continue;
    }

    // Split on any combination of semicolons, commas, or whitespace — this handles
    // both "a@b.com; c@d.com" (Outlook-style) and "a@b.com,c@d.com" (CSV-style).
    const addresses = r.email.split(/[;,\s]+/).map(s => s.trim()).filter(Boolean);

    // Log a warning for each individual invalid address — helpful when one address
    // in a multi-TO row is a typo and the others are fine.
    addresses.forEach(addr => {
      if (!EMAIL_REGEX.test(addr)) {
        log(`Row ${r._csvRow || (i + 2)}: invalid address skipped — "${addr}"`, "warning");
      }
    });

    // A row is valid if at least one of its addresses passes EMAIL_REGEX.
    // .some() stops as soon as it finds the first valid one — efficient.
    const hasValid = addresses.some(a => EMAIL_REGEX.test(a));
    if (hasValid) {
      valid.push(r); // Keep the row — sendBatchWithRetry will filter out individual bad addresses
    } else {
      invalid.push({ row: i + 2, email: r.email }); // All addresses were bad — report to user
    }
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
      // IMPORTANT: the client_id here MUST be the Azure app registration CLIENT_ID
      // (d06ae3cf-...), NOT the Teams manifest add-in ID (a3a648da-...).
      // An admin visiting this URL grants consent for the Azure app to call Graph on
      // behalf of users in the tenant. Using the wrong ID sends the admin to a dead URL.
      return "Admin consent is required before this add-in can send email. " +
             "Your Microsoft 365 administrator must grant permissions. Ask them to visit: " +
             "https://login.microsoftonline.com/common/adminconsent" +
             "?client_id=d06ae3cf-a7da-4264-b20e-ab8d70c06977";
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

/**
 * getAccessToken() — the main auth entry point for every Graph API call.
 *
 * Strategy (two-tier):
 *   Tier 1 — Office SSO (getAccessTokenAsync):
 *     Office silently acquires a token using the host app's (Outlook's) existing
 *     sign-in session. No popup, no redirect — the user never sees it. This works
 *     when the tenant's Entra ID is configured correctly and the user is signed in
 *     with a work/school account. It's the preferred path: fast and invisible.
 *
 *   Tier 2 — MSAL dialog fallback (getTokenViaDialog):
 *     When Office SSO fails for any reason (wrong IdP, federation misconfiguration,
 *     Duo MFA, legacy Outlook version, etc.), we fall back to opening an Office
 *     dialog popup that runs a full MSAL browser-based auth flow. The user sees
 *     the Microsoft sign-in page (and Duo if MFA is required), and the token comes
 *     back via a message from the dialog.
 *
 * Returns a Promise<string> that resolves to an access token string or rejects
 * with a human-readable Error for the user to act on.
 */
function getAccessToken() {
  // We wrap the callback-based Office API in a Promise so callers can use
  // async/await instead of nested callbacks — much cleaner call sites.
  return new Promise((resolve, reject) => {
    Office.context.auth.getAccessTokenAsync(
      {
        allowSignInPrompt: true,    // If the user isn't signed in, show the sign-in UI
        allowConsentPrompt: true,   // If the app needs new Graph scopes, show the consent UI
        forMSGraphAccess: true      // Tells Office we want a Graph-compatible token (v2 endpoint)
      },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          // Happy path: Office SSO worked. result.value is the raw JWT access token string.
          // Pass it straight through — the caller sends it as a Bearer token to Graph.
          resolve(result.value);
        } else {
          const code = result.error.code;

          // Office SSO error codes 13000-13013 all mean "SSO isn't going to work here"
          // for different reasons (wrong IdP, MFA required, not configured, etc.).
          // In every case the right move is to fall back to the MSAL dialog rather than
          // showing the user a cryptic error code. Only completely unexpected codes
          // (outside this range) surface as hard errors.
          //
          // Code 13000: Office can't complete the request — legacy Mac Outlook, missing
          //             admin consent, or Cisco Duo MFA blocking the silent flow.
          // Code 13001: User is not signed in or token expired. Dialog re-authenticates.
          // Code 13002: User cancelled the consent prompt. Dialog lets them try again.
          // Code 13003: Personal Microsoft account — dialog shows work/school sign-in.
          // Code 13004: App registration misconfigured (wrong redirect URI, missing scope).
          // Code 13005: User account not provisioned in tenant (JumpCloud CDI sync issues).
          // Code 13006: Internal Office error — fallback is safer than a hard failure.
          // Code 13007: Token expired; silent refresh failed — dialog re-authenticates.
          // Code 13008: Admin blocked consent prompts — admin must pre-consent via portal.
          // Code 13009: Admin consent required for the org before any user can sign in.
          // Code 13010: Navigation error in the SSO iframe — transient, dialog avoids it.
          // Code 13012: Outlook version too old to support SSO — dialog always works.
          // Code 13013: Too many auth requests in flight — dialog serialises the attempt.
          const fallbackCodes = [13000, 13001, 13002, 13003, 13004, 13005, 13006, 13007, 13008, 13009, 13010, 13012, 13013];

          if (fallbackCodes.includes(code)) {
            log("SSO unavailable (code " + code + ") — opening sign-in dialog…", "warning");
            // Chain the dialog result directly into this Promise — if the dialog
            // succeeds, we resolve with the token; if it fails, we reject with
            // its error. The caller doesn't need to know which path was taken.
            getTokenViaDialog().then(resolve).catch(reject);
          } else {
            // Unknown error code — surface a human-readable message and stop.
            // ssoErrorMessage() maps each known code to plain English; the default
            // case handles anything unforeseen with generic restart advice.
            reject(new Error(ssoErrorMessage(code)));
          }
        }
      }
    );
  });
}

/* ─── MSAL DIALOG FALLBACK (for JumpCloud/federation SSO failures) ─ */

/**
 * getTokenViaDialog() — MSAL browser-based auth inside an Office dialog popup.
 *
 * This is the Tier 2 fallback when Office SSO fails. It works like this:
 *   1. Opens auth-dialog.html in an Office dialog popup (a separate WKWebView on Mac,
 *      a separate window on Windows).
 *   2. auth-dialog.html runs a full MSAL auth flow: if a cached session exists it
 *      returns silently; otherwise it redirects through Microsoft's login page (and
 *      Duo MFA if required), then back to auth-dialog.html with a token.
 *   3. The token travels back to this taskpane via two parallel channels:
 *      a) Office.context.ui.messageParent() — the preferred path. The dialog calls
 *         this once it has the token, and DialogMessageReceived fires here.
 *      b) localStorage — the fallback for legacy Mac Outlook where messageParent
 *         is unavailable after the redirect chain. auth-dialog.html writes the token
 *         to localStorage; this function polls for it on a 500 ms interval.
 *
 * The settle() pattern ensures only ONE of these channels wins. Whichever fires
 * first calls settle(), which sets the `settled` flag, clears the poller, closes
 * the dialog (if not already closed), and calls resolve/reject exactly once.
 * The second channel hitting settle() is a no-op because `if (settled) return`.
 *
 * Returns a Promise<string> that resolves to the access token or rejects with Error.
 */
function getTokenViaDialog() {
  return new Promise((resolve, reject) => {
    // Build the auth dialog URL relative to the current page's origin + path.
    // Using the URL constructor is safer than string manipulation — it handles
    // edge cases like trailing slashes, query strings, or hash fragments already
    // present in window.location.href that would corrupt a naive concatenation.
    // The ?v=7 cache-bust query param forces Office (and browser caches) to fetch
    // a fresh copy of auth-dialog.html on every sign-in attempt, so users never
    // get served a stale cached version after a deployment.
    const dialogUrl = new URL('auth-dialog.html?v=7', window.location.href).href;

    // Remove any stale token from a previous auth attempt before opening the dialog.
    // If we didn't do this, the localStorage poller below might instantly pick up
    // a result from a previous session and resolve with an expired token.
    try { localStorage.removeItem('mm_auth_result'); } catch(e) {}

    // settled: guards against double-resolution. A Promise can only be resolved or
    // rejected once — calling resolve() or reject() a second time is silently ignored
    // by the Promise spec, but relying on that would leave the poller running and the
    // dialog unclosed. settle() is the single safe exit point.
    let settled = false;
    let lsPoller = null;   // Reference to the setInterval so we can cancel it
    let _dialog  = null;   // Reference to the Office dialog object so we can close it

    /**
     * settle(fn, value) — the single exit point for both auth channels.
     * Prevents double-resolution, cleans up the poller, and closes the dialog.
     * @param {Function} fn    - resolve or reject
     * @param {*}        value - the token string (for resolve) or Error (for reject)
     */
    function settle(fn, value) {
      if (settled) return; // Already resolved/rejected by the other channel — bail out
      settled = true;

      // Stop polling localStorage — no need once we have the result
      if (lsPoller) { clearInterval(lsPoller); lsPoller = null; }

      // Clean up the localStorage key so it doesn't persist as a security artefact.
      // The token has a 60-second TTL anyway, but explicit cleanup is best practice.
      try { localStorage.removeItem('mm_auth_result'); } catch(e) {}

      // Close the dialog from the taskpane side. This is specifically needed when the
      // localStorage path resolves first (legacy Mac Outlook), because the dialog's
      // window.close() call fires 1.2 s after writing to localStorage — the dialog
      // may still be open when settle() runs. Calling _dialog.close() here dismisses
      // it immediately rather than leaving it hanging on screen.
      if (_dialog) { try { _dialog.close(); } catch(e) {} _dialog = null; }

      // Call the actual resolve() or reject() with the token or error
      fn(value);
    }

    // ── Channel B: localStorage poller (legacy Mac Outlook fallback) ─────────
    // Legacy Mac Outlook's WKWebViews for taskpane and dialog are on the same origin
    // (leighton-grey.github.io), so they share localStorage. auth-dialog.html writes
    // the token to 'mm_auth_result' after the redirect flow completes. We poll every
    // 500 ms to detect it. This is the fallback for when messageParent (Channel A) is
    // unavailable — typically after a multi-hop redirect chain (Microsoft → Duo → back).
    lsPoller = setInterval(() => {
      try {
        const raw = localStorage.getItem('mm_auth_result');
        if (!raw) return; // Nothing written yet — keep polling

        const msg = JSON.parse(raw); // Expected: { type, token|message, ts }

        // Ignore results older than 60 seconds — they're from a previous auth attempt
        // that wasn't cleaned up. The 60 s window is generous enough to cover slow
        // Duo MFA flows without risking replay of genuinely stale tokens.
        if (!msg.ts || Date.now() - msg.ts > 60000) return;

        // Route based on whether the dialog succeeded or failed
        if (msg.type === 'token') {
          settle(resolve, msg.token); // Token successfully retrieved — resolve the Promise
        } else {
          settle(reject, new Error(msg.message || 'Authentication failed')); // Auth failed
        }
      } catch(e) {
        // JSON.parse or localStorage access failed — silently ignore and keep polling.
        // The dialog will eventually write a complete/valid JSON string.
      }
    }, 500); // Poll every 500 ms — fast enough to feel responsive, cheap enough to not matter

    // ── Open the Office dialog popup ─────────────────────────────────────────
    // displayDialogAsync opens auth-dialog.html in a new Office-managed popup window.
    // height and width are percentages of the screen dimensions, not pixels.
    // promptBeforeOpen: false skips the "Allow this add-in to open a dialog?" confirmation
    // that Office shows by default — users shouldn't need to click through that.
    Office.context.ui.displayDialogAsync(
      dialogUrl,
      { height: 60, width: 35, promptBeforeOpen: false },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          // Office couldn't open the dialog at all (e.g., popups blocked at OS level).
          // Settle as rejected — the lsPoller would run forever otherwise.
          settle(reject, new Error("Could not open sign-in dialog: " + asyncResult.error.message));
          return;
        }

        const dialog = asyncResult.value; // The Office dialog object — used to close it and listen for messages
        _dialog = dialog; // Expose to settle() so the localStorage path can close it

        // ── Channel A: messageParent listener ────────────────────────────────
        // When auth-dialog.html calls Office.context.ui.messageParent(payload), this
        // event fires here in the taskpane with arg.message = the JSON payload string.
        // This is the fast path for modern Outlook — no polling needed.
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          // The dialog has already called messageParent and is closing itself —
          // null out _dialog so settle() doesn't try to call .close() on a dead reference.
          _dialog = null;
          dialog.close(); // Tell Office to close the dialog window from the taskpane side

          try {
            const msg = JSON.parse(arg.message); // Parse the JSON payload from the dialog
            if (msg.type === "token") {
              settle(resolve, msg.token); // Token OK — resolve the outer Promise
            } else {
              settle(reject, new Error(msg.message || "Authentication failed")); // Auth error from dialog
            }
          } catch {
            // JSON.parse failed — the dialog sent something malformed. Shouldn't happen
            // in practice but we handle it gracefully.
            settle(reject, new Error("Invalid response from auth dialog"));
          }
        });

        // ── Dialog closed event ───────────────────────────────────────────────
        // DialogEventReceived fires when the dialog is closed for any reason OTHER than
        // messageParent — specifically:
        //   error 12006 = user manually closed the dialog (X button), OR window.close()
        //                 was called from inside the dialog (our fallback after messageParent
        //                 exhaustion). We report this as "Sign-in cancelled" to the user.
        //   other codes = unexpected closure (crash, navigation to an invalid domain, etc.)
        // Note: if settle() was already called by the localStorage poller, this is a no-op
        // because of the `if (settled) return` guard in settle(). That's the correct
        // behaviour — window.close() fires AFTER localStorage is written, so by the time
        // DialogEventReceived fires here, the Promise is already resolved.
        dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
          _dialog = null; // Dialog is gone — clear the reference to prevent a stale .close() call
          settle(reject, new Error(arg.error === 12006
            ? "Sign-in cancelled"
            : "Dialog closed unexpectedly (error " + arg.error + ")"
          ));
        });
      }
    );
  });
}

/* ─── PRE-FLIGHT CHECKS ────────────────────────────────────────── */

/**
 * checkPayloadSize(bodyTemplate) — warn the user if the email body is approaching the
 * Microsoft Graph API's 4 MB per-message size limit.
 *
 * The Graph /me/sendMail endpoint hard-rejects requests larger than 4 MB. Inline images
 * (embedded as base64 in the HTML body) are the most common cause — a single high-res
 * screenshot can push a template over the limit. This check runs on the template before
 * personalization, so the actual per-recipient payload may be slightly larger after
 * tokens are replaced, but it's a useful early warning.
 *
 * TextEncoder().encode() gives the exact UTF-8 byte count, which is what the network
 * transmits. Using .length on a JS string would give the character count, which
 * under-counts multi-byte Unicode characters.
 *
 * @param {string} bodyTemplate - the raw HTML body from the compose window
 */
function checkPayloadSize(bodyTemplate) {
  const bytes = new TextEncoder().encode(bodyTemplate).length;
  if (bytes > MAX_PAYLOAD_BYTES) {
    log(`⚠️ Email body is ~${(bytes / 1024 / 1024).toFixed(2)} MB — dangerously close to the 4 MB ` +
        `Graph API per-message limit. Compress or remove inline images before sending.`, "warning");
  }
}

/* ─── SENDING WINDOW (Feature 2) ──────────────────────────────── */

/**
 * msUntilWindowOpens() — returns the milliseconds to wait until the configured sending
 * window opens, or 0 if no window is configured / we are already inside the window.
 *
 * The sending window lets users restrict mail merge delivery to business hours (e.g.
 * Mon-Fri 09:00-17:00) to avoid recipients receiving bulk emails at midnight.
 *
 * Logic:
 *   - If the "Sending window" checkbox is unchecked, returns 0 immediately (no restriction).
 *   - If we are currently inside the window (weekday AND time between start and end), returns 0.
 *   - If today is a weekday but we haven't reached the window start yet, returns ms until today's start.
 *   - If we are past the window (or it's a weekend), returns ms until next weekday's window start.
 *
 * The while(true) loop advances `daysAhead` until it lands on a Mon-Fri. With a maximum
 * of 2 weekend days to skip, this loop always terminates in at most 2 iterations.
 *
 * @returns {number} milliseconds to wait (0 = send now)
 */
function msUntilWindowOpens() {
  const sendingWindowEl = document.getElementById("sendingWindowEnabled");
  const enabled = sendingWindowEl && sendingWindowEl.checked;
  if (!enabled) return 0; // Feature disabled — no restriction
  const now = new Date();
  const day = now.getDay(); // 0=Sun, 6=Sat
  const startParts = (document.getElementById("windowStart").value || "09:00").split(":").map(Number);
  const endParts = (document.getElementById("windowEnd").value || "17:00").split(":").map(Number);
  // Convert HH:MM to total minutes-since-midnight for easy comparison
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
  if (inWindow) return 0; // We're inside the window right now — send immediately

  // Calculate ms until next window open
  let next = new Date(now);
  next.setSeconds(0, 0); // zero out sub-minute precision so the wait is exact to the minute
  // If today is a weekday and before window start, wait until today's start
  if (isWeekday && nowMins < startMins) {
    next.setHours(startParts[0], startParts[1], 0, 0);
    return next - now;
  }
  // Otherwise (after end of window, or it's a weekend): find the next weekday
  let daysAhead = 1;
  while (true) {
    const nextDay = (day + daysAhead) % 7; // wraps 0-6 cyclically through the week
    if (nextDay >= 1 && nextDay <= 5) break; // Mon=1 … Fri=5
    daysAhead++;
  }
  next.setDate(next.getDate() + daysAhead);
  next.setHours(startParts[0], startParts[1], 0, 0);
  return next - now;
}

/* ─── RATE LIMITING (Feature 3) ───────────────────────────────── */

// In-memory counters for the current session. Persisted to localStorage by
// saveRateLimitState() so they survive a task pane reload within the same day/hour.
let emailsSentThisHour = 0; // count of emails sent within the current rolling 1-hour window
let emailsSentToday = 0;    // count of emails sent since midnight (day window)
let hourWindowStart = Date.now();                      // epoch ms when the current hourly window opened
let dayWindowStart = new Date().setHours(0, 0, 0, 0); // epoch ms at the start of today (midnight)

/* ─── RATE LIMIT STATE PERSISTENCE (Feature 8) ─────────────────── */

/**
 * loadRateLimitState() — restore rate limit counters from localStorage on startup.
 *
 * Without persistence, closing and reopening the task pane would reset the counters to 0,
 * allowing a user to trivially bypass the hourly cap by reloading. By saving to localStorage
 * on every batch and restoring here, the counters accurately reflect the actual send history
 * for the current hour and day even after a reload.
 *
 * Restore logic:
 *   - Hourly counter: only restored if the saved window started within the last 3600 s.
 *     If it's older, the window has expired and the counter should start fresh.
 *   - Daily counter: only restored if the saved day-window start is >= today's midnight.
 *     This resets correctly across midnight without needing a timer.
 */
function loadRateLimitState() {
  const saved = lsGet("mm_rate_state", null);
  if (saved) {
    const now = Date.now();
    // Only use the saved hourly count if we're still within the same 1-hour window
    if (saved.hourWindowStart && (now - saved.hourWindowStart) < 3600000) {
      emailsSentThisHour = saved.emailsSentThisHour || 0;
      hourWindowStart = saved.hourWindowStart;
    }
    // Only use the saved daily count if the saved window started today (same or later than midnight)
    const todayStart = new Date().setHours(0, 0, 0, 0);
    if (saved.dayWindowStart && saved.dayWindowStart >= todayStart) {
      emailsSentToday = saved.emailsSentToday || 0;
      dayWindowStart = saved.dayWindowStart;
    }
  }
}

/**
 * saveRateLimitState() — persist the current rate limit counters to localStorage.
 * Called once per batch (not per email) to minimise localStorage writes.
 */
function saveRateLimitState() {
  lsSet("mm_rate_state", JSON.stringify({ emailsSentThisHour, emailsSentToday, hourWindowStart, dayWindowStart }));
}

/**
 * checkRateLimits() — enforce hourly and daily send caps, pausing if needed.
 *
 * Called at the top of each batch before sending. Two checks:
 *   1. Daily cap: if emailsSentToday >= dailyCap, stops the merge entirely (returns false).
 *      Unlike the hourly limit, we don't wait — the daily cap is absolute for the day.
 *   2. Hourly cap: if emailsSentThisHour >= maxPerHour, waits until the hour window resets
 *      then returns true so the send continues. The await causes the UI to remain responsive
 *      while paused (unlike a synchronous sleep).
 *
 * Returns true = continue sending, false = stop entirely (daily cap hit).
 *
 * @async
 * @returns {Promise<boolean>}
 */
async function checkRateLimits() {
  // Refresh hour window if the last window opened more than 1 hour ago
  if (Date.now() - hourWindowStart >= 3600000) {
    hourWindowStart = Date.now();
    emailsSentThisHour = 0;
  }
  // Refresh day window if we've crossed midnight since the window was set
  if (Date.now() - dayWindowStart >= 86400000) {
    dayWindowStart = new Date().setHours(0, 0, 0, 0); // reset to today's midnight
    emailsSentToday = 0;
  }

  const maxPerHourEl2 = document.getElementById("maxPerHour");
  const dailyCapEl2   = document.getElementById("dailyCap");
  const maxPerHour = parseInt(maxPerHourEl2 ? maxPerHourEl2.value : "0", 10);
  const dailyCap   = parseInt(dailyCapEl2   ? dailyCapEl2.value   : "0", 10);

  if (dailyCap > 0 && emailsSentToday >= dailyCap) {
    // Daily cap reached — stop the entire merge run, don't just pause
    log(`Daily cap of ${dailyCap} reached. Stopping merge.`, "warning");
    return false;
  }

  if (maxPerHour > 0 && emailsSentThisHour >= maxPerHour) {
    // Hourly cap reached — calculate exact ms until the current window resets and wait
    const msUntilReset = hourWindowStart + 3600000 - Date.now();
    const waitMins = Math.ceil(msUntilReset / 60000);
    log(`⏸ Hourly limit of ${maxPerHour} reached. Pausing ${waitMins} min until hour resets…`, "warning");
    await new Promise(r => setTimeout(r, msUntilReset)); // async wait — UI stays responsive
    hourWindowStart = Date.now(); // start a fresh hour window
    emailsSentThisHour = 0;
    log("▶ Resuming sends (hour window reset).", "info");
  }

  return true;
}

/* ─── GRAPH BATCH SEND ─────────────────────────────────────────── */

/**
 * buildEmailRequest(id, toEmail, opts) — construct a single Graph $batch request object
 * for one recipient.
 *
 * The Graph $batch endpoint accepts an array of up to 20 individual request objects.
 * Each object has an `id` (for correlating responses), `method`, `url`, `headers`, and `body`.
 * This function builds exactly one such object for a single personalised email.
 *
 * The `id` parameter is the 1-based index within the current batch (not the global row number).
 * It's returned verbatim in each batch response so we can map responses back to recipients:
 *   response.id === "3" → recipients[2] (zero-indexed).
 *
 * draftsMode toggle: when draftsMode is true, the URL is /me/messages (creates a draft)
 * rather than /me/sendMail, and the body format changes accordingly:
 *   sendMail: { message: {...}, saveToSentItems: bool }
 *   drafts:   { ...message fields directly (no wrapper) }
 *
 * @param {number} id      - 1-based batch request ID (used to correlate responses)
 * @param {string} toEmail - semicolon-separated To address(es) from the recipient's email column
 * @param {Object} opts    - all email options (see destructuring below for full list)
 * @returns {Object} Graph $batch request object { id, method, url, headers, body }
 */
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

/**
 * effectiveBatchSize(attachmentSizeBytes) — calculate how many emails to put in each
 * Graph $batch request based on attachment size.
 *
 * The Graph $batch endpoint enforces a 4 MB limit per individual request AND a total
 * payload limit for the whole batch POST. Attachments are base64-encoded in the JSON body,
 * which inflates their size by ~37% (the 1.37 multiplier). We add 4096 bytes per email
 * as a rough overhead budget for all the JSON fields (subject, recipients, headers etc.).
 *
 * The formula finds the maximum number of emails N such that N * perEmailBytes <= 3.5 MB.
 * We use 3.5 MB as the target (not 4 MB) to leave a 0.5 MB safety margin. The result is
 * clamped to [1, BATCH_SIZE] so we never send 0 emails or more than the API's maximum.
 *
 * @param {number} attachmentSizeBytes - total attachment bytes (before base64 encoding) per email
 * @returns {number} number of emails per batch (1 to BATCH_SIZE)
 */
function effectiveBatchSize(attachmentSizeBytes) {
  if (!attachmentSizeBytes) return BATCH_SIZE; // No attachments — use the full configured batch size
  const base64Bytes   = Math.ceil(attachmentSizeBytes * 1.37); // base64 adds ~37% overhead
  const perEmailBytes = base64Bytes + 4096;                    // add JSON field overhead per email
  return Math.max(1, Math.min(BATCH_SIZE, Math.floor(3_500_000 / perEmailBytes)));
}

/**
 * buildUnsubHeaders(url) — generate the RFC 8058 List-Unsubscribe and List-Unsubscribe-Post
 * headers for one-click unsubscribe support.
 *
 * Gmail and other inbox providers use these headers to show a prominent "Unsubscribe" button
 * at the top of bulk emails. RFC 8058 defines the one-click format: a POST to the URL with
 * the body "List-Unsubscribe=One-Click". Without both headers, Gmail shows the standard
 * (slower, two-step) unsubscribe flow instead.
 *
 * @param {string} url - the fully-qualified unsubscribe URL (from listUnsubscribeTemplate after personalization)
 * @returns {{ name: string, value: string }[]} pair of Graph internetMessageHeader objects
 */
function buildUnsubHeaders(url) {
  // Strip CR/LF to prevent header injection (CRLF injection / RFC 5322 folding attack)
  const safeUrl = String(url).replace(/[\r\n]/g, "");
  return [
    { name: "List-Unsubscribe",      value: `<${safeUrl}>` },
    { name: "List-Unsubscribe-Post", value: "List-Unsubscribe=One-Click" }
  ];
}

/**
 * chunkArray(arr, size) — split an array into sub-arrays of at most `size` elements.
 * Used to split the full recipients array into batches of BATCH_SIZE for Graph $batch.
 */
function chunkArray(arr, size) {
  const chunks = [];
  for (let i = 0; i < arr.length; i += size) chunks.push(arr.slice(i, i + size));
  return chunks;
}

/**
 * delay(ms) — return a Promise that resolves after `ms` milliseconds.
 * Using await delay(ms) is the idiomatic async equivalent of a blocking sleep —
 * it yields execution back to the event loop so the UI stays responsive during waits.
 */
function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

/**
 * getJitterMs() — return a random jitter delay in milliseconds within the user's configured range.
 *
 * Jitter adds randomness between batches so that a large send doesn't hit the Graph API
 * with perfectly uniform timing. Exchange Online's throttling is partly time-based, so
 * spreading requests randomly (rather than at a fixed interval) reduces the chance of
 * triggering sustained throttling even when average send rate is within limits.
 *
 * The min/max values are in seconds in the UI — we multiply by 1000 to get milliseconds.
 * Math.min/max guards against the user entering min > max (swaps them silently).
 *
 * @returns {number} a random millisecond delay in [lo*1000, hi*1000), or 0 if jitter is disabled
 */
function getJitterMs() {
  const minEl = document.getElementById("jitterMinInput");
  const maxEl = document.getElementById("jitterMaxInput");
  const min = parseInt(minEl ? minEl.value : "0", 10);
  const max = parseInt(maxEl ? maxEl.value : "0", 10);
  if (!min && !max) return 0; // Both zero = jitter disabled
  const lo = Math.min(min, max); // Ensure lo <= hi even if user entered them backwards
  const hi = Math.max(min, max);
  return (lo + Math.random() * (hi - lo)) * 1000; // Random ms in [lo, hi)
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

/**
 * sendScheduledMessages(...) — send personalised emails with Exchange deferred delivery.
 *
 * Unlike immediate batch send (which uses /me/sendMail directly), scheduled send must
 * go through a two-step process:
 *   Step 1: POST to /me/messages — creates a draft with a singleValueExtendedProperty
 *           that sets the deferred delivery time (MAPI property "SystemTime 0x000F").
 *           This is the standard Exchange/Outlook "delay delivery" mechanism.
 *   Step 2: POST to /me/messages/{id}/send — instructs Exchange to send the draft.
 *           Exchange then holds it in the Outbox until the scheduled time, even if
 *           Outlook is closed. This is why "Save to Sent" can't be suppressed — Exchange
 *           always saves deferred sends.
 *
 * Because each email is two separate API calls (not batchable — /me/messages/{id}/send
 * isn't supported in Graph $batch for deferred delivery), scheduled send is inherently
 * slower than immediate batch send and runs one email at a time.
 *
 * Per-row send_at override: if a CSV row has a `send_at` column with a valid date/time,
 * that overrides the global scheduled time for that specific recipient. This allows
 * different recipients to receive their email at different times within one merge run.
 *
 * Auth error handling: if token acquisition fails mid-run, all remaining recipients are
 * pushed to failedRecipients for retry, and the loop breaks. A 403 error (Mail.ReadWrite
 * not granted) causes an immediate abort for all remaining recipients, since every
 * subsequent request would also fail.
 *
 * @async
 * @param {Array<Object>} recipients    - validated recipient rows
 * @param {string} subjectTemplate      - subject line with {{tokens}}
 * @param {string} bodyTemplate         - HTML body with {{tokens}}
 * @param {boolean} saveToSent          - ignored (Exchange always saves), but logged as warning
 * @param {string} replyTo              - Reply-To address
 * @param {string} sendAs               - Send As address (from field override)
 * @param {string} scheduledTimeISO     - ISO 8601 global scheduled delivery time
 * @param {string} importance           - "low" | "normal" | "high"
 * @param {boolean} isReadReceiptRequested
 * @param {boolean} isDeliveryReceiptRequested
 * @param {Array} customHeaders         - parsed custom MIME headers
 * @param {boolean} plainText           - strip HTML to plain text before sending
 * @param {string} sensitivity          - "normal" | "personal" | "private" | "confidential"
 * @param {string[]} categories         - Outlook category names
 * @param {boolean} bccSelf             - BCC the sender's own address
 * @param {string} selfEmail            - sender's email for BCC self
 * @param {boolean} flagged             - flag sent items for follow-up
 * @param {string|null} expiryISO       - ISO 8601 message expiry time
 * @param {string} listUnsubscribeTemplate - URL template for List-Unsubscribe header
 * @param {Array} inlineImagesArr       - inline CID images from the compose window
 * @param {number} globalOffset         - index of first recipient in the full valid array (for record_num)
 * @param {number} totalRecipientCount  - total valid recipient count across all sends (for record_count)
 * @returns {Promise<{totalSent: number, totalFailed: number}>}
 */
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

/**
 * sendImmediateBatch(...) — send personalised emails immediately using Graph $batch.
 *
 * This is the main send path for non-scheduled, non-broadcast sends (and also for
 * drafts mode). It batches up to BATCH_SIZE (20) emails per Graph $batch POST.
 *
 * High-level flow per batch:
 *   1. Pre-personalise all emails in the current batch and build Graph request objects.
 *   2. Acquire an OAuth2 token (re-authenticates if needed).
 *   3. POST the batch to Graph's $batch endpoint via sendBatchWithRetry().
 *   4. Parse the response array — 2xx = success, 429 = throttled (handled in sendBatchWithRetry),
 *      401 = token expired (abort remaining), other = failure.
 *   5. Update send counters, push failed recipients for retry, record outcomes for the report.
 *   6. Wait batchDelayMs + jitter before the next batch.
 *
 * The `allValid` parameter is the full valid recipient array (across all send paths).
 * It's used to compute the correct global `record_num` and `record_count` tokens for each
 * recipient even in mixed scheduled/immediate runs where `recipients` is a subset of `allValid`.
 *
 * effectiveBatchSize() reduces batch size when attachments are present to stay within
 * Graph's 4 MB per-batch limit — see that function for the calculation.
 *
 * @async
 * @param {Array<Object>} recipients          - recipients to send to in this call
 * @param {string} subjectTemplate            - subject with {{tokens}}
 * @param {string} emailBodyTemplate          - HTML body with {{tokens}}
 * @param {boolean} saveToSent                - include in Sent Items
 * @param {string} replyTo                    - Reply-To address
 * @param {string} sendAs                     - Send As address
 * @param {string} importance                 - "low" | "normal" | "high"
 * @param {boolean} isReadReceiptRequested
 * @param {boolean} isDeliveryReceiptRequested
 * @param {Array} globalCustomHeaders         - parsed custom MIME headers
 * @param {boolean} plainTextMode             - strip HTML
 * @param {string} sensitivity
 * @param {string[]} categories
 * @param {boolean} bccSelf
 * @param {string} selfEmail
 * @param {boolean} flagged
 * @param {string|null} expiryISO
 * @param {string} listUnsubscribeTemplate
 * @param {Array} inlineImagesArr
 * @param {Array<Object>} allValid            - full validated recipient list for record_num computation
 * @returns {Promise<{totalSent: number, totalFailed: number}>}
 */
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

/**
 * showPreSendConfirmation(recipients, batchDelayMs, scheduledCount) — show a summary modal
 * before sending, giving the user one last chance to review recipient count, first recipient
 * preview, and estimated time.
 *
 * Skipped for single sends (count <= 1) to avoid interrupting quick test sends.
 *
 * The time estimate accounts for both the inter-batch delay configured by the user AND
 * the round-trip time for each Graph API batch call (~2 s). Without the batch round-trip
 * estimate, the displayed time would be too low for large sends with no delay configured.
 *
 * A breakdown of immediate vs. scheduled recipients is shown when both are present
 * (i.e. some rows have a send_at column and some don't).
 *
 * Returns a Promise<boolean> that resolves true (confirm) or false (cancel) based on
 * which button the user clicks.
 *
 * @param {Array<Object>} recipients    - the full list of valid recipients about to be sent
 * @param {number} batchDelayMs         - configured delay between batches in milliseconds
 * @param {number} scheduledCount       - number of recipients with a send_at column value
 * @returns {Promise<boolean>}
 */
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
    try { firstSubjectPreview = personalize(subjectTemplate, first, false); } catch(e) { firstSubjectPreview = subjectTemplate; }
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

/**
 * trapFocus(modal) — constrain keyboard Tab focus to within a modal dialog.
 *
 * WCAG accessibility guideline 2.1.2 requires that keyboard focus not leave a modal
 * while it is open. Without a trap, pressing Tab from the last focusable element inside
 * the modal would move focus to elements behind it in the DOM — the modal might as well
 * not exist for keyboard users.
 *
 * Implementation: intercept Tab and Shift+Tab keydown events. When focus reaches the
 * first or last focusable element in the modal and the user Tabs further, we intercept
 * e.preventDefault() and manually move focus to the other end of the focusable list.
 *
 * Returns a cleanup function. Callers MUST call it when closing the modal to remove
 * the event listener — failing to do so would leave a handler attached to a hidden
 * element that fires on every Tab press.
 *
 * @param {HTMLElement} modal - the modal DOM element to trap focus within
 * @returns {Function} cleanup function that removes the keydown listener
 */
function trapFocus(modal) {
  const focusable = modal.querySelectorAll(
    'button:not([disabled]), input:not([disabled]), select:not([disabled]), textarea:not([disabled]), [tabindex="0"]'
  );
  if (!focusable.length) return function() {}; // Empty modal — nothing to trap, return no-op cleanup
  const first = focusable[0];
  const last  = focusable[focusable.length - 1];
  first.focus(); // Move focus into the modal immediately on open

  function handler(e) {
    if (e.key !== "Tab") return; // Only intercept Tab, pass everything else through
    if (e.shiftKey) {
      // Shift+Tab at the first element → wrap to the last
      if (document.activeElement === first) { e.preventDefault(); last.focus(); }
    } else {
      // Tab at the last element → wrap to the first
      if (document.activeElement === last) { e.preventDefault(); first.focus(); }
    }
  }
  modal.addEventListener("keydown", handler);
  return function() { modal.removeEventListener("keydown", handler); }; // cleanup function
}

// Map of modalId → { prevFocus, releaseTrap }
// Stores pre-open focus state and the cleanup function for each currently-open modal.
// Multiple modals can be open simultaneously (e.g. confirm inside check errors), so we
// use a Map rather than a single variable.
const _modalTrapState = new Map();

/**
 * _openModalWithTrap(modalId) — show a modal and set up its focus trap.
 * Saves the element that had focus before the modal opened so _closeModalWithTrap
 * can restore it when the modal is dismissed (required by WCAG 2.4.3).
 */
function _openModalWithTrap(modalId) {
  const modal = document.getElementById(modalId);
  if (!modal) return;
  const prevFocus = document.activeElement; // snapshot focus before modal steals it
  modal.classList.remove("hidden");
  const releaseTrap = trapFocus(modal); // sets up Tab trap and moves focus into modal
  _modalTrapState.set(modalId, { prevFocus, releaseTrap }); // store for cleanup
}

/**
 * _closeModalWithTrap(modalId) — hide a modal and release its focus trap.
 * Restores focus to the element that was active before the modal opened.
 */
function _closeModalWithTrap(modalId) {
  const modal = document.getElementById(modalId);
  if (modal) modal.classList.add("hidden");
  const state = _modalTrapState.get(modalId);
  if (state) {
    state.releaseTrap(); // remove the keydown event listener
    // Restore focus to wherever the user was before opening the modal
    if (state.prevFocus && typeof state.prevFocus.focus === "function") state.prevFocus.focus();
    _modalTrapState.delete(modalId); // clean up the Map entry
  }
}

/* ─── SIMPLE CONFIRM MODAL HELPER ─────────────────────────────── */

/**
 * showSimpleConfirm(message) — show the confirmModal with a generic Yes/No prompt.
 *
 * Reuses the main confirm modal but changes the send button label to "Yes" and
 * restores the original text when the modal closes. This lets non-send confirmations
 * (like "Clear opt-out list?") share the same modal without a separate HTML element.
 *
 * Resolves true on Yes, false on No/dismiss (via dismissModal / confirmSend).
 *
 * @param {string} message - the question to display in the modal body
 * @returns {Promise<boolean>}
 */
function showSimpleConfirm(message) {
  document.getElementById("confirmText").textContent = message;
  const sendBtn = document.getElementById("confirmSendBtn");
  const savedText = sendBtn.textContent; // remember original label (e.g. "Send" or "Save as drafts")
  sendBtn.textContent = "Yes";           // generic label for non-send confirmations
  _openModalWithTrap("confirmModal");
  return new Promise(function(resolve) {
    _confirmModalPrevFocusSaved = true;
    pendingMergeResolve = function(result) {
      sendBtn.textContent = savedText; // restore original label before resolving
      resolve(result);
    };
  });
}
let _confirmModalPrevFocusSaved = false;

/* ─── CONFIRMATION MODAL ───────────────────────────────────────── */

// pendingMergeResolve holds the Promise resolve function for the currently-open confirm modal.
// confirmSend() and dismissModal() call it to resolve the modal's Promise. Using a module-level
// variable (rather than closure) means only one confirm modal can be open at a time — which is
// correct, since we don't ever want two overlapping confirmations.
let pendingMergeResolve = null;

/**
 * showConfirmModal(rowCount, addressCount, scheduledTimeISO, broadcastMode) — show the
 * "are you sure you want to send?" confirmation modal before a merge or broadcast run.
 *
 * Builds a human-readable summary string that describes exactly what is about to happen:
 *   - Normal merge: "This will send 42 emails." (or "42 emails (reaching 50 addresses)" for multi-TO rows)
 *   - Scheduled merge: "This will schedule 42 emails to be sent at ..."
 *   - Drafts mode: "This will save 42 emails as drafts ..."
 *   - Broadcast mode: "Send 1 broadcast email to 42 recipients via BCC?"
 *
 * @param {number} rowCount           - number of individual emails (CSV rows) to send
 * @param {number} addressCount       - total number of unique To addresses (>rowCount for multi-TO rows)
 * @param {string|null} scheduledTimeISO - ISO 8601 scheduled delivery time, or null for immediate
 * @param {boolean} broadcastMode     - true for broadcast (BCC) sends
 * @returns {Promise<boolean>}        - true = user confirmed, false = user cancelled
 */
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

/**
 * confirmSend() — called by the "Send" button inside the confirmation modal.
 * Resolves the pending Promise with true and closes the modal.
 * We null out pendingMergeResolve before calling resolve() to prevent a second
 * click from triggering a double-resolution (which the Promise spec ignores but
 * would leave the modal in an inconsistent state).
 */
function confirmSend() {
  const resolve = pendingMergeResolve;
  pendingMergeResolve = null;
  _closeModalWithTrap("confirmModal"); // Feature 15: release trap + restore focus
  if (resolve) resolve(true);
}

/**
 * dismissModal() — called by the "Cancel" button or backdrop click in the confirmation modal.
 * Resolves the pending Promise with false (user cancelled the merge).
 */
function dismissModal() {
  const resolve = pendingMergeResolve;
  pendingMergeResolve = null;
  _closeModalWithTrap("confirmModal"); // Feature 15: release trap + restore focus
  if (resolve) resolve(false);
}

/* ─── CANCEL ───────────────────────────────────────────────────── */

/**
 * handleStop() — set the cancelRequested flag so the merge loop breaks at the next
 * batch boundary. We don't abort the current in-flight batch — that would leave some
 * emails sent and others not, with no way to know which. Instead we finish the current
 * batch cleanly and skip all subsequent batches.
 */
function handleStop() {
  cancelRequested = true;
  log("Stop requested — will halt after current batch.", "warning");
  document.getElementById("stopBtn").disabled = true; // Prevent double-stop clicks
}

/**
 * setMergeRunning(running) — toggle the UI between "merge in progress" and "idle" states.
 *
 * When running = true:
 *   - Disables the Send, Test, Drafts, and Broadcast buttons to prevent concurrent runs.
 *   - Changes the Send button text to indicate the merge is in progress.
 *   - Shows the Stop button so the user can abort.
 *   - Resets the progress bar to 0% and clears the _mergeCompletedSuccessfully flag.
 *
 * When running = false:
 *   - Re-enables all action buttons.
 *   - Hides the Stop button.
 *   - If _mergeCompletedSuccessfully is true (showMergeComplete was called), the progress bar
 *     stays visible showing "Done" — we don't call hideProgress() in that case.
 *
 * Called twice per merge: once at the start of sendBatchWithRetry (running=true) and once
 * in the finally block (running=false) to guarantee the UI is always restored even on error.
 */
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
    // Feature 5: only hide progress if we did NOT just call showMergeComplete —
    // if we did call it, the progress bar should stay at 100% showing the "Done" state
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
      showToast("No valid recipients found. Make sure your CSV has an 'email' column with valid addresses.", "error");
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
    showToast("No valid email addresses found. Check your recipient list.", "error");
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
    document.getElementById("recipientCount").textContent = "0 recipients";
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

/**
 * csvField(val) — safely format a value for inclusion in a CSV file (RFC 4180).
 *
 * A value must be quoted if it contains a comma (would split into multiple columns),
 * a double-quote (would break quoting), or a newline (would split into multiple rows).
 * When quoting, existing double-quotes inside the value are escaped by doubling them ("").
 *
 * This is used when building the simulation CSV download, the send report CSV,
 * and when importing contacts into the CSV textarea.
 *
 * @param {*} val - any value (will be coerced to string)
 * @returns {string} RFC 4180-safe CSV field
 */
function csvField(val) {
  const s = String(val !== null && val !== undefined ? val : "");
  if (s.includes(",") || s.includes('"') || s.includes("\n")) {
    // Wrap in double-quotes and double any existing double-quotes inside
    return '"' + s.replace(/"/g, '""') + '"';
  }
  return s; // No special characters — safe to use as-is
}

/* ─── GROUPED ROW MERGE (Feature 4) ───────────────────────────── */

/**
 * groupRecipientsByEmail(recipients) — merge multiple CSV rows with the same email address
 * into a single recipient object with a `_groupedRows` array.
 *
 * This is the "many-to-one" merge feature: a single email gets all rows for that recipient.
 * The primary (first) row's fields are used for the top-level token replacements; all rows
 * (including the primary) are available in _groupedRows for use with {{merge_table}}.
 *
 * Use case: a recipient purchased 3 items — each row has the same email but different
 * product names. The merge sends one email containing a table of all 3 products.
 *
 * The Map key is the lowercase email to match case-insensitively (foo@BAR.com and foo@bar.com
 * are the same person and should be grouped together).
 *
 * @param {Array<Object>} recipients - the validated recipient list
 * @returns {Array<Object>} deduplicated list with _groupedRows added to grouped recipients
 */
function groupRecipientsByEmail(recipients) {
  const map = new Map();
  for (const r of recipients) {
    const key = (r.email || "").toLowerCase().trim();
    if (!key) {
      // BUG 12: log skipped rows with whitespace-only email instead of silently dropping
      log(`Row skipped in group merge: empty email address (row data: ${JSON.stringify(r).slice(0, 80)}).`, "warning");
      continue;
    }
    if (!map.has(key)) map.set(key, { primary: r, rows: [r] }); // first occurrence = primary
    else map.get(key).rows.push(r);                              // subsequent occurrences added to group
  }
  // Spread primary row fields and add _groupedRows array containing all rows for that email
  return Array.from(map.values()).map(({ primary, rows }) => ({ ...primary, _groupedRows: rows }));
}

/**
 * buildMergeTable(rows, headers) — build an HTML table from the grouped rows of a recipient.
 *
 * Used by personalize() to replace {{merge_table}} when a recipient has _groupedRows.
 * Excludes operational columns (email, cc, bcc, etc.) that would clutter the table —
 * only data columns meaningful to the recipient are included.
 *
 * @param {Array<Object>} rows    - all CSV rows for one recipient (from _groupedRows)
 * @param {string[]} headers      - column names from the first grouped row
 * @returns {string} HTML table string safe to embed directly in the email body
 */
function buildMergeTable(rows, headers) {
  // Exclude meta columns that are operational (not content for the recipient)
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

// localStorage key for the persistent opt-out list (separate from the run-level suppression set
// which lives in suppressionSet). The opt-out list is permanent across sessions; the suppression
// set is loaded from a CSV and cleared when a new CSV is loaded.
const LS_KEY_OPTOUT = "mm_optout_list";

/**
 * getOptOutList() — return the persistent opt-out list as a lowercase Set for fast lookup.
 * Lowercase normalisation ensures foo@BAR.com matches foo@bar.com in the exclusion check.
 */
function getOptOutList() {
  return new Set((lsGet(LS_KEY_OPTOUT, []) || []).map(e => e.toLowerCase().trim()));
}

/**
 * saveOptOutList(set) — persist the opt-out Set to localStorage and re-render the UI list.
 * Always stores the full Set (not a delta) so it's safe to call with any state of the list.
 */
function saveOptOutList(set) {
  lsSet(LS_KEY_OPTOUT, JSON.stringify(Array.from(set)));
  renderOptOutList(); // update the visible list immediately after every change
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
      if (!parsedRecipients.length) { log("Load a CSV first.", "warning"); showToast("Load a CSV file first from the Recipients tab.", "warning"); return; }
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

/**
 * collectFillInValues(template) — scan the template for {{fill_in:prompt}} tokens,
 * show a modal asking the user to fill in each one, and return the values.
 *
 * The fill-in feature is for values that are the SAME for all recipients but are not
 * known ahead of time — e.g. a conference date, a promo code, or a personalised opening
 * line the sender wants to type fresh for each batch.
 *
 * How it works:
 *   1. Find all unique {{fill_in:prompt}} tokens in the combined subject+body template.
 *   2. Dynamically build form fields inside the fillInModal (one per unique prompt).
 *   3. Show the modal and await the user's OK/Cancel.
 *   4. Return a Map<promptName, value> so applyFillInValues() can substitute them.
 *
 * Returns null if the user cancels (which aborts the merge run).
 * Returns the Map even if the user left some fields blank — blanks substitute to "".
 *
 * @param {string} template - combined subject + body template string
 * @returns {Promise<Map<string, string>|null>} map of prompt→value, or null if cancelled
 */
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

/**
 * checkDomainDns(domain) — query Cloudflare's DNS-over-HTTPS API for SPF, DMARC, and DKIM
 * records for the sender's domain.
 *
 * Missing email authentication records are a common reason bulk emails get rejected or land
 * in spam. We check for all three records before large sends:
 *   SPF (Sender Policy Framework): TXT record at the domain root with "v=spf1"
 *   DMARC (Domain-based Message Auth): TXT record at _dmarc.domain with "v=DMARC1"
 *   DKIM (DomainKeys Identified Mail): CNAME records at selector1._domainkey.domain
 *                                      or selector2._domainkey.domain (Microsoft 365 convention)
 *
 * We use Cloudflare's DNS-over-HTTPS (cloudflare-dns.com/dns-query) because:
 *   - Office.js add-ins run in a sandboxed WKWebView / IE11 WebView with no direct DNS access.
 *   - The browser's built-in DNS resolver isn't accessible via JavaScript.
 *   - Cloudflare's DoH API is reliable, fast, and free for this usage pattern.
 *
 * Each check is in a try/catch — a network error on one check shouldn't block the others.
 * Results are advisory only (warnings in the log, not hard errors).
 *
 * @param {string} domain - e.g. "company.com" (extracted from the sender's email address)
 * @returns {Promise<{spf: boolean, dmarc: boolean, dkim: boolean}>}
 */
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

// Contacts picker state — module-level so filterContacts() and importSelectedContacts()
// can access the full list loaded by handleImportContacts() without re-fetching.
let contactsData = [];              // raw Graph contacts loaded from /me/contacts
let selectedContacts = new Set();   // email addresses of checked contacts in the picker UI

/**
 * handleImportContacts() — fetch the user's Outlook contacts from Graph and show the picker modal.
 *
 * Uses BUG 8's pagination fix: fetches up to 500 contacts per page and follows @odata.nextLink
 * until all contacts are loaded. Without pagination, contacts beyond the first 500 would be silently
 * omitted — easy to miss for users with large address books.
 *
 * Only contacts with at least one email address are included (contacts with only phone numbers
 * or physical addresses are not useful for mail merge).
 *
 * @async
 */
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
      `<p style="padding:8px;color:#eb5757;">Failed to load contacts: ${escapeHtml(err.message)}</p>`;
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

/**
 * recordSentEmails(emails) — record a batch of sent email addresses in localStorage
 * with the current timestamp, pruning entries older than 30 days.
 *
 * The send history is a simple { emailAddress: ISO-timestamp } object in localStorage.
 * Before every merge, checkDuplicateSendHistory() scans this object and warns the user
 * if any of their recipients appear in the history.
 *
 * Pruning prevents unbounded growth: we delete any entry where the stored ISO timestamp
 * is more than 30 days ago. Using numeric comparison (getTime() < cutoffMs) rather than
 * string comparison is important — ISO date strings do sort lexicographically correctly,
 * but being explicit about numeric comparison prevents subtle timezone bugs.
 *
 * @param {string[]} emails - array of raw email strings (may be semicolon-separated multi-TO)
 */
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

/**
 * handleBroadcast() — send one email to all recipients at once via BCC (not personalised).
 *
 * Unlike the mail merge run (one personalised email per row), broadcast sends a single
 * email with all recipients in the BCC field. This is appropriate for newsletters or
 * announcements that are the same for every recipient. Recipients cannot see each other
 * (BCC), which preserves privacy.
 *
 * Key differences from handleMergeClick():
 *   - Subject and body are NOT personalised — no {{tokens}} are replaced.
 *   - All recipients go in bccRecipients; the To field is set to the sender (or a
 *     "Undisclosed Recipients" placeholder if the sender's email is unavailable).
 *   - Per-recipient attachments are NOT supported (shared attachments and inline images are).
 *   - Uses a single /me/sendMail call (not $batch) — one HTTP request for all recipients.
 *
 * Hard limit: Graph's /me/sendMail endpoint caps bccRecipients at 999 addresses.
 * Exchange Online's own limit is 500. We warn at 500 and hard-stop at 999.
 *
 * broadcastInProgress re-entrancy guard: prevents a second broadcast from being launched
 * while one is in flight (analogous to mergeInProgress for the mail merge path).
 *
 * @async
 */
async function handleBroadcast() {
  if (broadcastInProgress) {
    log("A broadcast is already in progress.", "warning");
    return;
  }
  broadcastInProgress = true;
  try {
  const rawAll = getFilteredSortedRecipients();
  if (!rawAll.length) { log("No recipients loaded.", "warning"); showToast("No recipients loaded. Load a CSV file from the Recipients tab first.", "warning"); return; }

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

  if (!validEmails.length) { log("No valid recipients for broadcast.", "warning"); showToast("No valid recipients for broadcast. Check your recipient list.", "warning"); return; }

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

/**
 * updateSchedulingBadge() — show/hide the "Scheduling active" badge on the Options tab.
 * The badge gives the user a visible reminder that scheduled send or sending window is on,
 * so they don't forget and wonder why their merge is paused mid-run.
 */
function updateSchedulingBadge() {
  const active = document.getElementById("scheduleEnabled")?.checked ||
                 document.getElementById("sendingWindowEnabled")?.checked;
  const badge = document.getElementById("schedulingActiveBadge");
  if (badge) badge.classList.toggle("hidden", !active);
}

/**
 * updateRateLimitBadge() — show/hide the "Rate limit active" badge on the Options tab.
 * Visible whenever hourly or daily cap is set to a non-zero value.
 */
function updateRateLimitBadge() {
  const mphVal = parseInt(document.getElementById("maxPerHour")?.value || "0", 10);
  const dcVal  = parseInt(document.getElementById("dailyCap")?.value   || "0", 10);
  const active = mphVal > 0 || dcVal > 0;
  const badge  = document.getElementById("rateLimitActiveBadge");
  if (badge) badge.classList.toggle("hidden", !active);
}

/* ─── SIMULATE / DRY-RUN MODE (Feature 10) ─────────────────────── */

/**
 * handleSimulate() — run the personalisation logic for every recipient and export the
 * results as a CSV file for review, WITHOUT sending any emails.
 *
 * Each output row shows:
 *   record_num        - the row's position in the merged list
 *   email             - the recipient's email address
 *   display_name      - the recipient's name (display_name or first_name)
 *   subject_preview   - the personalised subject line (first 80 chars)
 *   body_tokens_missing - any {{tokens}} that remained unresolved (joined by "; ")
 *   skip_if_result    - "SKIP" if the row would be skipped, "send" otherwise
 *   attachment_found  - "check manually" if an attachment column exists, "n/a" otherwise
 *
 * The `await new Promise(r => setTimeout(r, 0))` yield every 50 rows prevents the UI
 * from freezing on large lists (P1 fix) — it gives the browser's event loop a chance to
 * process pending UI updates between batches of rows.
 *
 * @async
 */
async function handleSimulate() {
  if (!parsedRecipients.length) { log("Load a CSV first.", "warning"); showToast("Load a CSV file first from the Recipients tab.", "warning"); return; }
  const valid = getFilteredSortedRecipients();
  if (!valid.length) { log("No recipients match current filter.", "warning"); showToast("No recipients match the current filter. Try clearing your filters.", "warning"); return; }

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

/**
 * searchDirectory(query) — search the Microsoft 365 Global Address List (GAL) using the
 * Graph /me/people endpoint and render the results in the directory picker UI.
 *
 * /me/people uses Microsoft's "People API" which combines the GAL, contacts, and
 * frequent collaborators into a relevance-scored list. The ConsistencyLevel: "eventual"
 * header is required for $search queries — without it the API returns a 400 error.
 *
 * Requires the People.Read permission on the Entra app registration. If that permission
 * is missing (403), we show a specific error message rather than a generic "search failed"
 * so the IT admin knows exactly what to grant.
 *
 * Minimum 2 characters before querying to avoid triggering rate limits with single-letter
 * searches (and because the results for 1-char queries are typically not useful).
 *
 * @async
 * @param {string} query - the search string (at least 2 characters)
 */
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
    el.innerHTML = `<p style="padding:8px;color:#eb5757;">Search failed: ${escapeHtml(err.message)}</p>`;
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

/**
 * populateInsertFieldSelect(headers) — rebuild the "Insert field" dropdown with the
 * current CSV headers plus the built-in smart tokens.
 *
 * Called whenever the CSV is parsed or the edit table is saved so the dropdown stays
 * in sync with available fields. Smart tokens (greeting_line, today, etc.) are listed
 * first with a ✨ prefix to distinguish them from CSV column tokens.
 *
 * P2: builds all option HTML as a string then assigns to innerHTML once, rather than
 * calling appendChild() or innerHTML += in a loop (which causes repeated DOM reparse).
 *
 * @param {string[]} headers - CSV column names from the parsed recipient list
 */
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

/**
 * handleCheckErrors() — scan the current subject and body template for token issues
 * and display a detailed report modal before the user sends.
 *
 * Three types of issues are detected:
 *   1. Unknown tokens: {{tokens}} used in the template that don't match any CSV column
 *      or built-in smart token. These will appear literally in the sent email.
 *   2. Blank values: known tokens where some CSV rows have an empty value for that column.
 *      The token will be replaced with "" for those rows (the email goes out without the value).
 *   3. Conditional tokens ({{if:...}}) are parsed separately to extract the column name from
 *      the condition expression so they can also be checked for unknown/blank values.
 *
 * fill_in: tokens and merge_table are excluded from the unknown token check because
 * they are resolved at send time (fill_in: by user input, merge_table by grouped rows).
 *
 * The report is rendered as innerHTML in the checkErrorsModal. All user-generated content
 * (token names, email addresses) is passed through escapeHtml() to prevent XSS.
 *
 * @async
 */
async function handleCheckErrors() {
  if (!parsedRecipients.length) {
    log("Load a CSV first before checking for errors.", "warning");
    showToast("Load a CSV file from the Recipients tab before running the pre-flight check.", "warning");
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

/**
 * openEditTableModal() — copy the current parsedRecipients into editTableHeaders and
 * editTableRows (flat arrays for easier grid manipulation), then render the editable grid.
 *
 * The edit table is an in-memory copy — it doesn't modify parsedRecipients until the user
 * clicks Save. This means Cancel discards all changes cleanly without needing to undo them.
 *
 * Internal tracking columns that start with "_" (like _csvRow, _originalIndex) are excluded
 * from the editable grid since they're implementation details, not user-facing data.
 */
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

/**
 * renderEditTable() — build and inject the full HTML table grid into #editTableGrid.
 *
 * Each column header is an editable input (for renaming) with a delete button.
 * Each data cell is an editable input that writes back to editTableRows on "input" event.
 * Row numbers are displayed in the first column as read-only labels for reference.
 *
 * Performance note: the entire grid is rebuilt from scratch as a single innerHTML assignment
 * rather than incrementally updating individual cells. For a typical recipient list
 * (hundreds of rows) this is fast enough and much simpler than partial updates. If the
 * table gets very large (thousands of rows) this could be optimised with virtualisation.
 */
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
      const trimmed = this.value.trim().toLowerCase();
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

/**
 * flushEditTableInputs() — read the current live values from all column header and cell
 * inputs in the DOM and write them back into editTableHeaders and editTableRows.
 *
 * This is needed before any structural operation (add/delete column, save) because the
 * user may have typed in a cell without triggering a "change" event yet. Without flushing,
 * those edits would be lost when the grid is re-rendered.
 */
function flushEditTableInputs() {
  const grid = document.getElementById("editTableGrid");
  grid.querySelectorAll(".edit-col-header-input").forEach(function(input) {
    const col = parseInt(input.dataset.col, 10);
    const trimmed = input.value.trim().toLowerCase();
    if (trimmed) editTableHeaders[col] = trimmed;
  });
  grid.querySelectorAll(".edit-cell-input").forEach(function(input) {
    editTableRows[parseInt(input.dataset.row, 10)][parseInt(input.dataset.col, 10)] = input.value;
  });
}

/**
 * editTableAddColumn() — add a new empty column to the in-memory edit table.
 * Reads the column name from #newColNameInput, validates it (non-empty, no duplicate),
 * appends the column to every row in editTableRows, and re-renders the grid.
 * After rendering, scrolls horizontally to the new column and focuses its first cell.
 */
function editTableAddColumn() {
  const input = document.getElementById("newColNameInput");
  const name  = (input ? input.value.trim().toLowerCase() : "");
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

/**
 * saveEditTableChanges() — apply the edit table's in-memory state back to parsedRecipients
 * and refresh all dependent UI components.
 *
 * The original parsedRecipients is replaced entirely with new objects built from
 * editTableHeaders and editTableRows. Internal tracking fields (_originalIndex, _csvRow)
 * are preserved from the original rows so that row number references in send logs and
 * the preview table remain correct after editing.
 *
 * After saving, refreshAfterTableEdit() updates the recipient count, testRowSelect dropdown,
 * tag bar, filter/sort bar, and re-renders the preview table.
 */
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

/**
 * refreshAfterTableEdit() — update all UI components that depend on parsedRecipients
 * after the edit table is saved.
 *
 * Must be called after any operation that modifies parsedRecipients directly:
 *   - Edit table save
 *   - Any future in-place editing feature
 *
 * Resets paging state (previewTablePage = 0) and selectedRowIndices (null = all selected)
 * to avoid stale state after a structural change to the recipient list.
 */
function refreshAfterTableEdit() {
  const headers = editTableHeaders.slice();
  document.getElementById("recipientCount").textContent =
    parsedRecipients.length + ' recipient' + (parsedRecipients.length !== 1 ? 's' : '');
  const testRowSel = document.getElementById("testRowSelect");
  if (testRowSel) {
    testRowSel.innerHTML = parsedRecipients.map(function(r, i) {
      const label = escapeHtml((r.email || r.first_name || "").slice(0, 28));
      return '<option value="' + i + '">Row ' + (i + 1) + ': ' + label + '</option>';
    }).join("");
  }
  headers.forEach(function(h) {
    const tag = '{{' + h + '}}';
    if (!DEFAULT_TAGS.includes(tag)) addTagToBar(tag, undefined, false);
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
