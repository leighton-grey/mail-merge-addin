# Security Reference — Mail Merge Add-in

Last reviewed: August 2026

This document records security decisions, mitigations, and known trade-offs that would be relevant in an audit. It is not exhaustive — only items with real-world risk or a deliberate "accept" decision are listed.

---

## Fixes Applied

### S1 — XSS via error message in contacts panel
**File:** `taskpane.js`  
**Risk:** Error messages injected into `innerHTML` without sanitisation could execute attacker-controlled script if a Graph API error response contained HTML.  
**Fix:** Wrapped `err.message` in `escapeHtml()` before insertion.

### S2 — `unsafe-inline` in auth dialog CSP
**File:** `auth-dialog.html`, `auth-dialog.js`  
**Risk:** `'unsafe-inline'` in `script-src` allows any injected inline script to execute — it negates the XSS protection CSP is meant to provide.  
**Fix:** All inline `<script>` blocks extracted to `auth-dialog.js`. MSAL loaded dynamically from JS (no inline `onerror` attribute). `'unsafe-inline'` removed from `script-src`.  
**Note:** `style-src` still carries `'unsafe-inline'` — required because Duo's Universal Prompt injects `<style>` elements at runtime. This is acceptable given we control no user-supplied CSS.

### S4 — CRLF injection in List-Unsubscribe header
**File:** `taskpane.js` — `buildUnsubHeaders()`  
**Risk:** If the unsubscribe URL template contained `\r\n`, an attacker with control over that field could inject arbitrary email headers.  
**Fix:** `String(url).replace(/[\r\n]/g, "")` applied before the URL is placed into the header value. `parseCustomHeaders()` applies the same pattern to all custom X-headers.

### C1 — Double HTML-escaping in pre-send confirmation modal
**File:** `taskpane.js`  
**Risk:** Subject line showed `&lt;FirstName&gt;` instead of `<FirstName>` — merge values were escaped by `personalize()` then escaped again by `escapeHtml()` in the modal.  
**Fix:** `personalize(subjectTemplate, first, false)` — raw values returned, then `escapeHtml()` at the display layer handles escaping once.

---

## Accepted Trade-offs

### S3 — No SRI hash on SheetJS (cdnjs)
**File:** `taskpane.html`  
**What an auditor will flag:** The SheetJS `<script>` tag on cdnjs lacks an `integrity` attribute. SRI ensures the browser rejects the file if it has been tampered with at the CDN.  
**Why accepted:** cdnjs is Cloudflare-operated with strong supply-chain controls. The add-in is a private corporate tool — blast radius from a cdnjs compromise is limited and the library version is pinned. Adding SRI introduces maintenance overhead (hash must be updated on every version bump) for low practical gain.  
**If risk appetite changes:** Download `xlsx.full.min.js` locally, run `shasum -a 256`, encode as base64, and add `integrity="sha256-<hash>"` plus `crossorigin="anonymous"` to the `<script>` tag.

---

## Persistent Security Architecture

### XSS defence — `escapeHtml()`
All user-supplied or API-sourced strings inserted into `innerHTML` must go through `escapeHtml()`. The `personalize()` function accepts a third argument `escapeValues` (default `true`) that applies `escapeHtml()` to every merge token before substitution. Pass `false` only for plain-text contexts (subject line, header values) where the output will be HTML-escaped at the point of display.

### Header injection — `parseCustomHeaders()`
All custom X-headers entered by the user are validated through `parseCustomHeaders()`, which strips `\r\n` from both name and value. `buildUnsubHeaders()` does the same for the List-Unsubscribe URL.

### Token security — access tokens
Access tokens are stored in `sessionStorage` (MSAL cache) and `localStorage` (dual-channel auth bridge). Both are cleared on sign-out. Tokens are scoped to `Mail.Send` and `User.Read` only — no admin-level Graph permissions are requested.

### Auth dialog isolation
The auth dialog runs in a separate Office WebView process. Communication back to the taskpane uses two channels: `Office.context.ui.messageParent()` (primary) and a `localStorage` bridge (fallback for legacy Mac Outlook after a redirect chain). The taskpane's "settle" pattern ensures only the first channel to resolve is acted on — the second is silently ignored.

### LICENSE_ENFORCEMENT
`LICENSE_ENFORCEMENT = false` — the add-in is distributed via Admin Centre / EAC and does not enforce a licence check at runtime. **Do not change this without explicit instruction from the repository owner.** Enabling enforcement would break deployment for all managed tenants.

---

## Out of Scope (by design)

- **Server-side validation:** This is a client-side Office add-in. All sends go directly from the user's mailbox via Microsoft Graph. There is no backend to validate against.
- **Rate limiting:** Graph API enforces its own throttling. The add-in batches requests and jitters sends to stay within limits.
- **Data at rest:** No CSV data is persisted to a server. Recipients exist only in browser memory and `localStorage` for settings (greeting config, opt-out list). `localStorage` is scoped to the GitHub Pages origin.
