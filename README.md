# Classic Mail Merge Engine — New Outlook Add-in

A free, open-source, serverless mail merge add-in for **New Outlook** (Windows, Mac, Web). Replaces the classic Word-to-Outlook mail merge feature. Built with Office.js and Microsoft Graph API.

---

## How it works

1. User opens a new compose window in New Outlook
2. Clicks **"Run Merge"** in the toolbar to open the task pane
3. Types or pastes CSV recipient data directly into the pane
4. Composes their email with `{{placeholder}}` tags
5. Clicks **"Run mail merge"** — emails are sent individually via Microsoft Graph

All data is processed **100% locally in the user's browser**. Zero recipient data is transmitted to or stored on GitHub Pages servers.

---

## Prerequisites

- Microsoft 365 tenant with Exchange Online
- JumpCloud SSO federated with Microsoft 365 (SAML/OIDC)
- Microsoft Entra ID (Azure AD) admin access
- Microsoft 365 Admin Center access

---

## Step 1 — Host on GitHub Pages

1. Fork or clone this repository
2. Replace all `<YOUR-GITHUB-USERNAME>` placeholders in `manifest.json` with your GitHub username
3. Go to your repo **Settings → Pages**
4. Set source to **main branch / root folder**
5. Your add-in will be live at `https://<YOUR-GITHUB-USERNAME>.github.io/mail-merge-addin/`

---

## Step 2 — Register the App in Microsoft Entra ID

1. Go to [https://entra.microsoft.com](https://entra.microsoft.com)
2. Navigate to **Applications → App registrations → New registration**
3. Configure:
   - **Name:** Classic Mail Merge Engine
   - **Supported account types:** Accounts in any organizational directory (Multitenant)
   - **Redirect URI:** Single-page application (SPA) → `https://<YOUR-GITHUB-USERNAME>.github.io/mail-merge-addin/taskpane.html`
4. Click **Register**
5. Go to **API permissions → Add a permission → Microsoft Graph → Delegated**
6. Add `Mail.Send`
7. Click **Grant admin consent**
8. Copy the **Application (client) ID** from the Overview page — you will need this if implementing auth code flow in future versions

> **Note:** This add-in uses Office.js implicit SSO (`getAccessTokenAsync`) — the user's active Outlook session provides the token. No client secret is needed or stored.

---

## Step 3 — Deploy via M365 Admin Center to JumpCloud Groups

1. Log into [https://admin.microsoft.com](https://admin.microsoft.com)
2. Go to **Settings → Integrated apps → Add-ins**
3. Click **Deploy Add-in → Upload custom apps**
4. Upload your `manifest.json` file
5. On the **Assign users** screen:
   - Do **not** select "Everyone"
   - Search for the JumpCloud-synced group (e.g. `M365-MailMerge-Users`)
6. Set deployment to **Fixed (Defaults to On)**
7. Save and publish

> The add-in will appear automatically in New Outlook for assigned users within 24 hours — no user action required.

---

## CSV Format

The first row must be column headers. An `email` column is required. All other columns become available as `{{placeholder}}` tags.

```
email,first_name,last_name,company,title
john.smith@example.com,John,Smith,Acme Corp,Manager
sara.jones@example.com,Sara,Jones,Globex,Director
```

---

## Supported Placeholders

| Tag | Column name |
|---|---|
| `{{first_name}}` | first_name |
| `{{last_name}}` | last_name |
| `{{email}}` | email |
| `{{company}}` | company |
| `{{title}}` | title |
| `{{any_column}}` | any column header from your CSV |

---

## Batching & Throttle Protection

- Recipients are split into **batches of 20** (Microsoft Graph hard limit)
- A **1.5 second delay** is applied between batches
- 261 recipients = ~14 batches = ~21 seconds total
- Live activity log shows per-batch progress and any failures

---

## Security & Privacy Compliance

| Concern | Implementation |
|---|---|
| Authentication | Office.js SSO — short-lived bearer tokens via user's active M365 session. No passwords or secrets stored. |
| Authorisation | `Mail.Send` delegated scope only — principle of least privilege |
| Data privacy | All CSV data lives in browser RAM only. Never transmitted to GitHub Pages. Cleared when task pane closes. |
| XSS prevention | Strict Content Security Policy meta tag. All user input is HTML-escaped before injection into email body. |
| Transport | All Graph API calls use HTTPS/TLS only |
| Device trust | JumpCloud Conditional Access Policies remain fully active — if a device is untrusted, M365 token request is blocked at the JumpCloud layer |
| Audit trail | Every sent email passes through the user's Exchange outbox — visible in Message Trace and Content Search in M365 Admin Center |

---

## JumpCloud + M365 Federation Notes

- Authentication routes: **User → JumpCloud IdP → Microsoft 365 (SP) → Outlook → Office.js token**
- No JumpCloud configuration is required for the add-in itself
- Use JumpCloud user groups synced to M365 to control who receives the add-in via M365 Admin Center deployment
- Conditional Access Policies configured in JumpCloud automatically apply to all Graph API token requests made by this add-in

---

## License

MIT — free to use, modify, and distribute.
