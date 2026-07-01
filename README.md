# career-ops-plugin-outlook-interviews

A [career-ops](https://github.com/santifer/career-ops) community plugin that reads your Outlook mail via Microsoft Graph, detects interview invitation emails, and surfaces them in the career-ops pipeline.

Most enterprise recruiters (aerospace, banking, Big 4, media) send interview invites from Outlook/Exchange — this closes the gap for job seekers interviewing with them. Completes the mailbox trio: `google-calendar` (interviews from Calendar), `linkedin-alerts` (job alerts from Gmail), and this plugin (interview invites from Outlook).

## What it does

- Exchanges your OAuth refresh token for a short-lived access token (no SDK, pure REST — supports public and confidential Azure app registrations)
- Fetches recent messages via Microsoft Graph with read-only `Mail.Read` scope
- Filters by interview keywords in English and Chinese (interview, prescreen, hiring manager, 面试, …)
- Extracts company from the sender ("Marie Jansa (Royal Conservatory)" → Royal Conservatory, or from the domain), role from the subject, and the meeting link (Teams / Zoom / Meet / Whereby) from the body
- Returns `Job[]` for the career-ops `ingest` hook — tracker updates stay human-in-the-loop

## Install

```bash
node plugins.mjs install https://github.com/Schlaflied/career-ops-plugin-outlook-interviews
```

## Setup

### 1. Azure App Registration

1. [portal.azure.com](https://portal.azure.com/) → **App registrations** → New registration
2. Supported account types: *Personal Microsoft accounts* (or your org tenant)
3. Add platform **Mobile and desktop applications** (public client — no secret needed)
4. Note the **Application (client) ID**

### 2. Get a refresh token

Run any one-time OAuth2 authorization-code flow with scope:
```
https://graph.microsoft.com/Mail.Read offline_access
```
Copy the `refresh_token` from the token response.

### 3. Add to `.env`

```
MSGRAPH_CLIENT_ID=your-application-client-id
MSGRAPH_REFRESH_TOKEN=your-refresh-token
# Only for confidential (Web) app registrations:
# MSGRAPH_CLIENT_SECRET=your-client-secret
```

### 4. Enable

```bash
node plugins.mjs enable outlook-interviews --confirm
```

## Optional config (`config/plugins.yml`)

```yaml
outlook-interviews:
  tenant: common      # Azure AD tenant id (default: common — works for personal + work accounts)
  days: 7             # look-back window in days (default: 7)
  maxResults: 50      # max emails per fetch (default: 50)
  allEmails: false    # true = return all emails, skips interview-keyword filter
```

## Privacy

All API calls go through `ctx.fetch` and are limited to `login.microsoftonline.com` and `graph.microsoft.com`. Read-only mail scope. No data leaves your machine to any third-party service. Credentials stay in your local `.env`.

## License

MIT
