# career-ops-plugin-outlook-interviews

Detects interview invitation emails in Outlook via Microsoft Graph and surfaces them in the career-ops pipeline. Works with personal outlook.com accounts and enterprise Azure AD tenants.

## Setup

### 1. Register an Azure app

1. Go to [portal.azure.com](https://portal.azure.com/) → **App registrations** → New registration
2. Supported account types: *Personal Microsoft accounts* (or your org tenant for enterprise)
3. Add platform **Mobile and desktop applications** (public client — no secret needed)
4. Copy the **Application (client) ID**

### 2. Get a refresh token

Run a one-time OAuth2 authorization-code flow with scope:
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

## Optional settings (`config/plugins.yml`)

```yaml
outlook-interviews:
  tenant: common      # Azure AD tenant (default: common — personal + work accounts)
  days: 7             # how many days back to scan (default: 7)
  maxResults: 50      # max emails to fetch (default: 50)
  allEmails: false    # true = return all emails, false = interview keywords only
```

## How it works

The `ingest` hook:
1. Exchanges your refresh token for an access token (public or confidential client)
2. Fetches recent messages via Microsoft Graph (`Mail.Read`, read-only)
3. Filters by interview keywords (EN + ZH): interview, prescreen, phone screen, hiring manager, recruiter, etc.
4. Extracts company (sender display name / domain), role (subject line), and meeting link (Teams / Zoom / Meet / Whereby)
5. Returns `Job[]` for the career-ops pipeline
