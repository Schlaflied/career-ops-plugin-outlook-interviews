// career-ops-plugin-outlook-interviews
// Author: Schlaflied · https://github.com/Schlaflied
// License: MIT · https://github.com/Schlaflied/career-ops-plugin-outlook-interviews
//
// Reads recent Outlook mail via Microsoft Graph, detects interview
// invitation emails, extracts company / role / meeting link, and returns
// them as Job[] for the career-ops pipeline. Works with personal
// outlook.com accounts and enterprise Azure AD tenants.
// Network access only via ctx.fetch (engine enforces allowedHosts).

const AUTH_BASE  = 'https://login.microsoftonline.com';
const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const SCOPES     = 'https://graph.microsoft.com/Mail.Read offline_access';

// EN + ZH interview keywords (mirrors the google-calendar plugin)
const INTERVIEW_KEYWORDS = [
  'interview', 'prescreen', 'pre-screen', 'phone screen', 'screening',
  'hiring manager', 'recruiter', 'recruitment', 'talent acquisition',
  'video call', 'virtual interview', 'teams meeting', 'zoom invite',
  'we would like to speak', 'next steps in your application', 'panel',
  '面试', '筛选', '电话', '视频', '招聘',
];

async function getAccessToken(ctx) {
  const tenant = ctx.settings?.tenant ?? 'common';
  const params = {
    client_id:     ctx.env.MSGRAPH_CLIENT_ID,
    refresh_token: ctx.env.MSGRAPH_REFRESH_TOKEN,
    grant_type:    'refresh_token',
    scope:         SCOPES,
  };
  // Confidential clients send a secret; public (Desktop/Mobile) clients don't
  if (ctx.env.MSGRAPH_CLIENT_SECRET) params.client_secret = ctx.env.MSGRAPH_CLIENT_SECRET;

  const res = await ctx.fetch(`${AUTH_BASE}/${tenant}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams(params).toString(),
  });
  if (!res.ok) throw new Error(`Token exchange failed ${res.status}: ${await res.text()}`);
  const data = await res.json();
  if (!data.access_token) throw new Error('No access_token in response');
  return data.access_token;
}

async function listMessages(ctx, accessToken) {
  const days  = ctx.settings?.days ?? 7;
  const since = new Date(Date.now() - days * 24 * 60 * 60 * 1000).toISOString();

  // Keyword matching happens client-side: combining contains() chains with
  // $orderby trips Graph's InefficientFilter limits on the messages endpoint.
  const params = new URLSearchParams({
    '$filter':  `receivedDateTime ge ${since}`,
    '$select':  'subject,from,receivedDateTime,bodyPreview,webLink',
    '$orderby': 'receivedDateTime desc',
    '$top':     String(ctx.settings?.maxResults ?? 50),
  });

  const res = await ctx.fetch(`${GRAPH_BASE}/me/messages?${params}`, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) throw new Error(`Graph API ${res.status}: ${await res.text()}`);
  const data = await res.json();
  return data.value || [];
}

function isInterviewEmail(msg) {
  const subject = (msg.subject || '').toLowerCase();
  const preview = (msg.bodyPreview || '').toLowerCase();
  return INTERVIEW_KEYWORDS.some(kw => subject.includes(kw) || preview.includes(kw));
}

function extractMeetingLink(body) {
  const patterns = [
    /https:\/\/teams\.microsoft\.com\/[^\s"<>)]+/,
    /https:\/\/[a-z0-9]+\.zoom\.us\/[^\s"<>)]+/,
    /https:\/\/meet\.google\.com\/[^\s"<>)]+/,
    /https:\/\/whereby\.com\/[^\s"<>)]+/,
  ];
  for (const p of patterns) {
    const m = body.match(p);
    if (m) return m[0].split('"')[0].split(')')[0];
  }
  return '';
}

// "Marie Jansa (Royal Conservatory)" → Royal Conservatory; else derive from
// the sender's domain, skipping freemail and ATS relay domains.
function extractCompany(senderName, senderAddress) {
  const parenMatch = (senderName || '').match(/\(([^)]+)\)/);
  if (parenMatch) return parenMatch[1].trim();

  const domainMatch = (senderAddress || '').match(/@([a-z0-9.-]+\.[a-z]{2,})/i);
  if (!domainMatch) return senderName || senderAddress || '';
  const domain = domainMatch[1].toLowerCase();
  const relay = [
    'gmail.com', 'yahoo.com', 'hotmail.com', 'outlook.com',
    'greenhouse.io', 'lever.co', 'ashbyhq.com', 'myworkday.com', 'icims.com',
  ];
  if (relay.includes(domain)) return senderName || senderAddress.split('@')[0];
  return domain.replace(/\.(com|ca|io|org|net|co)$/, '').replace(/-/g, ' ');
}

function extractRole(subject) {
  const patterns = [
    /interview.*?(?:for|re:|:)\s*(.+?)(?:\s*[-–|@]|$)/i,
    /(?:role|position|opportunity):\s*(.+?)(?:\s*[-–|@]|$)/i,
    /(?:hiring|recruiting).*?for\s+(.+?)(?:\s*[-–|@]|$)/i,
    /^(.+?)\s+interview/i,
  ];
  for (const p of patterns) {
    const m = subject.match(p);
    if (m && m[1].length < 80) return m[1].trim();
  }
  const stripped = subject.replace(/interview|prescreen|phone screen/gi, '').trim();
  return stripped.slice(0, 60) || subject.trim();
}

function messageToJob(msg) {
  const senderName    = msg.from?.emailAddress?.name || '';
  const senderAddress = msg.from?.emailAddress?.address || '';
  const meetingLink   = extractMeetingLink(msg.bodyPreview || '');
  return {
    title:    extractRole(msg.subject || ''),
    url:      meetingLink || msg.webLink || '',
    company:  extractCompany(senderName, senderAddress),
    location: meetingLink ? 'Remote' : '',
    // Extra fields passed through for pipeline context
    receivedAt: msg.receivedDateTime || '',
    sender:     senderAddress,
  };
}

export default {
  async ingest(ctx) {
    const accessToken = await getAccessToken(ctx);
    const messages    = await listMessages(ctx, accessToken);

    const interviewEmails = ctx.settings?.allEmails
      ? messages
      : messages.filter(isInterviewEmail);

    return interviewEmails.map(messageToJob);
  },
};
