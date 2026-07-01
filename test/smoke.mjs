// Smoke test — verifies plugin contract without hitting Microsoft APIs
import plugin from '../index.mjs';

let passed = 0;
let failed = 0;

function assert(condition, msg) {
  if (condition) { console.log(`  ✅ ${msg}`); passed++; }
  else           { console.error(`  ❌ ${msg}`); failed++; }
}

const SAMPLE_MESSAGES = [
  {
    subject: 'Interview for Learning Operations Advisor',
    from: { emailAddress: { name: 'Olga Sivanesan (Bombardier)', address: 'olga.sivanesan@bombardier.com' } },
    receivedDateTime: '2026-07-01T14:00:00Z',
    bodyPreview: 'Hi, join us via Teams: https://teams.microsoft.com/l/meetup-join/abc123 on July 2 at 2:00 PM EST',
    webLink: 'https://outlook.live.com/mail/deeplink/1',
  },
  {
    subject: 'Prescreen — Talent Management Coordinator',
    from: { emailAddress: { name: 'Recruiting Team', address: 'talent@wildbrain.com' } },
    receivedDateTime: '2026-07-01T10:00:00Z',
    bodyPreview: 'We would like to speak with you. Zoom: https://us02.zoom.us/j/12345',
    webLink: 'https://outlook.live.com/mail/deeplink/2',
  },
  {
    subject: 'Your weekly newsletter digest',
    from: { emailAddress: { name: 'Substack', address: 'no-reply@substack.com' } },
    receivedDateTime: '2026-07-01T08:00:00Z',
    bodyPreview: 'Top stories this week in publishing and media.',
    webLink: 'https://outlook.live.com/mail/deeplink/3',
  },
];

function mockCtx(settings = {}) {
  let callCount = 0;
  let tokenBody = '';
  const httpStub = async (url, options) => {
    callCount++;
    if (url.includes('login.microsoftonline.com')) {
      tokenBody = options?.body || '';
      return { ok: true, json: async () => ({ access_token: 'test-token' }), text: async () => '' };
    }
    if (url.includes('graph.microsoft.com') && url.includes('/me/messages')) {
      return { ok: true, json: async () => ({ value: SAMPLE_MESSAGES }), text: async () => '' };
    }
    throw new Error(`Unexpected request: ${url}`);
  };
  return {
    env: {
      MSGRAPH_CLIENT_ID:     'client-id-test',
      MSGRAPH_REFRESH_TOKEN: 'refresh-token-test',
    },
    settings,
    fetch: httpStub,
    getCallCount: () => callCount,
    getTokenBody: () => tokenBody,
  };
}

console.log('career-ops-plugin-outlook-interviews smoke test\n');

// 1. Plugin shape
console.log('1. Plugin shape');
assert(typeof plugin === 'object',          'default export is object');
assert(typeof plugin.ingest === 'function', 'exports ingest hook');

// 2. ingest filters non-interview emails
console.log('\n2. ingest hook — keyword filtering');
const ctx  = mockCtx();
const jobs = await plugin.ingest(ctx);
assert(Array.isArray(jobs), 'returns array');
assert(jobs.length === 2,   'filters out newsletter email (got ' + jobs.length + ')');
assert(jobs.every(j => j.title !== undefined && j.url !== undefined && j.company !== undefined), 'jobs have title/url/company');

// 3. Meeting link extraction
console.log('\n3. Meeting link extraction');
assert(jobs[0].url === 'https://teams.microsoft.com/l/meetup-join/abc123', 'extracts Teams link from body');
assert(jobs[1].url === 'https://us02.zoom.us/j/12345',                     'extracts Zoom link from body');
assert(jobs[0].location === 'Remote', 'meeting link implies Remote location');

// 4. Company extraction
console.log('\n4. Company extraction');
assert(jobs[0].company === 'Bombardier', 'extracts company from sender display name parenthetical');
assert(jobs[1].company === 'wildbrain',  'derives company from sender domain');

// 5. Role extraction
console.log('\n5. Role extraction');
assert(jobs[0].title === 'Learning Operations Advisor',    'extracts role from "Interview for X"');
assert(jobs[1].title.includes('Talent Management Coordinator'), 'extracts role from prescreen subject');

// 6. allEmails setting bypasses filter
console.log('\n6. allEmails setting');
const allCtx  = mockCtx({ allEmails: true });
const allJobs = await plugin.ingest(allCtx);
assert(allJobs.length === 3, 'allEmails:true returns all messages');

// 7. OAuth flow
console.log('\n7. OAuth flow');
assert(ctx.getCallCount() >= 2, 'at least 2 http calls (token + messages)');
assert(!ctx.getTokenBody().includes('client_secret'), 'public client: no client_secret sent when env not set');

const secretCtx = mockCtx();
secretCtx.env.MSGRAPH_CLIENT_SECRET = 'shh';
await plugin.ingest(secretCtx);
assert(secretCtx.getTokenBody().includes('client_secret=shh'), 'confidential client: client_secret sent when env set');

console.log(`\n${passed + failed} checks — ${passed} passed, ${failed} failed`);
if (failed > 0) process.exit(1);
