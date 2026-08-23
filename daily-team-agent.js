/*
  daily-team-agent.js

  Agentic pipeline: GATHER -> PLAN -> DRAFT -> CRITIQUE -> ACT -> LOG

  Requirements:
  - Uses dotenv for secrets
  - Uses MSAL client credentials flow for Graph API
  - Calls Claude (Anthropic) for planning, drafting, critique
  - Logs each step to console and agent-log.json
  - Updates sre-topics-log.json to avoid repeats within 30 days

  Note: Fill environment variables in a .env file or via environment.
*/

const fs = require('fs-extra');
const path = require('path');
const axios = require('axios');
const { ConfidentialClientApplication } = require('@azure/msal-node');
require('dotenv').config();

// --- Config and helpers ---
const BASE_DIR = process.cwd();
const TEAM_CONTEXT_FILE = path.join(BASE_DIR, 'team-context.json');
const TOPICS_LOG_FILE = path.join(BASE_DIR, 'sre-topics-log.json');
const AGENT_LOG_FILE = path.join(BASE_DIR, 'agent-log.json');

const {
  AZURE_TENANT_ID,
  AZURE_CLIENT_ID,
  AZURE_CLIENT_SECRET,
  SENDER_EMAIL,
  RECIPIENT_EMAILS,
  ANTHROPIC_API_KEY,
  TEAM_NAME,
  COUNTRY_CODE
} = process.env;

const COUNTRY = COUNTRY_CODE || 'US';

function nowISO() { return new Date().toISOString(); }

async function readJsonSafe(filePath, defaultValue) {
  try {
    const exists = await fs.pathExists(filePath);
    if (!exists) return defaultValue;
    return await fs.readJson(filePath);
  } catch (err) {
    console.error(`Failed reading ${filePath}:`, err.message);
    return defaultValue;
  }
}

async function writeJsonSafe(filePath, data) {
  try {
    await fs.writeJson(filePath, data, { spaces: 2 });
  } catch (err) {
    console.error(`Failed writing ${filePath}:`, err.message);
  }
}

function consoleBlock(title, obj) {
  console.log('='.repeat(60));
  console.log(title);
  console.log('-'.repeat(60));
  console.log(typeof obj === 'string' ? obj : JSON.stringify(obj, null, 2));
  console.log('='.repeat(60));
}

// --- Anthropic Claude helper ---
// Simple wrapper: posts a prompt and expects text output. Retries once on parse errors.
async function callClaude(prompt, { model = 'claude-sonnet-4-6', max_tokens_to_sample = 800 } = {}) {
  if (!ANTHROPIC_API_KEY) throw new Error('ANTHROPIC_API_KEY not set in environment');
  const url = 'https://api.anthropic.com/v1/complete';
  const headers = {
    'x-api-key': ANTHROPIC_API_KEY,
    'Content-Type': 'application/json'
  };
  const body = {
    model,
    prompt,
    max_tokens_to_sample
  };

  try {
    const res = await axios.post(url, body, { headers, timeout: 20000 });
    return res.data;
  } catch (err) {
    console.error('Claude call failed:', err.message);
    throw err;
  }
}

// --- MSAL / Graph helpers ---
function getMsalClient() {
  if (!AZURE_TENANT_ID || !AZURE_CLIENT_ID || !AZURE_CLIENT_SECRET) {
    throw new Error('Azure credentials missing in environment');
  }
  const config = {
    auth: {
      clientId: AZURE_CLIENT_ID,
      authority: `https://login.microsoftonline.com/${AZURE_TENANT_ID}`,
      clientSecret: AZURE_CLIENT_SECRET
    }
  };
  return new ConfidentialClientApplication(config);
}

async function getGraphToken() {
  const client = getMsalClient();
  const tokenRequest = {
    scopes: ['https://graph.microsoft.com/.default']
  };
  const resp = await client.acquireTokenByClientCredential(tokenRequest);
  if (!resp || !resp.accessToken) throw new Error('Failed to acquire Graph token');
  return resp.accessToken;
}

async function sendEmailViaGraph(accessToken, senderEmail, recipientsCsv, subject, htmlBody) {
  const recipients = recipientsCsv.split(',').map(s => ({ emailAddress: { address: s.trim() } }));
  const url = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(senderEmail)}/sendMail`;
  const payload = {
    message: {
      subject,
      body: { contentType: 'HTML', content: htmlBody },
      toRecipients: recipients
    }
  };

  const headers = { Authorization: `Bearer ${accessToken}`, 'Content-Type': 'application/json' };
  const res = await axios.post(url, payload, { headers });
  return res.status === 202 || res.status === 200;
}

// --- Agent pipeline steps ---
async function gather() {
  console.log('\n[GATHER] Reading local context and topic log...');
  const teamContext = await readJsonSafe(TEAM_CONTEXT_FILE, { team: TEAM_NAME || 'Team', recentWins: [], notes: [] });
  const topicsLog = await readJsonSafe(TOPICS_LOG_FILE, []);

  // Optional holiday API call
  let holidayInfo = null;
  try {
    const year = new Date().getFullYear();
    const hurl = `https://date.nager.at/api/v3/PublicHolidays/${year}/${COUNTRY}`;
    console.log(`[GATHER] Checking public holidays for ${COUNTRY} (${year})`);
    const resp = await axios.get(hurl, { timeout: 5000 });
    const today = new Date().toISOString().slice(0,10);
    const todayHoliday = resp.data.find(h => h.date === today);
    if (todayHoliday) holidayInfo = todayHoliday;
  } catch (err) {
    console.warn('[GATHER] Holiday API failed; continuing without holiday info.');
  }

  const gathered = { teamContext, topicsLog, holidayInfo, timestamp: nowISO() };
  consoleBlock('[GATHER] Result', gathered);
  return gathered;
}

async function plan(gathered) {
  console.log('\n[PLAN] Asking Claude to choose angle and topic (structured JSON).');
  const recentTopics = (gathered.topicsLog || []).map(t => t.topic);
  // Build a prompt asking Claude to return structured JSON with 'angle' and 'chosen_topic'
  // If LOCAL_TEST is set, skip external Claude API and pick a fallback topic
  if (process.env.LOCAL_TEST === 'true') {
    const chosen = pickFallbackTopic(gathered.topicsLog);
    const parsed = { angle: 'Keep it positive and actionable', chosen_topic: chosen, reason: 'Local test mode fallback' };
    consoleBlock('[PLAN] (LOCAL_TEST) Decision', parsed);
    return parsed;
  }

  const prompt = `You are an assistant that helps select an angle for a short team greeting and choose an SRE topic.\nInput: ${JSON.stringify({ teamContext: gathered.teamContext, holiday: gathered.holidayInfo, recentTopics })}\nRespond with JSON exactly like: {"angle":"...","chosen_topic":"...","reason":"..."} and choose a topic not in recentTopics from the SRE topic space (SLOs/SLIs/error budgets, incident response, observability, chaos engineering, capacity planning, on-call practices, postmortems, toil reduction, automation, reliability patterns).`;

  // Call Claude and attempt to parse JSON. Retry once on parse failure.
  let responseText;
  try {
    const res = await callClaude(prompt, { max_tokens_to_sample: 400 });
    responseText = res.completion || res.text || JSON.stringify(res);
  } catch (err) {
    throw new Error('[PLAN] Claude call failed: ' + err.message);
  }

  // Try parse
  let parsed;
  try { parsed = JSON.parse(responseText); }
  catch (err) {
    // retry once by asking Claude to re-output only JSON
    console.warn('[PLAN] Failed to parse Claude JSON. Retrying once.');
    const retryPrompt = `Re-output ONLY the JSON object requested previously (no explanation). Previous output: ${responseText}`;
    const r = await callClaude(retryPrompt, { max_tokens_to_sample: 300 });
    const retryText = r.completion || r.text || JSON.stringify(r);
    try { parsed = JSON.parse(retryText); }
    catch (err2) { throw new Error('[PLAN] Failed to parse Claude JSON after retry'); }
  }

  consoleBlock('[PLAN] Decision', parsed);
  return parsed;
}

async function draft(planDecision, gathered) {
  console.log('\n[DRAFT] Asking Claude to write greeting and SRE tip.');
  // LOCAL_TEST mode: return simple deterministic draft without calling Claude
  if (process.env.LOCAL_TEST === 'true') {
    const topic = planDecision.chosen_topic || 'SRE Practice';
    const greeting = `Good morning ${process.env.TEAM_NAME || 'team'} — wishing you a productive day!`;
    const tip = `Topic: ${topic}\n\nThis is a short test tip used when running in local test mode. Focus: understand the basics, why it matters, and one small action item to try. Actionable takeaway: discuss this topic during your next standup.`;
    const parsed = { greeting, topic, tip };
    consoleBlock('[DRAFT] (LOCAL_TEST) Drafted Content', parsed);
    return parsed;
  }

  const prompt = `Using this plan: ${JSON.stringify(planDecision)} and team context: ${JSON.stringify(gathered.teamContext)}, produce a JSON object:{"greeting":"...","topic":"...","tip":"..."}. Greeting: <=50 words, warm. Tip: 150-250 words, explain concept, why it matters, one concrete takeaway. Topic field should match chosen_topic.`;

  const res = await callClaude(prompt, { max_tokens_to_sample: 1200 });
  const txt = res.completion || res.text || JSON.stringify(res);

  // parse with retry
  let parsed;
  try { parsed = JSON.parse(txt); }
  catch (err) {
    console.warn('[DRAFT] Failed to parse draft JSON; retrying once.');
    const retryPrompt = `Re-output ONLY the JSON object requested previously. Previous output: ${txt}`;
    const r = await callClaude(retryPrompt, { max_tokens_to_sample: 800 });
    const rt = r.completion || r.text || JSON.stringify(r);
    try { parsed = JSON.parse(rt); }
    catch (err2) { throw new Error('[DRAFT] Failed to parse draft JSON after retry'); }
  }

  consoleBlock('[DRAFT] Drafted Content', parsed);
  return parsed;
}

async function critique(draftObj) {
  console.log('\n[CRITIQUE] Asking Claude to critique and, if necessary, revise.');
  // LOCAL_TEST: skip critique and accept draft
  if (process.env.LOCAL_TEST === 'true') {
    consoleBlock('[CRITIQUE] (LOCAL_TEST) Finalized Content', draftObj);
    return draftObj;
  }

  const prompt = `Critique the following JSON with fields greeting, topic, tip against accuracy, clarity, conciseness, and actionable guidance. If any field should be revised, return the full revised JSON. Input: ${JSON.stringify(draftObj)}\nRespond with JSON only.`;

  const res = await callClaude(prompt, { max_tokens_to_sample: 800 });
  const txt = res.completion || res.text || JSON.stringify(res);

  // parse with retry
  let parsed;
  try { parsed = JSON.parse(txt); }
  catch (err) {
    console.warn('[CRITIQUE] Failed to parse critique JSON; retrying once.');
    const retryPrompt = `Re-output ONLY the JSON object requested previously. Previous output: ${txt}`;
    const r = await callClaude(retryPrompt, { max_tokens_to_sample: 500 });
    const rt = r.completion || r.text || JSON.stringify(r);
    try { parsed = JSON.parse(rt); }
    catch (err2) { throw new Error('[CRITIQUE] Failed to parse critique JSON after retry'); }
  }

  consoleBlock('[CRITIQUE] Finalized Content', parsed);
  return parsed;
}

async function act(finalContent, planDecision) {
  console.log('\n[ACT] Sending email via Microsoft Graph.');
  if (!AZURE_CLIENT_ID || !AZURE_CLIENT_SECRET || !AZURE_TENANT_ID) {
    throw new Error('[ACT] Azure credentials not configured');
  }
  if (!SENDER_EMAIL || !RECIPIENT_EMAILS) throw new Error('[ACT] SENDER_EMAIL or RECIPIENT_EMAILS not set');

  const accessToken = await getGraphToken();

  const subject = `${TEAM_NAME || 'Team'} — Morning SRE Tip: ${planDecision.chosen_topic || finalContent.topic}`;
  const html = `
  <html>
  <body style="font-family:Arial,Helvetica,sans-serif;line-height:1.4;color:#111">
    <div style="padding:16px;border-bottom:1px solid #eee">
      <p style="font-size:16px;margin:0">${finalContent.greeting}</p>
    </div>
    <div style="padding:16px">
      <h2 style="margin-top:0">SRE Tip — ${planDecision.chosen_topic || finalContent.topic}</h2>
      <div>${finalContent.tip.replace(/\n/g,'<br/>')}</div>
    </div>
    <div style="padding:12px;color:#777;font-size:12px;border-top:1px solid #f1f1f1">Sent by daily-team-agent at ${nowISO()}</div>
  </body>
  </html>`;

  const sent = await sendEmailViaGraph(accessToken, SENDER_EMAIL, RECIPIENT_EMAILS, subject, html);
  console.log('[ACT] Email send status:', sent);
  return sent;
}

async function logRun(gathered, planDecision, draft1, critiqueResult, sent) {
  console.log('\n[LOG] Appending topic log and agent audit log.');

  // Update sre-topics-log.json
  const topicsLog = await readJsonSafe(TOPICS_LOG_FILE, []);
  const chosen = planDecision.chosen_topic || (draft1.topic || 'Unknown');
  topicsLog.push({ topic: chosen, date: new Date().toISOString().slice(0,10) });
  await writeJsonSafe(TOPICS_LOG_FILE, topicsLog);

  // Append to agent-log.json
  const agentLog = await readJsonSafe(AGENT_LOG_FILE, []);
  agentLog.push({ timestamp: nowISO(), gathered, planDecision, draft: draft1, critique: critiqueResult, sent });
  await writeJsonSafe(AGENT_LOG_FILE, agentLog);

  consoleBlock('[LOG] Topic Log (last entries)', topicsLog.slice(-5));
}

// --- Utilities: pick fallback topic if Claude fails to choose ---
const FALLBACK_TOPICS = [
  'SLOs and SLIs', 'Error budgets', 'Incident response basics', 'Observability fundamentals', 'Chaos engineering',
  'Capacity planning', 'On-call best practices', 'Postmortems', 'Toil reduction', 'Automation for reliability', 'Reliability patterns'
];

function pickFallbackTopic(topicsLog) {
  const cutoff = new Date(); cutoff.setDate(cutoff.getDate() - 30);
  const recent = new Set((topicsLog || []).filter(t => new Date(t.date) >= cutoff).map(t => t.topic));
  const candidates = FALLBACK_TOPICS.filter(t => !recent.has(t));
  return candidates.length ? candidates[Math.floor(Math.random()*candidates.length)] : FALLBACK_TOPICS[Math.floor(Math.random()*FALLBACK_TOPICS.length)];
}

// --- Main run ---
async function run() {
  try {
    const gathered = await gather();

    let planDecision;
    try {
      planDecision = await plan(gathered);
      if (!planDecision.chosen_topic) throw new Error('No chosen_topic from Claude');
    } catch (err) {
      console.warn('[MAIN] Plan step failed or ambiguous; using fallback topic. Error:', err.message);
      const fb = pickFallbackTopic(gathered.topicsLog);
      planDecision = { angle: 'Keep it positive and actionable', chosen_topic: fb, reason: 'Fallback due to planning error' };
    }

    let draftObj;
    try {
      draftObj = await draft(planDecision, gathered);
    } catch (err) {
      console.warn('[MAIN] Draft step failed:', err.message);
      // create a minimal draft
      draftObj = { greeting: `Good morning ${TEAM_NAME || 'team'} — have a great day!`, topic: planDecision.chosen_topic, tip: `Today's topic is ${planDecision.chosen_topic}. (Fallback draft)` };
    }

    let finalContent;
    try { finalContent = await critique(draftObj); }
    catch (err) { console.warn('[MAIN] Critique step failed:', err.message); finalContent = draftObj; }

    let sent = false;
    try { sent = await act(finalContent, planDecision); }
    catch (err) { console.error('[MAIN] Act/send failed:', err.message); }

    await logRun(gathered, planDecision, draftObj, finalContent, sent);
    console.log('\n[MAIN] Run complete.');
  } catch (err) {
    console.error('[MAIN] Fatal error:', err.message);
  }
}

// Execute when run directly
if (require.main === module) {
  run();
}
