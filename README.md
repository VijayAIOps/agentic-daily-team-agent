**Daily Team Agent**

This repository contains `daily-team-agent`, a Node.js agentic script that sends a short morning greeting and an SRE Learning Tip of the Day to your team via Outlook (Microsoft Graph). It implements a 6-step agentic pipeline: GATHER -> PLAN -> DRAFT -> CRITIQUE -> ACT -> LOG.

**Why this is agentic**
- **Tool use**: the agent reads local context files, calls external APIs (holiday API), uses Anthropic Claude for planning/drafting/critiquing, and uses Microsoft Graph to act (send email).
- **Autonomous planning**: the PLAN step asks Claude to choose an angle and topic given gathered context.
- **Self-critique**: after drafting, the agent asks Claude to critique and revise the content.
- **Action & logging**: the agent sends email and logs the full decision trail for auditing.

Files:
- `daily-team-agent.js`: main script implementing the pipeline.
- `team-context.json`: local context (placeholder values).
- `sre-topics-log.json`: tracks which SRE topics were used (starts empty).
- `.env.example`: environment variable examples.
- `package.json`: dependencies and `start` script.

Requirements & Setup
1. Node.js (18+) and npm installed.
2. Copy `.env.example` to `.env` and fill in values:

  - `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET`: credentials for an Azure AD App registered in your tenant. Grant Application permission `Mail.Send` for Microsoft Graph and grant admin consent.
  - `SENDER_EMAIL`: the mailbox (user) that will send email (e.g., service account). When using application permissions you must call `/users/{sender}/sendMail`.
  - `RECIPIENT_EMAILS`: comma-separated list of recipients.
  - `ANTHROPIC_API_KEY`: API key for Claude.
  - `TEAM_NAME`, `COUNTRY_CODE` (optional).

How to register Azure AD App for Graph send (summary)
1. In Azure Portal -> Azure Active Directory -> App registrations -> New registration.
2. Record the Application (client) ID and Directory (tenant) ID.
3. Create a client secret under Certificates & secrets.
4. Under API permissions -> Add a permission -> Microsoft Graph -> Application permissions -> search `Mail.Send` and add it.
5. Click "Grant admin consent" for the tenant.

Install & Run
```powershell
npm install
# create .env from .env.example and fill values
npm start
```

Testing locally without external APIs (fast, safe)
- To run a safe local test that avoids calling Claude, Azure, or the holiday API, set `LOCAL_TEST=true` and `DRY_RUN=true` in your `.env` (or export them in PowerShell). Example to run with a specific test recipient:

```powershell
$env:RECIPIENT_EMAILS = "vijayalakshmi.bojja04@gmail.com"
$env:LOCAL_TEST = "true"
$env:DRY_RUN = "true"
npm start
```

- When ready to send real emails, set `LOCAL_TEST=false` (or unset), populate `ANTHROPIC_API_KEY`, and provide Azure credentials in `.env`. To add or update recipients in the future, update `RECIPIENT_EMAILS` (comma-separated) in `.env`.

Testing Manually
- Run `npm start` to execute the pipeline. The script prints each step's outputs to the console and writes `agent-log.json` and `sre-topics-log.json` in the repository.

Scheduling via Windows Task Scheduler (weekday mornings)
1. Open Task Scheduler -> Create Task.
2. On Triggers, create a new trigger: Weekly, Mon-Fri, set time (e.g., 08:30).
3. On Actions, add an action: Start a program. Program: `node` (full path to Node). Add arguments: `"<repo path>\\daily-team-agent.js"` and set Start in: the repository folder.
4. On Settings, allow task to run on demand.

Notes and safety
- The script logs decisions; keep logs private. Secrets should stay in `.env` and not be committed.
- If the holiday API or Claude calls fail, the script continues with fallback behavior.

If you'd like, I can:
- Run `npm install` here (requires network) and run a smoke test (without sending email) by toggling a dry-run flag.
- Add a Windows Task Scheduler XML export for easy import.
