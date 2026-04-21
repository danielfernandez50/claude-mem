# Tomorrow's Prep Agent

Runs every weekday morning on GitHub Actions. Reads your Outlook email, calendar, and Microsoft To Do list via Microsoft Graph, asks Claude (via your **Claude Max subscription** — no API key) to produce an action-oriented briefing, then emails it to you, creates any new tasks it suggests, and adds any new calendar events.

## One-time setup

### 1. Install dependencies locally

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### 2. Get a Microsoft Graph refresh token

Run the device-code flow once:

```bash
export AZURE_CLIENT_ID=bd79a7c4-f916-4c61-a9f4-9cb0f2cbdaad
export AZURE_TENANT_ID=3b492fa4-c883-4d83-9beb-41c571e378df
python auth_setup.py
```

Follow the URL + code, sign in with your work account, then copy the printed refresh token.

### 2b. Get a Claude Code OAuth token

Install Claude Code if you don't have it (`npm install -g @anthropic-ai/claude-code`), then run:

```bash
claude setup-token
```

Sign in with your Max subscription account when prompted. Copy the printed long-lived OAuth token (~1 year validity).

### 3. Create the GitHub repo

```bash
# From this directory, after `git init` is done:
gh repo create morning-briefing --private --source=. --remote=origin --push
```

(or create it manually on github.com and push.)

### 4. Add GitHub Actions secrets

In the repo → Settings → Secrets and variables → Actions → New repository secret:

| Secret              | Value                                                  |
| ------------------- | ------------------------------------------------------ |
| `AZURE_CLIENT_ID`   | `bd79a7c4-f916-4c61-a9f4-9cb0f2cbdaad`                 |
| `AZURE_TENANT_ID`   | `3b492fa4-c883-4d83-9beb-41c571e378df`                 |
| `MS_REFRESH_TOKEN`  | (from step 2)                                          |
| `CLAUDE_CODE_OAUTH_TOKEN` | (from step 2b)                                   |
| `TO_EMAIL`          | the email address to send the briefing to              |

### 5. Test it

Trigger manually: repo → Actions → Tomorrow's Prep → Run workflow.

## Local testing

```bash
cp .env.example .env
# fill in MS_REFRESH_TOKEN, ANTHROPIC_API_KEY, TO_EMAIL
set -a && source .env && set +a
python briefing.py
```

## Changing the schedule

Edit `.github/workflows/morning-briefing.yml`. Cron times are **UTC**. The default is `0 11 * * 1-5` (7am EDT / 6am EST, weekdays).

## Token expiry

- **Microsoft refresh token**: ~90 days, auto-extends on each use (which happens daily here). If the workflow fails with `invalid_grant`, re-run `auth_setup.py` and update the `MS_REFRESH_TOKEN` secret.
- **Claude Code OAuth token**: ~1 year. Set a calendar reminder — when it expires, re-run `claude setup-token` and update the `CLAUDE_CODE_OAUTH_TOKEN` secret.

## Files

- `auth_setup.py` — one-time device-code flow to obtain the refresh token.
- `briefing.py` — the daily job. Fetches Graph data, calls Claude, emails/creates tasks/events.
- `.github/workflows/morning-briefing.yml` — scheduled GitHub Actions run.
- `requirements.txt` — Python deps.
