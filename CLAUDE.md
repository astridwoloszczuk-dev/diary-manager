# Diary Manager

Conversational calendar management via Google Calendar API.
Calendar account: claude.w.lowndes@gmail.com
Claude manages events from this Claude Code session on request.

## Setup

```bash
cd "diary-manager"
python -m venv .venv
.venv/Scripts/pip install -r requirements.txt
.venv/Scripts/python auth.py   # opens browser, sign in as claude.w.lowndes@gmail.com
```

Token cached in `token.json` — refreshes automatically. Re-run `auth.py` only if it expires.

## Files

- `auth.py` — OAuth2 flow, token caching/refresh
- `calendar_manager.py` — Google Calendar API: list, find, add, update, delete events
- `credentials.json` — OAuth client credentials from Google Cloud Console (do NOT commit)
- `token.json` — auto-generated auth token (do NOT commit)
- `requirements.txt` — google-api-python-client, google-auth-oauthlib, python-dateutil

## Google Cloud Setup (already done)

- Project: `Diary Manager` on console.cloud.google.com
- API: Google Calendar API enabled
- OAuth credentials: Desktop app, downloaded as `credentials.json`
- Test user: claude.w.lowndes@gmail.com
- Timezone: Europe/Vienna

## Adding Google Calendar to Outlook

To see Claude-managed events in Outlook:
1. Go to calendar.google.com (sign in as claude.w.lowndes@gmail.com)
2. Settings → share this calendar with astrid.woloszczuk@googlemail.com (or add via Outlook)
3. In Outlook: Add calendar → From internet → paste the Google Calendar iCal URL

## Usage (via Claude Code chat)

Always confirm with user before executing any write operation.

Ask Claude to:
- "What's on the calendar this week?"
- "Add a dentist appointment Thursday 3pm for 1 hour"
- "Move my Friday call to Monday at 10am"
- "Delete the event called X"
- "Find all events with 'school' in the title"

## Planned: WebUntis Sync (Phase 2)

Script to check WebUntis for school events over the next 10 rolling days
and create/update/cancel matching calendar entries. Run via Windows Task Scheduler.
