"""
Morning briefing agent.

Pulls recent Outlook email, today's calendar, and open Microsoft To Do tasks,
asks Claude to produce an action-oriented briefing, then:
  - emails the briefing to TO_EMAIL
  - creates any new tasks Claude suggests
  - creates any new calendar events Claude suggests
"""
import json
import os
import subprocess
import sys
from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo

import requests

GRAPH = "https://graph.microsoft.com/v1.0"
SCOPES = [
    "Mail.ReadWrite",
    "Mail.Send",
    "Calendars.ReadWrite",
    "Tasks.ReadWrite",
    "User.Read",
]


def get_access_token() -> str:
    client_id = os.environ["AZURE_CLIENT_ID"]
    tenant_id = os.environ["AZURE_TENANT_ID"]
    refresh_token = os.environ["MS_REFRESH_TOKEN"]

    resp = requests.post(
        f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token",
        data={
            "client_id": client_id,
            "grant_type": "refresh_token",
            "refresh_token": refresh_token,
            "scope": " ".join(SCOPES),
        },
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()["access_token"]


def graph_get(token: str, path: str, params: dict | None = None) -> dict:
    r = requests.get(
        f"{GRAPH}{path}",
        headers={"Authorization": f"Bearer {token}"},
        params=params,
        timeout=30,
    )
    if not r.ok:
        print(f"Graph GET failed: {r.status_code} {r.url}", file=sys.stderr)
        print(f"Response: {r.text[:500]}", file=sys.stderr)
    r.raise_for_status()
    return r.json()


def graph_post(token: str, path: str, body: dict) -> dict:
    r = requests.post(
        f"{GRAPH}{path}",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json=body,
        timeout=30,
    )
    r.raise_for_status()
    return r.json() if r.text else {}


def fetch_recent_emails(token: str) -> list[dict]:
    since = (datetime.now(timezone.utc) - timedelta(hours=24)).strftime("%Y-%m-%dT%H:%M:%SZ")
    data = graph_get(
        token,
        "/me/mailFolders/inbox/messages",
        params={
            "$top": "30",
            "$orderby": "receivedDateTime desc",
            "$filter": f"receivedDateTime ge {since}",
            "$select": "subject,from,receivedDateTime,bodyPreview,isRead,importance,webLink",
        },
    )
    return data.get("value", [])


def fetch_today_calendar(token: str, tz: ZoneInfo) -> list[dict]:
    now_local = datetime.now(tz)
    start_local = now_local.replace(hour=0, minute=0, second=0, microsecond=0)
    end_local = start_local + timedelta(days=1)
    start_utc = start_local.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    end_utc = end_local.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    data = graph_get(
        token,
        "/me/calendarView",
        params={
            "startDateTime": start_utc,
            "endDateTime": end_utc,
            "$select": "subject,start,end,location,organizer,bodyPreview,isOnlineMeeting,onlineMeeting",
            "$orderby": "start/dateTime",
            "$top": "50",
        },
    )
    return data.get("value", [])


def fetch_default_tasklist_id(token: str) -> str | None:
    try:
        data = graph_get(token, "/me/todo/lists")
    except requests.HTTPError as e:
        print(f"WARN: /me/todo/lists failed: {e}. Continuing without tasks.", file=sys.stderr)
        return None
    lists = data.get("value", [])
    for lst in lists:
        print(f"  list: name={lst.get('displayName')!r} wellknown={lst.get('wellknownListName')} id={lst.get('id')}", file=sys.stderr)
    for lst in lists:
        if lst.get("wellknownListName") == "defaultList":
            return lst["id"]
    return lists[0]["id"] if lists else None


def fetch_open_tasks(token: str, list_id: str | None) -> list[dict]:
    if not list_id:
        return []
    try:
        data = graph_get(
            token,
            f"/me/todo/lists/{list_id}/tasks",
            params={"$top": "100"},
        )
    except requests.HTTPError as e:
        print(f"WARN: tasks fetch failed: {e}. Continuing without tasks.", file=sys.stderr)
        return []
    return [t for t in data.get("value", []) if t.get("status") != "completed"]


def build_claude_input(emails: list[dict], events: list[dict], tasks: list[dict], tz_name: str) -> str:
    def email_row(e: dict) -> str:
        sender = e.get("from", {}).get("emailAddress", {})
        return (
            f"- [{e.get('receivedDateTime', '')}] "
            f"{'UNREAD' if not e.get('isRead') else 'read  '} | "
            f"importance={e.get('importance', 'normal')} | "
            f"from: {sender.get('name', '?')} <{sender.get('address', '?')}> | "
            f"subject: {e.get('subject', '(none)')}\n"
            f"  preview: {(e.get('bodyPreview') or '')[:400].strip()}"
        )

    def event_row(ev: dict) -> str:
        start = ev.get("start", {}).get("dateTime", "")
        end = ev.get("end", {}).get("dateTime", "")
        organizer = ev.get("organizer", {}).get("emailAddress", {}).get("name", "?")
        loc = ev.get("location", {}).get("displayName") or ("online" if ev.get("isOnlineMeeting") else "")
        return (
            f"- {start} → {end} | {ev.get('subject', '(no subject)')} "
            f"| organizer: {organizer} | location: {loc}"
        )

    def task_row(t: dict) -> str:
        due = t.get("dueDateTime", {}).get("dateTime") if t.get("dueDateTime") else None
        return (
            f"- importance={t.get('importance', 'normal')} "
            f"| due={due or 'none'} "
            f"| {t.get('title', '(no title)')}"
        )

    return (
        f"TIMEZONE: {tz_name}\n"
        f"CURRENT_TIME: {datetime.now(ZoneInfo(tz_name)).isoformat()}\n\n"
        f"RECENT EMAILS (last 24h, {len(emails)}):\n"
        + ("\n".join(email_row(e) for e in emails) if emails else "(none)")
        + f"\n\nTODAY'S CALENDAR ({len(events)} events):\n"
        + ("\n".join(event_row(ev) for ev in events) if events else "(none)")
        + f"\n\nOPEN TASKS ({len(tasks)}):\n"
        + ("\n".join(task_row(t) for t in tasks) if tasks else "(none)")
    )


BRIEFING_SCHEMA = {
    "type": "object",
    "properties": {
        "email_subject": {"type": "string"},
        "email_html": {
            "type": "string",
            "description": (
                "Full HTML body of the briefing email. Include: top priorities, "
                "emails needing reply (grouped by urgency/tone), meeting prep notes, "
                "conflicts or gaps, and suggested focus blocks. Use <h2>, <ul>, <li>, "
                "<strong> — keep it scannable."
            ),
        },
        "new_tasks": {
            "type": "array",
            "description": "Tasks to add to Microsoft To Do. Only propose ones NOT already in the open-tasks list.",
            "items": {
                "type": "object",
                "properties": {
                    "title": {"type": "string"},
                    "notes": {"type": "string"},
                    "due_date": {
                        "type": "string",
                        "description": "ISO 8601 date (YYYY-MM-DD) or empty string.",
                    },
                    "importance": {"type": "string", "enum": ["low", "normal", "high"]},
                },
                "required": ["title", "notes", "due_date", "importance"],
            },
        },
        "new_events": {
            "type": "array",
            "description": "Calendar events to create — use sparingly, only for focus blocks or prep time tied to today's meetings.",
            "items": {
                "type": "object",
                "properties": {
                    "title": {"type": "string"},
                    "start": {"type": "string", "description": "ISO 8601 local datetime"},
                    "end": {"type": "string", "description": "ISO 8601 local datetime"},
                    "notes": {"type": "string"},
                },
                "required": ["title", "start", "end", "notes"],
            },
        },
    },
    "required": ["email_subject", "email_html", "new_tasks", "new_events"],
}


SYSTEM_PROMPT = (
    "You are a proactive executive assistant preparing a morning briefing.\n"
    "- Read the user's inbox, calendar, and task list.\n"
    "- Infer urgency from email tone, sender seniority, explicit deadlines, and how long emails have sat.\n"
    "- Produce ONE clear action list ordered by priority. Be specific (not 'review emails').\n"
    "- Flag meetings needing prep and what to prep.\n"
    "- Note conflicts, double-bookings, back-to-backs with no buffer.\n"
    "- Only propose new tasks that are NOT already in the open tasks list.\n"
    "- Only propose new events for concrete focus blocks or prep time tied to today's meetings.\n"
    "- Respond with JSON matching the provided schema."
)


def call_claude(user_input: str) -> dict:
    schema_json = json.dumps(BRIEFING_SCHEMA, indent=2)
    prompt = (
        f"{SYSTEM_PROMPT}\n\n"
        "Respond with a SINGLE JSON object matching this schema exactly. "
        "No prose, no code fences, no explanations — just the JSON object.\n\n"
        f"SCHEMA:\n{schema_json}\n\n"
        "---\n\n"
        f"{user_input}"
    )
    result = subprocess.run(
        [
            "claude",
            "-p",
            prompt,
            "--output-format", "json",
        ],
        capture_output=True,
        text=True,
        timeout=600,
    )
    if result.returncode != 0:
        raise RuntimeError(
            f"claude CLI failed (exit {result.returncode}). "
            f"stderr: {result.stderr[:1000]}"
        )

    try:
        raw = json.loads(result.stdout)
    except json.JSONDecodeError:
        print(f"ERROR: claude output is not JSON: {result.stdout[:1000]!r}", file=sys.stderr)
        raise

    inner = raw.get("result", raw) if isinstance(raw, dict) else raw
    if isinstance(inner, dict):
        return inner

    text = str(inner).strip()
    if text.startswith("```"):
        lines = text.splitlines()
        if lines[0].startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].startswith("```"):
            lines = lines[:-1]
        text = "\n".join(lines).strip()
    start = text.find("{")
    end = text.rfind("}")
    if start != -1 and end != -1:
        text = text[start:end + 1]

    try:
        return json.loads(text)
    except json.JSONDecodeError:
        print(f"ERROR: Claude did not return valid JSON. Raw 'result':\n{str(inner)[:2000]}", file=sys.stderr)
        raise


def create_tasks(token: str, list_id: str | None, new_tasks: list[dict]) -> int:
    if not list_id or not new_tasks:
        return 0
    created = 0
    for t in new_tasks:
        body = {
            "title": t["title"],
            "importance": t.get("importance", "normal"),
            "body": {"content": t.get("notes", ""), "contentType": "text"},
        }
        if t.get("due_date"):
            body["dueDateTime"] = {"dateTime": f"{t['due_date']}T09:00:00", "timeZone": os.environ.get("TIMEZONE", "UTC")}
        try:
            graph_post(token, f"/me/todo/lists/{list_id}/tasks", body)
            created += 1
        except requests.HTTPError as e:
            print(f"WARN: failed to create task {t['title']!r}: {e}", file=sys.stderr)
    return created


def create_events(token: str, new_events: list[dict], tz_name: str) -> int:
    created = 0
    for ev in new_events:
        body = {
            "subject": ev["title"],
            "start": {"dateTime": ev["start"], "timeZone": tz_name},
            "end": {"dateTime": ev["end"], "timeZone": tz_name},
            "body": {"contentType": "text", "content": ev.get("notes", "")},
        }
        graph_post(token, "/me/events", body)
        created += 1
    return created


def send_email(token: str, to: str, subject: str, html: str) -> None:
    graph_post(
        token,
        "/me/sendMail",
        {
            "message": {
                "subject": subject,
                "body": {"contentType": "HTML", "content": html},
                "toRecipients": [{"emailAddress": {"address": to}}],
            },
            "saveToSentItems": True,
        },
    )


def main() -> int:
    tz_name = os.environ.get("TIMEZONE", "America/New_York")
    tz = ZoneInfo(tz_name)
    to_email = os.environ["TO_EMAIL"]

    print(f"[{datetime.now(tz).isoformat()}] Getting access token...")
    token = get_access_token()

    print("Fetching inbox...")
    emails = fetch_recent_emails(token)
    print(f"  emails={len(emails)}")
    print("Fetching calendar...")
    events = fetch_today_calendar(token, tz)
    print(f"  events={len(events)}")
    print("Fetching tasks...")
    list_id = fetch_default_tasklist_id(token)
    tasks = fetch_open_tasks(token, list_id)
    print(f"  tasks={len(tasks)}")

    user_input = build_claude_input(emails, events, tasks, tz_name)
    print("Calling Claude...")
    result = call_claude(user_input)

    new_tasks = result.get("new_tasks", [])
    new_events = result.get("new_events", [])
    print(f"Claude proposed: {len(new_tasks)} tasks, {len(new_events)} events.")

    footer_html = (
        f"<hr><p style='color:#888;font-size:12px'>"
        f"Generated {datetime.now(tz).strftime('%Y-%m-%d %H:%M %Z')} · "
        f"{len(emails)} emails · {len(events)} meetings · {len(tasks)} open tasks · "
        f"{len(new_tasks)} tasks added · {len(new_events)} events added"
        f"</p>"
    )

    print("Sending briefing email...")
    send_email(token, to_email, result["email_subject"], result["email_html"] + footer_html)

    if new_tasks:
        print("Creating tasks...")
        create_tasks(token, list_id, new_tasks)
    if new_events:
        print("Creating events...")
        create_events(token, new_events, tz_name)

    print("Done.")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except requests.HTTPError as e:
        print(f"HTTP error: {e.response.status_code} {e.response.text}", file=sys.stderr)
        sys.exit(1)
