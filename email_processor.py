"""
Email-based calendar request processor.
Polls Gmail (claude.w.lowndes@gmail.com) for unread emails from Astrid,
parses calendar requests using Claude API, executes them, and replies with confirmation.

Run manually or via Windows Task Scheduler every 5 minutes.
"""

import base64
import json
import os
import re
from datetime import datetime, timedelta
from email.mime.text import MIMEText

import anthropic
from googleapiclient.discovery import build

from auth import get_credentials
from calendar_manager import add_event, delete_event, find_events, update_event

BASE_DIR = os.path.dirname(__file__)

with open(os.path.join(BASE_DIR, "contacts.json")) as f:
    CONTACTS = json.load(f)

AUTHORIZED_SENDERS = {
    "astrid.woloszczuk@outlook.com",
    "astrid.woloszczuk@googlemail.com",
}

# Load API key from .env file
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
if not ANTHROPIC_API_KEY:
    env_file = os.path.join(BASE_DIR, ".env")
    if os.path.exists(env_file):
        for line in open(env_file):
            if line.startswith("ANTHROPIC_API_KEY="):
                ANTHROPIC_API_KEY = line.strip().split("=", 1)[1].strip()


def _gmail():
    return build("gmail", "v1", credentials=get_credentials())


def resolve_attendees(text):
    """Convert name mentions to email addresses using contacts.json."""
    if not text:
        return []
    emails = []
    text_lower = text.lower()

    # Handle group aliases first
    for key in ("family", "kids", "boys"):
        if key in text_lower:
            value = CONTACTS.get(key, [])
            emails.extend(value if isinstance(value, list) else [value])
            text_lower = text_lower.replace(key, "")

    # Handle niko override
    if "niko private" in text_lower:
        emails.append(CONTACTS["niko_private"])
        text_lower = text_lower.replace("niko private", "")
    elif "niko" in text_lower:
        emails.append(CONTACTS["niko_work"])
        text_lower = text_lower.replace("niko", "")

    # Handle individual names
    for name in ("astrid", "me", "alex", "alexander", "victoria", "vicky", "maximilian", "max"):
        if name in text_lower:
            value = CONTACTS.get(name)
            if value and isinstance(value, str):
                emails.append(value)

    return list(dict.fromkeys(emails))  # deduplicate, preserve order


SKIP_SUBJECT_PREFIXES = ("accepted:", "declined:", "tentative:", "cancelled:", "re:", "fwd:", "fw:")

def get_unread_messages():
    svc = _gmail()
    sender_query = " OR ".join(f"from:{s}" for s in AUTHORIZED_SENDERS)
    result = svc.users().messages().list(
        userId="me", q=f"is:unread ({sender_query})", maxResults=10
    ).execute()
    return result.get("messages", [])


def get_message_content(msg_id):
    svc = _gmail()
    msg = svc.users().messages().get(userId="me", id=msg_id, format="full").execute()
    headers = {h["name"]: h["value"] for h in msg["payload"]["headers"]}
    subject = headers.get("Subject", "")
    sender = headers.get("From", "")
    thread_id = msg["threadId"]

    body = ""
    payload = msg["payload"]
    if "parts" in payload:
        for part in payload["parts"]:
            if part["mimeType"] == "text/plain":
                data = part["body"].get("data", "")
                if data:
                    body = base64.urlsafe_b64decode(data).decode("utf-8")
                    break
    elif payload["body"].get("data"):
        body = base64.urlsafe_b64decode(payload["body"]["data"]).decode("utf-8")

    return {
        "id": msg_id,
        "thread_id": thread_id,
        "subject": subject,
        "sender": sender,
        "body": body.strip(),
    }


def build_rrule(recurrence):
    """Convert a recurrence dict to a Google Calendar RRULE string."""
    if not recurrence:
        return None
    freq = recurrence.get("frequency", "WEEKLY").upper()
    parts = [f"FREQ={freq}"]
    days = recurrence.get("days")
    if days:
        day_map = {"monday":"MO","tuesday":"TU","wednesday":"WE","thursday":"TH",
                   "friday":"FR","saturday":"SA","sunday":"SU"}
        byday = ",".join(day_map.get(d.lower(), d.upper()[:2]) for d in days)
        parts.append(f"BYDAY={byday}")
    until = recurrence.get("until")
    if until:
        parts.append(f"UNTIL={until.replace('-', '')}T235959Z")
    count = recurrence.get("count")
    if count and not until:
        parts.append(f"COUNT={count}")
    return "RRULE:" + ";".join(parts)


def parse_request(subject, body):
    """Use Claude Haiku to parse a calendar request into structured JSON."""
    today = datetime.now()
    text = f"{subject} {body}".strip()

    prompt = f"""You are a calendar assistant. Parse this calendar request into a JSON action.

Today is {today.strftime('%A, %d %B %Y')}.

Request: {text}

Return ONLY a JSON object. For adding a one-time event:
{{"action":"add","title":"event title","date":"YYYY-MM-DD","start_time":"HH:MM","end_time":"HH:MM","location":"location or null","attendees_raw":"names as written or null","description":"extra notes or null","recurrence":null}}

For adding a recurring event, include a recurrence object:
{{"action":"add","title":"event title","date":"YYYY-MM-DD","start_time":"HH:MM","end_time":"HH:MM","location":"location or null","attendees_raw":"names as written or null","description":"extra notes or null","recurrence":{{"frequency":"WEEKLY","days":["friday"],"until":"YYYY-MM-DD"}}}}

Recurrence rules:
- frequency: DAILY, WEEKLY, or MONTHLY
- days: list of day names (only for WEEKLY), e.g. ["monday","wednesday"]
- until: end date as YYYY-MM-DD (use this if a specific end date is given)
- count: number of occurrences (use this only if no end date is given)

For moving/rescheduling:
{{"action":"move","search_term":"keyword to find event","date":"YYYY-MM-DD","start_time":"HH:MM","end_time":"HH:MM or null"}}

For cancelling/deleting:
{{"action":"delete","search_term":"keyword to find event"}}

If unclear:
{{"action":"unclear","message":"what is unclear"}}

Rules:
- "this friday" = the coming Friday, even if today is Friday
- "next tuesday" = Tuesday of next week
- If no end time, default to 1 hour after start
- If no title given, infer one from context
- date = first occurrence date"""

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    response = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=800,
        messages=[{"role": "user", "content": prompt}],
    )
    raw = response.content[0].text.strip()
    raw = re.sub(r"^```(?:json)?\n?", "", raw)
    raw = re.sub(r"\n?```$", "", raw)
    return json.loads(raw)


def send_reply(thread_id, to_email, subject, body_text):
    svc = _gmail()
    message = MIMEText(body_text)
    message["to"] = to_email
    message["subject"] = subject if subject.startswith("Re:") else f"Re: {subject}"
    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
    svc.users().messages().send(
        userId="me", body={"raw": raw, "threadId": thread_id}
    ).execute()


def mark_as_read(msg_id):
    _gmail().users().messages().modify(
        userId="me", id=msg_id, body={"removeLabelIds": ["UNREAD"]}
    ).execute()
    print(f"Marked as read: {msg_id}")


def process_email(msg_id):
    msg = get_message_content(msg_id)
    subject_lower = msg["subject"].lower()
    if any(subject_lower.startswith(p) for p in SKIP_SUBJECT_PREFIXES):
        print(f"Skipping notification: {msg['subject']}")
        mark_as_read(msg_id)
        return
    print(f"Processing: {(msg['subject'] or msg['body'])[:80]}")

    parsed = parse_request(msg["subject"], msg["body"])
    action = parsed.get("action")

    if action == "add":
        attendees = resolve_attendees(parsed.get("attendees_raw", ""))
        start = f"{parsed['date']}T{parsed['start_time']}:00"
        end_time = parsed.get("end_time") or (
            datetime.strptime(parsed["start_time"], "%H:%M") + timedelta(hours=1)
        ).strftime("%H:%M")
        end = f"{parsed['date']}T{end_time}:00"
        rrule = build_rrule(parsed.get("recurrence"))

        add_event(
            subject=parsed["title"],
            start=start,
            end=end,
            location=parsed.get("location"),
            description=parsed.get("description"),
            attendees=attendees or None,
            recurrence=rrule,
        )

        reply = f"Done!\n\n{parsed['title']}\n{parsed['date']}  {parsed['start_time']}–{end_time}"
        if parsed.get("location"):
            reply += f"\n{parsed['location']}"
        if rrule:
            rec = parsed["recurrence"]
            until = rec.get("until", "")
            days = ", ".join(rec.get("days", [])).title()
            reply += f"\nRepeats weekly on {days}" if days else "\nRepeating event"
            if until:
                reply += f" until {until}"
        if attendees:
            reply += f"\nInvites sent to: {', '.join(attendees)}"

    elif action == "move":
        events = find_events(parsed["search_term"])
        if not events:
            reply = f"Couldn't find an upcoming event matching '{parsed['search_term']}'. No changes made."
        else:
            event = events[0]
            start = f"{parsed['date']}T{parsed['start_time']}:00"
            if parsed.get("end_time"):
                end = f"{parsed['date']}T{parsed['end_time']}:00"
            else:
                old_s = event["start"].get("dateTime", "")
                old_e = event["end"].get("dateTime", "")
                if old_s and old_e:
                    duration = datetime.fromisoformat(old_e[:19]) - datetime.fromisoformat(old_s[:19])
                    end = (datetime.fromisoformat(start) + duration).isoformat()
                else:
                    end = f"{parsed['date']}T{(datetime.strptime(parsed['start_time'], '%H:%M') + timedelta(hours=1)).strftime('%H:%M')}:00"

            update_event(event["id"], start=start, end=end)
            reply = f"Done! Moved '{event['summary']}' to {parsed['date']} at {parsed['start_time']}."

    elif action == "delete":
        events = find_events(parsed["search_term"])
        if not events:
            reply = f"Couldn't find an upcoming event matching '{parsed['search_term']}'. Nothing deleted."
        else:
            event = events[0]
            delete_event(event["id"])
            reply = f"Done! Deleted '{event['summary']}'."

    else:
        reply = (
            "Sorry, I couldn't understand that request. Try something like:\n"
            "• add dentist tuesday 3pm\n"
            "• move alex tennis to saturday\n"
            "• cancel friday 4.30"
        )

    sender_match = re.search(r"[\w.+\-]+@[\w.\-]+", msg["sender"])
    sender_email = sender_match.group() if sender_match else next(iter(AUTHORIZED_SENDERS))

    send_reply(msg["thread_id"], sender_email, msg["subject"], reply)
    mark_as_read(msg_id)
    print(f"Replied to {sender_email}.")


def main():
    messages = get_unread_messages()
    if not messages:
        print("No new calendar requests.")
        return
    print(f"Found {len(messages)} unread message(s).")
    for msg in messages:
        try:
            process_email(msg["id"])
        except Exception as e:
            print(f"Error processing {msg['id']}: {e}")


if __name__ == "__main__":
    main()
