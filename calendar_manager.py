"""
Google Calendar operations via API.
Used by Claude to manage the calendar conversationally.
Calendar: claude.w.lowndes@gmail.com
"""

from datetime import datetime, timedelta
from googleapiclient.discovery import build
from auth import get_credentials

TIMEZONE = "Europe/Vienna"
CALENDAR_ID = "primary"


def _service():
    return build("calendar", "v3", credentials=get_credentials())


def list_events(days_ahead=7, days_back=0):
    """List calendar events in a date range."""
    now = datetime.utcnow()
    time_min = (now - timedelta(days=days_back)).isoformat() + "Z"
    time_max = (now + timedelta(days=days_ahead)).isoformat() + "Z"
    result = _service().events().list(
        calendarId=CALENDAR_ID,
        timeMin=time_min,
        timeMax=time_max,
        singleEvents=True,
        orderBy="startTime",
        maxResults=50,
    ).execute()
    return result.get("items", [])


def find_events(search_term, days_ahead=30):
    """Search for events by keyword in subject/description."""
    now = datetime.utcnow()
    time_min = now.isoformat() + "Z"
    time_max = (now + timedelta(days=days_ahead)).isoformat() + "Z"
    result = _service().events().list(
        calendarId=CALENDAR_ID,
        q=search_term,
        timeMin=time_min,
        timeMax=time_max,
        singleEvents=True,
        orderBy="startTime",
        maxResults=10,
    ).execute()
    return result.get("items", [])


def add_event(subject, start, end, location=None, description=None, attendees=None, all_day=False):
    """
    Add a calendar event.
    start/end: 'YYYY-MM-DDTHH:MM:SS' for timed events, 'YYYY-MM-DD' for all-day
    attendees: list of email addresses
    """
    if all_day:
        event = {
            "summary": subject,
            "start": {"date": start},
            "end": {"date": end},
        }
    else:
        event = {
            "summary": subject,
            "start": {"dateTime": start, "timeZone": TIMEZONE},
            "end": {"dateTime": end, "timeZone": TIMEZONE},
        }
    if location:
        event["location"] = location
    if description:
        event["description"] = description
    if attendees:
        event["attendees"] = [{"email": a} for a in attendees]

    created = _service().events().insert(calendarId=CALENDAR_ID, body=event).execute()
    print(f"Created: '{created['summary']}' — {created['start']}")
    return created


def update_event(event_id, subject=None, start=None, end=None, location=None, description=None, all_day=False):
    """Update fields on an existing event."""
    svc = _service()
    event = svc.events().get(calendarId=CALENDAR_ID, eventId=event_id).execute()
    if subject:
        event["summary"] = subject
    if start:
        if all_day:
            event["start"] = {"date": start}
        else:
            event["start"] = {"dateTime": start, "timeZone": TIMEZONE}
    if end:
        if all_day:
            event["end"] = {"date": end}
        else:
            event["end"] = {"dateTime": end, "timeZone": TIMEZONE}
    if location:
        event["location"] = location
    if description:
        event["description"] = description
    updated = svc.events().update(calendarId=CALENDAR_ID, eventId=event_id, body=event).execute()
    print(f"Updated: '{updated['summary']}'")
    return updated


def delete_event(event_id):
    """Delete an event by ID."""
    _service().events().delete(calendarId=CALENDAR_ID, eventId=event_id).execute()
    print("Event deleted.")


def format_events(events):
    """Pretty-print a list of events for display in chat."""
    if not events:
        return "No events found."
    lines = []
    for e in events:
        start = e["start"].get("dateTime", e["start"].get("date", ""))
        end = e["end"].get("dateTime", e["end"].get("date", ""))
        loc = e.get("location", "")
        line = f"• [{e['id'][:8]}...] **{e['summary']}** — {start[:16].replace('T', ' ')} → {end[:16].replace('T', ' ')}"
        if loc:
            line += f" @ {loc}"
        lines.append(line)
    return "\n".join(lines)
