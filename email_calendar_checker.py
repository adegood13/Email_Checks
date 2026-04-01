#!/usr/bin/env python3
"""
Email & Calendar Follow-up Checker

Checks Gmail for unreplied sent emails (past 21 days) and Google Calendar
for meetings without follow-up emails (past 3 weeks). Emails a summary
to andrew@askbobai.com.
"""

import base64
import os
import sys
from datetime import datetime, timedelta, timezone
from email.mime.text import MIMEText

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/calendar.readonly",
]

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TOKEN_PATH = os.path.join(SCRIPT_DIR, "token.json")
CREDENTIALS_PATH = os.path.join(SCRIPT_DIR, "credentials.json")

RECIPIENT_EMAIL = "andrew@askbobai.com"
LOOKBACK_DAYS = 21
FOLLOWUP_GRACE_DAYS = 3
CALENDAR_FOLLOWUP_DAYS = 5


def get_credentials():
    """Get or refresh Google API credentials."""
    creds = None
    if os.path.exists(TOKEN_PATH):
        creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDENTIALS_PATH):
                print(
                    f"ERROR: {CREDENTIALS_PATH} not found. "
                    "Please download OAuth credentials from Google Cloud Console."
                )
                sys.exit(1)
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_PATH, SCOPES)
            try:
                creds = flow.run_local_server(port=0)
            except Exception:
                # Fallback for headless environments
                flow.redirect_uri = "urn:ietf:wg:oauth:2.0:oob"
                auth_url, _ = flow.authorization_url(prompt="consent")
                print(f"\nPlease visit this URL to authorize:\n\n{auth_url}\n")
                code = input("Enter the authorization code: ")
                flow.fetch_token(code=code)
                creds = flow.credentials
        with open(TOKEN_PATH, "w") as token_file:
            token_file.write(creds.to_json())
    return creds


def get_my_email(gmail_service):
    """Get the authenticated user's email address."""
    profile = gmail_service.users().getProfile(userId="me").execute()
    return profile["emailAddress"]


def find_unreplied_sent_emails(gmail_service, my_email):
    """
    Find emails sent in the past 21 days where no one has replied,
    excluding threads where the user already followed up in the past 3 days.
    """
    now = datetime.now(timezone.utc)
    lookback_date = now - timedelta(days=LOOKBACK_DAYS)
    grace_date = now - timedelta(days=FOLLOWUP_GRACE_DAYS)

    after_str = lookback_date.strftime("%Y/%m/%d")
    query = f"in:sent after:{after_str}"

    results = gmail_service.users().messages().list(
        userId="me", q=query, maxResults=200
    ).execute()

    sent_messages = results.get("messages", [])
    unreplied = []
    seen_threads = set()

    for msg_info in sent_messages:
        thread_id = msg_info["threadId"]
        if thread_id in seen_threads:
            continue
        seen_threads.add(thread_id)

        thread = gmail_service.users().threads().get(
            userId="me", id=thread_id, format="metadata",
            metadataHeaders=["From", "To", "Subject", "Date"]
        ).execute()

        messages = thread.get("messages", [])
        if not messages:
            continue

        # Find the last sent message and check if anyone replied after it
        last_sent_idx = -1
        last_sent_date = None
        subject = ""
        recipients = set()

        for i, message in enumerate(messages):
            headers = {h["name"].lower(): h["value"] for h in message["payload"]["headers"]}
            from_addr = headers.get("from", "")

            if not subject:
                subject = headers.get("subject", "(no subject)")

            if my_email.lower() in from_addr.lower():
                last_sent_idx = i
                internal_date = int(message.get("internalDate", "0"))
                last_sent_date = datetime.fromtimestamp(internal_date / 1000, tz=timezone.utc)
                to_addr = headers.get("to", "")
                for addr in to_addr.split(","):
                    addr = addr.strip()
                    if addr and my_email.lower() not in addr.lower():
                        recipients.add(addr)

        if last_sent_idx == -1 or last_sent_date is None:
            continue

        # Check if someone else replied after the last sent message
        has_reply = False
        for message in messages[last_sent_idx + 1:]:
            headers = {h["name"].lower(): h["value"] for h in message["payload"]["headers"]}
            from_addr = headers.get("from", "")
            if my_email.lower() not in from_addr.lower():
                has_reply = True
                break

        if has_reply:
            continue

        # Exclude if the user already followed up in the past 3 days
        if last_sent_date >= grace_date:
            continue

        days_waiting = (now - last_sent_date).days
        unreplied.append({
            "subject": subject,
            "recipients": ", ".join(recipients) if recipients else "unknown",
            "sent_date": last_sent_date.strftime("%b %d, %Y"),
            "days_waiting": days_waiting,
        })

    return unreplied


def find_meetings_without_followup(gmail_service, calendar_service, my_email):
    """
    Find calendar meetings from the past 3 weeks where no follow-up email
    was sent within 5 days of the meeting.
    """
    now = datetime.now(timezone.utc)
    lookback_date = now - timedelta(days=LOOKBACK_DAYS)
    followup_threshold = now - timedelta(days=CALENDAR_FOLLOWUP_DAYS)

    events_result = calendar_service.events().list(
        calendarId="primary",
        timeMin=lookback_date.isoformat(),
        timeMax=now.isoformat(),
        maxResults=100,
        singleEvents=True,
        orderBy="startTime",
    ).execute()

    events = events_result.get("items", [])
    missing_followups = []

    for event in events:
        # Skip all-day events and cancelled events
        if event.get("status") == "cancelled":
            continue

        start = event["start"].get("dateTime")
        if not start:
            continue

        summary = event.get("summary", "(no title)")

        # Get attendees (excluding the user)
        attendees = event.get("attendees", [])
        other_attendees = [
            a.get("email", "")
            for a in attendees
            if a.get("email", "").lower() != my_email.lower()
            and not a.get("resource", False)
        ]

        if not other_attendees:
            continue

        # Parse the event start time
        from dateutil import parser as dateutil_parser
        event_time = dateutil_parser.isoparse(start)

        # Only flag meetings older than 5 days
        if event_time > followup_threshold:
            continue

        # Check if a follow-up email was sent to any attendee after the meeting
        has_followup = False
        meeting_date_str = event_time.strftime("%Y/%m/%d")

        for attendee_email in other_attendees:
            query = f"in:sent to:{attendee_email} after:{meeting_date_str}"
            results = gmail_service.users().messages().list(
                userId="me", q=query, maxResults=5
            ).execute()

            if results.get("messages"):
                has_followup = True
                break

        if not has_followup:
            days_since = (now - event_time).days
            missing_followups.append({
                "meeting": summary,
                "date": event_time.strftime("%b %d, %Y"),
                "attendees": ", ".join(other_attendees[:5]),
                "days_since": days_since,
            })

    return missing_followups


def build_email_body(unreplied_emails, missing_followups):
    """Build a formatted email body with the results."""
    today = datetime.now().strftime("%A, %B %d, %Y")
    lines = [f"Follow-up Reminder for {today}", "=" * 50, ""]

    if unreplied_emails:
        lines.append(f"UNREPLIED SENT EMAILS ({len(unreplied_emails)} found)")
        lines.append("-" * 40)
        lines.append("")
        for item in sorted(unreplied_emails, key=lambda x: x["days_waiting"], reverse=True):
            lines.append(f"  To: {item['recipients']}")
            lines.append(f"  Subject: {item['subject']}")
            lines.append(f"  Sent: {item['sent_date']} ({item['days_waiting']} days ago)")
            lines.append("")
    else:
        lines.append("UNREPLIED SENT EMAILS: None found - you're all caught up!")
        lines.append("")

    if missing_followups:
        lines.append(f"MEETINGS WITHOUT FOLLOW-UP ({len(missing_followups)} found)")
        lines.append("-" * 40)
        lines.append("")
        for item in sorted(missing_followups, key=lambda x: x["days_since"], reverse=True):
            lines.append(f"  Meeting: {item['meeting']}")
            lines.append(f"  Date: {item['date']} ({item['days_since']} days ago)")
            lines.append(f"  Attendees: {item['attendees']}")
            lines.append("")
    else:
        lines.append("MEETINGS WITHOUT FOLLOW-UP: None found - you're all caught up!")
        lines.append("")

    return "\n".join(lines)


def send_email(gmail_service, to_email, subject, body):
    """Send an email via Gmail API."""
    message = MIMEText(body)
    message["to"] = to_email
    message["subject"] = subject
    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
    gmail_service.users().messages().send(
        userId="me", body={"raw": raw}
    ).execute()
    print(f"Summary email sent to {to_email}")


def main():
    print("Authenticating with Google APIs...")
    creds = get_credentials()
    gmail_service = build("gmail", "v1", credentials=creds)
    calendar_service = build("calendar", "v3", credentials=creds)

    my_email = get_my_email(gmail_service)
    print(f"Authenticated as: {my_email}")

    print("Checking for unreplied sent emails (past 21 days)...")
    unreplied = find_unreplied_sent_emails(gmail_service, my_email)
    print(f"  Found {len(unreplied)} unreplied emails")

    print("Checking calendar meetings for missing follow-ups...")
    missing_followups = find_meetings_without_followup(
        gmail_service, calendar_service, my_email
    )
    print(f"  Found {len(missing_followups)} meetings without follow-up")

    body = build_email_body(unreplied, missing_followups)
    today_str = datetime.now().strftime("%b %d, %Y")
    subject = f"Follow-up Reminder - {today_str}"

    print(f"Sending summary to {RECIPIENT_EMAIL}...")
    send_email(gmail_service, RECIPIENT_EMAIL, subject, body)

    print("Done!")


if __name__ == "__main__":
    main()
