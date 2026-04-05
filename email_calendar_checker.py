#!/usr/bin/env python3
"""
Email & Calendar Follow-up Checker

Reads Gmail sent emails (past 21 days) and analyzes whether a follow-up is
actually needed based on email content AND recent contact history. Checks
Google Calendar for meetings where no follow-up occurred within 24 hours.
Skips people you've been in touch with recently on any channel.
Emails a summary to andrew@askbobai.com.
"""

import base64
import os
import re
import sys
from datetime import datetime, timedelta, timezone
from email.mime.text import MIMEText

from dateutil import parser as dateutil_parser
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
CALENDAR_FOLLOWUP_HOURS = 24
# How recently must contact have occurred (on any thread/meeting) to count
RECENT_CONTACT_DAYS = 5

# Patterns indicating the email is automated / no reply expected
NOREPLY_PATTERNS = [
    r"noreply", r"no-reply", r"no\.reply", r"donotreply", r"do-not-reply",
    r"mailer-daemon", r"postmaster", r"notifications?@", r"alerts?@",
    r"updates?@", r"info@", r"support@", r"billing@", r"receipts?@",
    r"news@", r"newsletter", r"marketing@", r"promo",
]

# Subject patterns for automated / transactional emails to skip
AUTO_SUBJECT_PATTERNS = [
    r"(order|shipping|delivery)\s*(confirm|notif|update|receipt)",
    r"(payment|invoice|receipt)\s*(confirm|received|processed)",
    r"(password|account)\s*(reset|confirm|verif)",
    r"(welcome\s+to|thanks?\s+for\s+(signing|registering|subscribing|your\s+(order|purchase)))",
    r"(unsubscribe|opt.out)",
    r"(out\s+of\s+office|auto.?reply|automatic\s+reply|ooo)",
    r"(calendar|invit(e|ation))\s*[:.]?\s*(accept|decline|tentative|update|cancel|team|sync|standup|meeting|review)",
    r"(daily|weekly|monthly)\s*(digest|summary|report|update|recap)",
    r"(security\s+alert|login\s+attempt|sign.in\s+attempt)",
]

# Patterns in the sent email body suggesting a reply IS expected
EXPECTING_REPLY_PATTERNS = [
    r"\?",                                    # contains a question
    r"(let\s+me\s+know|lmk)",                # asking for response
    r"(can\s+you|could\s+you|would\s+you)",   # making a request
    r"(please\s+(send|share|provide|confirm|review|check|update|advise|let))",
    r"(get\s+back\s+to\s+me|hear\s+(from|back))",
    r"(thoughts|feedback|input|opinion)\s*\??",
    r"(following\s+up|circling\s+back|checking\s+in|touching\s+base)",
    r"(when\s+(can|will|do|should|would))",
    r"(what\s+(do|did|is|are|was|were|should|would|will))",
    r"(any\s+(update|progress|news|thoughts|chance))",
    r"(action\s+items?|next\s+steps?|deliverables?)",
    r"(deadline|due\s+date|by\s+(end\s+of|eod|eow|monday|tuesday|wednesday|thursday|friday))",
    r"(asap|urgent|priority|time.sensitive)",
    r"(attached|attaching|see\s+attached).*\b(review|sign|fill|complete)",
    r"(schedule|set\s+up|book)\s+(a\s+)?(call|meeting|time|chat)",
    r"(proposal|quote|estimate|contract|agreement)\s*(attached|enclosed|below|for\s+your)",
    r"(interested\s+in|would\s+you\s+be)",
]

# Patterns suggesting the email is just informational / no reply needed
NO_REPLY_NEEDED_BODY = [
    r"^(thanks|thank\s+you|thx|ty)[.!]?\s*$",
    r"^(got\s+it|sounds?\s+good|perfect|great|noted|ack)[.!]?\s*$",
    r"(no\s+(reply|response|action)\s+(needed|required|necessary))",
    r"(fyi|for\s+your\s+(info|information|records?|reference))",
    r"(just\s+(a\s+)?(heads?\s+up|fyi|letting\s+you\s+know|sharing|forwarding))",
]


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


def extract_email_address(header_value):
    """Extract bare email address from a header like 'Name <email@example.com>'."""
    match = re.search(r"<([^>]+)>", header_value)
    if match:
        return match.group(1).lower()
    return header_value.strip().lower()


def get_message_body(message):
    """Extract plain text body from a Gmail message."""
    payload = message.get("payload", {})

    if payload.get("mimeType", "").startswith("text/plain"):
        data = payload.get("body", {}).get("data", "")
        if data:
            return base64.urlsafe_b64decode(data).decode("utf-8", errors="replace")

    parts = payload.get("parts", [])
    for part in parts:
        if part.get("mimeType") == "text/plain":
            data = part.get("body", {}).get("data", "")
            if data:
                return base64.urlsafe_b64decode(data).decode("utf-8", errors="replace")
        for subpart in part.get("parts", []):
            if subpart.get("mimeType") == "text/plain":
                data = subpart.get("body", {}).get("data", "")
                if data:
                    return base64.urlsafe_b64decode(data).decode("utf-8", errors="replace")

    for part in parts:
        if part.get("mimeType") == "text/html":
            data = part.get("body", {}).get("data", "")
            if data:
                html = base64.urlsafe_b64decode(data).decode("utf-8", errors="replace")
                return re.sub(r"<[^>]+>", " ", html)

    return ""


def is_automated_sender(recipients):
    """Check if the email was sent to an automated/noreply address."""
    for addr in recipients:
        for pattern in NOREPLY_PATTERNS:
            if re.search(pattern, addr, re.IGNORECASE):
                return True
    return False


def is_automated_subject(subject):
    """Check if the subject looks like an automated/transactional email."""
    for pattern in AUTO_SUBJECT_PATTERNS:
        if re.search(pattern, subject, re.IGNORECASE):
            return True
    return False


def email_needs_followup(subject, body):
    """
    Analyze the email subject and body to determine if a follow-up reply
    is actually expected. Returns (needs_followup: bool, reason: str).
    """
    body_lower = body.lower().strip()
    body_lines = [line.strip() for line in body_lower.split("\n") if line.strip()]

    short_body = " ".join(body_lines[:3]) if body_lines else ""
    for pattern in NO_REPLY_NEEDED_BODY:
        if re.search(pattern, short_body, re.IGNORECASE | re.MULTILINE):
            return False, ""

    # Strip quoted text and signature
    original_lines = []
    for line in body_lines:
        if line.startswith(">") or line.startswith("on ") and "wrote:" in line:
            break
        if re.match(r"^-{2,}$|^_{2,}$|^sent from", line, re.IGNORECASE):
            break
        original_lines.append(line)

    original_text = " ".join(original_lines)
    if not original_text:
        return False, ""

    reasons = []
    for pattern in EXPECTING_REPLY_PATTERNS:
        if re.search(pattern, original_text, re.IGNORECASE):
            match = re.search(pattern, original_text, re.IGNORECASE)
            reasons.append(match.group(0).strip())

    if reasons:
        unique = list(dict.fromkeys(reasons))
        reason_str = unique[0]
        if "?" in reason_str:
            return True, "Asked a question"
        if re.search(r"(follow|circle|check|touch)", reason_str, re.IGNORECASE):
            return True, "Following up on previous conversation"
        if re.search(r"(please|can you|could you|would you)", reason_str, re.IGNORECASE):
            return True, "Made a request"
        if re.search(r"(deadline|due|asap|urgent|eod|eow)", reason_str, re.IGNORECASE):
            return True, "Time-sensitive request"
        if re.search(r"(schedule|set up|book)", reason_str, re.IGNORECASE):
            return True, "Trying to schedule a meeting"
        if re.search(r"(proposal|quote|contract|estimate)", reason_str, re.IGNORECASE):
            return True, "Sent proposal/document for review"
        if re.search(r"(interest|would you be)", reason_str, re.IGNORECASE):
            return True, "Gauging interest"
        if re.search(r"(thoughts|feedback|input)", reason_str, re.IGNORECASE):
            return True, "Requested feedback"
        if re.search(r"(update|progress|news)", reason_str, re.IGNORECASE):
            return True, "Asked for an update"
        if re.search(r"(action|next steps|deliverable)", reason_str, re.IGNORECASE):
            return True, "Waiting on action items"
        if re.search(r"(attached|review|sign|fill)", reason_str, re.IGNORECASE):
            return True, "Sent attachment for review/action"
        return True, f"Likely expecting reply: \"{reason_str}\""

    return False, ""


def build_recent_contacts(gmail_service, calendar_service, my_email):
    """
    Build a set of email addresses you've had recent contact with across
    ALL channels: emails they sent you, emails you sent them (on any thread),
    and calendar meetings you both attended.

    Returns a dict of {email_address: last_contact_description}.
    """
    now = datetime.now(timezone.utc)
    recent_cutoff = now - timedelta(days=RECENT_CONTACT_DAYS)
    recent_cutoff_str = recent_cutoff.strftime("%Y/%m/%d")

    recent = {}  # email -> description of last contact

    # 1) People who emailed me recently (inbox, any thread)
    print("  Building recent contacts: checking incoming emails...")
    query = f"in:inbox after:{recent_cutoff_str}"
    results = gmail_service.users().messages().list(
        userId="me", q=query, maxResults=500
    ).execute()
    for msg_info in results.get("messages", []):
        msg = gmail_service.users().messages().get(
            userId="me", id=msg_info["id"], format="metadata",
            metadataHeaders=["From"]
        ).execute()
        headers = {h["name"].lower(): h["value"] for h in msg["payload"]["headers"]}
        from_addr = extract_email_address(headers.get("from", ""))
        if from_addr and from_addr != my_email.lower():
            if from_addr not in recent:
                recent[from_addr] = "They emailed you recently"

    # 2) People I emailed recently (on any thread, not just the flagged one)
    print("  Building recent contacts: checking outgoing emails...")
    query = f"in:sent after:{recent_cutoff_str}"
    results = gmail_service.users().messages().list(
        userId="me", q=query, maxResults=500
    ).execute()
    for msg_info in results.get("messages", []):
        msg = gmail_service.users().messages().get(
            userId="me", id=msg_info["id"], format="metadata",
            metadataHeaders=["To", "Cc"]
        ).execute()
        headers = {h["name"].lower(): h["value"] for h in msg["payload"]["headers"]}
        for field in ["to", "cc"]:
            for addr in headers.get(field, "").split(","):
                email_addr = extract_email_address(addr)
                if email_addr and email_addr != my_email.lower():
                    if email_addr not in recent:
                        recent[email_addr] = "You emailed them recently (other thread)"

    # 3) People I've had calendar meetings with recently
    print("  Building recent contacts: checking recent meetings...")
    events_result = calendar_service.events().list(
        calendarId="primary",
        timeMin=recent_cutoff.isoformat(),
        timeMax=now.isoformat(),
        maxResults=100,
        singleEvents=True,
        orderBy="startTime",
    ).execute()
    for event in events_result.get("items", []):
        if event.get("status") == "cancelled":
            continue
        # Skip events I declined
        for a in event.get("attendees", []):
            if a.get("self") and a.get("responseStatus") == "declined":
                break
        else:
            summary = event.get("summary", "a meeting")
            for a in event.get("attendees", []):
                email_addr = a.get("email", "").lower()
                if email_addr and email_addr != my_email.lower() and not a.get("resource", False):
                    if email_addr not in recent:
                        recent[email_addr] = f"Met in \"{summary}\" recently"

    print(f"  Found {len(recent)} recent contacts across all channels")
    return recent


def find_unreplied_sent_emails(gmail_service, my_email, recent_contacts):
    """
    Find emails sent in the past 21 days where:
    - No one has replied on that thread
    - The email content suggests a reply was expected
    - The user hasn't already followed up in the past 3 days
    - The recipient hasn't been in touch on ANY other channel recently
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
            userId="me", id=thread_id, format="full"
        ).execute()

        messages = thread.get("messages", [])
        if not messages:
            continue

        last_sent_idx = -1
        last_sent_date = None
        last_sent_msg = None
        subject = ""
        recipients = set()

        for i, message in enumerate(messages):
            headers = {h["name"].lower(): h["value"] for h in message["payload"]["headers"]}
            from_addr = headers.get("from", "")

            if not subject:
                subject = headers.get("subject", "(no subject)")

            if my_email.lower() in from_addr.lower():
                last_sent_idx = i
                last_sent_msg = message
                internal_date = int(message.get("internalDate", "0"))
                last_sent_date = datetime.fromtimestamp(internal_date / 1000, tz=timezone.utc)
                to_addr = headers.get("to", "")
                recipients = set()
                for addr in to_addr.split(","):
                    email_addr = extract_email_address(addr)
                    if email_addr and my_email.lower() not in email_addr:
                        recipients.add(email_addr)

        if last_sent_idx == -1 or last_sent_date is None or last_sent_msg is None:
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

        # Exclude if already followed up in the past 3 days
        if last_sent_date >= grace_date:
            continue

        # Skip automated/noreply recipients
        if is_automated_sender(recipients):
            continue

        # Skip automated/transactional subjects
        if is_automated_subject(subject):
            continue

        # Read the actual email body and determine if a reply is expected
        body = get_message_body(last_sent_msg)
        needs_followup, reason = email_needs_followup(subject, body)

        if not needs_followup:
            continue

        # Check if ALL recipients have been in recent contact on other channels
        # If every recipient is someone you've recently touched base with, skip
        recipients_with_recent_contact = []
        recipients_without_contact = []
        for r in recipients:
            if r in recent_contacts:
                recipients_with_recent_contact.append((r, recent_contacts[r]))
            else:
                recipients_without_contact.append(r)

        if not recipients_without_contact:
            # Everyone on this email has been in touch recently — skip
            continue

        days_waiting = (now - last_sent_date).days
        unreplied.append({
            "subject": subject,
            "recipients": ", ".join(recipients_without_contact),
            "sent_date": last_sent_date.strftime("%b %d, %Y"),
            "days_waiting": days_waiting,
            "reason": reason,
        })

    return unreplied


def find_meetings_without_followup(gmail_service, calendar_service, my_email,
                                   recent_contacts):
    """
    Find calendar meetings from the past 3 weeks where no follow-up email
    was sent to any attendee within 24 hours of the meeting AND you haven't
    been in contact with any attendee recently on other threads/meetings.
    """
    now = datetime.now(timezone.utc)
    lookback_date = now - timedelta(days=LOOKBACK_DAYS)
    followup_threshold = now - timedelta(hours=CALENDAR_FOLLOWUP_HOURS)

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
        if event.get("status") == "cancelled":
            continue

        start = event["start"].get("dateTime")
        if not start:
            continue

        summary = event.get("summary", "(no title)")

        attendees = event.get("attendees", [])
        my_status = None
        for a in attendees:
            if a.get("self"):
                my_status = a.get("responseStatus")
                break
        if my_status == "declined":
            continue

        other_attendees = [
            a.get("email", "").lower()
            for a in attendees
            if a.get("email", "").lower() != my_email.lower()
            and not a.get("resource", False)
        ]

        if not other_attendees:
            continue

        event_time = dateutil_parser.isoparse(start)

        # Only flag meetings older than 24 hours
        if event_time > followup_threshold:
            continue

        # Check if you've emailed ANY attendee after the meeting (on any thread)
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

        if has_followup:
            continue

        # Check if ANY attendee has been in recent contact (other emails, other meetings)
        # If you've been in touch with at least one key attendee, skip
        all_in_recent_contact = all(
            a in recent_contacts for a in other_attendees
        )
        any_in_recent_contact = any(
            a in recent_contacts for a in other_attendees
        )

        # If every attendee is someone you've recently been in touch with
        # via other channels, this meeting doesn't need a separate follow-up
        if all_in_recent_contact:
            continue

        # If you've been in touch with at least one attendee recently AND
        # the meeting had multiple attendees, likely covered via that contact
        if any_in_recent_contact and len(other_attendees) > 1:
            continue

        hours_since = int((now - event_time).total_seconds() / 3600)
        if hours_since < 48:
            time_label = f"{hours_since} hours ago"
        else:
            time_label = f"{hours_since // 24} days ago"

        # Only include attendees you HAVEN'T been in touch with
        stale_attendees = [a for a in other_attendees if a not in recent_contacts]

        missing_followups.append({
            "meeting": summary,
            "date": event_time.strftime("%b %d, %Y at %I:%M %p"),
            "attendees": ", ".join(stale_attendees[:5]),
            "time_label": time_label,
            "hours_since": hours_since,
        })

    return missing_followups


def build_email_body(unreplied_emails, missing_followups):
    """Build a formatted email body with the results."""
    today = datetime.now().strftime("%A, %B %d, %Y")
    lines = [f"Follow-up Reminder for {today}", "=" * 50, ""]

    if unreplied_emails:
        lines.append(f"EMAILS AWAITING REPLY ({len(unreplied_emails)} found)")
        lines.append("-" * 40)
        lines.append("")
        for item in sorted(unreplied_emails, key=lambda x: x["days_waiting"], reverse=True):
            lines.append(f"  To: {item['recipients']}")
            lines.append(f"  Subject: {item['subject']}")
            lines.append(f"  Sent: {item['sent_date']} ({item['days_waiting']} days ago)")
            lines.append(f"  Why: {item['reason']}")
            lines.append("")
    else:
        lines.append("EMAILS AWAITING REPLY: None - you're all caught up!")
        lines.append("")

    if missing_followups:
        lines.append(f"MEETINGS NEEDING FOLLOW-UP ({len(missing_followups)} found)")
        lines.append("-" * 40)
        lines.append("(No email sent to any attendee within 24 hours of meeting)")
        lines.append("")
        for item in sorted(missing_followups, key=lambda x: x["hours_since"], reverse=True):
            lines.append(f"  Meeting: {item['meeting']}")
            lines.append(f"  Date: {item['date']} ({item['time_label']})")
            lines.append(f"  Contact: {item['attendees']}")
            lines.append("")
    else:
        lines.append("MEETINGS NEEDING FOLLOW-UP: None - you're all caught up!")
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

    # Build a map of everyone you've been in contact with recently
    # across all channels (other email threads, calendar meetings)
    print("Building recent contact history (past 5 days)...")
    recent_contacts = build_recent_contacts(
        gmail_service, calendar_service, my_email
    )

    print("Checking for unreplied sent emails (past 21 days)...")
    unreplied = find_unreplied_sent_emails(gmail_service, my_email, recent_contacts)
    print(f"  Found {len(unreplied)} emails truly needing follow-up")

    print("Checking calendar meetings for missing follow-ups...")
    missing_followups = find_meetings_without_followup(
        gmail_service, calendar_service, my_email, recent_contacts
    )
    print(f"  Found {len(missing_followups)} meetings truly needing follow-up")

    body = build_email_body(unreplied, missing_followups)
    today_str = datetime.now().strftime("%b %d, %Y")
    subject = f"Follow-up Reminder - {today_str}"

    print(f"Sending summary to {RECIPIENT_EMAIL}...")
    send_email(gmail_service, RECIPIENT_EMAIL, subject, body)

    print("Done!")


if __name__ == "__main__":
    main()
