#!/usr/bin/env python3
"""
One-time helper to generate token.json for headless environments.

Usage:
  1. Run: python generate_token.py
  2. Open the printed URL in your browser
  3. Sign in and authorize
  4. You'll be redirected to a localhost URL that won't load — that's OK
  5. Copy the FULL URL from your browser's address bar
  6. Run: python generate_token.py "PASTE_THE_FULL_URL_HERE"
"""

import json
import os
import sys
from urllib.parse import urlparse, parse_qs

import requests as http_requests

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CREDENTIALS_PATH = os.path.join(SCRIPT_DIR, "credentials.json")
TOKEN_PATH = os.path.join(SCRIPT_DIR, "token.json")

SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/calendar.readonly",
]

REDIRECT_URI = "http://localhost"


def main():
    with open(CREDENTIALS_PATH) as f:
        cred_data = json.load(f)

    client_id = cred_data["installed"]["client_id"]
    client_secret = cred_data["installed"]["client_secret"]

    if len(sys.argv) < 2:
        # Step 1: Generate the auth URL
        scope_str = "+".join(s.replace(":", "%3A").replace("/", "%2F") for s in SCOPES)
        auth_url = (
            f"https://accounts.google.com/o/oauth2/auth"
            f"?response_type=code"
            f"&client_id={client_id}"
            f"&redirect_uri=http%3A%2F%2Flocalhost"
            f"&scope={scope_str}"
            f"&access_type=offline"
            f"&prompt=consent"
        )
        print(f"\nOpen this URL in your browser:\n")
        print(auth_url)
        print(f"\nAfter authorizing, your browser will redirect to a localhost URL that won't load.")
        print("Copy the FULL URL from the address bar and run:")
        print(f'  python generate_token.py "THE_FULL_URL"')
        return

    # Step 2: Exchange the code for a token
    redirect_url = sys.argv[1]
    parsed = urlparse(redirect_url)
    params = parse_qs(parsed.query)

    if "code" not in params:
        print("ERROR: No authorization code found in the URL.")
        print("Make sure you copied the full URL from the browser address bar.")
        sys.exit(1)

    code = params["code"][0]

    # Exchange code for token
    token_response = http_requests.post(
        "https://oauth2.googleapis.com/token",
        data={
            "code": code,
            "client_id": client_id,
            "client_secret": client_secret,
            "redirect_uri": REDIRECT_URI,
            "grant_type": "authorization_code",
        },
    )

    if token_response.status_code != 200:
        print(f"ERROR: Token exchange failed: {token_response.text}")
        sys.exit(1)

    token_data = token_response.json()

    # Save in the format google-auth expects
    token_json = {
        "token": token_data["access_token"],
        "refresh_token": token_data.get("refresh_token"),
        "token_uri": "https://oauth2.googleapis.com/token",
        "client_id": client_id,
        "client_secret": client_secret,
        "scopes": SCOPES,
    }

    with open(TOKEN_PATH, "w") as f:
        json.dump(token_json, f, indent=2)

    print(f"\ntoken.json saved successfully!")
    print("You can now run: python email_calendar_checker.py")


if __name__ == "__main__":
    main()
