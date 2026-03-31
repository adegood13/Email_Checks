# Email & Calendar Follow-up Checker

## Purpose
This project checks Gmail for unreplied sent emails (past 21 days) and Google Calendar for meetings without follow-up emails (past 3 weeks), then emails a summary to andrew@askbobai.com.

## How to Run
```bash
cd /home/user/Email_Checks
pip install -r requirements.txt
python email_calendar_checker.py
```

## Google API Setup
1. Create a Google Cloud project at https://console.cloud.google.com
2. Enable the Gmail API and Google Calendar API
3. Create OAuth 2.0 credentials (Desktop application)
4. Download the credentials JSON and save as `credentials.json` in this directory
5. On first run, a browser window will open for OAuth consent. The resulting token is saved as `token.json`

## Scheduled Execution (GitHub Actions)
A GitHub Actions workflow runs this script every weekday at 8:00 AM ET automatically.

### Setup: Add Repository Secrets
After completing the Google API setup above and running the script once locally to generate `token.json`, add these two secrets in GitHub (Settings > Secrets and variables > Actions):

1. **`GOOGLE_CREDENTIALS`** — Paste the full contents of your `credentials.json` file
2. **`GOOGLE_TOKEN`** — Paste the full contents of your `token.json` file

### Manual Run
You can also trigger the workflow manually from the Actions tab in GitHub using the "Run workflow" button.
