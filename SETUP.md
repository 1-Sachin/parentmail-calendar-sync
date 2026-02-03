# ParentMail Calendar Automation - Setup Guide

This guide walks you through setting up the automated daily sync from ParentMail to Google Calendar using GitHub Actions.

## Overview

The automation:
- Runs daily at 6:00 AM UK time
- Logs into ParentMail automatically
- Finds the latest newsletter
- Extracts school events from the Sway diary dates
- Filters for relevant events (YR, Y2, KS1, Red class, Yellow class)
- Syncs new events to your Google Calendar (avoiding duplicates)
- Color-codes events by child (Orange=Arvi, Blue=Rivan, Red=Both)

## Quick Setup (15-20 minutes)

### Step 1: Create a Private GitHub Repository

1. Go to [github.com/new](https://github.com/new)
2. Name it `parentmail-calendar-sync` (or similar)
3. **IMPORTANT**: Set visibility to **Private** (contains sensitive automation)
4. Click "Create repository"

### Step 2: Upload the Files

Upload these files to your repository:
```
parentmail-calendar-sync/
├── .github/
│   └── workflows/
│       └── daily-sync.yml
├── daily_sync.py
├── requirements.txt
└── README.md (optional)
```

**To upload via GitHub web:**
1. Click "Add file" → "Upload files"
2. Drag all files (maintaining folder structure for `.github/workflows/`)
3. Commit the changes

**Or via command line:**
```bash
git clone https://github.com/YOUR_USERNAME/parentmail-calendar-sync.git
cd parentmail-calendar-sync
# Copy all the files here
git add .
git commit -m "Initial setup"
git push
```

### Step 3: Set Up GitHub Secrets

This is where you store your credentials securely. GitHub encrypts these and never exposes them in logs.

1. Go to your repository on GitHub
2. Click **Settings** → **Secrets and variables** → **Actions**
3. Click **New repository secret** for each of the following:

#### Secret 1: PARENTMAIL_EMAIL
- **Name**: `PARENTMAIL_EMAIL`
- **Value**: `sachinsharma0787@gmail.com`

#### Secret 2: PARENTMAIL_PASSWORD
- **Name**: `PARENTMAIL_PASSWORD`
- **Value**: Your ParentMail password

#### Secret 3: GOOGLE_CALENDAR_TOKEN

This one requires extracting your token from `token.json`. Here's how:

1. Find your existing `token.json` file (from your previous setup)
2. Open it and copy the **entire contents** (it should look like JSON with `token`, `refresh_token`, `client_id`, etc.)
3. Create a new secret:
   - **Name**: `GOOGLE_CALENDAR_TOKEN`
   - **Value**: Paste the entire JSON content

**Example token.json content (yours will have different values):**
```json
{
  "token": "ya29.a0AfH6SM...",
  "refresh_token": "1//0eX1Y2Z3...",
  "token_uri": "https://oauth2.googleapis.com/token",
  "client_id": "368712147369-k50u2s5bf1155dr6liohvrl4no5l8ppf.apps.googleusercontent.com",
  "client_secret": "GOCSPX-...",
  "scopes": ["https://www.googleapis.com/auth/calendar"]
}
```

### Step 4: Test the Workflow

1. Go to **Actions** tab in your repository
2. Click on "Daily ParentMail Calendar Sync" workflow
3. Click "Run workflow" → "Run workflow" (green button)
4. Wait for it to complete (usually 1-2 minutes)
5. Check the logs to ensure it ran successfully

### Step 5: Verify It's Working

After running, check your Google Calendar for new events. The automation will:
- Add new events with prefixes like `[Arvi]`, `[Rivan]`, or `[School]`
- Skip any events that already exist (no duplicates)
- Set reminders (1 day and 1 hour before)

## Troubleshooting

### "Login failed" error
- Double-check your PARENTMAIL_PASSWORD secret
- Try logging into ParentMail manually to ensure the password is correct
- ParentMail may have updated their login page - open an issue for help

### "No valid Google credentials" error
- Ensure GOOGLE_CALENDAR_TOKEN secret contains valid JSON
- The token may have expired - you might need to refresh it locally and update the secret

### "Could not find newsletter"
- This is normal if no new newsletter exists
- Newsletters are typically sent on Fridays
- The script will try again the next day

### "Could not find Sway link"
- The newsletter format may have changed
- Check ParentMail manually to see the current format

### Workflow never runs
- Check that the `.github/workflows/daily-sync.yml` file is in the correct location
- Ensure GitHub Actions is enabled for your repository (Settings → Actions → General)

## How It Works

```
Daily at 6:00 AM UK time
         │
         ▼
┌─────────────────┐
│  Login to       │
│  ParentMail     │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  Find latest    │
│  newsletter     │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  Open Sway page │
│  with diary     │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  Extract events │
│  from table     │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  Filter for     │
│  YR/Y2/KS1      │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  Check Google   │
│  Calendar for   │
│  duplicates     │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  Create new     │
│  events only    │
└─────────────────┘
```

## Customization

### Change the schedule
Edit `.github/workflows/daily-sync.yml` and modify the cron expression:
```yaml
schedule:
  - cron: '0 6 * * *'  # Currently: 6:00 AM UTC daily
```

**Common schedules:**
- `0 6 * * *` - Daily at 6:00 AM
- `0 6 * * 1-5` - Weekdays only at 6:00 AM
- `0 6 * * 5` - Fridays only at 6:00 AM (when newsletters arrive)
- `0 18 * * 5` - Fridays at 6:00 PM

### Change event colors
Edit `daily_sync.py` and modify the color constants:
```python
COLOR_ARVI = '6'    # Orange
COLOR_RIVAN = '9'   # Blue
COLOR_BOTH = '11'   # Red
```

Google Calendar color IDs:
1. Lavender, 2. Sage, 3. Grape, 4. Flamingo, 5. Banana, 6. Tangerine,
7. Peacock, 8. Graphite, 9. Blueberry, 10. Basil, 11. Tomato

### Add more year groups
Edit the filter keywords in `daily_sync.py`:
```python
INCLUDE_KEYWORDS = ['yr', 'y2', 'ks1', 'red class', 'yellow class', 'reception', 'year 2']
```

## Security Notes

- **Repository**: Keep your repo **private** - it contains automation for your personal accounts
- **Secrets**: GitHub encrypts all secrets - they're never visible in logs
- **Token**: The Google Calendar token auto-refreshes, but may need manual refresh after ~6 months
- **Password**: Never commit credentials directly to code - always use secrets

## Monitoring

### Check past runs
1. Go to **Actions** tab
2. Click on any workflow run to see logs
3. Green ✓ = success, Red ✗ = failure

### Get notified of failures
1. Go to **Settings** → **Notifications**
2. Enable "Actions" notifications
3. You'll get an email if the workflow fails

## Support

If you encounter issues:
1. Check the workflow logs in the Actions tab
2. Verify your secrets are correctly set
3. Try running manually with "Run workflow"
4. Open an issue on the repository

---

**Last updated**: February 2026
