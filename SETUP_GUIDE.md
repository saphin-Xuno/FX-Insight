# FX Insights Automation — Setup Guide

## What this does
Every weekday after market close (NPT ~03:45), GitHub Actions will:
1. Run `main_script.py` to fetch live market data
2. Update `Forex_Insights.xlsx` with a new row
3. Upload the Excel file to your Google Drive folder
4. Send a Slack message in your existing format

---

## Step 1 — Add these files to your project

Copy these files into your repo:
- `notifier.py`        → root of the project
- `drive_uploader.py`  → root of the project
- `monitor.yml`        → `.github/workflows/monitor.yml`

---

## Step 2 — Edit main_script.py (2 small additions)

Open `main_script.py` and add these imports near the top:

```python
from notifier import send_slack_summary
from drive_uploader import upload_to_drive
```

Then at the end of the `run()` function, after `workbook.save(EXCEL_FILE)`, add:

```python
    upload_to_drive()
    send_slack_summary(state)
```

---

## Step 3 — Update requirements.txt

Add these lines to your `requirements.txt`:

```
google-api-python-client
google-auth
```

---

## Step 4 — Set up GitHub Secrets

Go to your GitHub repo → **Settings → Secrets and variables → Actions → New repository secret**

| Secret name | What it is |
|---|---|
| `SLACK_WEBHOOK_URL` | Your Slack webhook URL (you already have this) |
| `GOOGLE_SERVICE_ACCOUNT_JSON` | Full JSON content of your service account key (see Step 5) |
| `GOOGLE_DRIVE_FOLDER_ID` | The folder ID from your Google Drive URL (see Step 6) |

---

## Step 5 — Create a Google Service Account

1. Go to https://console.cloud.google.com
2. Create a new project (or use an existing one)
3. Go to **APIs & Services → Enable APIs** → search for **Google Drive API** → Enable it
4. Go to **APIs & Services → Credentials → Create Credentials → Service Account**
5. Give it any name (e.g. `fx-insights-bot`) → click **Done**
6. Click the service account you just created → **Keys** tab → **Add Key → Create new key → JSON**
7. It downloads a `.json` file — open it and copy the entire contents
8. Paste that into the `GOOGLE_SERVICE_ACCOUNT_JSON` GitHub secret

---

## Step 6 — Share your Google Drive folder with the service account

1. In Google Drive, open (or create) the folder where you want the Excel file
2. Copy the folder ID from the URL:
   `https://drive.google.com/drive/folders/`**`1aBcDeFgHiJkLmNoPqRsTuVwXyZ`**
3. Add that ID as the `GOOGLE_DRIVE_FOLDER_ID` GitHub secret
4. Right-click the folder → **Share** → paste the service account email
   (it looks like `fx-insights-bot@your-project.iam.gserviceaccount.com`)
5. Give it **Editor** access → click **Send**

---

## Step 7 — Test it manually

1. Go to your GitHub repo → **Actions** tab
2. Click **FX Insights daily run** on the left
3. Click **Run workflow** → **Run workflow**
4. Watch the run — it should go green ✅
5. Check your Slack for the message and your Drive folder for the Excel file

---

## Schedule timing

The workflow runs at **22:15 UTC** on weekdays, which is:
- **03:45 NPT** (Nepal Time) — after NY close in winter (EST)

During US summer (March–October, DST active), NY closes one hour earlier.
To adjust, change the cron in `monitor.yml`:
- Winter (Nov–Feb): `"15 22 * * 1-5"` = 22:15 UTC
- Summer (Mar–Oct): `"15 21 * * 1-5"` = 21:15 UTC

---

## Slack message format

The message will look exactly like your sample:

```
FX Insights (USD/NPR):
Opening 15th June (Monday): NPR 151.82
Closing 12th June (Friday): NPR 151.70
Percentage Change in Price: -0.08%

Commodity Market Movement:
• Gold: $4,309.56 (+2.15%)
• Silver: $69.98 (+2.85%)
• Crude Oil: $83.35 (-1.94%)

USA Capital Market Updates:
• NASDAQ 100: 30,543.92 points (+0.84%)
• S&P 500: 7,516.75 points (+0.50%)

FII / DII Activity (India):
• FII Net: ₹1,234.56 Cr
• DII Net: ₹2,345.67 Cr
```
