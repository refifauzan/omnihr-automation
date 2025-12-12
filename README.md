# OmniHR Integration

A Node.js tool to fetch leave data from OmniHR API and export it to Excel/Google Sheets, with automatic deployment to Google Apps Script.

## Setup

1. Install dependencies:

   ```bash
   npm install
   ```

2. Copy `.env.example` to `.env` and fill in your credentials:
   ```bash
   cp .env.example .env
   ```

## Scripts

| Command                 | Description                       |
| ----------------------- | --------------------------------- |
| `npm start`             | Run the main application          |
| `npm run fetch-leaves`  | Fetch leave data from OmniHR API  |
| `npm run update-excel`  | Update Excel file with leave data |
| `npm run export-sheets` | Export data to Google Sheets      |
| `npm run clasp:push`    | Deploy to Google Apps Script      |
| `npm run clasp:pull`    | Pull from Google Apps Script      |
| `npm run clasp:open`    | Open Apps Script in browser       |

---

## Google Apps Script

The `src/google-appscript/Code.gs` file contains a Google Apps Script for automatic leave data syncing in Google Sheets.

### Features

- Sync leave data from OmniHR API
- Full-day leave (RED) and half-day leave (ORANGE) highlighting
- Attendance hours tracking per project
- Time off override checkbox support
- Scheduled auto-sync (daily/monthly)

### Manual Setup

1. Open your Google Sheet
2. Go to **Extensions > Apps Script**
3. Copy contents of `src/google-appscript/Code.gs`
4. Run **OmniHR > Setup API Credentials**

---

## CI/CD Deployment

Automatically deploy to Google Apps Script via GitHub Actions.

### Prerequisites

1. Enable **Apps Script API** at [script.google.com/home/usersettings](https://script.google.com/home/usersettings)

### Setup OAuth Client ID (GCP)

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create/select a project
3. Go to **APIs & Services > Credentials**
4. Create **OAuth client ID** (Desktop app)
5. Download the JSON file

### Get Clasp Credentials

```bash
npm install -g @google/clasp
clasp login --creds /path/to/client_secret.json
cat ~/.clasprc.json
```

### Get Script ID

1. Open Google Sheet → **Extensions > Apps Script**
2. Go to **Project Settings** (⚙️)
3. Copy the **Script ID**

### Configure GitHub Secrets

Go to **GitHub Repo > Settings > Secrets > Actions** and add:

| Secret Name         | Value                         |
| ------------------- | ----------------------------- |
| `SCRIPT_ID`         | Your Apps Script ID           |
| `CLASP_CREDENTIALS` | Contents of `~/.clasprc.json` |

### Deploy

Push to `main` branch or manually trigger via **GitHub Actions > Run workflow**.

---

## Project Structure

```
├── src/
│   ├── google-appscript/
│   │   ├── Code.gs          # Google Apps Script code
│   │   └── appsscript.json  # Apps Script manifest
│   ├── data/                # Data files
│   ├── config.js            # Configuration
│   └── *.js                 # Node.js scripts
├── .github/workflows/       # CI/CD workflows
└── .env                     # Environment variables
```
