# OmniHR Integration

A Node.js tool to fetch leave data from OmniHR API and export it to Excel/Google Sheets.

## Setup

1. Install dependencies:

   ```bash
   npm install
   ```

2. Copy `.env.example` to `.env` and fill in your OmniHR credentials:
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

## Google Apps Script

The `src/google-appscript/Code.gs` file contains a Google Apps Script that can be deployed directly to Google Sheets for automatic leave data syncing.
