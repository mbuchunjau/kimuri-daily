# Kimuri Daily

Phone-friendly evening entry for Kimuri: stock and cash go straight into the Google Sheet (one tab per day), yesterday's closing stock becomes today's balance, and summaries go to Telegram.

## Structure
- `public/index.html` – daily entry form
- `public/summary.html` – view and send summaries and reports
- `public/style.css`, `public/app.js` – shared styling and helpers
- `api/gas.js` – Vercel function: checks the PIN, forwards to Apps Script
- `Code.gs` – Google Apps Script backend (paste into the sheet's Apps Script project)

## Setup
1. **Apps Script:** paste `Code.gs`. In Project Settings > Script properties add `TG_TOKEN`, `TG_CHAT_ID`, `API_KEY`. Deploy as a web app (Execute as: Me, Access: Anyone) and copy the `/exec` URL.
2. **Vercel:** import this repo (framework: Other). Add environment variables `GAS_URL` (the `/exec` URL), `GAS_KEY` (same value as `API_KEY`) and `APP_PIN`.
3. Open the Vercel URL on the cashier's phone and add it to the home screen.

## Secrets
No secrets belong in this repo. Tokens and keys live in Apps Script properties and Vercel environment variables only.
