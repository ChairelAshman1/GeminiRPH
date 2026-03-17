# Gemini RPH

## What this repo contains
- Root `index.html` + `geminiRPH.js` (web app for Render static site)
- `Gemini Extension Assistant/*` original extension source
- `code.gs.txt` Apps Script for writing/reading sheet via HTTP

## Deploy with Render (static site)
1. Go to https://render.com and sign in.
2. Click "New" -> "Static Site".
3. Connect GitHub and choose repo `ChairelAshman1/GeminiRPH`.
4. Configure:
   - Branch: `main`
   - Root directory: `.`
   - Build command: (empty)
   - Publish directory: `.`
5. Create site and wait for deploy.
6. Visit generated URL, e.g. `https://<appname>.onrender.com`.

## Google Apps Script backend setup
1. Open your Google Sheet.
2. Extensions -> Apps Script.
3. Copy `Gemini Extension Assistant/code.gs.txt` content into `Code.gs`.
4. Deploy -> New deployment -> Web app.
   - Execute as: `Me`
   - Who has access: `Anyone` (or Anyone with link)
5. Use the Web App URL in the web UI `Apps Script URL` field.

## Usage
1. In Render site, put:
   - Apps Script URL (from deployment)
   - `Sheet ID` optional (for different spreadsheet)
   - `Sheet Name`, `A2` range
2. Click `Sambung`.
3. Generate RPH.
4. Click `Hantar ke Sheet`.

## Git workflow
```bash
git add .
git commit -m "your message"
git push
```

## Notes
- For `index.html` referencing external resources, ensure relative path is correct.
- If you want the UI root at `/`, keep `index.html` (already done).
- If you want a separate webapp path, use Render `Publish Directory`.
