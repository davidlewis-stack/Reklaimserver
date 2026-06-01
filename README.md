# Reklaim Deck Server

Express API that generates Reklaim sales decks as PPTX files.

## Deploy to Railway (5 minutes)

1. Push this repo to GitHub
2. Go to railway.app → New Project → Deploy from GitHub repo
3. Select this repo — Railway auto-detects Node.js
4. Settings → Networking → Generate Domain
5. Copy your Railway URL (e.g. `reklaim-server-production.up.railway.app`)
6. Paste it into the React artifact (replace `http://localhost:3456`)

## Local development

```bash
npm install
node server.js
# Server runs on http://localhost:3456
```

## API

**POST /generate**

Body: `{ "config": { ...deck_config fields... } }`

Returns: PPTX file download

## Files

- `server.js` — Express API, handles config → PPTX pipeline
- `reklaim_deck_template.js` — PPTX rendering engine (do not edit per-pitch)
