# Oliver - Word Plug-in

A Microsoft Word add-in that integrates Oliver's core features into Word.

## Features

- **Summarize** - Get quick summaries of selected text
- **Compare** - Analyze differences between clauses  
- **Redraft** - Replace selected text with AI-powered rewritten draft
- **Chat** - Ask questions about your document

## Quick Start

### Setup
```bash
cd frontend
npm install
npm run dev
```

Open the dev server at `https://localhost:3000` (HTTPS required for Office add-ins).

### Load into Word
**Word Online (easiest):**
1. Open Word Online
2. Add-ins > My Add-ins > Upload My Add-in
3. Upload the `manifest.xml` file

**Word Desktop:**
1. File > Options > Trust Center > Trusted Add-in Catalogs
2. Add this project's directory
3. Restart Word, then Insert > My Add-ins > Oliver AI Assistant

### Backend
The frontend expects a FastAPI backend at `http://localhost:8000` with endpoints:
- `POST /api/summarize`
- `POST /api/compare` 
- `POST /api/redraft`
- `POST /api/analyze`

Change the URL in `src/services/apiService.ts` if needed.

## Tech stack

- **Frontend & styling**: React 19 + TypeScript
- **Build tool**: Vite
- **Word integration**: Office.js
- **Additional styling**: Plain CSS (no UI framework)

## Deployment

```bash
npm run build
```

Host the built files on HTTPS and update manifest.xml URLs to custom domain if required.

## Troubleshooting

- **"Office.js not loaded"** - Make sure that you are in Word, not a browser
- **HTTPS warnings** - Accept the self-signed cert by visiting https://localhost:3000