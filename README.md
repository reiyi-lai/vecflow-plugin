# Oliver AI Legal Assistant - Word Add-in

A Microsoft Word add-in that embeds Oliver's AI-powered legal workflow features directly into Word, enabling lawyers to summarize, compare, redraft, and analyze legal documents seamlessly.

## Features

- **📝 Summarize**: AI-powered text summarization of selected content
- **⚖️ Compare**: Side-by-side clause comparison with detailed analysis
- **✍️ Redraft**: AI-assisted text redrafting with inline suggestions
- **💬 Assistant**: Chat-style AI assistant for custom legal queries

## Architecture

```
┌─────────────────────────────────┐
│ Microsoft Word                  │
│ ┌─────────────────────────────┐ │
│ │ Task Pane (React App)       │ │
│ │ - Summarize                 │ │
│ │ - Compare                   │ │
│ │ - Redraft                   │ │
│ │ - Chat Assistant            │ │
│ └─────────────────────────────┘ │
│         │ Office.js API         │
└─────────────────────────────────┘
          │ HTTPS API calls
          ▼
┌─────────────────────────────────┐
│ FastAPI Backend                 │
│ - /api/summarize               │
│ - /api/compare                 │
│ - /api/redraft                 │
│ - /api/analyze                 │
└─────────────────────────────────┘
```

## Setup Instructions

### Prerequisites
- Node.js 18+ 
- npm or yarn
- Microsoft Word (Desktop or Online)

### 1. Install Dependencies
```bash
cd frontend
npm install
```

### 2. Start Development Server
```bash
npm run dev
```
This will start the server at `https://localhost:3000`

> **Note**: The server uses HTTPS as required by Office.js add-ins

### 3. Load Add-in in Word

#### Option A: Word Online (Recommended for Development)
1. Open Word Online
2. Go to **Insert** > **Add-ins** > **Upload My Add-in**
3. Upload the `manifest.xml` file from the root directory
4. The add-in will appear in the task pane

#### Option B: Word Desktop
1. Open Word Desktop
2. Go to **File** > **Options** > **Trust Center** > **Trust Center Settings** > **Trusted Add-in Catalogs**
3. Add the directory containing `manifest.xml`
4. Restart Word
5. Go to **Insert** > **My Add-ins** > **Shared Folder** > **Oliver AI Assistant**

### 4. Backend Setup (Required for Full Functionality)

The frontend expects a FastAPI backend running on `http://localhost:8000` with these endpoints:

```python
# Expected API endpoints
POST /api/summarize
POST /api/compare  
POST /api/redraft
POST /api/analyze
```

Update the `API_BASE_URL` in `src/services/apiService.ts` if your backend runs on a different port.

## Project Structure

```
├── frontend/
│   ├── src/
│   │   ├── components/
│   │   │   ├── Sidebar.tsx          # Main container
│   │   │   ├── SummarizeButton.tsx  # Summarization feature
│   │   │   ├── ComparePanel.tsx     # Clause comparison
│   │   │   ├── RedraftPanel.tsx     # Text redrafting
│   │   │   └── AssistantChat.tsx    # AI chat interface
│   │   ├── office/
│   │   │   └── wordApi.ts           # Office.js wrapper
│   │   ├── services/
│   │   │   └── apiService.ts        # API service layer
│   │   ├── App.tsx                  # Main app component
│   │   └── App.css                  # Styling
│   ├── package.json
│   └── vite.config.ts
├── manifest.xml                     # Office add-in manifest
└── README.md
```

## Usage

### Summarize
1. Select text in your Word document
2. Click **Summarize** tab
3. Click **Summarize Selected Text**
4. Review the AI-generated summary
5. Click **Insert into Document** to add it to your document

### Compare
1. Click **Compare** tab
2. Click **Select First Clause** and highlight the first clause
3. Click **Select Second Clause** and highlight the second clause
4. Click **Compare Clauses**
5. Review the detailed comparison analysis
6. Click **Insert into Document** to add the comparison

### Redraft
1. Click **Redraft** tab
2. Click **Select Text to Redraft** and highlight the text
3. Optionally add specific instructions
4. Click **Generate Redraft**
5. Review the AI suggestion side-by-side with original
6. Click **Accept & Replace** or **Reject**

### Assistant
1. Click **Assistant** tab
2. Configure options:
   - **Include selected text**: Include highlighted text in context
   - **Include entire document**: Include full document for context
3. Type your question or request
4. Review the AI response
5. Click **Insert into Document** to add the response

## Development

### Available Scripts
- `npm run dev` - Start development server
- `npm run build` - Build for production
- `npm run lint` - Run ESLint
- `npm run preview` - Preview production build

### Key Technologies
- **React 19** with TypeScript
- **Vite** for fast development
- **Office.js** for Word integration
- **CSS3** for styling (no external UI library)

## Deployment

### Building for Production
```bash
npm run build
```

### Hosting Requirements
- HTTPS hosting (required for Office add-ins)
- Update manifest.xml URLs to point to your production domain
- Ensure CORS is configured for your backend

## Troubleshooting

### Common Issues

**"Office.js not loaded" error**
- Ensure you're testing in Microsoft Word, not a regular browser
- Check that the manifest.xml is properly loaded

**API connection errors**
- Verify your backend is running on the correct port
- Check CORS settings on your backend
- Update API_BASE_URL in apiService.ts if needed

**HTTPS certificate warnings**
- For development, you may need to accept the self-signed certificate
- Visit https://localhost:3000 directly and accept the certificate

### Debug Mode
Open browser dev tools in Word Online to debug the add-in:
1. Right-click in the task pane
2. Select **Inspect** or **Inspect Element**
3. Use the console to debug JavaScript errors

## License

This project is part of Vecflow's Oliver AI Legal Assistant platform.