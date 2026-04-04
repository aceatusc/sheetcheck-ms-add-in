# Pista — Excel Verification Agent

An Excel add-in that uses AI to help verify and modify spreadsheet agent outputs.

## How to Run

1. Install dependencies (if not already installed):
   ```bash
   npm install -g http-server
   ```

2. Start the server:
   ```bash
   npx http-server --host 0.0.0.0 --port 8881
   ```

3. Update the `.SourceLocation` in [manifest.xml](manifest.xml) to your server URL (e.g., `http://your-ip:8881/index.html`)

4. Sideload the add-in in Excel through **Add-ins > My Add-ins > Upload My Add-in** and select `manifest.xml`

## Project Structure

### Core Modules
- **app.js** — Entry point; initializes Office and wires all modules together
- **officeInit.js** — Initializes Office JavaScript API and detects Excel
- **llmClient.js** — HTTP client for backend API endpoints (`/code`, `/edit`, `/chat`, etc.)

### DAG & Execution
- **dagStore.js** — Stores and manages the directed acyclic graph of worksheet states
- **dagRunner.js** — Navigates the DAG (forward/backward/branch); handles undo/redo and graph rendering
- **executionEngine.js** — Executes code changes on the worksheet

### UI & Chat
- **chatManager.js** — Chat UI and message orchestration
- **chatHistory.js** — Persists and retrieves chat messages
- **stepNavigator.js** — Visual DAG graph display for step history
- **aspectManager.js** — Manages verification specifications/rubrics panel
- **rubricManager.js** — Handles rubric creation and verification logic

### Worksheet Context
- **worksheetContext.js** — Utilities for reading/writing Excel range data
- **worksheetSnapshot.js** — Captures and restores worksheet state snapshots

### Styles
- **theme.css** — Main theme and layout
- **chat.css** — Chat panel styling
- **execution.css** — Execution panel styling
- **stepNavigator.css** — DAG visualization styling
- **aspect.css** — Specifications panel styling

## Backend
Expects a backend at `https://sackend.isi.edu/sheetcheck/backend/addin` with endpoints:
- `/code` — Generate code for changes
- `/edit` — Edit existing code
- `/chat` — Chat endpoint
- `/rubric/verify` — Verify against rubrics
- `/ask` — Q&A endpoint
