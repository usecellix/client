## Phase 0: Frontend ↔ Backend Contract

### Base URL configuration
- **Env var**: `VITE_API_BASE_URL`
- **Default (local)**: `http://localhost:4001`

Create `frontend/.env` from `frontend/.env.example` to point the add-in at a backend.

### Streaming endpoint
- **POST**: `/excel-ai/process`
- **Request body**: `{ prompt: string, sheetData: any[][] }`
- **Response**: `text/event-stream`

### SSE events (expected)
- **`chunk`**: `{ "text": string }` (append to assistant text)
- **`actions`**: `{ "actions": SheetAction[], "explanation": string }` (Excel apply payload)
- **`status`**: `{ "message": string }` (optional progress updates)
- **`error`**: `{ "message": string }`
- **end-of-stream**: either a `done` event or stream close

### Error envelope (non-200 responses)
When possible, backend returns JSON like:

```json
{
  "success": false,
  "traceId": "…",
  "error": { "code": "…", "message": "…", "details": {} }
}
```

Frontend behavior:
- **UI**: show `error.message`
- **Console**: log `traceId` for support/debug

