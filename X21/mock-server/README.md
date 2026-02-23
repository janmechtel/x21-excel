> [!WARNING]
> This project is unmaintained and no longer actively developed.

# Mock Server for X21 Web-UI Testing

A lightweight WebSocket server that simulates the Deno backend for testing the web-ui in isolation.

## Running the Mock Server

### Option 1: With Web-UI (Recommended for Development)

From the `web-ui` directory:

```bash
npm run dev:mock
```

This starts both the mock server AND the web-ui with **hot reloading enabled for both**.

### Option 2: Standalone (Mock Server Only)

From this directory:

```bash
# With auto-reload (watches for file changes)
deno run --allow-net --allow-env --watch mod.ts

# Without auto-reload
deno run --allow-net --allow-env mod.ts
```

The mock server will automatically find a free port (starting from 8085).

## Features

- ✅ **Auto-reload**: Changes to scenario files are detected automatically (when using `--watch`)
- ✅ **WebSocket streaming**: Simulates Claude API streaming responses
- ✅ **Multiple scenarios**: Pre-built test scenarios for different UI states
- ✅ **Tool approvals**: Simulates tool permission requests
- ✅ **UI requests**: Supports the `collect_input` tool for forms
- ✅ **CORS enabled**: Works with any web-ui port

## Available Scenarios

The mock server chooses scenarios based on keywords in your prompt:

| Trigger Keywords | Scenario | Description |
|------------------|----------|-------------|
| "sample questions", "sample question", "financial model" | Sample Questions Form | Boolean + segmented questions demo |
| "ui demo", "all controls", "showcase" | UI Controls Showcase | Demonstrates all control types |
| "form", "ui request", "amortization" | Amortization Form | Example financial form |
| "tool", "format", "write", "read" | Tool Approval | Excel operations requiring approval |
| (anything else) | Simple Response | Basic text response |

## Testing Changes

### Quick Test Loop

1. **Edit a scenario** in `scenarios/scenarios.ts`

   ```typescript
   export function getMyTestScenario(): StreamScenario {
     // your changes here
   }
   ```

2. **Add the trigger** in `websocket-handler.ts`

   ```typescript
   const shouldShowMyTest = prompt.toLowerCase().includes("my test");

   if (shouldShowMyTest) {
     await simulateClaudeStream(socket, getMyTestScenario());
   }
   ```

3. **Save the file** - the mock server will automatically reload (if running with `--watch` or via `npm run dev:mock`)

4. **Test in the UI** - type your trigger keyword and see the result instantly

No need to restart anything! 🎉

## How It Works

```
User types message in web-ui
    ↓
web-ui connects to ws://localhost:8085/ws
    ↓
mock-server receives the message
    ↓
Chooses scenario based on prompt keywords
    ↓
Streams back realistic Claude-style responses
    ↓
web-ui renders the response
```

## File Structure

```
mock-server/
├── mod.ts                      # Server entry point
├── websocket-handler.ts        # Message routing and scenario selection
├── scenarios/
│   └── scenarios.ts            # All test scenarios
└── utils/
    └── claude-stream.ts        # Streaming simulation logic
```

## Creating New Scenarios

See the main guide at `../web-ui/TESTING-UI-REQUESTS.md` for detailed instructions on:

- Capturing UI requests from the real backend
- Creating new test scenarios
- Writing Playwright tests
- All available control types

## Port Configuration

The mock server uses port **8085** by default. You can override this with an environment variable:

```bash
MOCK_SERVER_PORT=9000 deno run --allow-net --allow-env mod.ts
```

## Health Check

```bash
curl http://localhost:8085/health
```

Expected response:

```json
{
  "status": "ok",
  "mode": "mock",
  "port": 8085
}
```

## Auto-Reload Behavior

When running with `--watch` or via `npm run dev:mock`:

- ✅ Changes to any `.ts` file in the mock-server directory trigger a reload
- ✅ WebSocket connections are gracefully closed
- ✅ Server restarts on the same port
- ⚠️ You may need to refresh the web-ui page if the connection drops during reload

## Debugging

The mock server logs all WebSocket messages to the console:

```
📱 WebSocket client connected
📨 Received: stream:start { type: 'stream:start', payload: { prompt: '...' } }
🤖 Starting stream for prompt: "sample questions..."
✅ Stream completed
```

Look for these emoji indicators to track what's happening.
