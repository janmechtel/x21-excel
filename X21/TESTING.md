# X21 Testing Guide

This document covers all testing approaches for the X21 project: mock server testing, automated Playwright tests, and AI-driven TDD workflow.

## Architecture

```
┌─────────────────┐
│  Playwright     │
│  Test Runner    │
└────────┬────────┘
         │
         ▼
┌─────────────────┐      WebSocket      ┌─────────────────┐
│   Web-UI in     │◄────────────────────►│  Mock Server    │
│   Browser       │   ws://localhost:PORT│  (Deno)         │
│  + Mock         │                      │  Simulates:     │
│    WebView2     │                      │  - Claude API   │
│    Bridge       │                      │  - Streaming    │
└─────────────────┘                      │  - Tools        │
                                         └─────────────────┘
```

## Quick Start

### Option 1: Unified Dev Command (Recommended)

```bash
cd X21/web-ui
npm run dev:mock
```

This automatically:

- Starts mock-server on a free port
- Starts web-ui with correct port configuration
- Bypasses email authentication
- Handles cleanup on exit

### Option 2: Manual Setup

```bash
# Terminal 1: Start mock server
cd X21/mock-server
deno run --allow-net mod.ts

# Terminal 2: Start web-ui
cd X21/web-ui
npm run dev
```

## Running Tests

```bash
cd X21/web-ui

# All tests
npm test

# Test suites
npm run test:smoke          # Basic functionality
npm run test:visual         # Screenshot-based tests
npm run test:ui             # Interactive mode
npm run test:headed         # See browser
npm run test:debug          # Debug mode
```

## Test Suites

### Smoke Tests (`tests/smoke.spec.ts`)

Verifies:

- App loads successfully
- Can type and send messages
- Receives assistant responses
- Mock WebView2 bridge working
- WebSocket connection established
- No critical console errors

### Visual Tests (`tests/visual.spec.ts`)

Captures screenshots of:

- Empty state, user messages, streaming responses
- Tool approval UI, conversations
- Scrolled content, responsive layouts
- Dark mode, settings panel, edge cases

Screenshots saved to: `web-ui/tests/screenshots/`

## TDD Workflow (AI-Driven UI Development)

AI-powered Test-Driven Development for rapid UI iteration using Playwright and Claude's vision capabilities.

### Usage

```bash
# In Claude Code
/tdd-loop "center align the chat input box"
```

Claude will:

1. Ask for browser mode (headless/headed/attach)
2. Generate test in `tests/tdd/current.spec.ts`
3. Run test and capture screenshots
4. Evaluate screenshot using AI vision
5. Make code changes to achieve goal
6. Iterate until complete

### Browser Modes

**Headless (Fast)**

```bash
npm run test:tdd
```

No visible browser, fastest execution.

**Headed (Recommended for UI)**

```bash
npm run test:tdd:headed
```

Opens Chromium window, see changes in real-time.

**Attach to Chrome**

```bash
# 1. Start Chrome with debugging
chrome.exe --remote-debugging-port=9222 --user-data-dir="%TEMP%\chrome-debug"

# 2. Navigate to http://localhost:5173

# 3. Run test
npm run test:tdd:attach
```

Uses your actual Chrome browser with dev tools.

### TDD Best Practices

- Start with simple, specific goals
- Use headed mode for visual work
- Use headless for logic/functionality
- One goal at a time
- Stop early if stuck after 3-4 iterations
- Review changes and run full test suite after

### Custom Iteration

```bash
cross-env ITERATION=5 npm run test:tdd
```

Screenshots: `X21/web-ui/tests/tdd/screenshots/iteration-{N}.png`

## Mock Server Details

### Message Types Supported

`stream:start`, `stream:cancel`, `chat:restart`, `tool:permission:response`, `tool:approve`, `tool:reject`, `tool:view`, `tool:unview`, `tool:revert`, `tool:apply`, `score:score`, `score:feedback`, `user:email_response`

### Built-in Scenarios

Located in `mock-server/scenarios/scenarios.ts`:

- `getSimpleResponse()` - Basic text response
- `getResponseWithThinking()` - Response with thinking block
- `getToolApprovalScenario()` - Tool approval flow
- `getSingleToolScenario()` - Single tool use
- `getBatchToolScenario()` - Multiple tools
- `getUiRequestScenario()` - Amortization form
- `getUiControlsShowcaseScenario()` - All control types demo
- `getSampleQuestionsScenario()` - Sample questions form
- `getLongResponse()` - Long text for scrolling

### WebView2 Bridge Mock

`web-ui/public/mock-webview.js` provides:

- `window.chrome.webview` API
- `postMessage()` and `addEventListener()`
- Mock responses for: `getWorkbookName`, `getWebSocketUrl`, `getWorkbookPath`, `getSlashCommandsFromSheet`

Auto-detects if real WebView2 exists.

## Testing UI Requests (collect_input)

### Quick Test Sample Questions

```bash
cd X21/web-ui
npm run dev:mock
```

Type in UI: "sample questions" or "financial model"

### Capturing New UI Requests

**Option A: Browser DevTools**

1. Open DevTools (F12) → Network → WS
2. Click WebSocket connection
3. Find message with `type: "ui:request"`
4. Copy `payload` field

**Option B: Console Script**

1. Open DevTools → Console
2. Paste contents of `capture-ui-request.js`
3. Trigger prompt
4. Access via `window.__lastUiRequest`

### Adding New Scenarios

1. Create function in `mock-server/scenarios/scenarios.ts`
2. Import in `mock-server/websocket-handler.ts`
3. Add trigger condition in `handleStreamStart`
4. Test with trigger keyword
5. Create test file in `tests/`

### Available Control Types

- `boolean` - Yes/No questions
- `segmented` - Single choice from options
- `multi_choice` - Multiple selections
- `range_picker` - Excel range selection
- `text` - Free text input

## Test Helpers

Located in `web-ui/tests/helpers/`:

### `test-actions.ts`

`sendChatMessage`, `waitForAssistantResponse`, `waitForStreamingComplete`, `isToolApprovalVisible`, `approveAllTools`, `rejectAllTools`, `openSettings`, `toggleAutoApprove`, `clearChat`, `takeScreenshot`, `waitForAppReady`, `bypassAuthentication`

### `mock-server.ts`

`startMockServer`, `stopMockServer`, `isMockServerRunning`

## Troubleshooting

### Mock Server Won't Start

**Error:** Address already in use
**Solution:** Use `npm run dev:mock` (finds free port automatically)

### Tests Can't Connect

**Solution:**

- Ensure mock server running: `http://localhost:8085/health`
- Check firewall isn't blocking port
- Increase timeout in `tests/helpers/mock-server.ts`

### WebView2 Mock Not Loading

**Solution:**

- Check `index.html` includes `<script src="/mock-webview.js"></script>`
- Check browser console for mock loading message
- Verify `public/mock-webview.js` exists

### Dev Server Not Running (TDD)

**Error:** `page.goto: net::ERR_CONNECTION_REFUSED`
**Solution:**

```bash
cd X21/web-ui
npm run dev:mock
```

### Chrome Attach Not Working

**Solution:**

```bash
taskkill /F /IM chrome.exe
chrome.exe --remote-debugging-port=9222 --user-data-dir="%TEMP%\chrome-debug"
```

### Changes Not Visible in Headed Mode

**Cause:** Vite HMR delay
**Solution:** Wait 1-2 seconds or manually refresh browser

## File Structure

```
x21-x22/
├── .claude/commands/tdd-loop.md        # TDD slash command
├── X21/
│   ├── mock-server/
│   │   ├── mod.ts                      # Server entry point
│   │   ├── websocket-handler.ts        # WebSocket handler
│   │   └── scenarios/scenarios.ts      # Mock scenarios
│   └── web-ui/
│       ├── playwright.config.ts        # Main Playwright config
│       ├── playwright-tdd.config.ts    # TDD-specific config
│       ├── public/mock-webview.js      # WebView2 mock
│       └── tests/
│           ├── smoke.spec.ts           # Smoke tests
│           ├── visual.spec.ts          # Visual tests
│           ├── helpers/                # Test helpers
│           └── tdd/
│               ├── current.spec.ts     # Auto-generated TDD test
│               └── screenshots/        # TDD screenshots
└── TESTING.md                          # This file
```

## CI/CD Integration

```bash
# Install dependencies
cd X21/web-ui && npm ci

# Install browsers
npx playwright install --with-deps chromium

# Start services (in background)
cd X21/mock-server && deno run --allow-net mod.ts &
cd X21/web-ui && npm run dev &

# Run tests
cd X21/web-ui && npm test
```

## Manual Testing in Browser

```bash
cd X21/web-ui
npm run dev:mock
```

Open `http://localhost:5173` and test freely. WebView2 mock loads automatically.

## TDD Limitations

**Can do:** Layout, styling, visibility, CSS/Tailwind changes, simple component modifications
**Can't do:** Complex state logic, API integration, multi-page flows, performance optimization, accessibility

For complex changes, use TDD for visual parts, then add logic manually.

## Resources

- [Playwright Documentation](https://playwright.dev)
- [Deno WebSocket Guide](https://deno.land/manual/runtime/web_platform_apis/websockets)
- [Claude API Streaming](https://docs.anthropic.com/claude/reference/messages-streaming)
