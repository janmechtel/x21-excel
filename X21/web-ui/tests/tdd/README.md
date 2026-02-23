> [!WARNING]
> This project is unmaintained and no longer actively developed.

# TDD Loop Tests

This directory contains tests used by the `/tdd-loop` Claude Code command for
AI-driven UI development.

## Directory Structure

```
tests/tdd/
├── current.spec.ts          # Main test file (auto-generated/modified)
├── screenshots/             # Screenshots captured during iterations
│   ├── iteration-1.png
│   ├── iteration-2.png
│   └── ...
└── README.md               # This file
```

## How It Works

1. **Test Generation**: The `/tdd-loop` command generates or modifies
   `current.spec.ts` based on your goal
2. **Screenshot Capture**: Test runs and captures screenshots to
   `screenshots/iteration-{N}.png`
3. **AI Evaluation**: Claude analyzes screenshots using vision capabilities
4. **Code Changes**: Claude makes targeted changes to React components
5. **Iteration**: Process repeats until goal is achieved

## Browser Modes

### Headless (default)

```bash
npm run test:tdd
```

Fast, no visible browser.

### Headed (visible browser)

```bash
npm run test:tdd:headed
```

Opens Chromium window, see changes live.

### Attach (connect to existing Chrome)

```bash
# 1. Start Chrome with debugging port:
chrome.exe --remote-debugging-port=9222

# 2. Run test:
npm run test:tdd:attach
```

Connects to your already-running Chrome instance.

## Manual Testing

You can manually run the current test:

```bash
cd X21/web-ui

# Make sure dev server is running
npm run dev

# In another terminal, run TDD test
npm run test:tdd
```

## Files Created/Modified

- **current.spec.ts**: Auto-generated test (don't edit manually)
- **screenshots/*.png**: Iteration screenshots (git-ignored)

## Environment Variables

- `ITERATION`: Current iteration number (set automatically by Claude)
- `BROWSER_MODE`: 'headless' | 'headed' | 'attach'
- `CDP_ENDPOINT`: Chrome DevTools Protocol endpoint (default:
  <http://localhost:9222>)

## Notes

- Screenshots are git-ignored by default
- Tests assume dev server is running on `http://localhost:5173`
- Each iteration overwrites the previous screenshot
- Tests use mock authentication bypass (VITE_SKIP_AUTH=true)
