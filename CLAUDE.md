# X21 - AI Excel Add-in with LLM Integration Posthog and Langfuse for Analytics

## Architecture

- VSTO Add-in (C#/.NET 4.8)
- Deno backend port 8000 - LLM orchestration, calls VSTO endpoints
- React UI (Vite/TypeScript) - Embedded in Excel TaskPane via WebView2

Communication: React → Deno (port 8000) → VSTO (port 8080) → Excel

## Development

### Coding Standards

- C#: Follow existing NLog patterns for logging
- TypeScript: Use strict mode, functional patterns
- React: Functional components with hooks, shadcn/ui components

### Key Files

- `X21\vsto-addin\ThisAddIn.cs` - Add-in entry point, process management
- `X21\vsto-addin\Services\ExcelApiService.cs` - HTTP endpoints for Excel operations
- `X21\deno-server\src\tools\` - LLM tool definitions
- `X21\web-ui\src\App.tsx` - React root component
