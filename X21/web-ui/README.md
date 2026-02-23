> [!WARNING]
> This project is unmaintained and no longer actively developed.

# X21 web-ui

This is the React-based web interface for the X21 Excel add-in, served in a
WebView2 control.

## Development

### Pre-commit checks

The repository uses [pre-commit](https://pre-commit.com/) to run builds for the
backend, web UI, and VSTO add-in before allowing a commit.

Hooks run the following commands:

- `deno task build` (in `X21/deno-server`)
- `npm run build` (in `X21/web-ui`)
- `dotnet build X21.sln`

Install and enable the hooks:

```bash
pip install pre-commit
pre-commit install
```

You can skip these checks with `git commit --no-verify` or by setting
`SKIP=deno-build,web-ui-build,dotnet-build` for a specific commit.

### Prerequisites

- Node.js (v16 or higher)
- npm

### Setup

1. Install dependencies:

   ```bash
   npm install
   ```

2. Start development server:

   ```bash
   npm run dev
   ```

   This will start the Vite dev server on `http://localhost:5173`

### Development Workflow

- The development server runs on `http://localhost:5173`
- In DEBUG mode, the WebView2 will connect to this dev server
- In RELEASE mode, the WebView2 will serve the built files from
  `TaskPane/WebAssets`

## Building for Production

### Manual Build

```bash
npm run build
```

This will build the app and output files to `../TaskPane/WebAssets/`

### Clean Build

```bash
npm run build:clean
```

This will clean the output directory and rebuild from scratch.

### Watch Mode

```bash
npm run build:watch
```

This will watch for changes and rebuild automatically.

## Integration with Main Project

### Automatic Build

The main C# project includes MSBuild tasks that automatically build the WebUI:

- **Release builds**: web-ui is built automatically before the main project
  build
- **Debug builds**: web-ui build is attempted but won't fail the main build if
  it fails

### Manual Build Scripts

You can also use the provided build scripts from the main project directory:

- `build-webui.bat` - Windows batch script
- `build-webui.ps1` - PowerShell script (recommended)

### File Structure

```
X21/
├── WebUI/                 # Source code
│   ├── src/
│   ├── package.json
│   └── vite.config.ts
└── TaskPane/
    └── WebAssets/         # Built files (generated)
        ├── index.html
        ├── assets/
        │   ├── index.[hash].js
        │   └── index.[hash].css
        └── vite.svg
```

## Configuration

### Vite Configuration

The `vite.config.ts` is configured to:

- Build to `../TaskPane/WebAssets`
- Use relative paths (`base: './'`)
- Generate hashed filenames for cache busting
- Optimize for production (minification, no sourcemaps)

### WebView2 Integration

The WebView2 control in `WebView2TaskPaneHost.cs` handles:

- **DEBUG mode**: Connects to `http://localhost:5173`
- **RELEASE mode**: Serves from local `file://` URLs in `TaskPane/WebAssets`

## Troubleshooting

### Build Issues

1. Ensure Node.js and npm are installed and in PATH
2. Run `npm install` to install dependencies
3. Check for TypeScript compilation errors
4. Verify the output directory is writable

### Runtime Issues

1. Check that WebAssets files exist in the output directory
2. Verify file paths in the built `index.html`
3. Check WebView2 console for JavaScript errors
4. Ensure the WebView2 runtime is installed

### Development vs Production

- **Development**: Uses Vite dev server with hot reload
- **Production**: Uses static files with optimized builds
- Switch between modes by changing the build configuration in Visual Studio
