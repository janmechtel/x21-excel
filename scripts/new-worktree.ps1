param(
    [Parameter(Mandatory=$true)]
    [string]$feature,
    [Parameter(Mandatory=$false)]
    [switch]$SkipInstall,
    [Parameter(Mandatory=$false)]
    [switch]$SkipVSCode,
    [Parameter(Mandatory=$false)]
    [switch]$SkipDevServer
)

$ErrorActionPreference = 'Stop'

# Store the current directory
$sourceDir = Get-Location
$worktreePath = "../x21-$feature"

# Create git worktree
Write-Host "Creating git worktree for feature: $feature" -ForegroundColor Green
git worktree add $worktreePath -b $feature

if ($LASTEXITCODE -ne 0) {
    Write-Host "Failed to create worktree" -ForegroundColor Red
    exit 1
}

# Change to the new worktree directory
Set-Location $worktreePath
$worktreeAbsPath = Get-Location

Write-Host "Worktree created at: $worktreeAbsPath" -ForegroundColor Cyan

# Copy configuration files
Write-Host "`nCopying configuration files..." -ForegroundColor Green

$appsettingsSource = Join-Path $sourceDir "X21\vsto-addin\appsettings.json"
$appsettingsDest = "X21\vsto-addin\appsettings.json"
if (Test-Path $appsettingsSource) {
    Copy-Item $appsettingsSource $appsettingsDest -Force
    Write-Host "  ✓ Copied appsettings.json" -ForegroundColor Gray
} else {
    Write-Host "  ⚠ Warning: appsettings.json not found at $appsettingsSource" -ForegroundColor Yellow
}

$envSource = Join-Path $sourceDir "X21\deno-server\.env"
$envDest = "X21\deno-server\.env"
if (Test-Path $envSource) {
    Copy-Item $envSource $envDest -Force
    Write-Host "  ✓ Copied .env" -ForegroundColor Gray
} else {
    Write-Host "  ⚠ Warning: .env not found at $envSource" -ForegroundColor Yellow
}

$claudeSource = Join-Path $sourceDir ".claude"
$claudeDest = ".claude"
if (Test-Path $claudeSource) {
    Copy-Item $claudeSource $claudeDest -Recurse -Force
    Write-Host "  ✓ Copied .claude directory" -ForegroundColor Gray
} else {
    Write-Host "  ⚠ Warning: .claude directory not found at $claudeSource" -ForegroundColor Yellow
}

$codannaSource = Join-Path $sourceDir ".codanna"
$codannaDest = ".codanna"
if (Test-Path $codannaSource) {
    Copy-Item $codannaSource $codannaDest -Recurse -Force
    Write-Host "  ✓ Copied .codanna directory" -ForegroundColor Gray
} else {
    Write-Host "  ⚠ Warning: .codanna directory not found at $codannaSource" -ForegroundColor Yellow
}

# Install dependencies
if (-not $SkipInstall) {
    Write-Host "`nInstalling dependencies..." -ForegroundColor Green

    # Root npm ci
    Write-Host "  Installing root dependencies..." -ForegroundColor Cyan
    npm ci
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  ✗ Failed to install root dependencies" -ForegroundColor Red
        exit 1
    }
    Write-Host "  ✓ Root dependencies installed" -ForegroundColor Gray

    # Web-UI npm ci
    Write-Host "  Installing web-ui dependencies..." -ForegroundColor Cyan
    Push-Location "X21\web-ui"
    npm ci
    $webUiInstallResult = $LASTEXITCODE
    Pop-Location
    if ($webUiInstallResult -ne 0) {
        Write-Host "  ✗ Failed to install web-ui dependencies" -ForegroundColor Red
        exit 1
    }
    Write-Host "  ✓ Web-ui dependencies installed" -ForegroundColor Gray

    # Deno-Server npm ci
    Write-Host "  Installing deno-server dependencies..." -ForegroundColor Cyan
    Push-Location "X21\deno-server"
    npm ci
    $denoServerInstallResult = $LASTEXITCODE
    Pop-Location
    if ($denoServerInstallResult -ne 0) {
        Write-Host "  ✗ Failed to install deno-server dependencies" -ForegroundColor Red
        exit 1
    }
    Write-Host "  ✓ Deno-server dependencies installed" -ForegroundColor Gray
} else {
    Write-Host "`nSkipping dependency installation (use -SkipInstall to skip)" -ForegroundColor Yellow
}

# Open VS Code
if (-not $SkipVSCode) {
    Write-Host "`nOpening VS Code..." -ForegroundColor Green
    code .
    Write-Host "  ✓ VS Code opened" -ForegroundColor Gray
} else {
    Write-Host "`nSkipping VS Code launch" -ForegroundColor Yellow
}

# Start dev:mock
if (-not $SkipDevServer) {
    Write-Host "`nStarting dev:mock server..." -ForegroundColor Green
    Write-Host "  Branch: $feature" -ForegroundColor Cyan
    Write-Host "  Port will be auto-detected by dev:mock" -ForegroundColor Cyan
    Write-Host "  Press Ctrl+C to stop the server`n" -ForegroundColor Yellow

    # Change to web-ui directory
    Push-Location "X21\web-ui"
    npm run dev:mock
    Pop-Location
} else {
    Write-Host "`nSkipping dev server launch" -ForegroundColor Yellow
    Write-Host "To start the dev server manually, run:" -ForegroundColor Cyan
    Write-Host "  cd X21\web-ui" -ForegroundColor Gray
    Write-Host "  npm run dev:mock" -ForegroundColor Gray
}

Write-Host "`n✓ Worktree setup complete!" -ForegroundColor Green
Write-Host "Location: $worktreeAbsPath" -ForegroundColor Cyan
