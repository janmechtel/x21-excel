# Check if web-ui files are staged
git diff --cached --quiet -- X21/web-ui/ 2>&1 | Out-Null
if ($LASTEXITCODE -eq 0) {
    exit 0
}

Write-Host "🔧 Building web-ui..."
Push-Location X21/web-ui
if (-not $?) { exit 1 }

# Check if node_modules exists or package files changed
$nodeModulesExists = Test-Path "node_modules"
git diff --cached --quiet -- package.json package-lock.json 2>&1 | Out-Null
$packageChanged = $LASTEXITCODE -ne 0

if (-not $nodeModulesExists -or $packageChanged) {
    Write-Host "📦 Installing dependencies..."
    npm ci --silent
    if ($LASTEXITCODE -ne 0) { Pop-Location; exit 1 }
}

npm run build:output ../web-ui/bin/TaskPane/WebAssets
if ($LASTEXITCODE -ne 0) { Pop-Location; exit 1 }

Pop-Location
exit 0
