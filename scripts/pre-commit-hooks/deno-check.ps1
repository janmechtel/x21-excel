# Check if deno-server files are staged
git diff --cached --quiet -- X21/deno-server/ 2>&1 | Out-Null
if ($LASTEXITCODE -eq 0) {
    exit 0
}

Write-Host "🔧 Formatting and checking deno-server..."
Push-Location X21/deno-server
if (-not $?) { exit 1 }

deno fmt --check
if ($LASTEXITCODE -ne 0) {
    Write-Host "⚠️  Formatting issues found. Auto-formatting..."
    deno fmt
    if ($LASTEXITCODE -ne 0) { Pop-Location; exit 1 }
    Write-Host "✅ Files have been auto-formatted. Please review and stage the changes, then commit again."
    Pop-Location
    exit 1
}

deno lint
if ($LASTEXITCODE -ne 0) { Pop-Location; exit 1 }

Pop-Location
exit 0
