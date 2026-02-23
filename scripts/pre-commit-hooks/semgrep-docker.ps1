param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Files
)

$ErrorActionPreference = "Stop"

$repoPath = (Get-Location).Path
$repoPathNormalized = $repoPath -replace '\\', '/'

$normalizedFiles = @()
if ($Files) {
    foreach ($file in $Files) {
        $fullPath = if ([System.IO.Path]::IsPathRooted($file)) {
            (Resolve-Path $file).Path
        } else {
            (Resolve-Path (Join-Path $repoPath $file)).Path
        }

        $fullPathNormalized = $fullPath -replace '\\', '/'
        if ($fullPathNormalized.StartsWith("$repoPathNormalized/")) {
            $normalizedFiles += $fullPathNormalized.Substring($repoPathNormalized.Length + 1)
        } else {
            $normalizedFiles += $fullPathNormalized
        }
    }
}

$dockerArgs = @(
    "run",
    "--rm",
    "-v", "${repoPathNormalized}:/src",
    "-w", "/src",
    "returntocorp/semgrep",
    "semgrep",
    "--config", "p/ci",
    "--config", "p/r2c-security-audit",
    "--config", "p/owasp-top-ten",
    "--config", "p/typescript",
    "--config", "p/javascript",
    "--config", "p/react",
    "--config", "p/csharp",
    "--config", "p/python",
    "--config", "p/github-actions",
    "--error"
)

if ($normalizedFiles.Count -gt 0) {
    $dockerArgs += $normalizedFiles
}

& docker @dockerArgs
exit $LASTEXITCODE
