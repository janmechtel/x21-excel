param(
    [string]$Configuration = "Debug"
)

$dllPath = (Get-Location).Path + "\bin\$Configuration\X21.dll"

# Function to get ApplicationVersion from csproj file
function Get-ApplicationVersionFromCsproj {
    $csprojPath = (Get-Location).Path + "\X21.csproj"
    if (Test-Path $csprojPath) {
        $content = Get-Content $csprojPath -Raw
        if ($content -match '<ApplicationVersion>([^<]+)</ApplicationVersion>') {
            return $matches[1]
        }
    }
    return "1.1.0.0"  # Ultimate fallback
}

# Check if the DLL exists
if (-not (Test-Path $dllPath)) {
    $fallbackVersion = Get-ApplicationVersionFromCsproj
    Write-Output $fallbackVersion
    exit 0
}

try {
    $assembly = [System.Reflection.Assembly]::LoadFile($dllPath)
    $version = $assembly.GetName().Version
    Write-Output $version.ToString()
} catch {
    $fallbackVersion = Get-ApplicationVersionFromCsproj
    Write-Output $fallbackVersion
}
