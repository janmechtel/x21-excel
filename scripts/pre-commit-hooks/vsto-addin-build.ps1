# Check if vsto-addin files are staged
git diff --cached --quiet -- X21/vsto-addin/ 2>&1 | Out-Null
if ($LASTEXITCODE -eq 0) {
    exit 0
}

Write-Host "🔧 Building vsto-addin..."

# Find MSBuild on Windows
$paths = @(
    "${env:ProgramFiles}\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe",
    "${env:ProgramFiles}\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
    "${env:ProgramFiles}\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe",
    "${env:ProgramFiles}\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe",
    "${env:ProgramFiles}\Microsoft Visual Studio\2019\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
    "${env:ProgramFiles}\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe"
)

$MSBUILD_PATH = $null
foreach ($path in $paths) {
    if (Test-Path $path) {
        $MSBUILD_PATH = $path
        break
    }
}

if ($MSBUILD_PATH) {
    Write-Host "Found MSBuild at: $MSBUILD_PATH"
    & $MSBUILD_PATH X21/vsto-addin/X21.csproj -p:Configuration=Debug -p:Platform=AnyCPU -verbosity:minimal
    if ($LASTEXITCODE -ne 0) { exit 1 }
    exit 0
} else {
    Write-Host "❌ MSBuild not found. Please ensure Visual Studio is installed."
    exit 1
}
