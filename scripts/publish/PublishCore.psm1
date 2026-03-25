# PublishCore.psm1
# Core publishing functions shared between local and CI/CD contexts

# Script configuration constants
$script:ScriptConfig = @{
    ProjectName = "X21"
    ProjectPath = "X21\vsto-addin\X21.csproj"
    SolutionPath = "X21.sln"
    LogFile = "publish\publish.log"
    RcloneConfigName = "x21"
}

<#
.SYNOPSIS
    Detects the execution context (local vs CI/CD)

.DESCRIPTION
    Determines whether the script is running locally or in a GitHub Actions environment

.OUTPUTS
    String - "Local" or "GitHubActions"
#>
function Test-ExecutionContext {
    if ($env:GITHUB_ACTIONS -eq "true") {
        return "GitHubActions"
    } else {
        return "Local"
    }
}

<#
.SYNOPSIS
    Gets environment configuration

.DESCRIPTION
    Loads environment configuration from JSON file or legacy hashtable

.PARAMETER Environment
    The environment name (Dev, Internal, Staging, Production, ProductionLocal)

.OUTPUTS
    Hashtable containing environment configuration
#>
function Get-EnvironmentConfig {
    param(
        [Parameter(Mandatory=$true)]
        [ValidateSet("Dev", "Internal", "Staging", "Production", "ProductionLocal")]
        [string]$Environment
    )

    # Single source of truth: JSON config
    $configPath = Join-Path $PSScriptRoot "..\config\environments.json"

    if (-not (Test-Path $configPath)) {
        throw "Environment config not found at: $configPath"
    }

    Write-PublishLog "Loading environment configuration from JSON: $configPath" "Info"
    $jsonContent = Get-Content $configPath -Raw | ConvertFrom-Json

    if (-not ($jsonContent.PSObject.Properties.Name -contains $Environment)) {
        $available = ($jsonContent.PSObject.Properties.Name | Sort-Object) -join ", "
        throw "Environment '$Environment' not found in $configPath. Available: $available"
    }

    # Convert PSCustomObject to Hashtable for consistency
    $config = @{}
    $jsonContent.$Environment.PSObject.Properties | ForEach-Object {
        $config[$_.Name] = $_.Value
    }

    Write-PublishLog "Loaded configuration for $Environment from JSON" "Info"
    return $config
}

<#
.SYNOPSIS
    Writes a log message

.DESCRIPTION
    Centralized logging function that writes to both console and file

.PARAMETER Message
    The message to log

.PARAMETER Level
    The log level (Info, Warning, Error, Success)
#>
function Write-PublishLog {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,

        [Parameter(Mandatory=$false)]
        [ValidateSet("Info", "Warning", "Error", "Success")]
        [string]$Level = "Info"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"

    switch ($Level) {
        "Info" { Write-Host $logMessage -ForegroundColor White }
        "Warning" { Write-Host $logMessage -ForegroundColor Yellow }
        "Error" { Write-Host $logMessage -ForegroundColor Red }
        "Success" { Write-Host $logMessage -ForegroundColor Green }
    }

    # Ensure log directory exists
    $logDir = Split-Path $script:ScriptConfig.LogFile -Parent
    if (-not (Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }

    Add-Content -Path $script:ScriptConfig.LogFile -Value $logMessage
}

<#
.SYNOPSIS
    Validates prerequisites for publishing

.DESCRIPTION
    Checks for required tools and files

.PARAMETER Environment
    The environment being published to

.OUTPUTS
    String - Path to MSBuild.exe
#>
function Test-Prerequisites {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Environment
    )

    Write-PublishLog "Checking prerequisites..." "Info"

    # Check if we're in the right directory
    if (-not (Test-Path $script:ScriptConfig.ProjectPath)) {
        throw "Project file not found. Please run this script from the solution root directory."
    }

    # Check for MSBuild
    $msbuildPath = $null
    $possiblePaths = @(
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles}\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles}\Microsoft Visual Studio\2019\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles}\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe"
    )

    foreach ($path in $possiblePaths) {
        if (Test-Path $path) {
            $msbuildPath = $path
            break
        }
    }

    if (-not $msbuildPath) {
        throw "MSBuild not found. Please ensure Visual Studio is installed."
    }

    Write-PublishLog "Found MSBuild at: $msbuildPath" "Info"

    # Check if rclone is available (only if not skipping upload)
    $envConfig = Get-EnvironmentConfig -Environment $Environment
    if (-not $envConfig.SkipUpload) {
        try {
            $rcloneVersion = rclone version 2>&1 | Select-Object -First 1
            Write-PublishLog "Using rclone: $rcloneVersion" "Info"
        } catch {
            throw "rclone is not installed or not available in PATH. Please install rclone first."
        }
    }

    Write-PublishLog "Prerequisites check completed." "Success"
    return $msbuildPath
}

<#
.SYNOPSIS
    Gets the current version from the project file

.DESCRIPTION
    Extracts ApplicationVersion from the .csproj file

.OUTPUTS
    String - Current version (e.g., "1.6.0.6")
#>
function Get-CurrentVersion {
    Write-PublishLog "Reading current version from project file..." "Info"

    $projectContent = Get-Content $script:ScriptConfig.ProjectPath -Raw
    $versionMatch = [regex]::Match($projectContent, '<ApplicationVersion>([^<]+)</ApplicationVersion>')

    if ($versionMatch.Success) {
        $currentVersion = $versionMatch.Groups[1].Value
        Write-PublishLog "Current version: $currentVersion" "Info"
        return $currentVersion
    } else {
        throw "Could not find ApplicationVersion in project file."
    }
}

<#
.SYNOPSIS
    Updates the version in the project file

.DESCRIPTION
    Increments the version revision if AutoIncrementRevision is enabled

.PARAMETER CurrentVersion
    The current version string

.PARAMETER AutoIncrementRevision
    Whether to auto-increment the revision number

.OUTPUTS
    String - New version (e.g., "1.6.0.7")
#>
function Update-Version {
    param(
        [Parameter(Mandatory=$true)]
        [string]$CurrentVersion,

        [Parameter(Mandatory=$false)]
        [bool]$AutoIncrementRevision = $true
    )

    Write-PublishLog "Updating version..." "Info"

    $versionParts = $CurrentVersion.Split('.')
    if ($versionParts.Length -ne 4) {
        throw "Invalid version format. Expected format: Major.Minor.Build.Revision"
    }

    $major = [int]$versionParts[0]
    $minor = [int]$versionParts[1]
    $build = [int]$versionParts[2]
    $revision = [int]$versionParts[3]

    # Auto-increment revision if enabled
    if ($AutoIncrementRevision) {
        $revision++
        Write-PublishLog "Auto-incrementing revision to: $revision" "Info"
    }

    $newVersion = "$major.$minor.$build.$revision"

    # Update project file
    $projectContent = Get-Content $script:ScriptConfig.ProjectPath -Raw
    $projectContent = $projectContent -replace '<ApplicationVersion>[^<]+</ApplicationVersion>', "<ApplicationVersion>$newVersion</ApplicationVersion>"
    # Trim trailing whitespace/newlines and write without adding extra newline
    $projectContent = $projectContent.TrimEnd()
    Set-Content $script:ScriptConfig.ProjectPath $projectContent -Encoding UTF8 -NoNewline

    Write-PublishLog "Version updated to: $newVersion" "Success"
    return $newVersion
}

<#
.SYNOPSIS
    Confirms installation for testing

.DESCRIPTION
    Prompts user to confirm they have installed the current version for auto-update testing

.PARAMETER Environment
    The environment being published to
#>
function Confirm-InstallationForTesting {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Environment
    )

    Write-Host ""
    Write-Host "=== IMPORTANT ===" -ForegroundColor Yellow
    Write-Host "Have you installed the current version so that afterwards you can test the auto-update?" -ForegroundColor Yellow
    Write-Host "Please type 'yes' to confirm: " -ForegroundColor Yellow -NoNewline
    $response = Read-Host

    if ($response -ne "yes") {
        throw "Publishing cancelled. Please install the current version first to enable auto-update testing."
    }

    Write-PublishLog "Installation confirmed by user" "Success"
}

# Export functions
Export-ModuleMember -Function @(
    'Test-ExecutionContext',
    'Get-EnvironmentConfig',
    'Write-PublishLog',
    'Test-Prerequisites',
    'Get-CurrentVersion',
    'Update-Version',
    'Confirm-InstallationForTesting'
)
