# BuildOperations.psm1
# Build orchestration for React UI, Deno backend, and MSBuild ClickOnce
# Requires PublishCore.psm1 to be imported first for Write-PublishLog function

<#
.SYNOPSIS
    Publishes ClickOnce application using MSBuild

.DESCRIPTION
    Cleans and publishes the VSTO project with ClickOnce manifests

.PARAMETER MSBuildPath
    Path to MSBuild.exe

.PARAMETER Configuration
    Build configuration (Release/Debug)

.PARAMETER Properties
    Hashtable of MSBuild properties

.PARAMETER VerboseOutput
    Enable verbose MSBuild output
#>
function Publish-ClickOnce {
    param(
        [Parameter(Mandatory=$true)]
        [string]$MSBuildPath,

        [Parameter(Mandatory=$true)]
        [string]$Configuration,

        [Parameter(Mandatory=$false)]
        [hashtable]$Properties = @{},

        [Parameter(Mandatory=$false)]
        [switch]$VerboseOutput
    )

    Write-PublishLog "Publishing ClickOnce application..." "Info"

    $projectPath = "X21\vsto-addin\X21.csproj"

    # Clean the project first
    Write-PublishLog "Cleaning project before publish..." "Info"
    $cleanArgs = @(
        $projectPath,
        "/t:Clean",
        "/p:Configuration=$Configuration",
        "/p:Platform=AnyCPU"
    )

    if ($VerboseOutput) {
        $cleanArgs += "/verbosity:diagnostic"
    } else {
        $cleanArgs += "/verbosity:minimal"
    }

    Write-PublishLog "MSBuild clean command: $MSBuildPath $($cleanArgs -join ' ')" "Info"

    $cleanProcess = Start-Process -FilePath $MSBuildPath -ArgumentList $cleanArgs -Wait -PassThru -NoNewWindow

    if ($cleanProcess.ExitCode -ne 0) {
        throw "MSBuild clean failed with exit code: $($cleanProcess.ExitCode)"
    }

    Write-PublishLog "Project cleaned successfully." "Success"

    # Prepare MSBuild properties for publishing
    $publishProperties = @{
        "Platform" = "AnyCPU"
    }

    # Merge with provided properties
    foreach ($key in $Properties.Keys) {
        $publishProperties[$key] = $Properties[$key]
    }

    # Log certificate thumbprint if present
    if ($publishProperties.ContainsKey("ManifestCertificateThumbprint")) {
        Write-PublishLog "Using certificate thumbprint for ClickOnce signing: $($publishProperties['ManifestCertificateThumbprint'])" "Info"
    }

    # Prepare MSBuild arguments
    $propertyArgs = @()
    foreach ($key in $publishProperties.Keys) {
        $propertyArgs += "/p:$key=`"$($publishProperties[$key])`""
    }

    $publishArgs = @(
        $projectPath,
        "/p:Configuration=$Configuration",
        "/t:Publish"
    ) + $propertyArgs

    if ($VerboseOutput) {
        $publishArgs += "/verbosity:diagnostic"
    } else {
        $publishArgs += "/verbosity:minimal"
    }

    # Sort arguments for consistency
    $publishArgs = $publishArgs | Sort-Object
    Write-PublishLog "MSBuild publish command: $MSBuildPath $($publishArgs -join ' ')" "Info"

    # Output each argument on a numbered line for comparison
    Write-PublishLog "Publish arguments:" "Info"
    for ($i = 0; $i -lt $publishArgs.Count; $i++) {
        Write-PublishLog "  $($i + 1). $($publishArgs[$i])" "Info"
    }

    $process = Start-Process -FilePath $MSBuildPath -ArgumentList $publishArgs -Wait -PassThru -NoNewWindow

    if ($process.ExitCode -ne 0) {
        throw "MSBuild publish failed with exit code: $($process.ExitCode)"
    }

    Write-PublishLog "ClickOnce publishing completed successfully." "Success"
}

<#
.SYNOPSIS
    Renames setup.exe to versioned name

.DESCRIPTION
    Renames the generated setup.exe to include version number

.PARAMETER PublishDir
    The publish directory containing setup.exe

.PARAMETER Version
    The version string to include in the filename
#>
function Rename-SetupExecutable {
    param(
        [Parameter(Mandatory=$true)]
        [string]$PublishDir,

        [Parameter(Mandatory=$true)]
        [string]$Version
    )

    $setupExePath = Join-Path $PublishDir "setup.exe"
    if (Test-Path $setupExePath) {
        $versionedSetupName = "x21-setup-v$Version.exe"
        $versionedSetupPath = Join-Path $PublishDir $versionedSetupName

        # Remove existing versioned setup if it exists
        if (Test-Path $versionedSetupPath) {
            Remove-Item $versionedSetupPath -Force
            Write-PublishLog "Removed existing versioned setup: $versionedSetupPath" "Info"
        }

        # Rename setup.exe to versioned name
        Rename-Item $setupExePath $versionedSetupName -Force
        Write-PublishLog "Renamed setup.exe to versioned setup file: $versionedSetupName" "Success"
    } else {
        Write-PublishLog "Warning: setup.exe not found at $setupExePath" "Warning"
    }
}

<#
.SYNOPSIS
    Creates a generic copy of the environment-specific .vsto file

.DESCRIPTION
    Copies X21-{Environment}.vsto to X21.vsto for easier generic deployment

.PARAMETER PublishDir
    The publish directory containing the .vsto file

.PARAMETER Environment
    The environment name (e.g., Staging, Internal, Production)
#>
function Copy-GenericVstoFile {
    param(
        [Parameter(Mandatory=$true)]
        [string]$PublishDir,

        [Parameter(Mandatory=$true)]
        [string]$Environment
    )

    $envVstoPath = Join-Path $PublishDir "X21-$Environment.vsto"
    $genericVstoPath = Join-Path $PublishDir "X21.vsto"

    if (Test-Path $envVstoPath) {
        # Copy environment-specific .vsto to generic name
        Copy-Item $envVstoPath $genericVstoPath -Force
        Write-PublishLog "Copied $envVstoPath to X21.vsto" "Success"
    } else {
        Write-PublishLog "Warning: Environment-specific .vsto file not found at $envVstoPath" "Warning"
    }
}

<#
.SYNOPSIS
    Clears the publish directory

.DESCRIPTION
    Removes all files from the publish directory if ClearBeforeBuild is enabled

.PARAMETER PublishDir
    The directory to clear
#>
function Clear-PublishDirectory {
    param(
        [Parameter(Mandatory=$true)]
        [string]$PublishDir
    )

    if (Test-Path $PublishDir) {
        Write-PublishLog "Clearing publish directory before build: $PublishDir" "Info"
        try {
            Remove-Item -Path "$PublishDir\*" -Recurse -Force -ErrorAction Stop
            Write-PublishLog "Successfully cleared publish directory" "Success"
        } catch {
            Write-PublishLog "Warning: Could not fully clear directory. Some files may be in use." "Warning"
            Write-PublishLog "Error details: $($_.Exception.Message)" "Warning"
        }
    }
}

# Export functions
Export-ModuleMember -Function @(
    'Publish-ClickOnce',
    'Rename-SetupExecutable',
    'Copy-GenericVstoFile',
    'Clear-PublishDirectory'
)
