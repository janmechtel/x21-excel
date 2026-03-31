# X21 ClickOnce Publishing Script (Refactored)
# This script orchestrates the publishing process using modular PowerShell components
# Works in both local and CI/CD (GitHub Actions) contexts

param(
    [Parameter(Mandatory=$false)]
    [ValidateSet("Dev", "Internal", "Staging", "Production", "ProductionLocal", "NetworkShare")]
    [string]$Environment = "Dev",

    [Parameter(Mandatory=$false)]
    [switch]$VerboseOutput
)

# Normalize environment casing to keep add-in identity stable
function Normalize-Environment {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Env
    )

    $envValue = $Env.Trim()
    switch -Regex ($envValue) {
        '^dev$' { return "Dev" }
        '^internal$' { return "Internal" }
        '^staging$' { return "Staging" }
        '^productionlocal$' { return "ProductionLocal" }
        '^production$' { return "Production" }
        default { return $envValue }
    }
}

$Environment = Normalize-Environment -Env $Environment

# Set error action preference
$ErrorActionPreference = "Stop"

# Set console encoding to UTF-8 for proper display of Unicode characters
# Store original encoding to restore later
$Global:OriginalOutputEncoding = [Console]::OutputEncoding
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Import publishing modules (order matters - PublishCore must be first)
$modulePath = Join-Path $PSScriptRoot "scripts\publish"
Import-Module (Join-Path $modulePath "PublishCore.psm1") -Force -Global
Import-Module (Join-Path $modulePath "BuildOperations.psm1") -Force -Global
Import-Module (Join-Path $modulePath "DeployOperations.psm1") -Force -Global
Import-Module (Join-Path $modulePath "CertificateOps.psm1") -Force -Global
Import-Module (Join-Path $modulePath "GitOperations.psm1") -Force -Global

# Main execution
try {
    # Initialize log file
    $logFile = "publish\publish.log"
    $logDir = Split-Path $logFile -Parent
    if (-not (Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
    $null = New-Item -ItemType File -Path $logFile -Force

    Write-PublishLog "Starting X21 ClickOnce publishing process..." "Info"
    Write-PublishLog "Execution context: $(Test-ExecutionContext)" "Info"

    # Get environment configuration
    $envConfig = Get-EnvironmentConfig -Environment $Environment

    if (-not $envConfig) {
        throw "Invalid environment: $Environment"
    }

    Write-PublishLog "Environment: $Environment" "Info"
    if ($envConfig.ManifestCertificateThumbprint) {
        Write-PublishLog "Config ManifestCertificateThumbprint (from env config): $($envConfig.ManifestCertificateThumbprint)" "Info"
    }

    $Configuration = $envConfig.Configuration
    Write-PublishLog "Configuration: $Configuration" "Info"

    # Validate Git state
    $requiredBranch = if ($envConfig.ContainsKey("RequiredBranch")) { $envConfig.RequiredBranch } else { $null }
    Test-GitState -RequiredBranch $requiredBranch

    # Create publish directory structure early
    $publishBaseDir = "publish"
    if (-not (Test-Path $publishBaseDir)) {
        New-Item -ItemType Directory -Path $publishBaseDir -Force | Out-Null
        Write-PublishLog "Created publish base directory: $publishBaseDir" "Info"
    }

    # Validate prerequisites
    $msbuildPath = Test-Prerequisites -Environment $Environment

    # Confirm installation for auto-update testing (for stable releases)
    if ($Environment -in @("Staging", "Internal", "Production")) {
        Confirm-InstallationForTesting -Environment $Environment
    }

    # Convert relative publish URLs to absolute paths
    $scriptRoot = Get-Location
    Write-PublishLog "PublishDir (original): $($envConfig.PublishDir)" "Info"

    # Handle both relative and absolute paths properly
    if ([System.IO.Path]::IsPathRooted($envConfig.PublishDir)) {
        # Already an absolute path
        $envConfig.PublishDir = [System.IO.Path]::GetFullPath($envConfig.PublishDir).TrimEnd('\')
    } else {
        # Relative path - combine with script root
        $envConfig.PublishDir = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($scriptRoot, $envConfig.PublishDir)).TrimEnd('\')
    }

    Write-PublishLog "PublishDir (resolved): $($envConfig.PublishDir)" "Info"

    # Clear publish directory if ClearBeforeBuild is enabled
    if ($envConfig.ClearBeforeBuild) {
        Clear-PublishDirectory -PublishDir $envConfig.PublishDir
    }

    # Get current version
    $currentVersion = Get-CurrentVersion
    $newVersion = Update-Version -CurrentVersion $currentVersion -AutoIncrementRevision $envConfig.AutoIncrementRevision

    # Install frontend dependencies before building
    Install-FrontendDependencies

    # Publish ClickOnce application
    $publishProperties = @{
        "PublishDir" = $envConfig.PublishDir
        "PublishUrl" = $envConfig.PublishUrl
        "InstallUrl" = $envConfig.InstallUrl
        "UpdateUrl" = $envConfig.UpdateUrl
        "SupportUrl" = $envConfig.SupportUrl
        "ApplicationVersion" = $newVersion
        "DeployEnvironment" = $Environment
    }

    # Add certificate thumbprint to publish properties if available
    if ($envConfig.ManifestCertificateThumbprint) {
        $publishProperties["ManifestCertificateThumbprint"] = $envConfig.ManifestCertificateThumbprint
        # Avoid csproj-level ManifestKeyFile interfering with thumbprint-based signing.
        # When MSBuild sees ManifestKeyFile it may try to use it, leading to signing errors.
        $publishProperties["ManifestKeyFile"] = ""
        Write-PublishLog "Including certificate thumbprint in ClickOnce publish: $($envConfig.ManifestCertificateThumbprint)" "Info"

        # Proactively verify the cert exists locally so MSBuild signing doesn't fail late with MSB3482.
        # Note: MSBuild/ClickOnce may use either CurrentUser\My or LocalMachine\My depending on how the cert was installed.
        Write-PublishLog "Pre-flight: verifying ClickOnce signing certificate is available before MSBuild publish..." "Info"
        try {
            $cert = Get-SigningCertificate -Thumbprint $envConfig.ManifestCertificateThumbprint
            Write-PublishLog "Pre-flight: signing cert OK. Subject='$($cert.Subject)'; Store=CurrentUser\\My; Thumbprint=$($cert.Thumbprint)" "Success"
        } catch {
            # Fallback: also check LocalMachine\My to provide better guidance.
            $tp = ($envConfig.ManifestCertificateThumbprint -replace "\s", "").ToUpperInvariant()
            $lmPath = "Cert:\\LocalMachine\\My\\$tp"
            if (Test-Path $lmPath) {
                Write-PublishLog "Certificate is installed in LocalMachine\\My but not in CurrentUser\\My." "Error"
                Write-PublishLog "MSBuild signing in your context likely expects CurrentUser\\My." "Error"
                Write-PublishLog "Fix: re-import the PFX for the CURRENT USER (Personal store), or export+import from LocalMachine to CurrentUser." "Info"
            }

            Write-PublishLog "Pre-flight certificate check FAILED. Aborting publish before MSBuild." "Error"
            exit 1
        }
    }

    Publish-ClickOnce -MSBuildPath $msbuildPath -Configuration $Configuration -Properties $publishProperties -VerboseOutput:$VerboseOutput

    # Rename setup.exe to versioned name after publishing but before upload
    Rename-SetupExecutable -PublishDir $envConfig.PublishDir -Version $newVersion

    # existing users have this generic name for the .vsto manifest I'm unusure if we can update them from one path to another. doesn't matter much for now if we use this to also update the X21.vsto in the folder
    # Copy environment-specific .vsto to generic X21.vsto
    Copy-GenericVstoFile -PublishDir $envConfig.PublishDir -Environment $Environment

    # Invoke Git operations after successful publishing but before deployment
    if (-not $envConfig.SkipGitOperations) {
        Invoke-GitOperations -Version $newVersion -Environment $Environment
    } else {
        Write-PublishLog "Skipping Git operations as configured for $Environment environment" "Info"
    }

    # Create zip file if configured (regardless of upload setting)
    $zipFilePath = $null
    if ($envConfig.CreateZip) {
        Write-PublishLog "Creating zip file..." "Info"

        # Create zip file in publish directory
        $tempZipDir = Join-Path $scriptRoot "publish"
        $zipFilePath = New-PublishZip -SourcePath $envConfig.PublishDir -Version $newVersion -OutputDirectory $tempZipDir
    }

    # Upload using rclone if not skipped
    if (-not $envConfig.SkipUpload) {
        if ($envConfig.CreateZip -and $zipFilePath) {
            # Upload the zip file to R2
            Write-PublishLog "Uploading zip file to R2 storage..." "Info"
            Upload-FileWithRclone -FilePath $zipFilePath -ConfigName "x21" -BucketPath $envConfig.RcloneBucket -VerboseOutput:$VerboseOutput
            Write-PublishLog "Zip file upload completed successfully." "Success"
        } else {
            # Standard deployment for other environments
            Deploy-WithRclone -SourcePath $envConfig.PublishDir -ConfigName "x21" -BucketName $envConfig.RcloneBucket -VerboseOutput:$VerboseOutput
        }
    } else {
        Write-PublishLog "Skipping upload as configured for $Environment environment" "Info"
    }



    Write-PublishLog "Publishing process completed successfully!" "Success"
    Write-PublishLog "Version: $newVersion" "Info"
    Write-PublishLog "Published to: $($envConfig.PublishUrl)" "Info"
    if ($envConfig.InstallUrl) {
        Write-PublishLog "Install URL: $($envConfig.InstallUrl)" "Info"
    }
    if ($envConfig.RcloneBucket) {
        Write-PublishLog "Deployed to bucket: $($envConfig.RcloneBucket)" "Info"
    }

    # Display summary
    Write-Host ""
    Write-Host "=== PUBLISHING SUMMARY ===" -ForegroundColor Cyan
    Write-Host "Version: $newVersion" -ForegroundColor White
    Write-Host "Environment: $Environment" -ForegroundColor White
    Write-Host "Configuration: $Configuration" -ForegroundColor White
    Write-Host "Publish Directory: $($envConfig.PublishDir)" -ForegroundColor White
    Write-Host "Deploy URL: $($envConfig.PublishUrl)" -ForegroundColor White
    if ($envConfig.InstallUrl) {
        Write-Host "Install URL: $($envConfig.InstallUrl)" -ForegroundColor White
    }

    # Construct and display download URL
    if ($envConfig.CreateZip -and $zipFilePath) {
        # If we didn't upload (e.g. ProductionLocal), show local file path.
        if ($envConfig.SkipUpload) {
            $downloadUrl = $zipFilePath
        } else {
            $zipFileName = Split-Path $zipFilePath -Leaf
            $downloadUrl = "https://dl.kontext21.com/local/$zipFileName"
        }
    } else {
        # For other environments, show setup.exe
        $downloadUrl = "$($envConfig.PublishUrl.TrimEnd('/'))/setup.exe"
    }
    Write-Host ""
    Write-Host "Download URL: $downloadUrl" -ForegroundColor Green
    Write-Host ""

    if ($envConfig.ManifestCertificateThumbprint) {
        Write-Host "Certificate Thumbprint: $($envConfig.ManifestCertificateThumbprint)" -ForegroundColor White
    }
    if ($envConfig.RcloneBucket) {
        Write-Host "Rclone Bucket: $($envConfig.RcloneBucket)" -ForegroundColor White
    }
    Write-Host "Log File: $logFile" -ForegroundColor White
    Write-Host "=========================" -ForegroundColor Cyan
}
catch {
    Write-PublishLog "Publishing failed: $($_.Exception.Message)" "Error"
    Write-PublishLog "Stack trace: $($_.ScriptStackTrace)" "Error"
    exit 1
}
finally {
    Write-PublishLog "Publishing process finished." "Info"

    # Restore original console encoding
    if ($Global:OriginalOutputEncoding) {
        [Console]::OutputEncoding = $Global:OriginalOutputEncoding
    }
}
