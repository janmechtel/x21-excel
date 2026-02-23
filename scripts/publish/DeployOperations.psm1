# DeployOperations.psm1
# Deployment operations using rclone to Cloudflare R2
# Requires PublishCore.psm1 to be imported first for Write-PublishLog function

<#
.SYNOPSIS
    Deploys files using rclone

.DESCRIPTION
    Copies a directory to R2 bucket using rclone

.PARAMETER SourcePath
    The source directory to upload

.PARAMETER ConfigName
    The rclone config name (e.g., "x21")

.PARAMETER BucketName
    The bucket path (e.g., "x21/staging")

.PARAMETER DryRun
    If specified, performs a dry run without actually uploading

.PARAMETER VerboseOutput
    Enable verbose rclone output
#>
function Deploy-WithRclone {
    param(
        [Parameter(Mandatory=$true)]
        [string]$SourcePath,

        [Parameter(Mandatory=$true)]
        [string]$ConfigName,

        [Parameter(Mandatory=$true)]
        [string]$BucketName,

        [Parameter(Mandatory=$false)]
        [switch]$DryRun = $false,

        [Parameter(Mandatory=$false)]
        [switch]$VerboseOutput
    )

    Write-PublishLog "Deploying files using rclone..." "Info"
    Write-PublishLog "Source: $SourcePath" "Info"
    Write-PublishLog "Config: $ConfigName" "Info"
    Write-PublishLog "Bucket: $BucketName" "Info"
    Write-PublishLog "DryRun: $DryRun" "Info"

    if (-not (Test-Path $SourcePath)) {
        throw "Source path does not exist: $SourcePath"
    }

    # Build rclone command
    $rcloneArgs = @(
        "copy",
        $SourcePath,
        "$ConfigName`:$BucketName/"
    )

    # Add dry-run flag if specified
    if ($DryRun) {
        $rcloneArgs += "--dry-run"
        Write-PublishLog "DRY RUN MODE - No files will be actually uploaded" "Warning"
    }

    # Add verbose output if requested
    if ($VerboseOutput) {
        $rcloneArgs += "--verbose"
    }

    $rcloneCommand = "rclone $($rcloneArgs -join ' ')"
    Write-PublishLog "Executing: $rcloneCommand" "Info"

    # Execute rclone command
    $process = Start-Process -FilePath "rclone" -ArgumentList $rcloneArgs -Wait -PassThru -NoNewWindow

    if ($process.ExitCode -ne 0) {
        throw "rclone deployment failed with exit code: $($process.ExitCode)"
    }

    Write-PublishLog "rclone deployment completed successfully." "Success"
}

<#
.SYNOPSIS
    Uploads a single file using rclone

.DESCRIPTION
    Copies a single file to R2 bucket using rclone copyto

.PARAMETER FilePath
    The file to upload

.PARAMETER ConfigName
    The rclone config name (e.g., "x21")

.PARAMETER BucketPath
    The bucket path (e.g., "x21/staging")

.PARAMETER DryRun
    If specified, performs a dry run without actually uploading

.PARAMETER VerboseOutput
    Enable verbose rclone output
#>
function Upload-FileWithRclone {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath,

        [Parameter(Mandatory=$true)]
        [string]$ConfigName,

        [Parameter(Mandatory=$true)]
        [string]$BucketPath,

        [Parameter(Mandatory=$false)]
        [switch]$DryRun = $false,

        [Parameter(Mandatory=$false)]
        [switch]$VerboseOutput
    )

    Write-PublishLog "Uploading file using rclone..." "Info"
    Write-PublishLog "File: $FilePath" "Info"
    Write-PublishLog "Config: $ConfigName" "Info"
    Write-PublishLog "Bucket Path: $BucketPath" "Info"
    Write-PublishLog "DryRun: $DryRun" "Info"

    if (-not (Test-Path $FilePath)) {
        throw "File does not exist: $FilePath"
    }

    # Build rclone command to copy a single file
    $rcloneArgs = @(
        "copyto",
        $FilePath,
        "$ConfigName`:$BucketPath/$([System.IO.Path]::GetFileName($FilePath))",
        "--s3-no-check-bucket"
    )

    # Add dry-run flag if specified
    if ($DryRun) {
        $rcloneArgs += "--dry-run"
        Write-PublishLog "DRY RUN MODE - No files will be actually uploaded" "Warning"
    }

    # Add verbose output if requested
    if ($VerboseOutput) {
        $rcloneArgs += "--verbose"
    }

    $rcloneCommand = "rclone $($rcloneArgs -join ' ')"
    Write-PublishLog "Executing: $rcloneCommand" "Info"

    # Execute rclone command
    $process = Start-Process -FilePath "rclone" -ArgumentList $rcloneArgs -Wait -PassThru -NoNewWindow

    if ($process.ExitCode -ne 0) {
        throw "rclone upload failed with exit code: $($process.ExitCode)"
    }

    Write-PublishLog "rclone upload completed successfully." "Success"
}

<#
.SYNOPSIS
    Creates a zip file from publish directory

.DESCRIPTION
    Compresses the publish directory into a versioned zip file

.PARAMETER SourcePath
    The source directory to zip

.PARAMETER Version
    The version string to include in the filename

.PARAMETER OutputDirectory
    The directory where the zip file will be created

.OUTPUTS
    String - Path to the created zip file
#>
function New-PublishZip {
    param(
        [Parameter(Mandatory=$true)]
        [string]$SourcePath,

        [Parameter(Mandatory=$true)]
        [string]$Version,

        [Parameter(Mandatory=$true)]
        [string]$OutputDirectory
    )

    Write-PublishLog "Creating zip file from published files..." "Info"
    Write-PublishLog "Source: $SourcePath" "Info"
    Write-PublishLog "Version: $Version" "Info"
    Write-PublishLog "Output Directory: $OutputDirectory" "Info"

    if (-not (Test-Path $SourcePath)) {
        throw "Source path does not exist: $SourcePath"
    }

    # Ensure output directory exists
    if (-not (Test-Path $OutputDirectory)) {
        New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
        Write-PublishLog "Created output directory: $OutputDirectory" "Info"
    }

    # Create zip file name
    $zipFileName = "x21-setup-v$Version.zip"
    $zipFilePath = Join-Path $OutputDirectory $zipFileName

    # Remove existing zip file if it exists
    if (Test-Path $zipFilePath) {
        Remove-Item $zipFilePath -Force
        Write-PublishLog "Removed existing zip file: $zipFilePath" "Info"
    }

    # Create zip file using PowerShell's Compress-Archive
    # Zip the folder itself (not just its contents) to preserve folder structure
    Write-PublishLog "Compressing files to: $zipFilePath" "Info"
    $parentPath = Split-Path $SourcePath -Parent
    $folderName = Split-Path $SourcePath -Leaf

    # Zip from parent directory to include the folder name in the archive
    Push-Location $parentPath
    try {
        Compress-Archive -Path $folderName -DestinationPath $zipFilePath -CompressionLevel Optimal
    } finally {
        Pop-Location
    }

    if (-not (Test-Path $zipFilePath)) {
        throw "Failed to create zip file: $zipFilePath"
    }

    $zipSize = (Get-Item $zipFilePath).Length / 1MB
    Write-PublishLog "Zip file created successfully: $zipFilePath (Size: $([math]::Round($zipSize, 2)) MB)" "Success"

    return $zipFilePath
}

<#
.SYNOPSIS
    Tests rclone configuration

.DESCRIPTION
    Validates that rclone is available and configured

.PARAMETER ConfigName
    The rclone config name to test

.OUTPUTS
    Boolean - True if config is valid
#>
function Test-RcloneConfig {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ConfigName
    )

    try {
        # Check if rclone is available
        $rcloneVersion = rclone version 2>&1 | Select-Object -First 1
        Write-PublishLog "rclone found: $rcloneVersion" "Info"

        # List configured remotes
        $remotes = rclone listremotes 2>&1
        if ($remotes -match "$ConfigName`:") {
            Write-PublishLog "rclone config '$ConfigName' found" "Success"
            return $true
        } else {
            Write-PublishLog "rclone config '$ConfigName' not found" "Warning"
            return $false
        }
    } catch {
        Write-PublishLog "rclone test failed: $($_.Exception.Message)" "Error"
        return $false
    }
}

# Export functions
Export-ModuleMember -Function @(
    'Deploy-WithRclone',
    'Upload-FileWithRclone',
    'New-PublishZip',
    'Test-RcloneConfig'
)
