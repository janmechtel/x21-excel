# X21 Promotion Script
# This script promotes releases between environments
# Part 1 of the two-step release process for Staging and Production

param(
    [Parameter(Mandatory=$true)]
    [ValidateSet("Staging", "Production")]
    [string]$Environment,

    [Parameter(Mandatory=$false)]
    [switch]$VerboseOutput
)

# Set error action preference
$ErrorActionPreference = "Stop"

# Script configuration
$ScriptConfig = @{
    ProjectPath = "X21\vsto-addin\X21.csproj"
    LogFile = "publish\promote.log"
}

# Environment-specific configurations
$EnvironmentConfigs = @{
    Staging = @{
        TargetBranch = "staging"
        SourceTagSuffix = "-internal"
        SourceEnvironment = "Internal"
    }
    Production = @{
        TargetBranch = "production"
        SourceTagSuffix = "-staging"
        SourceEnvironment = "Staging"
    }
}

# Initialize logging
function Write-Log {
    param(
        [string]$Message,
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
    $logDir = Split-Path $ScriptConfig.LogFile -Parent
    if (-not (Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }

    Add-Content -Path $ScriptConfig.LogFile -Value $logMessage
}

# Function to validate Git repository state
function Test-GitState {
    Write-Log "Validating Git repository state..." "Info"

    # Check if we're in a Git repository
    try {
        $gitStatus = git status --porcelain
        if ($LASTEXITCODE -ne 0) {
            throw "Not in a Git repository. Please run this script from a Git repository."
        }
    } catch {
        throw "Git is not available or not in a Git repository. Please ensure Git is installed and you're in a Git repository."
    }

    # Check for uncommitted changes
    $uncommittedChanges = git status --porcelain
    if ($uncommittedChanges) {
        Write-Log "You have uncommitted changes:" "Warning"
        Write-Log $uncommittedChanges "Warning"

        $response = Read-Host "Continue anyway? Uncommitted changes will be stashed. (y/N)"
        if ($response -ne "y" -and $response -ne "Y") {
            throw "Promotion cancelled. Please commit or stash your changes before promoting."
        }

        # Stash changes
        Write-Log "Stashing uncommitted changes..." "Info"
        git stash push -m "Auto-stash before $Environment promotion"
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to stash changes."
        }
        Write-Log "Changes stashed successfully." "Success"
        return $true  # Indicate that we stashed changes
    }

    Write-Log "Git repository state validation completed." "Success"
    return $false  # No stash was created
}

# Function to get available tags for source environment
function Get-SourceTags {
    param(
        [string]$TagSuffix,
        [string]$SourceEnvironment
    )

    Write-Log "Fetching $SourceEnvironment release tags..." "Info"

    # Fetch latest tags from remote
    git fetch --tags
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to fetch tags from remote."
    }

    # Get all tags with the specified suffix
    $tags = git tag -l "*$TagSuffix" | Sort-Object -Descending

    if (-not $tags -or $tags.Count -eq 0) {
        throw "No $SourceEnvironment release tags found. Please publish to $SourceEnvironment environment first."
    }

    Write-Log "Found $($tags.Count) $SourceEnvironment release tags." "Success"
    return $tags
}

# Function to select a tag
function Select-Tag {
    param(
        [array]$Tags,
        [string]$SourceEnvironment,
        [string]$TargetEnvironment
    )

    Write-Host ""
    Write-Host "=== Available $SourceEnvironment Releases ===" -ForegroundColor Cyan
    for ($i = 0; $i -lt $Tags.Count; $i++) {
        Write-Host "$($i + 1). $($Tags[$i])" -ForegroundColor White
    }
    Write-Host "===================================" -ForegroundColor Cyan
    Write-Host ""

    do {
        $selection = Read-Host "Select a tag to promote to $TargetEnvironment (1-$($Tags.Count))"
        $selectionNum = $selection -as [int]
    } while ($selectionNum -lt 1 -or $selectionNum -gt $Tags.Count)

    $selectedTag = $Tags[$selectionNum - 1]
    Write-Log "Selected tag: $selectedTag" "Success"

    return $selectedTag
}

# Function to get version from tag
function Get-VersionFromTag {
    param(
        [string]$Tag,
        [string]$TagSuffix
    )

    # Remove the 'v' prefix and tag suffix to get version
    $version = $Tag -replace '^v', '' -replace "$TagSuffix$", ''
    return $version
}

# Function to get version from project file
function Get-ProjectVersion {
    Write-Log "Reading current version from project file..." "Info"

    $projectContent = Get-Content $ScriptConfig.ProjectPath -Raw
    $versionMatch = [regex]::Match($projectContent, '<ApplicationVersion>([^<]+)</ApplicationVersion>')

    if ($versionMatch.Success) {
        $currentVersion = $versionMatch.Groups[1].Value
        Write-Log "Current project version: $currentVersion" "Info"
        return $currentVersion
    } else {
        throw "Could not find ApplicationVersion in project file."
    }
}

# Function to update version in project file
function Update-ProjectVersion {
    param(
        [string]$NewVersion
    )

    Write-Log "Updating project version to: $NewVersion" "Info"

    $projectContent = Get-Content $ScriptConfig.ProjectPath -Raw
    $projectContent = $projectContent -replace '<ApplicationVersion>[^<]+</ApplicationVersion>', "<ApplicationVersion>$NewVersion</ApplicationVersion>"
    Set-Content $ScriptConfig.ProjectPath $projectContent -Encoding UTF8

    Write-Log "Project version updated successfully." "Success"
}

# Function to prompt for version adjustment
function Confirm-VersionAdjustment {
    param(
        [string]$CurrentVersion,
        [string]$SourceEnvironment
    )

    Write-Host ""
    Write-Host "=== Version Configuration ===" -ForegroundColor Cyan
    Write-Host "Version from $SourceEnvironment tag: $CurrentVersion" -ForegroundColor White
    Write-Host "Press Enter to keep this version, or type a new version (format: Major.Minor.Build.Revision)" -ForegroundColor Yellow
    Write-Host "=============================" -ForegroundColor Cyan
    Write-Host ""

    $newVersion = Read-Host "Version [$CurrentVersion]"

    # If user just pressed Enter, keep the current version
    if ([string]::IsNullOrWhiteSpace($newVersion)) {
        Write-Log "Keeping version: $CurrentVersion" "Info"
        return $CurrentVersion
    }

    # Validate version format
    if ($newVersion -notmatch '^\d+\.\d+\.\d+\.\d+$') {
        Write-Log "Invalid version format. Keeping original version: $CurrentVersion" "Warning"
        return $CurrentVersion
    }

    Write-Log "Version will be changed to: $newVersion" "Info"
    return $newVersion
}

# Function to promote tag to target branch
function Invoke-Promotion {
    param(
        [string]$Tag,
        [string]$Version,
        [string]$TargetBranch,
        [string]$Environment
    )

    Write-Log "Starting promotion of $Tag to $TargetBranch branch..." "Info"

    # Fetch latest changes
    Write-Log "Fetching latest changes from remote..." "Info"
    git fetch origin
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to fetch from remote."
    }

    # Check if target branch exists locally
    $localBranchExists = git branch --list $TargetBranch

    # Check if target branch exists on remote
    $remoteBranchExists = git branch -r --list "origin/$TargetBranch"

    if ($remoteBranchExists) {
        # Remote branch exists - check it out
        if ($localBranchExists) {
            Write-Log "Checking out existing local $TargetBranch branch..." "Info"
            git checkout $TargetBranch
        } else {
            Write-Log "Checking out $TargetBranch branch from remote..." "Info"
            git checkout -b $TargetBranch "origin/$TargetBranch"
        }
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to checkout $TargetBranch branch."
        }

        # Pull latest changes
        Write-Log "Pulling latest changes..." "Info"
        git pull origin $TargetBranch
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to pull latest changes."
        }
    } else {
        # Remote branch doesn't exist - create new branch
        Write-Log "Creating new $TargetBranch branch..." "Info"
        git checkout -b $TargetBranch
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to create $TargetBranch branch."
        }
    }

    # Merge the tag into target branch using theirs strategy for auto-conflict resolution
    Write-Log "Merging tag $Tag into $TargetBranch branch (using theirs strategy for conflicts)..." "Info"
    git merge $Tag -X theirs --no-ff -m "Promote $Tag to $Environment"

    if ($LASTEXITCODE -ne 0) {
        Write-Log "Merge failed." "Error"
        throw "Failed to merge tag into $TargetBranch branch."
    }

    Write-Log "Tag merged successfully." "Success"

    # Update version in project file if it changed
    $currentProjectVersion = Get-ProjectVersion
    if ($currentProjectVersion -ne $Version) {
        Write-Log "Updating project version from $currentProjectVersion to $Version..." "Info"
        Update-ProjectVersion -NewVersion $Version

        # Commit the version change
        Write-Log "Committing version update..." "Info"
        git add $ScriptConfig.ProjectPath
        git commit -m "Update version to $Version for $Environment"
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to commit version update."
        }
        Write-Log "Version update committed." "Success"
    }

    Write-Log "Promotion completed successfully." "Success"
}

# Main execution
try {
    # Initialize log file
    $null = New-Item -ItemType File -Path $ScriptConfig.LogFile -Force

    Write-Log "Starting X21 $Environment promotion process..." "Info"

    # Get environment configuration
    $envConfig = $EnvironmentConfigs[$Environment]

    if (-not $envConfig) {
        throw "Invalid environment: $Environment"
    }

    # Validate Git state
    $hasStash = Test-GitState

    # Get available source tags
    $tags = Get-SourceTags -TagSuffix $envConfig.SourceTagSuffix -SourceEnvironment $envConfig.SourceEnvironment

    # Let user select a tag
    $selectedTag = Select-Tag -Tags $tags -SourceEnvironment $envConfig.SourceEnvironment -TargetEnvironment $Environment

    # Get version from tag
    $tagVersion = Get-VersionFromTag -Tag $selectedTag -TagSuffix $envConfig.SourceTagSuffix

    # Ask if user wants to adjust version
    $finalVersion = Confirm-VersionAdjustment -CurrentVersion $tagVersion -SourceEnvironment $envConfig.SourceEnvironment

    # Confirm promotion
    Write-Host ""
    Write-Host "=== PROMOTION SUMMARY ===" -ForegroundColor Cyan
    Write-Host "Source: $($envConfig.SourceEnvironment)" -ForegroundColor White
    Write-Host "Target: $Environment" -ForegroundColor White
    Write-Host "Tag to promote: $selectedTag" -ForegroundColor White
    Write-Host "Target branch: $($envConfig.TargetBranch)" -ForegroundColor White
    Write-Host "Version: $finalVersion" -ForegroundColor White
    Write-Host "=========================" -ForegroundColor Cyan
    Write-Host ""
    $confirmation = Read-Host "Proceed with promotion? (y/N)"

    if ($confirmation -ne "y" -and $confirmation -ne "Y") {
        throw "Promotion cancelled by user."
    }

    # Perform promotion
    Invoke-Promotion -Tag $selectedTag -Version $finalVersion -TargetBranch $envConfig.TargetBranch -Environment $Environment

    # Offer to push to remote
    Write-Host ""
    $pushResponse = Read-Host "Push $($envConfig.TargetBranch) branch to remote? (y/N)"
    if ($pushResponse -eq "y" -or $pushResponse -eq "Y") {
        Write-Log "Pushing $($envConfig.TargetBranch) branch to remote..." "Info"
        git push origin $envConfig.TargetBranch
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to push to remote."
        }
        Write-Log "Successfully pushed to remote." "Success"
    } else {
        Write-Log "Skipping push to remote. Remember to push manually before publishing." "Warning"
    }

    # Restore stash if we created one
    if ($hasStash) {
        Write-Log "Restoring stashed changes..." "Info"
        git stash pop
        if ($LASTEXITCODE -ne 0) {
            Write-Log "Warning: Could not automatically restore stashed changes. Use 'git stash pop' manually." "Warning"
        } else {
            Write-Log "Stashed changes restored successfully." "Success"
        }
    }

    Write-Host ""
    Write-Host "=== PROMOTION COMPLETED ===" -ForegroundColor Green
    Write-Host "Tag $selectedTag has been promoted to the $($envConfig.TargetBranch) branch." -ForegroundColor White
    Write-Host "Version: $finalVersion" -ForegroundColor White
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Yellow
    Write-Host "  1. Test the changes on the $($envConfig.TargetBranch) branch" -ForegroundColor White
    Write-Host "  2. When ready, run: .\Publish.ps1 -Environment $Environment" -ForegroundColor White
    Write-Host "===========================" -ForegroundColor Green
    Write-Host ""

} catch {
    Write-Log "Promotion failed: $($_.Exception.Message)" "Error"
    Write-Log "Stack trace: $($_.ScriptStackTrace)" "Error"

    Write-Host ""
    Write-Host "Promotion failed. You may need to:" -ForegroundColor Red
    Write-Host "  - Resolve any merge conflicts" -ForegroundColor Yellow
    Write-Host "  - Run 'git merge --abort' to cancel the merge" -ForegroundColor Yellow
    Write-Host "  - Check the log file for details: $($ScriptConfig.LogFile)" -ForegroundColor Yellow
    Write-Host ""

    exit 1
} finally {
    Write-Log "Promotion process finished." "Info"
}
