# X21 PR Testing Script
# This script helps you test multiple pull request branches together locally
# Creates a temporary branch, merges PRs, and provides cleanup instructions

param(
    [Parameter(Mandatory=$false)]
    [string]$BaseBranch = "dev",

    [Parameter(Mandatory=$false)]
    [switch]$VerboseOutput
)

# Set error action preference
$ErrorActionPreference = "Stop"

# Script configuration
$ScriptConfig = @{
    TestBranchPrefix = "test-prs-"
    LogFile = "test-prs.log"
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

    # Log to file if needed
    if ($VerboseOutput) {
        Add-Content -Path $ScriptConfig.LogFile -Value $logMessage -ErrorAction SilentlyContinue
    }
}

# Function to validate Git repository state
function Test-GitState {
    Write-Log "Validating Git repository state..." "Info"

    # Check if we're in a Git repository
    try {
        $gitStatus = git status --porcelain 2>&1
        if ($LASTEXITCODE -ne 0) {
            throw "Not in a Git repository."
        }
    } catch {
        throw "Git is not available or not in a Git repository."
    }

    # Check for uncommitted changes
    $uncommittedChanges = git status --porcelain
    if ($uncommittedChanges) {
        Write-Log "You have uncommitted changes:" "Warning"
        Write-Log $uncommittedChanges "Warning"

        $response = Read-Host "Continue anyway? Uncommitted changes will be stashed. (y/N)"
        if ($response -ne "y" -and $response -ne "Y") {
            throw "Operation cancelled. Please commit or stash your changes first."
        }

        # Stash changes
        Write-Log "Stashing uncommitted changes..." "Info"
        git stash push -m "Auto-stash before PR testing"
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to stash changes."
        }
        Write-Log "Changes stashed successfully." "Success"
        return $true
    }

    Write-Log "Git repository state is clean." "Success"
    return $false
}

# Function to get available pull requests
function Get-AvailablePullRequests {
    Write-Log "Fetching open pull requests..." "Info"

    # Check if gh CLI is installed
    try {
        $ghVersion = gh --version 2>&1 | Select-Object -First 1
        Write-Log "Using GitHub CLI: $ghVersion" "Info"
    } catch {
        throw "GitHub CLI (gh) is not installed. Please install it from: https://cli.github.com/"
    }

    # Fetch all branches first
    git fetch --all --prune 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to fetch branches from remote."
    }

    # Get all open pull requests
    $prListJson = gh pr list --json number,title,headRefName --limit 100 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to fetch pull requests. Error: $prListJson"
    }

    $pullRequests = $prListJson | ConvertFrom-Json

    if (-not $pullRequests -or $pullRequests.Count -eq 0) {
        throw "No open pull requests found."
    }

    Write-Log "Found $($pullRequests.Count) open pull requests." "Success"
    return $pullRequests
}

# Function to select pull requests
function Select-PullRequests {
    param(
        [array]$PullRequests
    )

    Write-Host ""
    Write-Host "=== Open Pull Requests ===" -ForegroundColor Cyan
    for ($i = 0; $i -lt $PullRequests.Count; $i++) {
        $pr = $PullRequests[$i]
        $truncatedTitle = if ($pr.title.Length -gt 60) {
            $pr.title.Substring(0, 57) + "..."
        } else {
            $pr.title
        }
        Write-Host "$($i + 1). PR #$($pr.number): $truncatedTitle" -ForegroundColor White
        Write-Host "    Branch: $($pr.headRefName)" -ForegroundColor Gray
    }
    Write-Host "==========================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Enter PR numbers to test (comma-separated, e.g., 1,3,5):" -ForegroundColor Yellow
    Write-Host "Or enter 'a' to test all PRs" -ForegroundColor Yellow

    $selection = Read-Host "Selection"

    if ($selection -eq "a" -or $selection -eq "A") {
        Write-Log "Selected all PRs for testing." "Info"
        $selectedBranches = $PullRequests | ForEach-Object { $_.headRefName }
        return $selectedBranches
    }

    # Parse selection
    $selectedIndices = $selection -split ',' | ForEach-Object {
        $_.Trim() -as [int]
    } | Where-Object {
        $_ -ge 1 -and $_ -le $PullRequests.Count
    }

    if (-not $selectedIndices -or $selectedIndices.Count -eq 0) {
        throw "No valid PRs selected."
    }

    $selectedBranches = $selectedIndices | ForEach-Object {
        $PullRequests[$_ - 1].headRefName
    }

    Write-Log "Selected $($selectedBranches.Count) PRs for testing:" "Success"
    foreach ($i in $selectedIndices) {
        $pr = $PullRequests[$i - 1]
        Write-Log "  - PR #$($pr.number): $($pr.headRefName)" "Info"
    }

    return $selectedBranches
}

# Function to create test branch
function New-TestBranch {
    param(
        [string]$BaseBranch
    )

    # Generate test branch name with timestamp
    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $testBranch = "$($ScriptConfig.TestBranchPrefix)$timestamp"

    Write-Log "Creating test branch '$testBranch' from '$BaseBranch'..." "Info"

    # Checkout base branch and pull latest
    git checkout $BaseBranch 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to checkout base branch '$BaseBranch'."
    }

    git pull origin $BaseBranch 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Write-Log "Warning: Could not pull latest changes from '$BaseBranch'." "Warning"
    }

    # Create test branch
    git checkout -b $testBranch 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to create test branch '$testBranch'."
    }

    Write-Log "Test branch '$testBranch' created successfully." "Success"
    return $testBranch
}

# Function to merge branches
function Merge-Branches {
    param(
        [array]$Branches,
        [string]$TestBranch
    )

    Write-Log "Merging selected branches into '$TestBranch'..." "Info"

    $mergedBranches = @()
    $failedBranches = @()

    foreach ($branch in $Branches) {
        Write-Log "Merging 'origin/$branch'..." "Info"

        $mergeOutput = git merge "origin/$branch" --no-edit 2>&1

        if ($LASTEXITCODE -eq 0) {
            Write-Log "Successfully merged '$branch'." "Success"
            $mergedBranches += $branch
        } else {
            Write-Log "Failed to merge '$branch'. Merge conflicts may exist." "Error"
            Write-Log "Merge output: $mergeOutput" "Error"

            # Check if there are conflicts
            $conflictFiles = git diff --name-only --diff-filter=U
            if ($conflictFiles) {
                Write-Log "Conflicting files:" "Error"
                foreach ($file in $conflictFiles) {
                    Write-Log "  - $file" "Error"
                }
            }

            $failedBranches += $branch

            # Ask user what to do
            Write-Host ""
            Write-Host "What would you like to do?" -ForegroundColor Yellow
            Write-Host "  a - Abort merge and continue with next branch" -ForegroundColor White
            Write-Host "  s - Stop testing and exit" -ForegroundColor White
            Write-Host "  r - Resolve conflicts manually (script will pause)" -ForegroundColor White

            $response = Read-Host "Choice (a/s/r)"

            if ($response -eq "s" -or $response -eq "S") {
                git merge --abort 2>&1 | Out-Null
                throw "Testing stopped by user due to merge conflict in '$branch'."
            } elseif ($response -eq "r" -or $response -eq "R") {
                Write-Log "Please resolve conflicts and then run: git merge --continue" "Info"
                Write-Log "After resolving, press Enter to continue with next branch..." "Info"
                Read-Host

                # Check if merge was completed
                $mergeInProgress = Test-Path ".git/MERGE_HEAD"
                if ($mergeInProgress) {
                    Write-Log "Merge still in progress. Aborting current merge." "Warning"
                    git merge --abort 2>&1 | Out-Null
                    $failedBranches += $branch
                } else {
                    Write-Log "Merge completed." "Success"
                    $mergedBranches += $branch
                }
            } else {
                # Abort this merge and continue
                git merge --abort 2>&1 | Out-Null
                Write-Log "Aborted merge for '$branch'. Continuing with next branch." "Info"
            }
        }
    }

    return @{
        Merged = $mergedBranches
        Failed = $failedBranches
    }
}

# Main execution
try {
    Write-Host ""
    Write-Host "=== X21 PR Testing Tool ===" -ForegroundColor Cyan
    Write-Host "This tool helps you test multiple PRs together locally." -ForegroundColor White
    Write-Host "==============================" -ForegroundColor Cyan
    Write-Host ""

    # Validate Git state
    $hasStash = Test-GitState

    # Get available pull requests
    $pullRequests = Get-AvailablePullRequests

    # Let user select PRs
    $selectedBranches = Select-PullRequests -PullRequests $pullRequests

    # Confirm base branch
    Write-Host ""
    Write-Host "Base branch: $BaseBranch" -ForegroundColor White
    $changeBase = Read-Host "Change base branch? (y/N)"
    if ($changeBase -eq "y" -or $changeBase -eq "Y") {
        $BaseBranch = Read-Host "Enter base branch name"
    }

    # Create test branch
    $testBranch = New-TestBranch -BaseBranch $BaseBranch

    # Merge selected branches
    $result = Merge-Branches -Branches $selectedBranches -TestBranch $testBranch

    # Display summary
    Write-Host ""
    Write-Host "=== TEST BRANCH READY ===" -ForegroundColor Green
    Write-Host "Test branch: $testBranch" -ForegroundColor White
    Write-Host "Base branch: $BaseBranch" -ForegroundColor White
    Write-Host ""

    if ($result.Merged.Count -gt 0) {
        Write-Host "Successfully merged ($($result.Merged.Count)):" -ForegroundColor Green
        foreach ($branch in $result.Merged) {
            Write-Host "  ✓ $branch" -ForegroundColor Green
        }
    }

    if ($result.Failed.Count -gt 0) {
        Write-Host ""
        Write-Host "Failed to merge ($($result.Failed.Count)):" -ForegroundColor Red
        foreach ($branch in $result.Failed) {
            Write-Host "  ✗ $branch" -ForegroundColor Red
        }
    }

    Write-Host ""
    Write-Host "=== NEXT STEPS ===" -ForegroundColor Yellow
    Write-Host "1. Test your changes in this branch" -ForegroundColor White
    Write-Host "2. When done, cleanup with these commands:" -ForegroundColor White
    Write-Host ""
    Write-Host "   git checkout $BaseBranch" -ForegroundColor Cyan
    Write-Host "   git branch -D $testBranch" -ForegroundColor Cyan

    if ($hasStash) {
        Write-Host "   git stash pop" -ForegroundColor Cyan
    }

    Write-Host ""
    Write-Host "========================" -ForegroundColor Yellow
    Write-Host ""

} catch {
    Write-Log "PR testing failed: $($_.Exception.Message)" "Error"

    # Try to get back to base branch
    try {
        $currentBranch = git branch --show-current
        if ($currentBranch -match '^test-prs-') {
            Write-Log "Attempting to return to $BaseBranch..." "Info"
            git checkout $BaseBranch 2>&1 | Out-Null

            # Offer to delete the failed test branch
            $deleteTest = Read-Host "Delete failed test branch '$currentBranch'? (y/N)"
            if ($deleteTest -eq "y" -or $deleteTest -eq "Y") {
                git branch -D $currentBranch 2>&1 | Out-Null
                Write-Log "Deleted test branch '$currentBranch'." "Info"
            }
        }
    } catch {
        Write-Log "Could not automatically cleanup. You may need to manually switch branches." "Warning"
    }

    exit 1
} finally {
    Write-Log "PR testing process finished." "Info"
}
