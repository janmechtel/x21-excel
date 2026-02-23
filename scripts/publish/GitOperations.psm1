# GitOperations.psm1
# Git operations for version control (commits, tags, pushes)
# Requires PublishCore.psm1 to be imported first for Write-PublishLog and Test-ExecutionContext functions

<#
.SYNOPSIS
    Validates Git repository state

.DESCRIPTION
    Checks Git status, current branch, and synchronization with remote

.PARAMETER RequiredBranch
    The branch required for this operation (or "*" for any branch)
#>
function Test-GitState {
    param(
        [Parameter(Mandatory=$false)]
        [string]$RequiredBranch = "*"
    )

    Write-PublishLog "Validating Git repository state..." "Info"

    # Check if we're in a Git repository
    try {
        $gitStatus = git status --porcelain 2>&1
        if ($LASTEXITCODE -ne 0) {
            throw "Not in a Git repository. Please run this script from a Git repository."
        }
    } catch {
        throw "Git is not available or not in a Git repository. Please ensure Git is installed and you're in a Git repository."
    }

    # Get current branch
    $currentBranch = git branch --show-current 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to get current branch."
    }

    Write-PublishLog "Current branch: $currentBranch" "Info"

    # Skip branch validation if RequiredBranch is not specified or is "*"
    if ([string]::IsNullOrEmpty($RequiredBranch) -or $RequiredBranch -eq "*") {
        Write-PublishLog "Branch validation skipped (no specific branch required)" "Info"
        return
    } else {
        Write-PublishLog "Required branch: $RequiredBranch" "Info"

        # Check if we're on the required branch
        if ($currentBranch -ne $RequiredBranch) {
            throw "You are on branch '$currentBranch' but this operation requires branch '$RequiredBranch'."
        }
    }

    # Check if branch is up to date with remote
    try {
        git fetch origin 2>&1 | Out-Null
        $localCommit = git rev-parse HEAD 2>&1
        $remoteCommit = git rev-parse "origin/$currentBranch" 2>&1

        if ($LASTEXITCODE -eq 0 -and $localCommit -ne $remoteCommit) {
            Write-PublishLog "WARNING: Your local branch is not up to date with remote. Please pull latest changes." "Warning"
        }
    } catch {
        Write-PublishLog "Could not verify remote branch status. Continuing..." "Warning"
    }

    Write-PublishLog "Git repository state validation completed." "Success"
}

<#
.SYNOPSIS
    Commits version changes

.DESCRIPTION
    Stages and commits changes (typically .csproj version updates)

.PARAMETER Version
    The version being published

.PARAMETER Environment
    The environment being published to

.PARAMETER Interactive
    If true, prompts user before committing (default for local)
#>
function Invoke-GitCommit {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Version,

        [Parameter(Mandatory=$true)]
        [string]$Environment,

        [Parameter(Mandatory=$false)]
        [bool]$Interactive = $true
    )

    Write-PublishLog "Handling Git commit operations..." "Info"

    $context = Test-ExecutionContext

    # Check if there are any changes to commit
    $gitStatus = git status --porcelain 2>&1
    if (-not $gitStatus) {
        Write-PublishLog "No changes detected in Git repository." "Info"
        return
    }

    Write-PublishLog "Changes detected in Git repository:" "Info"
    Write-PublishLog ($gitStatus -join "`n") "Info"


    # Check if only .csproj file has changes
    $changedFiles = $gitStatus -split "`n" | Where-Object { $_ -match "^\s*[AM]\s+(.+)$" } | ForEach-Object { $matches[1] }
    $nonCsprojFiles = $changedFiles | Where-Object { $_ -notmatch "\.csproj$" }

    if ($nonCsprojFiles -and $Interactive -and $context -eq "Local") {
        Write-PublishLog "Warning: Changes detected in files other than .csproj:" "Warning"
        foreach ($file in $nonCsprojFiles) {
            Write-PublishLog "  - $file" "Warning"
        }

        $response = Read-Host "Continue with commit anyway? (y/N)"
        if ($response -ne "y" -and $response -ne "Y") {
            throw "Git operations cancelled by user due to unexpected file changes."
        }
    }

    # Stage all changes
    Write-PublishLog "Staging changes..." "Info"
    git add . 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to stage changes with git add."
    }

    # Commit changes
    $commitMessage = "Publish version $Version to $Environment environment"
    Write-PublishLog "Committing changes with message: $commitMessage" "Info"

    git commit -m $commitMessage 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to commit changes."
    }

    Write-PublishLog "Changes committed successfully." "Success"
}

<#
.SYNOPSIS
    Creates a Git tag

.DESCRIPTION
    Creates a version tag with environment suffix

.PARAMETER Version
    The version to tag

.PARAMETER Environment
    The environment suffix for the tag

.OUTPUTS
    String - The created tag name
#>
function New-GitTag {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Version,

        [Parameter(Mandatory=$true)]
        [string]$Environment
    )

    Write-PublishLog "Creating Git tag..." "Info"

    # Create tag with environment suffix
    $envSuffix = $Environment.ToLower()
    $tagName = "v$Version-$envSuffix"

    Write-PublishLog "Tag name: $tagName" "Info"

    git tag $tagName 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) {
        # Check if tag already exists
        $existingTag = git tag -l $tagName 2>&1
        if ($existingTag) {
            Write-PublishLog "Tag $tagName already exists, skipping tag creation" "Warning"
        } else {
            throw "Failed to create tag $tagName."
        }
    } else {
        Write-PublishLog "Tag $tagName created successfully" "Success"
    }

    return $tagName
}

<#
.SYNOPSIS
    Pushes commits and tags to remote

.DESCRIPTION
    Pushes local commits and tags to the remote repository

.PARAMETER TagName
    Optional specific tag name to push

.PARAMETER Interactive
    If true, prompts user before pushing (default for local)
#>
function Push-GitChanges {
    param(
        [Parameter(Mandatory=$false)]
        [string]$TagName = $null,

        [Parameter(Mandatory=$false)]
        [bool]$Interactive = $true
    )

    $context = Test-ExecutionContext

    # In CI/CD, automatically push unless explicitly disabled
    # In local context, ask for confirmation
    if ($context -eq "Local" -and $Interactive) {
        $pushResponse = Read-Host "Push changes and tag to remote repository? (y/N)"
        if ($pushResponse -ne "y" -and $pushResponse -ne "Y") {
            Write-PublishLog "Skipping push to remote repository." "Info"
            return
        }
    }

    Write-PublishLog "Pushing changes to remote..." "Info"
    git push 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Write-PublishLog "Warning: Failed to push changes to remote" "Warning"
    } else {
        Write-PublishLog "Changes pushed successfully" "Success"
    }

    if ($TagName) {
        Write-PublishLog "Pushing tag $TagName to remote..." "Info"
        git push origin $TagName 2>&1 | Out-Null
        if ($LASTEXITCODE -ne 0) {
            Write-PublishLog "Warning: Failed to push tag to remote" "Warning"
        } else {
            Write-PublishLog "Tag pushed successfully" "Success"
        }
    }
}

<#
.SYNOPSIS
    Performs all Git operations for a publish

.DESCRIPTION
    Convenience function that handles commit, tag, and push

.PARAMETER Version
    The version being published

.PARAMETER Environment
    The environment being published to

.PARAMETER SkipPush
    If true, skips pushing to remote
#>
function Invoke-GitOperations {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Version,

        [Parameter(Mandatory=$true)]
        [string]$Environment,

        [Parameter(Mandatory=$false)]
        [switch]$SkipPush
    )

    Write-PublishLog "Performing Git operations..." "Info"

    # Commit changes
    Invoke-GitCommit -Version $Version -Environment $Environment

    # Create tag
    $tagName = New-GitTag -Version $Version -Environment $Environment

    # Push changes and tag
    if (-not $SkipPush) {
        Push-GitChanges -TagName $tagName
    }

    Write-PublishLog "Git operations completed successfully" "Success"
}

# Export functions
Export-ModuleMember -Function @(
    'Test-GitState',
    'Invoke-GitCommit',
    'New-GitTag',
    'Push-GitChanges',
    'Invoke-GitOperations'
)
