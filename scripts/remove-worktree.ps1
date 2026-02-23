param(
    [Parameter(Mandatory=$false)]
    [string]$worktree,
    [Parameter(Mandatory=$false)]
    [switch]$Force,
    [Parameter(Mandatory=$false)]
    [switch]$List
)

$ErrorActionPreference = 'Stop'

# Function to list all worktrees
function Get-Worktrees {
    $worktrees = git worktree list --porcelain
    $result = @()
    $current = @{}

    foreach ($line in $worktrees) {
        if ($line -match '^worktree (.+)$') {
            if ($current.Count -gt 0) {
                $result += $current
            }
            $current = @{ Path = $matches[1] }
        }
        elseif ($line -match '^HEAD (.+)$') {
            $current.HEAD = $matches[1]
        }
        elseif ($line -match '^branch (.+)$') {
            $current.Branch = $matches[1] -replace '^refs/heads/', ''
        }
        elseif ($line -match '^detached') {
            $current.Branch = 'DETACHED'
        }
    }
    if ($current.Count -gt 0) {
        $result += $current
    }
    return $result
}

# Function to display worktrees in a nice format
function Show-Worktrees {
    $worktrees = Get-Worktrees
    Write-Host "`nCurrent worktrees:" -ForegroundColor Cyan
    Write-Host ("=" * 80) -ForegroundColor Gray

    foreach ($wt in $worktrees) {
        $branch = if ($wt.Branch) { $wt.Branch } else { "N/A" }
        $path = $wt.Path
        $isMain = $path -match '[\\/]x21$'

        if ($isMain) {
            Write-Host "  [MAIN] " -NoNewline -ForegroundColor Green
        } else {
            Write-Host "  " -NoNewline
        }
        Write-Host "$branch" -NoNewline -ForegroundColor Yellow
        Write-Host " -> " -NoNewline -ForegroundColor Gray
        Write-Host "$path" -ForegroundColor Cyan
    }
    Write-Host ("=" * 80) -ForegroundColor Gray
    Write-Host ""
}

# If -List flag is provided, just show worktrees and exit
if ($List) {
    Show-Worktrees
    exit 0
}

# If no worktree specified, show list and prompt
if (-not $worktree) {
    Show-Worktrees
    Write-Host "Usage: .\remove-worktree.ps1 [-worktree] <name-or-path> [-Force] [-List]" -ForegroundColor Yellow
    Write-Host "`nExamples:" -ForegroundColor Cyan
    Write-Host "  .\remove-worktree.ps1 feature-xyz" -ForegroundColor Gray
    Write-Host "  .\remove-worktree.ps1 ..\x21-feature-xyz" -ForegroundColor Gray
    Write-Host "  .\remove-worktree.ps1 -List" -ForegroundColor Gray
    Write-Host "`nOptions:" -ForegroundColor Cyan
    Write-Host "  -Force    Force removal even with uncommitted changes" -ForegroundColor Gray
    Write-Host "  -List     List all worktrees" -ForegroundColor Gray
    exit 0
}

# Resolve worktree path
$worktreePath = $worktree

# If it's not a full path, try to construct it
if (-not (Test-Path $worktreePath)) {
    # Try as x21-<name> in parent directory
    $candidatePath = Join-Path (Split-Path $PSScriptRoot -Parent) "..\x21-$worktree"
    if (Test-Path $candidatePath) {
        $worktreePath = $candidatePath
    } else {
        # Try the worktree value as-is if it looks like a branch name
        $candidatePath2 = Join-Path (Split-Path $PSScriptRoot -Parent) "..\$worktree"
        if (Test-Path $candidatePath2) {
            $worktreePath = $candidatePath2
        }
    }
}

# Verify path exists
if (-not (Test-Path $worktreePath)) {
    Write-Host "Error: Worktree not found at: $worktreePath" -ForegroundColor Red
    Write-Host "`nAvailable worktrees:" -ForegroundColor Yellow
    Show-Worktrees
    exit 1
}

# Get absolute path
$worktreeAbsPath = Resolve-Path $worktreePath

# Verify it's actually a git worktree
$allWorktrees = Get-Worktrees
$isWorktree = $false
$branchName = $null
foreach ($wt in $allWorktrees) {
    if ((Resolve-Path $wt.Path).Path -eq $worktreeAbsPath.Path) {
        $isWorktree = $true
        $branchName = $wt.Branch
        break
    }
}

if (-not $isWorktree) {
    Write-Host "Error: '$worktreeAbsPath' is not a registered git worktree" -ForegroundColor Red
    Write-Host "`nAvailable worktrees:" -ForegroundColor Yellow
    Show-Worktrees
    exit 1
}

# Don't allow removing the main worktree
if ($worktreeAbsPath.Path -match '[\\/]x21$') {
    Write-Host "Error: Cannot remove the main worktree!" -ForegroundColor Red
    exit 1
}

# Confirm deletion
Write-Host "`nAbout to remove worktree:" -ForegroundColor Yellow
Write-Host "  Path: $worktreeAbsPath" -ForegroundColor Cyan
Write-Host "  Branch: $branchName" -ForegroundColor Cyan

if (-not $Force) {
    $confirmation = Read-Host "`nAre you sure you want to remove this worktree? (y/N)"
    if ($confirmation -ne 'y' -and $confirmation -ne 'Y') {
        Write-Host "Cancelled." -ForegroundColor Yellow
        exit 0
    }
}

# Remove the worktree
Write-Host "`nRemoving worktree..." -ForegroundColor Green

$removeArgs = @('worktree', 'remove', $worktreeAbsPath.Path)
if ($Force) {
    $removeArgs += '--force'
}

& git $removeArgs

if ($LASTEXITCODE -eq 0) {
    Write-Host "✓ Worktree removed successfully" -ForegroundColor Green

    # Show remaining worktrees
    Write-Host "`nRemaining worktrees:" -ForegroundColor Cyan
    Show-Worktrees
} else {
    Write-Host "✗ Failed to remove worktree" -ForegroundColor Red
    Write-Host "`nTip: Use -Force flag to force removal if there are uncommitted changes" -ForegroundColor Yellow
    exit 1
}
