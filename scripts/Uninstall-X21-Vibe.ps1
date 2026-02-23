# X21 Application Uninstaller Script
# This script removes the X21 Excel add-in completely from the system
# Use this as an alternative to Windows Add/Remove Programs

param(
    [Parameter(Mandatory=$false)]
    [switch]$Force,

    [Parameter(Mandatory=$false)]
    [switch]$VerboseOutput,

    [Parameter(Mandatory=$false)]
    [switch]$WhatIf
)

# Set error action preference
$ErrorActionPreference = "Stop"

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
}

# Function to check if running as administrator
function Test-IsAdmin {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Function to find ClickOnce installed applications
function Get-ClickOnceApps {
    Write-Log "Searching for X21 ClickOnce installations..." "Info"

    $clickOnceApps = @()

    # Common ClickOnce installation paths
    $clickOncePaths = @(
        "$env:LOCALAPPDATA\Apps\2.0",
        "$env:USERPROFILE\AppData\Local\Apps\2.0"
    )

    foreach ($basePath in $clickOncePaths) {
        if (Test-Path $basePath) {
            Write-Log "Checking ClickOnce path: $basePath" "Info"

            # Search for X21 installation folders
            $x21Folders = Get-ChildItem -Path $basePath -Recurse -Directory -Name "X21*" -ErrorAction SilentlyContinue

            foreach ($folder in $x21Folders) {
                $fullPath = Join-Path $basePath $folder
                $manifestFiles = Get-ChildItem -Path $fullPath -Filter "*.application" -Recurse -ErrorAction SilentlyContinue

                foreach ($manifest in $manifestFiles) {
                    $clickOnceApps += @{
                        Name = "X21"
                        Path = $fullPath
                        ManifestFile = $manifest.FullName
                    }
                    Write-Log "Found X21 installation at: $fullPath" "Success"
                }
            }
        }
    }

    return $clickOnceApps
}

# Function to uninstall ClickOnce application using rundll32
function Uninstall-ClickOnceApp {
    param(
        [string]$ManifestPath
    )

    Write-Log "Attempting to uninstall ClickOnce app using manifest: $ManifestPath" "Info"

    if ($WhatIf) {
        Write-Log "WHAT-IF: Would uninstall ClickOnce app: $ManifestPath" "Warning"
        return $true
    }

    try {
        # Use rundll32 to uninstall the ClickOnce application
        $arguments = "dfshim.dll,ShArpMaintain $ManifestPath"
        Write-Log "Executing: rundll32.exe $arguments" "Info"

        $process = Start-Process -FilePath "rundll32.exe" -ArgumentList $arguments -Wait -PassThru -WindowStyle Hidden

        if ($process.ExitCode -eq 0) {
            Write-Log "ClickOnce uninstall completed successfully" "Success"
            return $true
        } else {
            Write-Log "ClickOnce uninstall returned exit code: $($process.ExitCode)" "Warning"
            return $false
        }
    } catch {
        Write-Log "Error during ClickOnce uninstall: $($_.Exception.Message)" "Error"
        return $false
    }
}

# Function to remove registry entries
function Remove-RegistryEntries {
    Write-Log "Removing registry entries..." "Info"

    # First, let's check the specific X21 key we know exists
    $x21AddinKey = "HKCU:\Software\Microsoft\Office\Excel\Addins\X21"
    Write-Log "Specifically testing for X21 add-in key: $x21AddinKey" "Info"

    if (Test-Path $x21AddinKey) {
        Write-Log "CONFIRMED: X21 add-in registry key EXISTS at: $x21AddinKey" "Warning"
        try {
            $keyInfo = Get-Item $x21AddinKey -ErrorAction SilentlyContinue
            Write-Log "Key details: $($keyInfo.Name)" "Info"
            $properties = Get-ItemProperty $x21AddinKey -ErrorAction SilentlyContinue
            Write-Log "Key properties: $($properties | Format-List | Out-String)" "Info"
        } catch {
            Write-Log "Could not get key details: $($_.Exception.Message)" "Warning"
        }
    } else {
        Write-Log "X21 add-in registry key NOT FOUND at: $x21AddinKey" "Info"
    }

    $registryPaths = @(
        "HKCU:\Software\Microsoft\Office\Excel\Addins\X21",
        "HKCU:\Software\Microsoft\VSTO\Security\Inclusion Lists\http*kontext21*",
        "HKCU:\Software\Microsoft\Office\16.0\User Settings\X21*",
        "HKCU:\Software\Microsoft\Office\15.0\User Settings\X21*",
        "HKCU:\Software\Classes\Software\Microsoft\Windows\CurrentVersion\Deployment\SideBySide\2.0\X21*",
        "HKLM:\Software\Classes\Software\Microsoft\Windows\CurrentVersion\Deployment\SideBySide\2.0\X21*",
        "HKCU:\Software\Microsoft\Office\Excel\Addins\X21*",
        "HKCU:\Software\Microsoft\VSTO_LoaderLock\X21*",
        "HKCU:\Software\Microsoft\VSTO\Security\Inclusion Lists\*X21*",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*X21*",
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*X21*"
    )

    $removedAny = $false

    foreach ($regPath in $registryPaths) {
        try {
            Write-Log "Processing registry path: $regPath" "Info"
            # Handle wildcards by getting parent path and searching
            if ($regPath -like "*`*") {
                $parentPath = $regPath -replace '\\\*.*$', ''
                $searchPattern = ($regPath -split '\\')[-1]

                if (Test-Path $parentPath) {
                    $childItems = Get-ChildItem -Path $parentPath -ErrorAction SilentlyContinue | Where-Object { $_.Name -like $searchPattern }

                    foreach ($item in $childItems) {
                        Write-Log "Found matching registry key: $($item.PSPath)" "Info"
                        if ($WhatIf) {
                            Write-Log "WHAT-IF: Would remove registry key: $($item.PSPath)" "Warning"
                        } else {
                            try {
                                Remove-Item -Path $item.PSPath -Recurse -Force -ErrorAction Stop
                                Write-Log "Successfully removed registry key: $($item.PSPath)" "Success"
                                $removedAny = $true
                            } catch {
                                Write-Log "Failed to remove registry key $($item.PSPath)`: $($_.Exception.Message)" "Error"
                            }
                        }
                    }
                }
            } else {
                # Direct path
                Write-Log "Checking registry path: $regPath" "Info"
                if (Test-Path $regPath) {
                    Write-Log "Registry key exists: $regPath" "Info"
                    if ($WhatIf) {
                        Write-Log "WHAT-IF: Would remove registry key: $regPath" "Warning"
                    } else {
                        try {
                            Remove-Item -Path $regPath -Recurse -Force -ErrorAction Stop
                            Write-Log "Successfully removed registry key: $regPath" "Success"
                            $removedAny = $true
                        } catch {
                            Write-Log "Failed to remove registry key $regPath`: $($_.Exception.Message)" "Error"
                        }
                    }
                } else {
                    Write-Log "Registry key not found: $regPath" "Info"
                }
            }
        } catch {
            Write-Log "Error removing registry key $regPath`: $($_.Exception.Message)" "Warning"
        }
    }

    # Force removal of the specific X21 key we confirmed exists
    if (Test-Path $x21AddinKey) {
        Write-Log "Force removing confirmed X21 registry key..." "Warning"
        if ($WhatIf) {
            Write-Log "WHAT-IF: Would force remove: $x21AddinKey" "Warning"
            $removedAny = $true
        } else {
            try {
                Remove-Item -Path $x21AddinKey -Recurse -Force -ErrorAction Stop
                Write-Log "SUCCESS: Force removed X21 registry key: $x21AddinKey" "Success"
                $removedAny = $true
            } catch {
                Write-Log "FAILED: Could not force remove X21 registry key: $($_.Exception.Message)" "Error"
            }
        }
    }

    if (-not $removedAny) {
        Write-Log "No registry entries found to remove" "Info"
    }
}

# Function to remove file system remnants
function Remove-FileSystemRemnants {
    Write-Log "Removing file system remnants..." "Info"

    $pathsToCheck = @(
        "$env:LOCALAPPDATA\Apps\2.0",
        "$env:LOCALAPPDATA\assembly\dl3",
        "$env:APPDATA\Microsoft\Excel\XLSTART",
        "$env:TEMP\VSTOInstaller*",
        "$env:TEMP\*X21*"
    )

    $removedAny = $false

    foreach ($basePath in $pathsToCheck) {
        try {
            if ($basePath -like "*`*") {
                # Handle wildcards
                $parentPath = Split-Path $basePath -Parent
                $searchPattern = Split-Path $basePath -Leaf

                if (Test-Path $parentPath) {
                    $items = Get-ChildItem -Path $parentPath -Filter $searchPattern -ErrorAction SilentlyContinue

                    foreach ($item in $items) {
                        if ($WhatIf) {
                            Write-Log "WHAT-IF: Would remove: $($item.FullName)" "Warning"
                        } else {
                            Remove-Item -Path $item.FullName -Recurse -Force -ErrorAction SilentlyContinue
                            Write-Log "Removed: $($item.FullName)" "Success"
                            $removedAny = $true
                        }
                    }
                }
            } else {
                # Search for X21-related items in the path
                if (Test-Path $basePath) {
                    $x21Items = Get-ChildItem -Path $basePath -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*X21*" -or $_.Name -like "*kontext21*" }

                    foreach ($item in $x21Items) {
                        if ($WhatIf) {
                            Write-Log "WHAT-IF: Would remove: $($item.FullName)" "Warning"
                        } else {
                            Remove-Item -Path $item.FullName -Recurse -Force -ErrorAction SilentlyContinue
                            Write-Log "Removed: $($item.FullName)" "Success"
                            $removedAny = $true
                        }
                    }
                }
            }
        } catch {
            Write-Log "Error checking path $basePath`: $($_.Exception.Message)" "Warning"
        }
    }

    if (-not $removedAny) {
        Write-Log "No file system remnants found to remove" "Info"
    }
}

# Function to close Excel processes
function Stop-ExcelProcesses {
    Write-Log "Checking for running Excel processes..." "Info"

    $excelProcesses = Get-Process -Name "excel" -ErrorAction SilentlyContinue

    if ($excelProcesses) {
        if ($Force -or $WhatIf) {
            foreach ($process in $excelProcesses) {
                if ($WhatIf) {
                    Write-Log "WHAT-IF: Would stop Excel process (PID: $($process.Id))" "Warning"
                } else {
                    Write-Log "Stopping Excel process (PID: $($process.Id))" "Info"
                    $process | Stop-Process -Force
                    Write-Log "Excel process stopped" "Success"
                }
            }
        } else {
            Write-Log "Excel is currently running. Please close Excel and run again, or use -Force parameter" "Warning"
            $response = Read-Host "Close Excel automatically? (y/N)"
            if ($response -eq "y" -or $response -eq "Y") {
                foreach ($process in $excelProcesses) {
                    Write-Log "Stopping Excel process (PID: $($process.Id))" "Info"
                    $process | Stop-Process -Force
                }
            } else {
                throw "Excel must be closed to proceed with uninstallation"
            }
        }
    } else {
        Write-Log "No Excel processes found running" "Info"
    }
}

# Function to clear Office add-in cache
function Clear-OfficeAddinCache {
    Write-Log "Clearing Office add-in cache..." "Info"

    $cachePaths = @(
        "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef",
        "$env:LOCALAPPDATA\Microsoft\Office\15.0\Wef",
        "$env:APPDATA\Microsoft\Office\Recent",
        "$env:APPDATA\Microsoft\Templates"
    )

    foreach ($cachePath in $cachePaths) {
        if (Test-Path $cachePath) {
            try {
                $x21CacheFiles = Get-ChildItem -Path $cachePath -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*X21*" -or $_.Name -like "*kontext21*" }

                foreach ($file in $x21CacheFiles) {
                    if ($WhatIf) {
                        Write-Log "WHAT-IF: Would remove cache file: $($file.FullName)" "Warning"
                    } else {
                        Remove-Item -Path $file.FullName -Force -ErrorAction SilentlyContinue
                        Write-Log "Removed cache file: $($file.FullName)" "Success"
                    }
                }
            } catch {
                Write-Log "Error clearing cache from $cachePath`: $($_.Exception.Message)" "Warning"
            }
        }
    }
}

# Main execution
try {
    Write-Host ""
    Write-Host "=== X21 APPLICATION UNINSTALLER ===" -ForegroundColor Cyan
    Write-Host "This script will completely remove X21 from your system" -ForegroundColor White
    Write-Host ""

    if ($WhatIf) {
        Write-Host "RUNNING IN WHAT-IF MODE - No changes will be made" -ForegroundColor Yellow
        Write-Host ""
    }

    # Check if running as admin for registry operations
    if (-not (Test-IsAdmin)) {
        Write-Log "Not running as administrator. Some registry operations may fail." "Warning"
        if (-not $Force) {
            $response = Read-Host "Continue anyway? (y/N)"
            if ($response -ne "y" -and $response -ne "Y") {
                Write-Log "Operation cancelled by user" "Info"
                exit 0
            }
        }
    }

    Write-Log "Starting X21 uninstallation process..." "Info"

    # Step 1: Close Excel processes
    Stop-ExcelProcesses

    # Step 2: Find and uninstall ClickOnce applications
    $clickOnceApps = Get-ClickOnceApps

    if ($clickOnceApps.Count -eq 0) {
        Write-Log "No X21 ClickOnce installations found" "Warning"
    } else {
        Write-Log "Found $($clickOnceApps.Count) X21 installation(s)" "Info"

        foreach ($app in $clickOnceApps) {
            $success = Uninstall-ClickOnceApp -ManifestPath $app.ManifestFile
            if (-not $success -and -not $WhatIf) {
                Write-Log "ClickOnce uninstall may have failed, continuing with manual cleanup..." "Warning"
            }
        }
    }

    # Step 3: Remove registry entries
    Remove-RegistryEntries

    # Step 4: Remove file system remnants
    Remove-FileSystemRemnants

    # Step 5: Clear Office add-in cache
    Clear-OfficeAddinCache

    # Step 6: Final verification
    Write-Log "Performing final verification..." "Info"
    $remainingApps = Get-ClickOnceApps

    if ($remainingApps.Count -eq 0 -or $WhatIf) {
        Write-Log "X21 uninstallation completed successfully!" "Success"

        if (-not $WhatIf) {
            Write-Log "Please restart Excel to ensure all components are fully removed" "Info"
            Write-Log "If you experience any issues, you may need to restart your computer" "Info"
        }
    } else {
        Write-Log "Some X21 components may still be present. Manual removal may be required." "Warning"
    }

    # Display summary
    Write-Host ""
    Write-Host "=== UNINSTALL SUMMARY ===" -ForegroundColor Cyan
    if ($WhatIf) {
        Write-Host "Mode: WHAT-IF (no changes made)" -ForegroundColor Yellow
    } else {
        Write-Host "Mode: EXECUTE" -ForegroundColor Green
    }
    Write-Host "ClickOnce apps found: $($clickOnceApps.Count)" -ForegroundColor White
    Write-Host "Registry cleanup: Completed" -ForegroundColor White
    Write-Host "File system cleanup: Completed" -ForegroundColor White
    Write-Host "Office cache cleanup: Completed" -ForegroundColor White
    Write-Host "=========================" -ForegroundColor Cyan
    Write-Host ""

    if (-not $WhatIf) {
        Write-Host "X21 has been uninstalled. Please restart Excel." -ForegroundColor Green
    }

} catch {
    Write-Log "Uninstallation failed: $($_.Exception.Message)" "Error"
    Write-Log "You may need to run this script as administrator or manually remove remaining components" "Error"
    exit 1
}
