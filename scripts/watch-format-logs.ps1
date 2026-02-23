# Real-time log watcher for write_format optimization testing
# Monitors the log file and highlights optimization indicators

param(
    [string]$LogPath = "$env:LOCALAPPDATA\X21\X21-deno-Debug\Logs\deno-ha-desktop.log"
)

Clear-Host
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Format Optimization - Live Log Monitor" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Monitoring: $LogPath" -ForegroundColor Yellow
Write-Host "Press Ctrl+C to stop" -ForegroundColor DarkGray
Write-Host ""

if (-not (Test-Path $LogPath)) {
    Write-Host "ERROR: Log file not found!" -ForegroundColor Red
    Write-Host "Path: $LogPath" -ForegroundColor Red
    exit 1
}

# Track statistics
$stats = @{
    WriteFormatCalls = 0
    ParallelOps = 0
    SelectiveOps = 0
    PerCellLogs = 0
    ReaderCounts = @{}
}

# Function to display stats
function Show-Stats {
    Write-Host ""
    Write-Host "--- Statistics ---" -ForegroundColor Cyan
    Write-Host "write_format calls: $($stats.WriteFormatCalls)" -ForegroundColor White
    Write-Host "Parallel operations: $($stats.ParallelOps)" -ForegroundColor Green
    Write-Host "Selective operations: $($stats.SelectiveOps)" -ForegroundColor Green
    Write-Host "Per-cell logs (bad): $($stats.PerCellLogs)" -ForegroundColor $(if ($stats.PerCellLogs -gt 0) { 'Red' } else { 'Green' })

    if ($stats.ReaderCounts.Count -gt 0) {
        Write-Host "Reader usage:" -ForegroundColor Yellow
        foreach ($readers in ($stats.ReaderCounts.Keys | Sort-Object)) {
            $count = $stats.ReaderCounts[$readers]
            Write-Host "  $readers/7 readers: $count times" -ForegroundColor Cyan
        }
    }
    Write-Host ""
}

# Get current file position
$lastPosition = (Get-Item $LogPath).Length

Write-Host "Waiting for format operations..." -ForegroundColor Yellow
Write-Host "(Use X21 to format some cells in Excel)" -ForegroundColor DarkGray
Write-Host ""

# Monitor loop
$lastStatsTime = Get-Date
try {
    while ($true) {
        # Check for new content
        $currentSize = (Get-Item $LogPath).Length

        if ($currentSize -gt $lastPosition) {
            # Read new content
            $stream = [System.IO.File]::Open($LogPath, 'Open', 'Read', 'ReadWrite')
            $stream.Seek($lastPosition, [System.IO.SeekOrigin]::Begin) | Out-Null
            $reader = New-Object System.IO.StreamReader($stream)

            while (-not $reader.EndOfStream) {
                $line = $reader.ReadLine()

                # Parse and colorize the line
                $color = "Gray"
                $prefix = ""

                # Check for various indicators
                if ($line -match "Executing write format tool") {
                    $stats.WriteFormatCalls++
                    $color = "Cyan"
                    $prefix = "📝 "
                }
                elseif ($line -match "Reading formats using parallel processing with (\d+) readers") {
                    $stats.ParallelOps++
                    $readers = [int]$matches[1]
                    if (-not $stats.ReaderCounts.ContainsKey($readers)) {
                        $stats.ReaderCounts[$readers] = 0
                    }
                    $stats.ReaderCounts[$readers]++
                    $color = "Green"
                    $prefix = "⚡ "
                }
                elseif ($line -match "selective:\s*\[?([^\]]+)" -and $line -match "Reading format") {
                    $stats.SelectiveOps++
                    $color = "Yellow"
                    $prefix = "🎯 "
                }
                elseif ($line -match "Filtered (\d+)/\d+ readers") {
                    $color = "Cyan"
                    $prefix = "🔍 "
                }
                elseif ($line -match "Completed.*FormatReader") {
                    $color = "DarkGreen"
                    $prefix = "✓ "
                }
                elseif ($line -match "Cell #\d+:") {
                    $stats.PerCellLogs++
                    $color = "Red"
                    $prefix = "⚠️ "
                }
                elseif ($line -match "error|exception|failed" -and $line -notmatch "Error reading.*for cell") {
                    $color = "Red"
                    $prefix = "❌ "
                }

                # Extract just the message part (remove timestamp and log level)
                if ($line -match "\|\s*(.*)") {
                    $message = $matches[1]
                    Write-Host "$prefix$message" -ForegroundColor $color
                } else {
                    Write-Host "$prefix$line" -ForegroundColor $color
                }
            }

            $reader.Close()
            $stream.Close()
            $lastPosition = $currentSize
        }

        # Show stats periodically
        $now = Get-Date
        if (($now - $lastStatsTime).TotalSeconds -gt 30) {
            Show-Stats
            $lastStatsTime = $now
        }

        # Sleep briefly
        Start-Sleep -Milliseconds 100
    }
}
catch {
    Write-Host ""
    Write-Host "Monitoring stopped." -ForegroundColor Yellow
}
finally {
    Show-Stats
    Write-Host "Final statistics shown above." -ForegroundColor Green
}
