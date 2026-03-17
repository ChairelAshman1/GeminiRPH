# auto-sync-git.ps1
# Usage: run in PowerShell from the repository root:
#   .\auto-sync-git.ps1
# It monitors files for changes and performs git add/commit/push automatically.
# NOTE: commit messages are auto-generated; adjust logic as needed.

$repo = (Get-Location).Path
$filter = '*.*'
$outputLog = "$repo\auto-sync-git.log"

Write-Host "[AutoSync] Started in $repo" -ForegroundColor Green
Write-Host "[AutoSync] Logging to $outputLog" -ForegroundColor Green

$fsw = New-Object System.IO.FileSystemWatcher $repo, $filter
$fsw.IncludeSubdirectories = $true
$fsw.EnableRaisingEvents = $true
$fsw.NotifyFilter = [System.IO.NotifyFilters]::FileName -bor [System.IO.NotifyFilters]::LastWrite -bor [System.IO.NotifyFilters]::DirectoryName

$timer = $null
$pending = $false

function Sync-Git {
    if ($pending) { return }
    $pending = $true
    Start-Sleep -Milliseconds 500

    $status = git status --porcelain
    if ([string]::IsNullOrWhiteSpace($status)) {
        $pending = $false
        return
    }

    try {
        git add --all
        $message = "Auto-sync at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        git commit -m $message
        git push origin main
        $entry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Synced $message`n"
        Add-Content -Path $outputLog -Value $entry
        Write-Host "[AutoSync] $message" -ForegroundColor Cyan
    } catch {
        Write-Host "[AutoSync] ERROR: $_" -ForegroundColor Red
        Add-Content -Path $outputLog -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ERROR: $_`n"
    } finally {
        $pending = $false
    }
}

Register-ObjectEvent $fsw Changed -Action { Sync-Git } | Out-Null
Register-ObjectEvent $fsw Created -Action { Sync-Git } | Out-Null
Register-ObjectEvent $fsw Deleted -Action { Sync-Git } | Out-Null
Register-ObjectEvent $fsw Renamed -Action { Sync-Git } | Out-Null

Write-Host "[AutoSync] Monitoring file changes. Press Ctrl+C to stop." -ForegroundColor Green

try {
    while ($true) { Start-Sleep -Seconds 1 }
} finally {
    Unregister-Event -SourceIdentifier * | Out-Null
    $fsw.Dispose()
    Write-Host "[AutoSync] Stopped." -ForegroundColor Yellow
}
