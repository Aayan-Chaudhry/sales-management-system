$ErrorActionPreference = "Continue"

$folder = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $folder

$port = 5000
$url = "http://127.0.0.1:$port/login"
$logFile = Join-Path $folder "launcher_log.txt"

function Write-LauncherLog {
    param([string]$Message)
    try {
        $stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Add-Content -Path $logFile -Value "[$stamp] $Message"
    } catch {}
}

function Get-TrackerProcessIds {
    $connections = Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue
    if (-not $connections) { return @() }
    return @($connections | Select-Object -ExpandProperty OwningProcess -Unique)
}

function Stop-TrackerServer {
    $trackerPids = Get-TrackerProcessIds
    foreach ($trackerPid in $trackerPids) {
        try {
            Stop-Process -Id $trackerPid -Force -ErrorAction SilentlyContinue
            Write-LauncherLog "Stopped tracker process $trackerPid"
        } catch {}
    }
}

function Test-TrackerRunning {
    $trackerPids = Get-TrackerProcessIds
    return ($trackerPids.Count -gt 0)
}

function Find-Browser {
    $edge = @(
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1

    if ($edge) { return $edge }

    $chrome = @(
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
        "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1

    if ($chrome) { return $chrome }

    return $null
}

Write-LauncherLog "Launcher opened from $folder"

# Always try to create a database backup when the VBS/app launcher is opened.
# This still works when the tracker is already running.
try {
    if (Test-Path (Join-Path $folder "backup_db.py")) {
        Write-LauncherLog "Running backup_db.py"
        $backupOutput = & python (Join-Path $folder "backup_db.py") 2>&1
        foreach ($line in $backupOutput) { Write-LauncherLog "backup_db.py: $line" }
    } else {
        Write-LauncherLog "backup_db.py not found in launcher folder"
    }
} catch {
    Write-LauncherLog "backup_db.py failed: $($_.Exception.Message)"
}

$startedServer = $false

# If a tracker server is already running, do not start another one.
# Just open the login page, which clears the saved Admin/Employee session.
if (-not (Test-TrackerRunning)) {
    $startedServer = $true
    Write-LauncherLog "No tracker server found. Starting START.bat hidden."
    Start-Process -FilePath "cmd.exe" -ArgumentList "/c", "`"$folder\START.bat`"" -WorkingDirectory $folder -WindowStyle Hidden | Out-Null

    # Wait up to 15 seconds for Flask to start.
    $deadline = (Get-Date).AddSeconds(15)
    while ((Get-Date) -lt $deadline) {
        if (Test-TrackerRunning) { break }
        Start-Sleep -Milliseconds 300
    }
} else {
    Write-LauncherLog "Tracker server already running on port $port. Opening login page instead of reusing old session."
}

$browser = Find-Browser

if ($browser) {
    Write-LauncherLog "Opening browser app window using $browser at $url"
    # Dedicated app profile so this window behaves more like its own app.
    $profile = Join-Path $env:TEMP "RoudhaTrackerAppProfile"
    New-Item -ItemType Directory -Force -Path $profile | Out-Null

    $args = @(
        "--app=$url",
        "--user-data-dir=$profile",
        "--no-first-run"
    )

    $browserProcess = Start-Process -FilePath $browser -ArgumentList $args -PassThru

    # Wait for the dedicated app window process to close.
    if ($browserProcess) {
        Wait-Process -Id $browserProcess.Id -ErrorAction SilentlyContinue
    }
} else {
    Write-LauncherLog "No Edge/Chrome found. Opening default browser at $url."
    Start-Process $url
    Start-Sleep -Seconds 3600
}

# If this launcher started the server, stop it when the app window closes.
if ($startedServer) {
    Write-LauncherLog "App window closed. Stopping server that this launcher started."
    Stop-TrackerServer
} else {
    Write-LauncherLog "App window closed. Server was already running before launcher opened, so leaving it running."
}
