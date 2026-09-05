# install_scheduled_task.ps1
# ---------------------------------------------------------------------------
# One-time Task Scheduler registration for the ZOHO daily auto-merge script.
#
# What this creates:
#   Task name  : ZOHO_Daily_AutoMerge
#   Trigger    : Every hour from 08:00 to 12:00, Monday-Saturday
#   Action     : C:\anaconda3\python.exe daily_auto_merge.py  (from the repo)
#   Runs as    : the interactive user, only when logged in
#                (Outlook desktop must be open for COM to work)
#   Idle/net   : no idle/network requirements; runs even on battery
#
# Re-running this script overwrites the existing task in place.
# To uninstall: Unregister-ScheduledTask -TaskName ZOHO_Daily_AutoMerge -Confirm:$false
# ---------------------------------------------------------------------------

$ErrorActionPreference = 'Stop'

$TaskName    = 'ZOHO_Daily_AutoMerge'
$ScriptDir   = Split-Path -Parent $MyInvocation.MyCommand.Path
$ScriptPath  = Join-Path $ScriptDir 'daily_auto_merge.py'
$PythonExe   = 'C:\anaconda3\python.exe'   # Kshitij's Anaconda Python
$LogsDir     = 'C:\Users\k.buch\OneDrive - Transasia Bio Medicals Ltd\TBM 2026 Onwards\Pending Calls\Raw Files'

# ── Sanity checks ──────────────────────────────────────────────────────────
if (-not (Test-Path $ScriptPath)) {
    throw "daily_auto_merge.py not found at: $ScriptPath"
}
if (-not (Test-Path $PythonExe)) {
    throw "Anaconda Python not found at: $PythonExe.  Update `$PythonExe in this script if your install is elsewhere."
}
if (-not (Test-Path $LogsDir)) {
    throw "Raw Files folder not found at: $LogsDir.  OneDrive may not be signed in."
}

Write-Host ""
Write-Host "Installing scheduled task '$TaskName'..." -ForegroundColor Cyan
Write-Host "  Script  : $ScriptPath"
Write-Host "  Python  : $PythonExe"
Write-Host "  Logs to : $LogsDir\auto_merge.log"
Write-Host ""

# ── Remove any prior version so re-running upgrades cleanly ────────────────
$existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existing) {
    Write-Host "  (replacing existing '$TaskName' task)" -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}

# ── Action ─────────────────────────────────────────────────────────────────
# -Argument is quoted so the path with a space works.
$action = New-ScheduledTaskAction `
    -Execute  $PythonExe `
    -Argument "`"$ScriptPath`"" `
    -WorkingDirectory $ScriptDir

# ── Trigger: hourly Mon-Sat, 08:00 → 12:00 ─────────────────────────────────
# Register-ScheduledTask does not have a native "hourly for N hours on
# specific weekdays" one-liner, so we build 5 hourly triggers and constrain
# each to Mon-Sat via -DaysOfWeek on a Weekly trigger.
$daysOfWeek = 'Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'
$triggers = @()
foreach ($hour in 8..12) {
    $t = New-ScheduledTaskTrigger `
        -Weekly -DaysOfWeek $daysOfWeek `
        -At ([datetime]::Today.AddHours($hour))
    $triggers += $t
}

# ── Settings ───────────────────────────────────────────────────────────────
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 15)

# ── Principal: current interactive user, only when logged in ───────────────
$principal = New-ScheduledTaskPrincipal `
    -UserId ([Security.Principal.WindowsIdentity]::GetCurrent().Name) `
    -LogonType Interactive `
    -RunLevel Limited

# ── Register ───────────────────────────────────────────────────────────────
Register-ScheduledTask `
    -TaskName    $TaskName `
    -Description 'Daily ZOHO/ERP Pending Calls auto-merge (see AUTOMATION.md).' `
    -Action      $action `
    -Trigger     $triggers `
    -Settings    $settings `
    -Principal   $principal | Out-Null

Write-Host ""
Write-Host "Done. Task '$TaskName' registered." -ForegroundColor Green
Write-Host ""
Write-Host "Run once now (recommended, to prime today's merge):" -ForegroundColor Cyan
Write-Host "  Start-ScheduledTask -TaskName $TaskName"
Write-Host ""
Write-Host "See current status:" -ForegroundColor Cyan
Write-Host "  Get-ScheduledTaskInfo -TaskName $TaskName"
Write-Host ""
Write-Host "Pause (won't fire again until re-enabled):" -ForegroundColor Cyan
Write-Host "  Disable-ScheduledTask -TaskName $TaskName"
Write-Host ""
Write-Host "Remove:" -ForegroundColor Cyan
Write-Host "  Unregister-ScheduledTask -TaskName $TaskName -Confirm:`$false"
Write-Host ""
