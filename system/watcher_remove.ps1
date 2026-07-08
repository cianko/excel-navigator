# =====================================================================
#  Excel Navigator  -  (c) 2026 Ahmet Zan  -  MIT License
#  watcher_remove.ps1  -  Stops the auto-watcher and removes its auto-start
#  (Task Scheduler task AND/OR Startup-folder shortcut) for this folder.
#  Run from inside the  system  folder:
#    powershell -ExecutionPolicy Bypass -File .\watcher_remove.ps1
# =====================================================================
$here = $PSScriptRoot
if (-not $here) { $here = Split-Path -Parent $MyInvocation.MyCommand.Path }
$stem = ($here -replace '[^A-Za-z0-9]', '_')
if ($stem.Length -gt 40) { $stem = $stem.Substring($stem.Length - 40) }
$taskName = "ExcelNavWatcher_$stem"

# 1) Stop the running watcher
$stopped = 0
Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" |
    Where-Object { $_.CommandLine -like "*excel_watcher.ps1*" } |
    ForEach-Object { try { Stop-Process -Id $_.ProcessId -Force; $stopped++ } catch {} }
Write-Host "Watcher processes stopped: $stopped" -ForegroundColor Cyan

# 2) Remove the Task Scheduler task (if any)
try {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction Stop
    Write-Host "Task removed: $taskName" -ForegroundColor Green
} catch {
    Write-Host "No task to remove ($taskName)." -ForegroundColor DarkGray
}

# 3) Remove the Startup-folder shortcut (if any)
$lnk = Join-Path ([Environment]::GetFolderPath('Startup')) "ExcelNavWatcher_$stem.lnk"
if (Test-Path $lnk) {
    Remove-Item $lnk -Force
    Write-Host "Startup shortcut removed." -ForegroundColor Green
} else {
    Write-Host "No startup shortcut." -ForegroundColor DarkGray
}
