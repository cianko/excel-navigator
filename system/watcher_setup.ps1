# =====================================================================
#  Excel Navigator  -  (c) 2026 Ahmet Zan  -  MIT License
#  watcher_setup.ps1  -  Starts the auto-watcher in the BACKGROUND and
#  registers it to start automatically at every logon.
#  For auto-start it FIRST tries Task Scheduler; if that needs admin
#  rights (Access denied) it falls back to the STARTUP FOLDER method
#  (no admin needed, always works).
#
#  Run this from inside the  system  folder:
#    powershell -ExecutionPolicy Bypass -File .\watcher_setup.ps1
# =====================================================================
$ErrorActionPreference = "Stop"
$here = $PSScriptRoot
if (-not $here) { $here = Split-Path -Parent $MyInvocation.MyCommand.Path }

$vbs = Join-Path $here "launcher.vbs"
$ps1 = Join-Path $here "excel_watcher.ps1"

# Are the required files next to this script?
foreach ($f in @($vbs, $ps1, (Join-Path $here "kur.ps1"), (Join-Path $here "modNavigasyon.bas"))) {
    if (-not (Test-Path $f)) {
        Write-Host "ERROR: $([System.IO.Path]::GetFileName($f)) is not in this folder." -ForegroundColor Red
        Write-Host "kur.ps1, modNavigasyon.bas, excel_watcher.ps1, launcher.vbs must be together." -ForegroundColor Yellow
        exit 1
    }
}

$stem = ($here -replace '[^A-Za-z0-9]', '_')
if ($stem.Length -gt 40) { $stem = $stem.Substring($stem.Length - 40) }
$taskName = "ExcelNavWatcher_$stem"

# Stop any watcher already running for this folder (avoid duplicates)
Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" |
    Where-Object { $_.CommandLine -like "*excel_watcher.ps1*" } |
    ForEach-Object { try { Stop-Process -Id $_.ProcessId -Force } catch {} }

# --- Auto-start: try Task Scheduler first, else the Startup folder ---
$method = ""
try {
    $action    = New-ScheduledTaskAction -Execute "wscript.exe" -Argument ('"{0}"' -f $vbs)
    $trigger   = New-ScheduledTaskTrigger -AtLogOn
    $settings  = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -ExecutionTimeLimit ([TimeSpan]::Zero)
    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Force -ErrorAction Stop | Out-Null
    $method = "Task Scheduler"
} catch {
    # Access denied etc. -> Startup-folder shortcut (no admin needed)
    try {
        $startup = [Environment]::GetFolderPath('Startup')
        $lnk = Join-Path $startup "ExcelNavWatcher_$stem.lnk"
        $wsh = New-Object -ComObject WScript.Shell
        $sc = $wsh.CreateShortcut($lnk)
        $sc.TargetPath       = "wscript.exe"
        $sc.Arguments        = '"' + $vbs + '"'
        $sc.WorkingDirectory = $here
        $sc.WindowStyle      = 7
        $sc.Description       = "Excel Navigation auto-watcher"
        $sc.Save()
        $method = "Startup folder"
    } catch {
        $method = ""
    }
}

$project = Split-Path -Parent $here    # watched folder = project folder (parent of 'system')

if ($method) {
    Write-Host "Auto-start installed (method: $method). It will start at every logon." -ForegroundColor Green
} else {
    Write-Host "WARNING: Auto-start could not be installed. The watcher runs now, but at" -ForegroundColor Yellow
    Write-Host "logon it won't start itself; double-click launcher.vbs to start it." -ForegroundColor Yellow
}

# --- Start it now (in the background, invisible) ---
Start-Process "wscript.exe" -ArgumentList ('"{0}"' -f $vbs)
Start-Sleep -Seconds 2
$running = Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" |
    Where-Object { $_.CommandLine -like "*excel_watcher.ps1*" }
if ($running) {
    Write-Host "Watcher is RUNNING. It watches the project folder:" -ForegroundColor Cyan
    Write-Host "   $project" -ForegroundColor Cyan
    Write-Host "Any .xlsx dropped there is converted to .xlsm + injected within ~2 s." -ForegroundColor Cyan
} else {
    Write-Host "Watcher does not seem to have started; try double-clicking launcher.vbs." -ForegroundColor Yellow
}
Write-Host ""
Write-Host "To stop/remove: STOP_WATCHER.cmd (or watcher_remove.ps1)" -ForegroundColor DarkGray
Write-Host "IMPORTANT: For the panel to appear in files, make the PROJECT folder" -ForegroundColor DarkGray
Write-Host "a Trusted Location in Excel (see HOW_TO_INSTALL.txt step 1B):" -ForegroundColor DarkGray
Write-Host "   $project" -ForegroundColor DarkGray
