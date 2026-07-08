# =====================================================================
#  Excel Navigator  -  (c) 2026 Ahmet Zan  -  MIT License
#  excel_watcher.ps1  -  Auto .xlsx -> .xlsm + navigation injection
#  Lives in the  system  subfolder and watches its PARENT folder (the
#  project folder where your Excel files are). For every .xlsx dropped
#  there it:
#    1) converts it to .xlsm (52 = macro-enabled)
#    2) deletes the original .xlsx
#    3) injects the navigation system (module + panel + events) by
#       calling the kur.ps1 next to it (-File mode)
#  So a file you drop into the project folder joins the system by itself.
#
#  NOTE: event-based paths (Register-ObjectEvent / WaitForChanged) did not
#  fire in some launch contexts, so a simple 2-second POLLING scan is used
#  (practically 0% CPU for a small folder).
# =====================================================================
$scriptDir = $PSScriptRoot
if (-not $scriptDir) { $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
$folder  = Split-Path -Parent $scriptDir        # watched folder = PARENT (project / Excel folder)
$kurYolu = Join-Path $scriptDir "kur.ps1"        # injection: the kur.ps1 next to this script

# Processes one .xlsx file (convert + delete original + inject).
function Process-File {
    param([string]$path, [string]$name)

    # 1) Ignore Excel temp files (start with ~$)
    if ($name -like "~`$*") { return }

    # 2) Ignore the default empty file from right-click "New > Excel Worksheet"
    #    (once renamed it is processed under the new name)
    if ($name -like "Yeni Microsoft Excel*" -or $name -like "New Microsoft Excel Worksheet*") { return }

    if (-not (Test-Path $path)) { return }

    # Don't touch it until writing settles: last change must be >800ms ago
    try {
        $fi = Get-Item -LiteralPath $path -ErrorAction Stop
        if (((Get-Date) - $fi.LastWriteTime).TotalMilliseconds -lt 800) { return }
    } catch { return }

    # Lock check: if still open/being written by another process, skip this round
    try {
        $fsCheck = [System.IO.File]::Open($path, 'Open', 'Read', 'None')
        $fsCheck.Close()
    } catch { return }

    # Convert with Excel COM
    $excel = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Open($path)
        $newPath = [System.IO.Path]::ChangeExtension($path, ".xlsm")

        $workbook.SaveAs($newPath, 52)   # 52 = xlOpenXMLWorkbookMacroEnabled
        $workbook.Close($false)
        $excel.Quit()

        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        $excel = $null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()

        if (Test-Path $path) { Remove-Item $path -Force }

        # Inject the navigation system into the new .xlsm
        if (Test-Path $kurYolu) {
            Start-Process powershell.exe -WindowStyle Hidden -Wait -ArgumentList @(
                '-ExecutionPolicy','Bypass','-File',$kurYolu,'-File',$newPath
            )
        }
    } catch {
        if ($excel -ne $null) {
            try { $excel.Quit() } catch {}
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
        [System.GC]::Collect()
    }
}

# Main loop: every 2 seconds scan the folder and process any pending .xlsx.
while ($true) {
    Start-Sleep -Seconds 2
    foreach ($f in @(Get-ChildItem -LiteralPath $folder -Filter *.xlsx -File -ErrorAction SilentlyContinue)) {
        Process-File $f.FullName $f.Name
    }
}
