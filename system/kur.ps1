# =====================================================================
# Excel Navigator  -  (c) 2026 Ahmet Zan  -  MIT License
# kur.ps1 - installs the floating navigation panel into Excel files.
# Into every .xlsm it injects:
#   1) modNavigasyon (standard module) - from the .bas file
#   2) frmNav (UserForm) - built programmatically via VBIDE
#   3) ThisWorkbook events (Open / SheetActivate / SheetSelectionChange)
#   4) Removes old worksheet nav shapes and unfreezes panes.
# Idempotent. REQUIRES: "Trust access to the VBA project object model" ON.
#
# USAGE:  .\kur.ps1 -Folder "C:\your\project"            (whole folder)
#         .\kur.ps1 -File   "C:\your\project\one.xlsm"    (single file)
# =====================================================================
param(
    [string]$Folder,          # folder mode: install into every .xlsm
    [string]$File,            # single-file mode: install into just this .xlsm (the watcher calls this)
    [switch]$Xlsx
)
if (-not $Folder -and -not $File) {
    Write-Host "ERROR: pass -Folder <folder> or -File <file.xlsm>." -ForegroundColor Red
    exit 1
}
$ErrorActionPreference = "Stop"
$MODULE_NAME = "modNavigasyon"
$FORM_NAME   = "frmNav"
$XL_MACRO_FMT = 52
$BAS_PATH = Join-Path $PSScriptRoot "modNavigasyon.bas"

# VBIDE constants
$ct_StdModule = 1; $ct_MSForm = 3
$pk_Proc = 0

$formCode = @'
Private Sub UserForm_Initialize()
    On Error Resume Next
    Me.Caption = "Navigation"
    Me.Width = 156
    Me.Height = 146
End Sub
Private Sub btnGeri_Click()
    Application.OnTime Now, "NavGeri"
End Sub
Private Sub btnIleri_Click()
    Application.OnTime Now, "NavIleri"
End Sub
Private Sub btnLink_Click()
    NavBaglantiKur
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    NavPanelKapat
End Sub
'@

$sheetFormCode = @'
Private Sub UserForm_Initialize()
    On Error Resume Next
    Me.Caption = "Select sheet"
    Me.Width = 204
    Me.Height = 226
End Sub
Private Sub btnOK_Click()
    If Me.lstSheets.ListIndex >= 0 Then modNavigasyon.gSecilenSayfa = CStr(Me.lstSheets.Value)
    Me.Hide
End Sub
Private Sub btnIptal_Click()
    modNavigasyon.gSecilenSayfa = ""
    Me.Hide
End Sub
Private Sub lstSheets_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    btnOK_Click
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    modNavigasyon.gSecilenSayfa = ""
End Sub
'@

$twbCode = @'
Private Sub Workbook_Open()
    NavInit
End Sub
Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    NavPathGuncelle
    NavKonumUygula
End Sub
Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    NavPathGuncelle
    NavKonumIzle
End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    NavPanelUnload
End Sub
'@

$geriCap  = "$([char]0x25C0) Back"
$ileriCap = "Forward $([char]0x25B6)"

if (-not (Test-Path $BAS_PATH)) { Write-Host "ERROR: $BAS_PATH not found." -ForegroundColor Red; exit 1 }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false    # don't trigger Workbook_Open during install

function ProcSil($cm, $ad) {
    try {
        $s = $cm.ProcStartLine($ad, $pk_Proc)
        $c = $cm.ProcCountLines($ad, $pk_Proc)
        if ($c -gt 0) { $cm.DeleteLines($s, $c) }
    } catch {}
}

function KurDosya($wb) {
    $script:kurAdim = "VBProject access"
    $vb = $wb.VBProject

    # --- 1) modNavigasyon module ---
    $script:kurAdim = "import module (modNavigasyon)"
    try { $vb.VBComponents.Remove($vb.VBComponents.Item($MODULE_NAME)) } catch {}
    $vb.VBComponents.Import($BAS_PATH) | Out-Null

    # --- 2) frmNav UserForm (reuse if present; Remove+Add errors in MSForms) ---
    $script:kurAdim = "build frmNav"
    $uf = $null
    try { $uf = $vb.VBComponents.Item($FORM_NAME) } catch {}
    if ($null -eq $uf) {
        $uf = $vb.VBComponents.Add($ct_MSForm)
        $uf.Name = $FORM_NAME
    }
    try { $uf.Properties.Item("Caption").Value = "Navigation" } catch {}
    try { $uf.Properties.Item("Width").Value  = 156 } catch {}
    try { $uf.Properties.Item("Height").Value = 146 } catch {}
    try { $uf.Properties.Item("StartUpPosition").Value = 0 } catch {}

    $des = $uf.Designer
    $g = $null; try { $g = $des.Controls.Item("btnGeri") } catch {}
    if ($null -eq $g) { $g = $des.Controls.Add("Forms.CommandButton.1"); $g.Name = "btnGeri" }
    $g.Caption = $geriCap; $g.Left = 6; $g.Top = 6; $g.Width = 66; $g.Height = 22
    $i = $null; try { $i = $des.Controls.Item("btnIleri") } catch {}
    if ($null -eq $i) { $i = $des.Controls.Add("Forms.CommandButton.1"); $i.Name = "btnIleri" }
    $i.Caption = $ileriCap; $i.Left = 78; $i.Top = 6; $i.Width = 66; $i.Height = 22
    $l = $null; try { $l = $des.Controls.Item("lblPath") } catch {}
    if ($null -eq $l) { $l = $des.Controls.Add("Forms.Label.1"); $l.Name = "lblPath" }
    $l.Caption = ""; $l.Left = 6; $l.Top = 32; $l.Width = 144; $l.Height = 40
    try { $l.Font.Size = 10 } catch {}
    try { $l.WordWrap = $true } catch {}
    $lk = $null; try { $lk = $des.Controls.Item("btnLink") } catch {}
    if ($null -eq $lk) { $lk = $des.Controls.Add("Forms.CommandButton.1"); $lk.Name = "btnLink" }
    $lk.Caption = "Link Cell"; $lk.Left = 6; $lk.Top = 74; $lk.Width = 138; $lk.Height = 20
    # Remove the old "Scan for New Files" button if left over from an older version
    try { $des.Controls.Remove("btnScan") } catch {}
    # "Developed by Ahmet Zan" credit line at the bottom of the panel
    $dv = $null; try { $dv = $des.Controls.Item("lblDev") } catch {}
    if ($null -eq $dv) { $dv = $des.Controls.Add("Forms.Label.1"); $dv.Name = "lblDev" }
    $dv.Caption = "Developed by Ahmet Zan"; $dv.Left = 6; $dv.Top = 100; $dv.Width = 138; $dv.Height = 12
    try { $dv.Font.Size = 7 } catch {}
    try { $dv.ForeColor = 8421504 } catch {}   # gray
    try { $dv.TextAlign = 2 } catch {}          # center

    $cm2 = $uf.CodeModule
    if ($cm2.CountOfLines -gt 0) { $cm2.DeleteLines(1, $cm2.CountOfLines) }
    $cm2.AddFromString($formCode)

    # --- 2b) frmSheetSec (sheet picker used by Link Cell) ---
    $script:kurAdim = "build frmSheetSec"
    $uf2 = $null
    try { $uf2 = $vb.VBComponents.Item("frmSheetSec") } catch {}
    if ($null -eq $uf2) { $uf2 = $vb.VBComponents.Add($ct_MSForm); $uf2.Name = "frmSheetSec" }
    try { $uf2.Properties.Item("Caption").Value = "Select sheet" } catch {}
    try { $uf2.Properties.Item("Width").Value = 204 } catch {}
    try { $uf2.Properties.Item("Height").Value = 224 } catch {}
    try { $uf2.Properties.Item("StartUpPosition").Value = 1 } catch {}
    $des2 = $uf2.Designer
    $lbl2 = $null; try { $lbl2 = $des2.Controls.Item("lblInfo") } catch {}
    if ($null -eq $lbl2) { $lbl2 = $des2.Controls.Add("Forms.Label.1"); $lbl2.Name = "lblInfo" }
    $lbl2.Caption = "Select target sheet:"; $lbl2.Left = 8; $lbl2.Top = 6; $lbl2.Width = 184; $lbl2.Height = 14
    $lst = $null; try { $lst = $des2.Controls.Item("lstSheets") } catch {}
    if ($null -eq $lst) { $lst = $des2.Controls.Add("Forms.ListBox.1"); $lst.Name = "lstSheets" }
    $lst.Left = 8; $lst.Top = 24; $lst.Width = 184; $lst.Height = 140
    $ok = $null; try { $ok = $des2.Controls.Item("btnOK") } catch {}
    if ($null -eq $ok) { $ok = $des2.Controls.Add("Forms.CommandButton.1"); $ok.Name = "btnOK" }
    $ok.Caption = "OK"; $ok.Left = 32; $ok.Top = 172; $ok.Width = 60; $ok.Height = 22
    $cn = $null; try { $cn = $des2.Controls.Item("btnIptal") } catch {}
    if ($null -eq $cn) { $cn = $des2.Controls.Add("Forms.CommandButton.1"); $cn.Name = "btnIptal" }
    $cn.Caption = "Cancel"; $cn.Left = 108; $cn.Top = 172; $cn.Width = 60; $cn.Height = 22
    $cm3 = $uf2.CodeModule
    if ($cm3.CountOfLines -gt 0) { $cm3.DeleteLines(1, $cm3.CountOfLines) }
    $cm3.AddFromString($sheetFormCode)

    # --- 3) ThisWorkbook events (locale-safe: try "ThisWorkbook", then CodeName) ---
    $script:kurAdim = "ThisWorkbook events"
    $twb = $null
    try { $twb = $vb.VBComponents.Item("ThisWorkbook") } catch {}
    if ($null -eq $twb) { try { $twb = $vb.VBComponents.Item($wb.CodeName) } catch {} }
    if ($null -eq $twb) {
        # last resort: find the Document-type (100) component that is not a Worksheet
        foreach ($c in $vb.VBComponents) {
            if ($c.Type -eq 100 -and $c.Name -notlike "Sheet*" -and $c.Name -notlike "Sayfa*") { $twb = $c; break }
        }
    }
    if ($null -ne $twb) {
        $cm = $twb.CodeModule
        ProcSil $cm "Workbook_Open"
        ProcSil $cm "Workbook_SheetActivate"
        ProcSil $cm "Workbook_SheetSelectionChange"
        ProcSil $cm "Workbook_BeforeClose"
        $cm.AddFromString($twbCode)
    }

    # --- 4) Clean up old worksheet nav shapes + unfreeze (NON-critical -> never breaks injection) ---
    $script:kurAdim = "cleanup"
    try {
        $win = $null; try { $win = $wb.Windows.Item(1) } catch {}
        foreach ($ws in $wb.Worksheets) {
            try {
                foreach ($sh in @($ws.Shapes)) {
                    $n = $sh.Name
                    if ($n -eq "btnGeri" -or $n -eq "btnIleri" -or $n -eq "lblPath" -or $n -like "nav|*") { $sh.Delete() }
                }
            } catch {}
            try { $ws.Activate() } catch {}
            try { if ($win -ne $null -and $win.FreezePanes) { $win.FreezePanes = $false } } catch {}
        }
        try { $wb.Worksheets.Item(1).Activate() } catch {}
    } catch {}
}

try {
    if ($File) {
        # --- SINGLE-FILE MODE (the watcher calls this when a new .xlsm appears) ---
        $d = Get-Item -LiteralPath $File
        try {
            $wb = $excel.Workbooks.Open($d.FullName)
            try { $null = $wb.VBProject } catch {
                $wb.Close($false)
                throw "Cannot access the VBA project. 'Trust access to the VBA project object model' must be ON."
            }
            KurDosya $wb
            $wb.Save(); $wb.Close($true)
            Write-Host "OK: $($d.Name)" -ForegroundColor Green
        } catch {
            Write-Host "ERROR $($d.Name) [step: $script:kurAdim]: $($_.Exception.Message)" -ForegroundColor Red
            if ($wb) { try { $wb.Close($false) } catch {} }
        }
    }
    else {
        if ($Xlsx) {
            foreach ($x in (Get-ChildItem -Path $Folder -Recurse -Filter *.xlsx)) {
                $yeni = [System.IO.Path]::ChangeExtension($x.FullName, ".xlsm")
                try {
                    if (-not (Test-Path $yeni)) {
                        $wb = $excel.Workbooks.Open($x.FullName)
                        $wb.SaveAs($yeni, $XL_MACRO_FMT); $wb.Close($false)
                        Write-Host "CONVERTED -> $([System.IO.Path]::GetFileName($yeni))" -ForegroundColor Green
                    }
                    # .xlsm ready -> delete the original .xlsx (keep the folder tidy)
                    if (Test-Path $yeni) {
                        Remove-Item -LiteralPath $x.FullName -Force
                        Write-Host "DELETED (original .xlsx) -> $($x.Name)" -ForegroundColor DarkYellow
                    }
                } catch {
                    Write-Host "CONVERT ERROR $($x.Name): $($_.Exception.Message)" -ForegroundColor Red
                    Write-Host "  -> Fix: (1) CLOSE the file/Excel. (2) If the auto-watcher runs in this" -ForegroundColor Yellow
                    Write-Host "     folder, stop it (STOP_WATCHER.cmd) - it CONFLICTS with INSTALL." -ForegroundColor Yellow
                    Write-Host "     (3) If OneDrive is 'online-only': right-click > 'Always keep on this device'." -ForegroundColor Yellow
                }
            }
        }

        foreach ($d in (Get-ChildItem -Path $Folder -Recurse -Filter *.xlsm)) {
            try {
                $wb = $excel.Workbooks.Open($d.FullName)
                try { $null = $wb.VBProject } catch {
                    $wb.Close($false)
                    throw "Cannot access the VBA project. 'Trust access to the VBA project object model' must be ON."
                }
                KurDosya $wb
                $wb.Save(); $wb.Close($true)
                Write-Host "OK: $($d.Name)" -ForegroundColor Green
            } catch {
                Write-Host "ERROR $($d.Name) [step: $script:kurAdim]: $($_.Exception.Message)" -ForegroundColor Red
                if ($wb) { try { $wb.Close($false) } catch {} }
            }
        }
    }
}
finally {
    $excel.EnableEvents = $true
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    Write-Host "Finished." -ForegroundColor Cyan
}
