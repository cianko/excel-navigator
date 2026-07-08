' Excel Navigator - (c) 2026 Ahmet Zan - MIT License
' launcher.vbs - starts excel_watcher.ps1 INVISIBLY (no console window).
' Runs the excel_watcher.ps1 next to it; portable.
Dim fso, sFolder, sPs1
Set fso = CreateObject("Scripting.FileSystemObject")
sFolder = fso.GetParentFolderName(WScript.ScriptFullName)
sPs1 = sFolder & "\excel_watcher.ps1"
CreateObject("WScript.Shell").Run _
  "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & sPs1 & """", 0, False
