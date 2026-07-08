@echo off
REM =====================================================================
REM  Excel Navigator  -  (c) 2026 Ahmet Zan  -  MIT License
REM  INSTALL.cmd  -  ONE-CLICK setup. Double-click this file to:
REM    1) install the navigation system into every Excel file here, and
REM    2) turn ON the auto-watcher, so any .xlsx you drop into this folder
REM       later is converted to .xlsm + set up automatically.
REM
REM  HOW TO USE:
REM    1) Copy  INSTALL.cmd  AND the whole  system  folder next to it into
REM       the SAME folder as your Excel files.
REM    2) Double-click INSTALL.cmd. That's it - no typing.
REM
REM  (Two one-time Excel settings are still needed on each PC - see
REM   system\HOW_TO_INSTALL.txt: enable VBA access, and make this folder
REM   a Trusted Location so the panel shows.)
REM =====================================================================
setlocal
set "HERE=%~dp0"
set "HERE=%HERE:~0,-1%"

if not exist "%HERE%\system\kur.ps1" (
    echo ERROR: "system\kur.ps1" not found next to INSTALL.cmd.
    echo Copy INSTALL.cmd together with the whole "system" folder.
    pause
    exit /b 1
)

echo ============================================================
echo   Excel Navigation - one-click setup
echo   Folder: %HERE%
echo ============================================================
echo.

echo [1/3] Stopping any running watcher (to avoid conflicts)...
powershell -NoProfile -ExecutionPolicy Bypass -File "%HERE%\system\watcher_remove.ps1"
echo.

echo [2/3] Installing the navigation system into your Excel files...
powershell -NoProfile -ExecutionPolicy Bypass -File "%HERE%\system\kur.ps1" -Folder "%HERE%" -Xlsx
echo.

echo [3/3] Turning on the auto-watcher...
powershell -NoProfile -ExecutionPolicy Bypass -File "%HERE%\system\watcher_setup.ps1"
echo.

echo ============================================================
echo   Done. New .xlsx files you drop into this folder will be
echo   converted and set up automatically.
echo.
echo   If the panel does NOT show when you open a file, make this
echo   folder a Trusted Location in Excel
echo   (see  system\HOW_TO_INSTALL.txt , step 1B).
echo ============================================================
echo.
pause
