@echo off
REM =====================================================================
REM  START_WATCHER.cmd  (inside the "system" folder)
REM  Double-click to turn ON automatic conversion. After this, any .xlsx
REM  you drop into the PROJECT folder (the folder that contains this
REM  "system" folder) is auto-converted to .xlsm and gets the navigation
REM  system within ~2 seconds. It also starts itself at every logon.
REM
REM  Requires INSTALL.cmd + this whole "system" folder to be inside your
REM  project folder (next to your Excel files).
REM =====================================================================
setlocal
set "HERE=%~dp0"
set "HERE=%HERE:~0,-1%"

echo Starting the auto-watcher...
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%HERE%\watcher_setup.ps1"

echo.
echo Done. You can close this window.
pause
