@echo off
REM =====================================================================
REM  STOP_WATCHER.cmd  -  Double-click to turn OFF automatic conversion.
REM  Stops the running watcher and removes its auto-start (Task Scheduler
REM  task and/or Startup-folder shortcut) for THIS folder.
REM =====================================================================
setlocal
set "HERE=%~dp0"
set "HERE=%HERE:~0,-1%"

echo Stopping the auto-watcher...
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%HERE%\watcher_remove.ps1"

echo.
echo Done. You can close this window.
pause
