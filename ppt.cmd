@echo off
rem PPT workspace entry. Usage:  ppt build templates\academic-defense.md -Open
rem Runs scripts\ppt.ps1 with ExecutionPolicy Bypass so it works on locked-down machines.
setlocal
set "PS=powershell"
where pwsh >nul 2>nul && set "PS=pwsh"
"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\ppt.ps1" %*
