@echo off
rem Mini ADS Tool - GUI. Extra arguments are passed through, e.g. AdsTool.bat -NetId 5.12.34.56.1.1
start "" powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0AdsTool.ps1" %*
