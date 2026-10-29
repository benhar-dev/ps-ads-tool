@echo off
rem Mini ADS Tool - console mode. Extra arguments are passed through, e.g. AdsTool-Console.bat -Port 801
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0AdsTool.ps1" -Cli %*
