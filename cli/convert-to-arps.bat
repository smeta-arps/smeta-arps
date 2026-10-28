@echo off
rem Drag a .docx / .xlsx / .csv estimate onto this file: the .arp for Estimate appears next to it.
set PYTHONUTF8=1
chcp 65001 >nul
if "%~1"=="" (
  echo Drag a .docx or .xlsx estimate onto this file.
  pause
  exit /b 1
)
python -P "%~dp0excel2arps.py" "%~1" "%~dpn1.arp"
echo.
pause
