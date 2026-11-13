@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
if exist "%~dp0.venv\Scripts\python.exe" (
  "%~dp0.venv\Scripts\python.exe" "21 TME报表合并.py" %*
) else if exist "%USERPROFILE%\.venv\Scripts\python.exe" (
  "%USERPROFILE%\.venv\Scripts\python.exe" "21 TME报表合并.py" %*
) else (
  py -3 "21 TME报表合并.py" %*
)
pause
endlocal
