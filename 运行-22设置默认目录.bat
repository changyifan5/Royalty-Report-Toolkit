@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
if exist "%~dp0.venv\Scripts\python.exe" (
  "%~dp0.venv\Scripts\python.exe" "22 设置默认目录.py"
) else if exist "%USERPROFILE%\.venv\Scripts\python.exe" (
  "%USERPROFILE%\.venv\Scripts\python.exe" "22 设置默认目录.py"
) else (
  py -3 "22 设置默认目录.py"
)
pause
endlocal
