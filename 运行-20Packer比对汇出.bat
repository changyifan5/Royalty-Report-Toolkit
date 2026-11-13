@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
if exist "%~dp0.venv\Scripts\python.exe" (
  "%~dp0.venv\Scripts\python.exe" "20 Packer比对汇出.py" %*
  goto done
)
if exist "%USERPROFILE%\.venv\Scripts\python.exe" (
  "%USERPROFILE%\.venv\Scripts\python.exe" "20 Packer比对汇出.py" %*
  goto done
)
where py >nul 2>nul
if not errorlevel 1 (
  py -3 "20 Packer比对汇出.py" %*
  goto done
)
where python >nul 2>nul
if not errorlevel 1 (
  python "20 Packer比对汇出.py" %*
  goto done
)
echo 请先安装 Python 3.10 或以上版本，并安装 requirements.txt 中的依赖。
:done
echo.
pause
endlocal
