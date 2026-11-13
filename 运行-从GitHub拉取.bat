@echo off
chcp 65001 >nul
cd /d "%~dp0"
title 从 GitHub 拉取

echo === 从 GitHub 拉取最新代码 ===
echo.
git status -sb
echo.

for /f "delims=" %%b in ('git branch --show-current 2^>nul') do set "BRANCH=%%b"
if not defined BRANCH set "BRANCH=master"

git diff --quiet
set "DIRTY=%ERRORLEVEL%"
git diff --cached --quiet
set /a STAGED=%ERRORLEVEL%

if %DIRTY% neq 0 (
    set "HAS_CHANGES=1"
) else if %STAGED% neq 0 (
    set "HAS_CHANGES=1"
) else (
    git ls-files --others --exclude-standard | findstr /r "." >nul
    if not errorlevel 1 set "HAS_CHANGES=1"
)

if defined HAS_CHANGES (
    echo 检测到本地有未提交改动，直接拉取会失败。
    echo.
    echo   1 = 先暂存本地改动，拉取后再恢复（推荐）
    echo   2 = 取消
    echo.
    set "CHOICE="
    set /p CHOICE=请选择 [1/2，直接回车=1]: 
    if "%CHOICE%"=="" set "CHOICE=1"
    if /i "%CHOICE%"=="2" (
        echo 已取消。
        pause
        exit /b 0
    )
    echo.
    echo 正在暂存本地改动...
    git stash push -u -m "拉取前自动暂存"
    if errorlevel 1 (
        echo 暂存失败，请手动处理后再试。
        pause
        exit /b 1
    )
    set "DID_STASH=1"
)

echo 正在拉取 origin %BRANCH% ...
echo.
git pull origin %BRANCH%
if errorlevel 1 (
    echo.
    echo 拉取失败。
    if defined DID_STASH (
        echo 正在恢复刚才暂存的本地改动...
        git stash pop
    )
    pause
    exit /b 1
)

if defined DID_STASH (
    echo.
    echo 正在恢复本地改动...
    git stash pop
    if errorlevel 1 (
        echo.
        echo 恢复时出现冲突，请手动解决后执行：git stash drop
        pause
        exit /b 1
    )
)

echo.
echo 拉取完成，当前状态：
git status -sb
echo.
pause
