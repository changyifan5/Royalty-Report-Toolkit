@echo off
chcp 65001 >nul
cd /d "%~dp0"
title 同步到 GitHub

echo === 同步到 GitHub ===
echo.
git status -sb
echo.

set "MSG="
set /p MSG=这次改了什么？（直接回车 =「更新脚本」）: 
if "%MSG%"=="" set "MSG=更新脚本"

git pull origin master
git add -A
git commit -m "%MSG%"
if errorlevel 1 (
    echo.
    echo 没有新改动，尝试直接推送已有提交…
)

git push origin master
echo.
echo 完成。
pause
