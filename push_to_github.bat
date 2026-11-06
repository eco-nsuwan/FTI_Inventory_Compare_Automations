@echo off
setlocal
title Push FTI Inventory Compare Automations to GitHub

echo ===================================================
echo   Pushing FTI Inventory Compare to GitHub...
echo ===================================================
echo.

cd /d "%~dp0"

git add .
git status --short

set /p COMMIT_MSG="Enter commit message (or press Enter for default): "
if "%COMMIT_MSG%"=="" set COMMIT_MSG=update: latest inventory compare automation improvements

git commit -m "%COMMIT_MSG%"
git push origin main

echo.
echo View repository at: https://github.com/eco-nsuwan/FTI_Inventory_Compare_Automations
pause
