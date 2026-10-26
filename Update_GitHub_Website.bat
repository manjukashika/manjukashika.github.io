@echo off
title Push Website Updates to GitHub Pages
cd /d "%~dp0"
echo ========================================================
echo   1. Syncing GoldfishMemory.xlsx to Website...
echo ========================================================
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0sync_excel.ps1"

echo.
echo ========================================================
echo   2. Pushing Changes to GitHub Pages...
echo ========================================================
git add .
git commit -m "Update website and study notes"
git push

echo.
echo ========================================================
echo   Done! Your GitHub Pages site will update in ~1 minute.
echo ========================================================
pause
