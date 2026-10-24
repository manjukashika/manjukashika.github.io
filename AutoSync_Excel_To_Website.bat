@echo off
title Goldfish Memory - Excel to Website Auto-Sync
echo ========================================================
echo   Syncing GoldfishMemory.xlsx to Website...
echo ========================================================
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0sync_excel.ps1" -Watch
