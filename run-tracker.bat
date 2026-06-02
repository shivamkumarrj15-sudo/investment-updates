@echo off
title Investment Tracker Automation Bot
cd /d "%~dp0"
echo ===================================================
echo 📈 Starting Weekly Investment Tracker Automation
echo ===================================================
node tracker.js
echo ===================================================
echo Automation process finished.
echo ===================================================
pause
