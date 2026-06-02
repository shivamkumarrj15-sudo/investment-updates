@echo off
title Stock Portfolio Manager Portal
cd /d "%~dp0"
echo ===================================================
echo 🖥️ Starting Local Stock Portfolio Manager Server
echo ===================================================
echo Opening web interface...
start "" "http://localhost:3000"
node server.js
echo ===================================================
echo Web server stopped.
echo ===================================================
pause
