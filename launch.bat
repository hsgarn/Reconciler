@echo off
cd /d "%~dp0"

:: Start Flask in a new background window
start "Bank Reconciler Server" python app.py

:: Wait for the server to be ready
:wait
timeout /t 1 /nobreak >nul
curl -s http://127.0.0.1:5000 >nul 2>&1
if errorlevel 1 goto wait

:: Open Chrome
start "" "chrome.exe" "http://127.0.0.1:5000"
