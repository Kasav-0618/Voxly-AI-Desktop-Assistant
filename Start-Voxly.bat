@echo off
title Voxly AI Voice Assistant
cd /d "%~dp0"

echo ========================================================
echo           Starting Voxly AI Voice Assistant
echo ========================================================
echo.

REM Check if server is already responding on port 5173
powershell -NoProfile -Command "try { $r = Invoke-WebRequest -Uri 'http://localhost:5173' -TimeoutSec 1 -UseBasicParsing; exit 0 } catch { exit 1 }" >nul 2>&1
if %ERRORLEVEL% equ 0 (
    echo [OK] Voxly server is already active on http://localhost:5173
    goto launch_app
)

echo [*] Starting Voxly background system bridge and dev server...
start /b "" cmd /c "npm run dev > nul 2>&1"

echo [*] Waiting for server to initialize...
timeout /t 3 /nobreak >nul

:launch_app
echo [*] Launching Voxly in Desktop Voice Mode...

REM 1. Try launching via Microsoft Edge in Native App Mode
if exist "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" (
    start "" "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --app="http://localhost:5173" --start-maximized
    echo [OK] Voxly is now running as a desktop voice application!
    exit /b 0
)

REM 2. Try Google Chrome in Native App Mode
if exist "C:\Program Files\Google\Chrome\Application\chrome.exe" (
    start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" --app="http://localhost:5173" --start-maximized
    echo [OK] Voxly is now running as a desktop voice application!
    exit /b 0
)

REM 3. Fallback to default system browser
start "" "http://localhost:5173"
echo [OK] Voxly launched in default browser!
exit /b 0
