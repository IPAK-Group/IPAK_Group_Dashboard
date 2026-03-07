@echo off
title KPI Dashboard - Update & Deploy
color 0A

echo ================================================
echo   KPI DASHBOARD - UPDATE AND PUBLISH
echo ================================================
echo.

REM ── Step 1: Convert Excel data to data.js ────────
echo [1/3] Converting Excel data...
python convert_data.py
if %ERRORLEVEL% NEQ 0 (
    color 0C
    echo.
    echo ERROR: Data conversion failed. Check your Excel file.
    echo Nothing was published. Please fix the error and try again.
    pause
    exit /b 1
)
echo       Done!
echo.

REM ── Step 2: Stage all changes ────────────────────
echo [2/3] Preparing files for upload...
git add dashboard_clean.html data.js
if %ERRORLEVEL% NEQ 0 (
    color 0C
    echo ERROR: Git staging failed. Is Git installed?
    pause
    exit /b 1
)
echo       Done!
echo.

REM ── Step 3: Commit and push ───────────────────────
echo [3/3] Publishing to GitHub...
REM Get current date/time for the commit message
for /f "tokens=1-3 delims=/ " %%a in ("%date%") do set TODAY=%%c-%%b-%%a
for /f "tokens=1-2 delims=: " %%a in ("%time%") do set NOW=%%a:%%b
git commit -m "Dashboard update: %TODAY% %NOW%"
git push origin main
if %ERRORLEVEL% NEQ 0 (
    color 0C
    echo.
    echo ERROR: Failed to push to GitHub.
    echo Make sure you are connected to the internet.
    pause
    exit /b 1
)

echo.
echo ================================================
color 0B
echo   SUCCESS! Dashboard published!
echo.
echo   Your live link is ready. Data will update
echo   on the shared link within 1-2 minutes.
echo ================================================
echo.
pause
