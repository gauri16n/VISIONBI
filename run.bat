@echo off
setlocal
title VisionBI - AI Data Analyst
cd /d "%~dp0"

echo ==============================================
echo   VisionBI - AI Data Analyst
echo ==============================================
echo.

REM --- Check Node.js is installed ---
where node >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Node.js was not found.
    echo         Download it from https://nodejs.org and re-run this file.
    echo.
    pause
    exit /b 1
)

REM --- Install dependencies the first time ---
if not exist node_modules (
    echo [SETUP] First run detected - installing dependencies...
    echo         This can take a minute or two.
    echo.
    call npm install
    if errorlevel 1 (
        echo [ERROR] npm install failed. Check your internet connection.
        echo.
        pause
        exit /b 1
    )
    echo.
)

REM --- Create .env from template if it does not exist ---
if not exist .env (
    if exist .env.example (
        copy /y .env.example .env >nul
        echo [INFO] Created .env from .env.example
        echo        Open .env and add your Anthropic API key to enable live AI.
        echo.
    )
)

echo [START] Opening VisionBI in your browser...
echo [START] Press Ctrl+C in this window to stop the server.
echo.
start "" "http://localhost:5173"
call npm run dev

echo.
echo Server stopped.
pause
