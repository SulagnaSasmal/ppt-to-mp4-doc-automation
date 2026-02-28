@echo off
title PPT to Video - Demo Launcher
color 0A

echo.
echo  ============================================
echo   PPT to Video Demo - Starting up...
echo  ============================================
echo.

REM Check .env file exists
if not exist ".env" (
    echo  ERROR: Missing .env file!
    echo  Copy .env.example to .env and fill in your Azure keys.
    echo.
    pause
    exit /b 1
)

REM Check virtual environment exists
if not exist ".venv\Scripts\python.exe" (
    echo  ERROR: Virtual environment not found!
    echo  See DEMO-GUIDE.md - First Time Setup section.
    echo.
    pause
    exit /b 1
)

echo  [1/3] Starting the app server...
start "PPT Video - App Server" cmd /k "title PPT Video - App Server && color 0B && .venv\Scripts\python.exe -m uvicorn app:app --host 127.0.0.1 --port 8000"

echo  [2/3] Waiting for server to start...
timeout /t 4 /nobreak >nul

echo  [3/3] Starting Cloudflare public tunnel...
start "PPT Video - Public URL" cmd /k "title PPT Video - Public URL && color 0E && echo. && echo  Look for a line that says: && echo  Your quick Tunnel has been created! Visit it at: && echo  https://something.trycloudflare.com && echo. && cloudflared tunnel --url http://localhost:8000"

echo.
echo  ============================================
echo   Two windows just opened:
echo.
echo   BLUE window  = App server (keep it running)
echo   YELLOW window = Your public demo URL
echo.
echo   Look at the YELLOW window for the URL.
echo   It looks like: https://xxxx.trycloudflare.com
echo.
echo   Share that URL for your demo!
echo  ============================================
echo.
echo  Opening local version in your browser...
timeout /t 3 /nobreak >nul
start http://localhost:8000

echo.
echo  Press any key to close this launcher window.
echo  (Keep the BLUE and YELLOW windows open during demo)
pause >nul
