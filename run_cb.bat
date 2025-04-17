@echo off
cd /d "%~dp0"

:: Suppress all background errors (stderr), but keep main program output (stdout)
python main.py 2>nul

:: Optional: pause to keep window open
echo.
echo ✅ Script finished. Press any key to exit...
pause >nul
