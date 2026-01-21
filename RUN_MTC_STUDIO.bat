@echo off
echo ==========================================
echo   MTC Report Studio - Startup Script
echo ==========================================

echo [1/2] Starting Backend Server...
start "MTC Backend" cmd /k "cd backend && python main.py"

echo [2/2] Starting Frontend Dashboard...
echo IMPORTANT: If the windows look stuck, click inside them and press ESCAPE.
echo Windows terminals sometimes pause if you click them by mistake.
echo.
start "MTC Frontend" cmd /k "cd frontend && npm run dev"

echo.
echo Both servers are starting in separate windows.
echo PLEASE DO NOT CLOSE THE NEW WINDOWS while using the app.
echo.
echo Opening your browser to: http://localhost:5173
echo.
start http://localhost:5173
echo.
echo If the page doesnt load, check the "MTC Frontend" window. 
echo If it says 'Select' in the title, press ESCAPE to resume.
echo.
pause
