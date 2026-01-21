@echo off
echo Starting MTC Report Studio V2 (Modern Architecture)...

:: Check for backend port availability or just start
echo [1/2] Starting Backend (FastAPI)...
start /B cmd /c "cd backend && python main.py"

:: Check for frontend port
echo [2/2] Starting Frontend (React/Vite)...
cd frontend
npm run dev

pause
