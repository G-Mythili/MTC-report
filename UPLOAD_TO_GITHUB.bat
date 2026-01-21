@echo off
echo ========================================
echo   UPLOADING MTC REPORT TO GITHUB
echo ========================================
echo.

cd /d "e:\2026 Report Famat"

:: Initialize Git if not already done
if not exist ".git" (
    echo [1/5] Initializing Git...
    git init
) else (
    echo [1/5] Git already initialized.
)

:: Add all files
echo [2/5] Adding files...
git add .

:: Commit
echo [3/5] Committing changes...
git commit -m "Initialize MTC Report Studio for Codespaces"

:: Rename branch to main
git branch -M main

:: Add remote (ignore error if already exists)
echo [4/5] Connecting to GitHub...
git remote add origin https://github.com/G-Mythili/MTC-report.git 2>nul
git remote set-url origin https://github.com/G-Mythili/MTC-report.git

:: Push
echo [5/5] Pushing to GitHub...
echo.
echo IMPORTANT: If a login window appears, please sign in.
echo.
git push -u origin main

echo.
echo ========================================
echo   UPLOAD COMPLETE! 
echo   Check your GitHub page now.
echo ========================================
pause
