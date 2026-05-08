@echo off
echo.
echo  Push Code Puppy Saved Work to GitHub
echo  ====================================
echo.
echo  IMPORTANT: Close Excel, PowerPoint, and Word files in this folder first.
echo.
cd /d "%~dp0"

echo Current repo:
git remote -v
echo.

echo [1/4] Checking status...
git status --short
echo.

echo [2/4] Adding files...
git add .
if errorlevel 1 (
    echo.
    echo ERROR: git add failed. A file is probably open/locked by Office or OneDrive.
    echo Close any open Excel/PowerPoint/Word files and try again.
    pause
    exit /b 1
)

echo.
echo [3/4] Committing changes...
git commit -m "Update NextGen PMO saved work"
if errorlevel 1 (
    echo No new changes to commit, or commit failed. Continuing to push existing commits...
)

echo.
echo [4/4] Pushing to GitHub...
git branch -M main
git push -u origin main
if errorlevel 1 (
    echo.
    echo ERROR: git push failed.
    echo If GitHub asks for sign-in, complete the browser/device authentication and run this again.
    pause
    exit /b 1
)

echo.
echo SUCCESS: Files pushed to GitHub.
echo https://github.com/kourtneymiller-sys/NextGen-PMO
echo.
pause
