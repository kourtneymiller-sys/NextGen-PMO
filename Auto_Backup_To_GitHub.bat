@echo off
setlocal enabledelayedexpansion

set "REPO_DIR=%~dp0"
set "LOG_DIR=%REPO_DIR%backup_logs"
if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"

for /f "tokens=1-4 delims=/ " %%a in ("%date%") do set "TODAY=%%d-%%b-%%c"
for /f "tokens=1-2 delims=: " %%a in ("%time%") do set "NOW=%%a%%b"
set "LOG_FILE=%LOG_DIR%\github_backup_%TODAY%_%NOW%.log"

echo ================================================== >> "%LOG_FILE%"
echo GitHub Backup Started: %date% %time% >> "%LOG_FILE%"
echo Repo: %REPO_DIR% >> "%LOG_FILE%"
echo ================================================== >> "%LOG_FILE%"

cd /d "%REPO_DIR%" >> "%LOG_FILE%" 2>&1
if errorlevel 1 (
  echo ERROR: Could not cd to repo directory. >> "%LOG_FILE%"
  exit /b 1
)

echo. >> "%LOG_FILE%"
echo [1/6] Git version >> "%LOG_FILE%"
git --version >> "%LOG_FILE%" 2>&1

echo. >> "%LOG_FILE%"
echo [2/6] Current status before add >> "%LOG_FILE%"
git status --short >> "%LOG_FILE%" 2>&1

echo. >> "%LOG_FILE%"
echo [3/6] Staging changes. Locked Office files may be skipped. >> "%LOG_FILE%"
git add --ignore-errors . >> "%LOG_FILE%" 2>&1

echo. >> "%LOG_FILE%"
echo [4/6] Commit staged changes if any >> "%LOG_FILE%"
git diff --cached --quiet
if errorlevel 1 (
  git commit -m "Automated backup %date% %time%" >> "%LOG_FILE%" 2>&1
) else (
  echo No staged changes to commit. >> "%LOG_FILE%"
)

echo. >> "%LOG_FILE%"
echo [5/6] Pull/rebase latest remote changes >> "%LOG_FILE%"
git pull --rebase origin main >> "%LOG_FILE%" 2>&1
if errorlevel 1 (
  echo WARNING: Pull/rebase failed. Push may fail; check log. >> "%LOG_FILE%"
)

echo. >> "%LOG_FILE%"
echo [6/6] Push to GitHub >> "%LOG_FILE%"
git push origin main >> "%LOG_FILE%" 2>&1
if errorlevel 1 (
  echo ERROR: Push failed. Authentication, network, or conflicts may need attention. >> "%LOG_FILE%"
  echo Finished with errors: %date% %time% >> "%LOG_FILE%"
  exit /b 1
)

echo. >> "%LOG_FILE%"
echo Final status >> "%LOG_FILE%"
git status --short >> "%LOG_FILE%" 2>&1

echo. >> "%LOG_FILE%"
echo SUCCESS: GitHub backup completed: %date% %time% >> "%LOG_FILE%"
exit /b 0
