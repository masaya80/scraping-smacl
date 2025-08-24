@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul 2>&1

REM Determine absolute path to run script
SET "SCRIPT_DIR=%~dp0"
SET "RUN_BAT=%SCRIPT_DIR%run_prod_windows.bat"
SET "TASK_NAME=SMCL_Delivery_Prod"

echo Registering Windows Task Scheduler job: %TASK_NAME%
echo Target: %RUN_BAT%

REM Remove existing task if present
schtasks /Query /TN "%TASK_NAME%" >nul 2>&1
IF %ERRORLEVEL% EQU 0 (
    echo [INFO] Existing task found. Deleting...
    schtasks /Delete /TN "%TASK_NAME%" /F >nul 2>&1
)

REM Create daily 10:00 task for current user (runs when user is logged on)
schtasks /Create ^
  /TN "%TASK_NAME%" ^
  /TR "\"%RUN_BAT%\"" ^
  /SC DAILY ^
  /ST 10:00 ^
  /RU "%USERNAME%" ^
  /RL LIMITED ^
  /F

IF %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Failed to create the scheduled task. ErrorLevel=%ERRORLEVEL%
    exit /b %ERRORLEVEL%
)

echo [SUCCESS] Task created.
echo To run now: schtasks /Run /TN "%TASK_NAME%"
echo To check status: schtasks /Query /TN "%TASK_NAME%" /V /FO LIST
echo To delete: schtasks /Delete /TN "%TASK_NAME%" /F

exit /b 0


