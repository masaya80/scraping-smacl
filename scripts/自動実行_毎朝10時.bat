@echo off
chcp 65001 >nul 2>&1
SETLOCAL ENABLEDELAYEDEXPANSION
SET "HERE=%~dp0"

REM ========================================
REM SMCL Auto Execution Script for Task Scheduler
REM Daily execution at 10:00 AM
REM ========================================

REM ---------- Create log directory ----------
SET "LOG_DIR=%HERE%logs"
IF NOT EXIST "%LOG_DIR%" mkdir "%LOG_DIR%"

REM ---------- Set log file ----------
SET "LOG_FILE=%LOG_DIR%\batch_%date:~0,4%%date:~5,2%%date:~8,2%_%time:~0,2%%time:~3,2%.log"

REM ---------- Start logging ----------
echo [%date% %time%] Auto execution started > "%LOG_FILE%"
echo [%date% %time%] Working directory: %HERE% >> "%LOG_FILE%"

REM ---------- Change to script directory ----------
cd /d "%HERE%"
IF ERRORLEVEL 1 (
    echo [%date% %time%] ERROR: Failed to change directory to %HERE% >> "%LOG_FILE%"
    EXIT /B 1
)

REM ---------- Execute main batch with auto flag ----------
echo [%date% %time%] Executing main batch script... >> "%LOG_FILE%"
CALL "%HERE%納品リスト処理バッチ.bat" auto >> "%LOG_FILE%" 2>&1
SET BATCH_EXIT_CODE=!ERRORLEVEL!

REM ---------- Log results ----------
IF !BATCH_EXIT_CODE! EQU 0 (
    echo [%date% %time%] Auto execution completed successfully >> "%LOG_FILE%"
) ELSE (
    echo [%date% %time%] Auto execution failed with exit code: !BATCH_EXIT_CODE! >> "%LOG_FILE%"
)

REM ---------- Cleanup old logs (keep last 7 days) ----------
forfiles /p "%LOG_DIR%" /m "batch_*.log" /d -7 /c "cmd /c del @path" 2>nul
forfiles /p "%LOG_DIR%" /m "*.log" /s +10485760 /c "cmd /c del @path" 2>nul

REM ---------- Exit with same code as main batch ----------
EXIT /B !BATCH_EXIT_CODE!