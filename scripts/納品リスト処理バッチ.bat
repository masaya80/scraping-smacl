@echo off
chcp 65001 >nul 2>&1
SETLOCAL ENABLEDELAYEDEXPANSION
SET "HERE=%~dp0"

REM ========================================
REM SMCL Delivery List Processing System
REM Automated processing executed daily at 10:00 AM
REM ========================================

echo [%date% %time%] SMCL Delivery List Processing System Starting
echo.

REM ---------- Python Version Check ----------
echo [%time%] Checking Python version...
py --version 2>nul | findstr /R "3\.[9-9]\|3\.1[0-9]" >nul
IF ERRORLEVEL 1 (
    echo [ERROR] Python 3.9 or higher is required
    py --version 2>nul
    pause
    EXIT /B 1
)

REM ---------- Create venv if not exists ----------
IF NOT EXIST "%HERE%.venv\Scripts\activate.bat" (
    echo [%time%] Virtual environment not found. Creating...
    py -m venv "%HERE%.venv"
    IF ERRORLEVEL 1 (
        echo [ERROR] Failed to create virtual environment
        pause
        EXIT /B 1
    )
    echo [%time%] Virtual environment created
    SET "VENV_CREATED=1"
) ELSE (
    REM ---------- venv health check ----------
    CALL "%HERE%.venv\Scripts\activate.bat" 2>nul
    IF ERRORLEVEL 1 (
        echo [%time%] Virtual environment is corrupted. Recreating...
        rmdir /s /q "%HERE%.venv" 2>nul
        py -m venv "%HERE%.venv"
        IF ERRORLEVEL 1 (
            echo [ERROR] Failed to recreate virtual environment
            pause
            EXIT /B 1
        )
        echo [%time%] Virtual environment recreated
        SET "VENV_CREATED=1"
    )
)

REM ---------- Activate venv ----------
echo [%time%] Activating virtual environment...
CALL "%HERE%.venv\Scripts\activate.bat"
IF ERRORLEVEL 1 (
    echo [ERROR] Failed to activate virtual environment
    pause
    EXIT /B 1
)

REM ---------- Dependencies Installation ----------
echo [%time%] Checking dependencies...

REM Always install for new venv
IF DEFINED VENV_CREATED (
    echo [%time%] Installing dependencies for new virtual environment...
    pip install --upgrade pip --quiet
    pip install -r requirements.txt
    IF ERRORLEVEL 1 (
        echo [ERROR] Failed to install dependencies
        pause
        EXIT /B 1
    )
    echo [%time%] Dependencies installation completed
) ELSE (
    REM Check for updates in existing venv
    pip install -r requirements.txt --dry-run --quiet 2>nul | findstr "Would install" >nul
    IF NOT ERRORLEVEL 1 (
        echo [%time%] Dependencies update required. Installing...
        pip install -r requirements.txt --quiet
        IF ERRORLEVEL 1 (
            echo [WARNING] Dependencies installation had warnings
        ) ELSE (
            echo [%time%] Dependencies update completed
        )
    ) ELSE (
        echo [%time%] Dependencies are up to date
    )
)

REM ---------- Network Drive Check ----------
echo [%time%] Checking network drive connection...
IF NOT EXIST "\\sv001\Userdata\納品リスト処理" (
    echo [WARNING] Cannot access network drive \\sv001\Userdata\納品リスト処理
    echo [INFO] Using local output folder
)

REM ---------- Execute Python Script ----------
echo [%time%] Starting delivery list processing...
echo.
python main.py
SET PYTHON_EXIT_CODE=!ERRORLEVEL!

REM ---------- Display Results ----------
echo.
IF !PYTHON_EXIT_CODE! EQU 0 (
    echo [%time%] SUCCESS: Processing completed successfully
) ELSE (
    echo [%time%] ERROR: Processing failed (exit code: !PYTHON_EXIT_CODE!)
)

REM ---------- Save Logs ----------
echo [%time%] Saving logs...
IF EXIST "logs\app.log" (
    copy "logs\app.log" "logs\app_%date:~0,4%%date:~5,2%%date:~8,2%_%time:~0,2%%time:~3,2%.log" >nul 2>&1
)

REM ---------- No pause for auto execution ----------
IF "%1"=="auto" (
    echo [%time%] Exiting in auto execution mode
    EXIT /B !PYTHON_EXIT_CODE!
)

REM ---------- Wait for user input in manual execution ----------
echo.
echo Press any key to exit...
pause >nul
EXIT /B !PYTHON_EXIT_CODE!