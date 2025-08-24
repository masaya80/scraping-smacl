@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul 2>&1

REM Resolve script directory and move there
SET "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

REM Prepare logs directory and logfile name (sanitize date/time for filename)
IF NOT EXIST "logs" mkdir "logs"
SET "DATESTR=%DATE%"
SET "TIMESTR=%TIME%"
SET "DATESTR=%DATESTR:/=-%"
SET "DATESTR=%DATESTR:.=-%"
SET "TIMESTR=%TIMESTR::=-%"
SET "TIMESTR=%TIMESTR: =0%"
SET "LOGFILE=logs\run_%DATESTR%_%TIMESTR%.log"

echo ========================================>> "%LOGFILE%" 2>&1
echo SMCL Production Run >> "%LOGFILE%" 2>&1
echo ========================================>> "%LOGFILE%" 2>&1
echo Started at %DATE% %TIME% >> "%LOGFILE%" 2>&1
echo Working directory: %CD% >> "%LOGFILE%" 2>&1

REM Check if we need to recreate venv (if RECREATE_VENV is set or venv is corrupted)
IF "%RECREATE_VENV%"=="1" (
    echo [INFO] RECREATE_VENV flag detected. Removing existing venv... >> "%LOGFILE%" 2>&1
    IF EXIST ".venv" rmdir /s /q ".venv" >> "%LOGFILE%" 2>&1
)

REM Create venv on first run (setup) and install dependencies
IF NOT EXIST ".venv\Scripts\activate.bat" (
    echo [INFO] Creating virtual environment... >> "%LOGFILE%" 2>&1
    py -m venv .venv >> "%LOGFILE%" 2>&1
    IF ERRORLEVEL 1 (
        echo [ERROR] Failed to create venv with 'py'. Trying 'python'... >> "%LOGFILE%" 2>&1
        python -m venv .venv >> "%LOGFILE%" 2>&1
        IF ERRORLEVEL 1 (
            echo [FATAL] Failed to create virtual environment >> "%LOGFILE%" 2>&1
            EXIT /B 1
        )
    )
    echo [INFO] Activating venv and installing requirements... >> "%LOGFILE%" 2>&1
    CALL ".venv\Scripts\activate.bat" >> "%LOGFILE%" 2>&1
    IF ERRORLEVEL 1 (
        echo [ERROR] Failed to activate virtual environment >> "%LOGFILE%" 2>&1
        EXIT /B 1
    )
    python -m ensurepip --upgrade >> "%LOGFILE%" 2>&1
    python -m pip install --upgrade pip >> "%LOGFILE%" 2>&1
    IF ERRORLEVEL 1 (
        echo [ERROR] Failed to upgrade pip >> "%LOGFILE%" 2>&1
        EXIT /B 1
    )
    echo [INFO] Installing requirements.txt... >> "%LOGFILE%" 2>&1
    python -m pip install -r requirements.txt >> "%LOGFILE%" 2>&1
    IF ERRORLEVEL 1 (
        echo [ERROR] Failed to install requirements. Trying individual packages... >> "%LOGFILE%" 2>&1
        echo [INFO] Installing selenium individually... >> "%LOGFILE%" 2>&1
        python -m pip install selenium>=4.17.2 >> "%LOGFILE%" 2>&1
        echo [INFO] Installing webdriver-manager individually... >> "%LOGFILE%" 2>&1
        python -m pip install webdriver-manager>=4.0.2 >> "%LOGFILE%" 2>&1
        echo [INFO] Installing other essential packages... >> "%LOGFILE%" 2>&1
        python -m pip install pandas>=2.0.0 openpyxl>=3.1.0 requests>=2.31.0 >> "%LOGFILE%" 2>&1
    )
    echo [INFO] Verifying selenium installation... >> "%LOGFILE%" 2>&1
    python -c "import selenium; print('Selenium version:', selenium.__version__)" >> "%LOGFILE%" 2>&1
    IF ERRORLEVEL 1 (
        echo [ERROR] Selenium import failed after installation >> "%LOGFILE%" 2>&1
        EXIT /B 1
    )
)

REM Activate venv for run
echo [INFO] Activating virtual environment for execution... >> "%LOGFILE%" 2>&1
CALL ".venv\Scripts\activate.bat" >> "%LOGFILE%" 2>&1
IF ERRORLEVEL 1 (
    echo [ERROR] Failed to activate virtual environment for execution >> "%LOGFILE%" 2>&1
    EXIT /B 1
)

REM Verify environment before running main.py
echo [INFO] Verifying Python environment... >> "%LOGFILE%" 2>&1
python --version >> "%LOGFILE%" 2>&1
python -c "import sys; print('Python executable:', sys.executable)" >> "%LOGFILE%" 2>&1
echo [INFO] Checking selenium availability... >> "%LOGFILE%" 2>&1
python -c "import selenium; print('Selenium version:', selenium.__version__)" >> "%LOGFILE%" 2>&1
IF ERRORLEVEL 1 (
    echo [ERROR] Selenium not available. Attempting to reinstall... >> "%LOGFILE%" 2>&1
    python -m pip install --force-reinstall selenium>=4.17.2 >> "%LOGFILE%" 2>&1
    IF ERRORLEVEL 1 (
        echo [FATAL] Failed to reinstall selenium >> "%LOGFILE%" 2>&1
        EXIT /B 1
    )
    echo [INFO] Selenium reinstalled successfully >> "%LOGFILE%" 2>&1
)

echo [INFO] Starting main.py >> "%LOGFILE%" 2>&1
python main.py >> "%LOGFILE%" 2>&1
SET "EXITCODE=%ERRORLEVEL%"

echo [INFO] Exit code: %EXITCODE% >> "%LOGFILE%" 2>&1
echo Finished at %DATE% %TIME% >> "%LOGFILE%" 2>&1

exit /b %EXITCODE%


