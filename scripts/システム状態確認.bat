@echo off
chcp 65001 >nul 2>&1
SETLOCAL
SET "HERE=%~dp0"

REM ========================================
REM SMCL System Status Check
REM ========================================

echo ========================================
echo SMCL System Status Check
echo ========================================
echo.

REM ---------- Python Version Check ----------
echo [1] Python Version Check
echo ----------------------------------------
py --version 2>nul
IF ERRORLEVEL 1 (
    echo [ERROR] Python is not installed or not in PATH
    echo Please install Python 3.9 or higher
) ELSE (
    py --version | findstr /R "3\.[9-9]\|3\.1[0-9]" >nul
    IF ERRORLEVEL 1 (
        echo [WARNING] Python 3.9 or higher is recommended
    ) ELSE (
        echo [OK] Python version is compatible
    )
)
echo.

REM ---------- Virtual Environment Check ----------
echo [2] Virtual Environment Check
echo ----------------------------------------
IF EXIST "%HERE%.venv\Scripts\activate.bat" (
    echo [OK] Virtual environment exists
    CALL "%HERE%.venv\Scripts\activate.bat" 2>nul
    IF ERRORLEVEL 1 (
        echo [ERROR] Virtual environment is corrupted
    ) ELSE (
        echo [OK] Virtual environment is functional
        python --version 2>nul
    )
) ELSE (
    echo [WARNING] Virtual environment not found
    echo Run the main batch file to create it automatically
)
echo.

REM ---------- Dependencies Check ----------
echo [3] Dependencies Check
echo ----------------------------------------
IF EXIST "%HERE%requirements.txt" (
    echo [OK] requirements.txt found
    IF EXIST "%HERE%.venv\Scripts\activate.bat" (
        CALL "%HERE%.venv\Scripts\activate.bat" 2>nul
        pip list --format=freeze > temp_installed.txt 2>nul
        IF EXIST temp_installed.txt (
            echo [INFO] Checking installed packages...
            findstr /i "selenium" temp_installed.txt >nul && echo [OK] selenium installed || echo [WARNING] selenium not found
            findstr /i "pandas" temp_installed.txt >nul && echo [OK] pandas installed || echo [WARNING] pandas not found
            findstr /i "openpyxl" temp_installed.txt >nul && echo [OK] openpyxl installed || echo [WARNING] openpyxl not found
            del temp_installed.txt 2>nul
        )
    )
) ELSE (
    echo [ERROR] requirements.txt not found
)
echo.

REM ---------- Configuration Files Check ----------
echo [4] Configuration Files Check
echo ----------------------------------------
IF EXIST "%HERE%main.py" (
    echo [OK] main.py found
) ELSE (
    echo [ERROR] main.py not found
)

IF EXIST "%HERE%services\core\config.py" (
    echo [OK] config.py found
) ELSE (
    echo [ERROR] config.py not found
)

IF EXIST "%HERE%services\docs\角上魚類.xlsx" (
    echo [OK] Master Excel file found
) ELSE (
    echo [WARNING] Master Excel file not found
    echo Expected: services\docs\角上魚類.xlsx
)
echo.

REM ---------- Network Drive Check ----------
echo [5] Network Drive Check
echo ----------------------------------------
IF EXIST "\\sv001\Userdata\" (
    echo [OK] Network drive \\sv001\Userdata\ accessible
    IF EXIST "\\sv001\Userdata\納品リスト処理" (
        echo [OK] Target folder \\sv001\Userdata\納品リスト処理 exists
    ) ELSE (
        echo [WARNING] Target folder \\sv001\Userdata\納品リスト処理 not found
        echo Will use local output folder
    )
) ELSE (
    echo [WARNING] Network drive \\sv001\Userdata\ not accessible
    echo Will use local output folder
)
echo.

REM ---------- Log Directory Check ----------
echo [6] Log Directory Check
echo ----------------------------------------
IF EXIST "%HERE%logs" (
    echo [OK] logs directory exists
    dir "%HERE%logs\*.log" /b 2>nul | findstr . >nul
    IF NOT ERRORLEVEL 1 (
        echo [INFO] Log files found
    ) ELSE (
        echo [INFO] No log files found (first run)
    )
) ELSE (
    echo [INFO] logs directory will be created on first run
)
echo.

REM ---------- Task Scheduler Check ----------
echo [7] Task Scheduler Check
echo ----------------------------------------
schtasks /query /tn "SMCL納品リスト処理_毎朝10時" >nul 2>&1
IF ERRORLEVEL 1 (
    echo [WARNING] Scheduled task not found
    echo Run タスクスケジューラ設定.ps1 to create it
) ELSE (
    echo [OK] Scheduled task exists
    schtasks /query /tn "SMCL納品リスト処理_毎朝10時" /fo LIST | findstr "Next Run Time"
)
echo.

REM ---------- Summary ----------
echo ========================================
echo System Status Summary
echo ========================================
echo If you see any [ERROR] items above, please fix them before running the system.
echo [WARNING] items are optional but recommended to fix.
echo.
echo To run the system manually:
echo   納品リスト処理バッチ.bat
echo.
echo To set up automatic execution:
echo   Run PowerShell as Administrator and execute:
echo   .\タスクスケジューラ設定.ps1
echo.

pause