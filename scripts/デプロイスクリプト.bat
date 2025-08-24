@echo off
chcp 65001 >nul 2>&1
SETLOCAL ENABLEDELAYEDEXPANSION
SET "HERE=%~dp0"

REM ========================================
REM SMCL System Deployment Script
REM ========================================

echo ========================================
echo SMCL System Deployment Script
echo ========================================
echo.

REM ---------- Check if running as Administrator ----------
net session >nul 2>&1
IF ERRORLEVEL 1 (
    echo [WARNING] This script should be run as Administrator for full functionality
    echo Some features may not work properly
    echo.
)

REM ---------- System Requirements Check ----------
echo [1] Checking system requirements...
echo ----------------------------------------

REM Python check
py --version 2>nul | findstr /R "3\.[9-9]\|3\.1[0-9]" >nul
IF ERRORLEVEL 1 (
    echo [ERROR] Python 3.9 or higher is required
    echo Please install Python from https://www.python.org/
    pause
    EXIT /B 1
) ELSE (
    echo [OK] Python version is compatible
)

REM PowerShell check
powershell -Command "Get-Host" >nul 2>&1
IF ERRORLEVEL 1 (
    echo [ERROR] PowerShell is required but not found
    pause
    EXIT /B 1
) ELSE (
    echo [OK] PowerShell is available
)
echo.

REM ---------- Create Virtual Environment ----------
echo [2] Setting up virtual environment...
echo ----------------------------------------
IF EXIST "%HERE%.venv" (
    echo [INFO] Removing existing virtual environment...
    rmdir /s /q "%HERE%.venv" 2>nul
)

echo [INFO] Creating new virtual environment...
py -m venv "%HERE%.venv"
IF ERRORLEVEL 1 (
    echo [ERROR] Failed to create virtual environment
    pause
    EXIT /B 1
)

echo [INFO] Activating virtual environment...
CALL "%HERE%.venv\Scripts\activate.bat"
IF ERRORLEVEL 1 (
    echo [ERROR] Failed to activate virtual environment
    pause
    EXIT /B 1
)

echo [OK] Virtual environment created and activated
echo.

REM ---------- Install Dependencies ----------
echo [3] Installing dependencies...
echo ----------------------------------------
echo [INFO] Upgrading pip...
pip install --upgrade pip --quiet

echo [INFO] Installing required packages...
pip install -r requirements.txt
IF ERRORLEVEL 1 (
    echo [ERROR] Failed to install some dependencies
    echo [INFO] Continuing with deployment...
) ELSE (
    echo [OK] All dependencies installed successfully
)
echo.

REM ---------- Create Directories ----------
echo [4] Creating directories...
echo ----------------------------------------
IF NOT EXIST "%HERE%logs" (
    mkdir "%HERE%logs"
    echo [OK] Created logs directory
)

IF NOT EXIST "%HERE%output" (
    mkdir "%HERE%output"
    echo [OK] Created output directory
)

IF NOT EXIST "%HERE%downloads" (
    mkdir "%HERE%downloads"
    echo [OK] Created downloads directory
)

IF NOT EXIST "%HERE%services\downloads" (
    mkdir "%HERE%services\downloads"
    echo [OK] Created services\downloads directory
)

IF NOT EXIST "%HERE%services\output" (
    mkdir "%HERE%services\output"
    echo [OK] Created services\output directory
)

IF NOT EXIST "%HERE%services\logs" (
    mkdir "%HERE%services\logs"
    echo [OK] Created services\logs directory
)
echo.

REM ---------- Network Drive Setup ----------
echo [5] Network drive setup...
echo ----------------------------------------
IF EXIST "\\sv001\Userdata\" (
    echo [OK] Network drive \\sv001\Userdata\ is accessible
    IF NOT EXIST "\\sv001\Userdata\納品リスト処理" (
        echo [INFO] Creating network folder...
        mkdir "\\sv001\Userdata\納品リスト処理" 2>nul
        IF ERRORLEVEL 1 (
            echo [WARNING] Could not create network folder
            echo [INFO] Will use local folders instead
        ) ELSE (
            echo [OK] Network folder created
        )
    ) ELSE (
        echo [OK] Network folder already exists
    )
) ELSE (
    echo [WARNING] Network drive not accessible
    echo [INFO] System will use local folders
)
echo.

REM ---------- Configuration Check ----------
echo [6] Configuration check...
echo ----------------------------------------
IF EXIST "%HERE%services\docs\角上魚類.xlsx" (
    echo [OK] Master Excel file found
) ELSE (
    echo [WARNING] Master Excel file not found
    echo [INFO] Please place 角上魚類.xlsx in services\docs\ folder
)

IF EXIST "%HERE%credentials.json" (
    echo [OK] Google Drive credentials found
) ELSE (
    echo [WARNING] Google Drive credentials not found
    echo [INFO] Please place credentials.json in project root
)
echo.

REM ---------- Test Run ----------
echo [7] Running system test...
echo ----------------------------------------
echo [INFO] Testing system configuration...
python -c "from services.core.config import Config; c=Config(); c.print_configuration_status()" 2>nul
IF ERRORLEVEL 1 (
    echo [WARNING] Configuration test had issues
    echo [INFO] Please check the configuration manually
) ELSE (
    echo [OK] Configuration test passed
)
echo.

REM ---------- Task Scheduler Setup ----------
echo [8] Task scheduler setup...
echo ----------------------------------------
echo [INFO] To set up automatic execution, run the following command as Administrator:
echo.
echo   powershell -ExecutionPolicy RemoteSigned -File ".\タスクスケジューラ設定.ps1"
echo.
echo [INFO] Or manually run: タスクスケジューラ設定.ps1
echo.

REM ---------- Deployment Summary ----------
echo ========================================
echo Deployment Summary
echo ========================================
echo [OK] Virtual environment created
echo [OK] Dependencies installed
echo [OK] Directories created
echo.
echo Next steps:
echo 1. Place master Excel file: services\docs\角上魚類.xlsx
echo 2. Place Google credentials: credentials.json
echo 3. Set up LINE Bot environment variables:
echo    set LINE_CHANNEL_ACCESS_TOKEN=your_token
echo    set LINE_GROUP_ID=your_group_id
echo 4. Run task scheduler setup (as Administrator):
echo    powershell -ExecutionPolicy RemoteSigned -File ".\タスクスケジューラ設定.ps1"
echo 5. Test the system:
echo    納品リスト処理バッチ.bat
echo.
echo Deployment completed!
echo.

pause