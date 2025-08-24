@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul 2>&1
echo ========================================
echo SMCL Delivery List Processing System
echo ========================================
echo.
echo Starting system...
echo.

REM Get script directory and handle UNC paths
SET "SCRIPT_DIR=%~dp0"
echo Script directory: %SCRIPT_DIR%

REM If path starts with \\ (UNC), use pushd to map a temporary drive
SET "USING_PUSHD="
IF "%SCRIPT_DIR:~0,2%"=="\\" (
    echo [INFO] UNC path detected. Using pushd to create a temporary mapping...
    pushd "%SCRIPT_DIR%" >nul 2>&1
    IF ERRORLEVEL 1 (
        echo [WARNING] pushd failed. Falling back to copy files to temp directory...
        SET "TEMP_DIR=%TEMP%\smcl_temp"
        IF NOT EXIST "%TEMP_DIR%" mkdir "%TEMP_DIR%"
        xcopy "%SCRIPT_DIR%*" "%TEMP_DIR%\" /E /Y /Q >nul 2>&1
        cd /d "%TEMP_DIR%"
        SET "USING_TEMP=1"
    ) ELSE (
        SET "USING_PUSHD=1"
        echo [INFO] Working directory after pushd: %CD%
    )
) ELSE (
    REM Normal path, just change directory
    cd /d "%SCRIPT_DIR%"
)

REM Verify we can access the current directory
IF NOT EXIST "requirements.txt" (
    echo [ERROR] Cannot find requirements.txt in current directory
    echo Current directory: %CD%
    echo Please ensure the script is run from the correct location
    pause
    EXIT /B 1
)

REM Create virtual environment if not exists
IF NOT EXIST ".venv\Scripts\activate.bat" (
    echo Creating virtual environment...
    py -m venv .venv
    IF ERRORLEVEL 1 (
        echo [ERROR] Failed to create virtual environment
        pause
        EXIT /B 1
    )
    echo Virtual environment created
    SET "VENV_CREATED=1"
)

REM Activate virtual environment
echo Activating virtual environment...
CALL ".venv\Scripts\activate.bat"
IF ERRORLEVEL 1 (
    echo [ERROR] Failed to activate virtual environment
    pause
    EXIT /B 1
)

REM Install dependencies if needed
IF DEFINED VENV_CREATED (
    echo Installing dependencies...
    echo [INFO] Upgrading pip...
    python -m pip install --upgrade pip --quiet
    IF ERRORLEVEL 1 (
        echo [WARNING] Failed to upgrade pip, continuing with current version
    )
    
    echo [INFO] Installing packages from requirements.txt...
    pip install -r requirements.txt --verbose
    IF ERRORLEVEL 1 (
        echo [ERROR] Failed to install dependencies
        echo [INFO] Trying alternative installation method...
        pip install -r requirements.txt --no-cache-dir --force-reinstall
        IF ERRORLEVEL 1 (
            echo [ERROR] All installation attempts failed
            echo [INFO] Please check your internet connection and try again
            pause
            EXIT /B 1
        )
    )
    echo [SUCCESS] Dependencies installed successfully
)

REM Execute Python script
echo Starting delivery list processing...
echo.
python main.py
SET PYTHON_EXIT_CODE=%ERRORLEVEL%

REM Display results
echo.
IF %PYTHON_EXIT_CODE% EQU 0 (
    echo [SUCCESS] Processing completed successfully
) ELSE (
    echo [ERROR] Processing failed (exit code: %PYTHON_EXIT_CODE%)
)

REM Cleanup temporary directory if used
IF DEFINED USING_TEMP (
    echo [INFO] Copying output files back to original location...
    IF EXIST "output" (
        xcopy "output\*" "%SCRIPT_DIR%output\" /E /Y /Q >nul 2>&1
    )
    IF EXIST "logs" (
        xcopy "logs\*" "%SCRIPT_DIR%logs\" /E /Y /Q >nul 2>&1
    )
    echo [INFO] Cleaning up temporary files...
    cd /d "%TEMP%"
    rmdir /s /q "%TEMP_DIR%" >nul 2>&1
)

REM If we used pushd for UNC, restore previous directory and unmap temp drive
IF DEFINED USING_PUSHD (
    popd >nul 2>&1
)

echo.
echo Press any key to exit...
pause >nul
EXIT /B %PYTHON_EXIT_CODE%