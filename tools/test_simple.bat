@echo off
chcp 65001 >nul 2>&1
SETLOCAL
SET "HERE=%~dp0"

echo ========================================
echo Simple Test Batch
echo ========================================
echo.

echo Current directory: %HERE%
echo Python version check:
py --version

echo.
echo Virtual environment check:
IF EXIST "%HERE%.venv\Scripts\activate.bat" (
    echo Virtual environment exists
) ELSE (
    echo Virtual environment not found
)

echo.
echo Test completed successfully
pause
