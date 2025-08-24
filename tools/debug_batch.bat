@echo off
echo ========================================
echo Debug Batch Test
echo ========================================
echo.

echo [1] Basic echo test:
echo This message should appear
echo.

echo [2] Python version check:
py --version
echo Error level after py command: %ERRORLEVEL%
echo.

echo [3] Current directory:
echo %CD%
echo.

echo [4] PATH check:
echo PATH contains Python?
echo %PATH% | findstr /i python
echo.

echo [5] File existence check:
IF EXIST "main.py" (
    echo [OK] main.py exists
) ELSE (
    echo [ERROR] main.py not found
)

IF EXIST "requirements.txt" (
    echo [OK] requirements.txt exists
) ELSE (
    echo [ERROR] requirements.txt not found
)
echo.

echo ========================================
echo Debug completed - Press any key to exit
echo ========================================
pause >nul
