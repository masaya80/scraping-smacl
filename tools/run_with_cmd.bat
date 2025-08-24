@echo off
echo ========================================
echo Running System Check with CMD
echo ========================================
echo.

echo This will open a new command prompt window
echo and run the system check there.
echo.
echo Press any key to continue...
pause >nul

REM Open new CMD window and run system check
start cmd /k "cd /d "%~dp0" && システム状態確認.bat"

echo.
echo A new command prompt window should have opened.
echo If not, try running the batch files directly from command prompt:
echo.
echo 1. Open Command Prompt (cmd)
echo 2. Navigate to this folder: cd /d "%~dp0"
echo 3. Run: システム状態確認.bat
echo.
pause
