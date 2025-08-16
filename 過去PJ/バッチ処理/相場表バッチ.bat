@echo off
SETLOCAL
SET "HERE=%~dp0"
REM ---------- venv が無ければ作成 ----------
IF NOT EXIST "%HERE%.venv\Scripts\activate.bat" (
    echo [.venv not found] creating one...
    py -m venv "%HERE%.venv"
)
REM ---------- venv を有効化 ----------
CALL "%HERE%.venv\Scripts\activate.bat"

REM ---------- Python スクリプト実行 ----------
python drive_test.py %*

REM ---------- 終了コードを返す ----------
EXIT /B %ERRORLEVEL%
