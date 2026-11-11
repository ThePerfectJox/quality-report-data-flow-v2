@echo off
rem Double-click to run the server, the client or the config editor.
setlocal
cd /d "%~dp0"
title Quality Report Data Flow

rem Use the project's virtual environment if there is one, otherwise Python on PATH.
set "PYTHON="
if exist ".venv\Scripts\python.exe" set "PYTHON=.venv\Scripts\python.exe"
if not defined PYTHON where python >nul 2>nul && set "PYTHON=python"
if not defined PYTHON where py >nul 2>nul && set "PYTHON=py -3"
if not defined PYTHON (
    echo Python was not found. Install Python 3.10 or newer from https://www.python.org/
    echo and tick "Add python.exe to PATH" during the installation.
    echo.
    pause
    exit /b 1
)

:menu
cls
echo  Quality Report Data Flow
echo  ------------------------
echo.
echo   1. Run server         (server device: answers the clients' report requests)
echo   2. Run client         (client device: gets the report and appends it to the master data)
echo   3. Run config editor  (client device: master data workbook and inspec requirements)
echo   4. Exit
echo.
choice /c 1234 /n /m " Choose 1, 2, 3 or 4: "
if errorlevel 4 goto :end
if errorlevel 3 goto :editor
if errorlevel 2 goto :client
goto :server

:server
cls
echo Starting the server. Press Ctrl+C to stop it, then N to come back to this menu.
echo.
%PYTHON% src\run_server.py
goto :finished

:client
cls
%PYTHON% src\run_client.py
goto :finished

:editor
cls
echo Starting the config editor. It opens in your browser.
echo Keep this window open while you edit. Press Ctrl+C to close it, then N to come back to this menu.
echo.
%PYTHON% src\run_config_editor.py
goto :finished

:finished
echo.
if errorlevel 1 (echo It stopped with an error. Read the message above.) else (echo Done.)
echo.
pause
goto :menu

:end
endlocal
