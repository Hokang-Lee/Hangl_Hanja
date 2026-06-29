@echo off
chcp 65001 > nul
echo Korean to Hanja Converter
echo.

if "%~1"=="" (
    echo Usage: Drag and drop a .docx file onto this bat file
    pause
    exit /b
)

set "PYTHON_CMD="
where py >nul 2>&1
if "%ERRORLEVEL%"=="0" set "PYTHON_CMD=py -3"
if not defined PYTHON_CMD (
    where python >nul 2>&1
    if "%ERRORLEVEL%"=="0" set "PYTHON_CMD=python"
)
if not defined PYTHON_CMD (
    if exist "%LOCALAPPDATA%\Microsoft\WindowsApps\python.exe" set "PYTHON_CMD="%LOCALAPPDATA%\Microsoft\WindowsApps\python.exe""
)

if not defined PYTHON_CMD (
    echo ERROR: Python was not found.
    echo Please install Python or add it to PATH.
    pause
    exit /b 1
)

echo Using: %PYTHON_CMD%
%PYTHON_CMD% "%~dp0convert_hanja.py" %*

if errorlevel 1 (
    echo.
    echo ERROR occurred. Please check:
    echo   1. pip install python-docx lxml openpyxl docx2pdf
    echo   2. Dictionary Excel file is in same folder as this script
    pause
)
