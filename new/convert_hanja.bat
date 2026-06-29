@echo off
chcp 65001 > nul
echo Korean to Hanja Converter
echo.

if "%~1"=="" (
    echo Usage: Drag and drop a .docx file onto this bat file
    pause
    exit /b
)

set PYTHON=%LOCALAPPDATA%\Microsoft\WindowsApps\python.exe

echo Using: %PYTHON%
"%PYTHON%" "%~dp0convert_hanja.py" %*

if errorlevel 1 (
    echo.
    echo ERROR occurred. Please check:
    echo   1. pip install python-docx lxml openpyxl docx2pdf
    echo   2. Dictionary Excel file is in same folder as this script
    pause
)
