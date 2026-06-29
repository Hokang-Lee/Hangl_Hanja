@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

set "PYTHON=%LOCALAPPDATA%\Microsoft\WindowsApps\python.exe"
set "SCRIPT=%~dp0convert_hanja.py"
set "OUT_DIR=%~dp0output"

for /f %%T in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set "STAMP=%%T"

if "%~1"=="" (
  echo [ERROR] 変換したい .docx をこのbatにドラッグ＆ドロップしてください。
  pause
  exit /b 1
)

set "IN_FILE=%~1"
set "IN_EXT=%~x1"
set "IN_BASE=%~n1"

if /I not "%IN_EXT%"==".docx" (
  echo [ERROR] .docxのみ対応です: "%IN_FILE%"
  pause
  exit /b 1
)

if not exist "%PYTHON%" (
  echo [ERROR] python.exe が見つかりません: "%PYTHON%"
  pause
  exit /b 1
)

if not exist "%SCRIPT%" (
  echo [ERROR] convert_hanja.py が見つかりません: "%SCRIPT%"
  pause
  exit /b 1
)

if not exist "%OUT_DIR%" mkdir "%OUT_DIR%" >nul 2>&1

set "OUT_DOCX=%OUT_DIR%\%IN_BASE%_hanja_%STAMP:~0,8%_%STAMP:~8,4%.docx"
set "OUT_PDF=%OUT_DIR%\%IN_BASE%_hanja_%STAMP:~0,8%_%STAMP:~8,4%.pdf"

cls
echo ========================================
echo Hanja Converter
echo ========================================
echo Input : "%IN_FILE%"
echo Output: "%OUT_DOCX%"
echo PDF   : "%OUT_PDF%"
echo ----------------------------------------
echo [1/2] 変換中（漢字変換→docx保存）...
echo ----------------------------------------

rem ===== 通常実行：画面に進行を出す（ログは作らない）=====
"%PYTHON%" "%SCRIPT%" "%IN_FILE%" "%OUT_DOCX%"
set "RC=%ERRORLEVEL%"
if not "%RC%"=="0" goto MAKELOG_AND_FAIL

echo.
echo ----------------------------------------
echo [2/2] PDF作成待ち（最大120秒）...
echo ----------------------------------------

rem ===== PDF必須：最大120秒待つ（待っている間に表示）=====
set /a WAIT=0
:WAITPDF
if exist "%OUT_PDF%" goto OK
set /a WAIT+=1
set /a MOD=WAIT%%5
if !MOD! EQU 0 echo   ... !WAIT! 秒経過
timeout /t 1 /nobreak >nul
if !WAIT! GEQ 120 (
  set "RC=2"
  goto MAKELOG_AND_FAIL
)
goto WAITPDF

:MAKELOG_AND_FAIL
rem ===== 失敗時のみログ作成（再実行して出力をログに取る）=====
set "LOG_FILE=%OUT_DIR%\error_%IN_BASE%_%STAMP:~0,8%_%STAMP:~8,4%.log"

(
  echo ===== ERROR LOG =====
  echo TIME  : %date% %time%
  echo BAT   : %~f0
  echo PYTHON: "%PYTHON%"
  echo SCRIPT: "%SCRIPT%"
  echo INPUT : "%IN_FILE%"
  echo OUTDOCX: "%OUT_DOCX%"
  echo OUTPDF : "%OUT_PDF%"
  echo.
  echo --- Python version ---
  "%PYTHON%" --version
  echo.
  echo --- Run script (captured) ---
  "%PYTHON%" "%SCRIPT%" "%IN_FILE%" "%OUT_DOCX%"
  echo.
  echo ERRORLEVEL(after run)=%ERRORLEVEL%
  echo.
  if "%RC%"=="2" (
    echo [ERROR] PDFが120秒以内に作成されませんでした（PDF必須）
  ) else (
    echo [ERROR] Python処理が失敗しました
  )
) > "%LOG_FILE%" 2>&1

echo.
echo [ERROR] 失敗しました。ログを作成しました:
echo "%LOG_FILE%"
pause
exit /b %RC%

:OK
echo.
echo [OK] 完了しました。まもなく閉じます...
timeout /t 2 /nobreak >nul
exit /b 0
endlocal