@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

rem ===== Pythonコマンドを探索 =====
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

rem ===== 設定（必要ならここだけ変更）=====
set "SCRIPT=%~dp0convert_hanja.py"
set "DICT="
set "OUT_DIR=%~dp0output"

echo ----------------------------------------
echo Hanja Converter (docx) - Drag and Drop
echo ----------------------------------------
echo.

rem ===== ドラッグ＆ドロップ引数チェック =====
if "%~1"=="" (
  echo [ERROR] 変換したい .docx ファイルを、この bat にドラッグ＆ドロップしてください。
  echo.
  pause
  exit /b 1
)

rem ===== 入力ファイル情報 =====
set "IN_FILE=%~1"
set "IN_EXT=%~x1"
set "IN_NAME=%~nx1"
set "IN_BASE=%~n1"

rem ===== 拡張子チェック（.docxのみ）=====
if /I not "%IN_EXT%"==".docx" (
  echo [ERROR] 対象は .docx のみです: %IN_NAME%
  echo.
  pause
  exit /b 1
)

rem ===== 出力フォルダ作成 =====
if not exist "%OUT_DIR%" (
  mkdir "%OUT_DIR%" >nul 2>&1
)

rem ===== スクリプト存在チェック =====
if not exist "%SCRIPT%" (
  echo [ERROR] convert_hanja.py が見つかりません。
  echo bat と同じフォルダに convert_hanja.py を置いてください。
  echo.
  pause
  exit /b 1
)

rem ===== Python存在チェック =====
if not defined PYTHON_CMD (
  echo [ERROR] Python が見つかりません。
  echo py -3、python、Microsoft Store版Python のいずれも利用できませんでした。
  echo 管理者に Python のインストール／設定を依頼してください。
  echo.
  pause
  exit /b 1
)

rem ===== 出力ファイルパス =====
set "OUT_FILE=%OUT_DIR%\%IN_BASE%_hanja.docx"

echo [INFO] Python: %PYTHON_CMD%
echo [INFO] 入力ファイル: %IN_NAME%
echo [INFO] 出力ファイル: %OUT_FILE%
echo.

echo ----------------------------------------
echo [INFO] 変換を開始します
echo ----------------------------------------

if defined DICT (
  %PYTHON_CMD% "%SCRIPT%" "%IN_FILE%" "%OUT_FILE%" "%DICT%"
) else (
  %PYTHON_CMD% "%SCRIPT%" "%IN_FILE%" "%OUT_FILE%"
)

if errorlevel 1 (
  echo.
  echo [ERROR] 変換に失敗しました。
  echo 画面のエラーメッセージを管理者に共有してください。
  echo.
  pause
  exit /b 1
)

echo.
echo ========================================
echo [DONE] 変換が完了しました。
echo 出力フォルダ: %OUT_DIR%
echo ========================================
echo.
pause
endlocal
