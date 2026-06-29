@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

rem ===== Pythonのパス（指定どおり）=====
set "PYTHON=%LOCALAPPDATA%\Microsoft\WindowsApps\python.exe"

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

rem ===== Python存在チェック（指定パス）=====
if not exist "%PYTHON%" (
  echo [ERROR] Python が見つかりません: %PYTHON%
  echo Microsoft Store版Pythonが未インストール、またはWindowsAppsが無効の可能性があります。
  echo 管理者に Python のインストール／設定を依頼してください。
  echo.
  pause
  exit /b 1
)

rem ===== 出力ファイルパス =====
set "OUT_FILE=%OUT_DIR%\%IN_BASE%_hanja.docx"

echo [INFO] Python: %PYTHON%
echo [INFO] 入力ファイル: %IN_NAME%
echo [INFO] 出力ファイル: %OUT_FILE%
echo.

echo ----------------------------------------
echo [INFO] 変換を開始します
echo ----------------------------------------

if defined DICT (
  "%PYTHON%" "%SCRIPT%" "%IN_FILE%" "%OUT_FILE%" "%DICT%"
) else (
  "%PYTHON%" "%SCRIPT%" "%IN_FILE%" "%OUT_FILE%"
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