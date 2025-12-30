@echo off
REM ================================
REM Step1: PasteMD 初期セットアップ
REM ================================

REM PasteMD 設定フォルダ
set "PASTEMD_CONFIG_DIR=%APPDATA%\PasteMD"

REM 1) Mermaid フィルターをグローバルインストール
npm install --global mermaid-filter

REM 2) 設定フォルダがなければ作成
if not exist "%PASTEMD_CONFIG_DIR%" (
    mkdir "%PASTEMD_CONFIG_DIR%"
)

REM 3) config.json をコピー
copy /Y "PasteMD\config.json" "%PASTEMD_CONFIG_DIR%\config.json"

echo Setup completed.
pause
