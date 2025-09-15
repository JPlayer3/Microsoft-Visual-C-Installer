ChatGPTに作成してもらった。 
 
Microsoft 公式配布リンクを使用。 
各バージョンごとに x86 / x64 両方を対象。  
サイレントインストール /quiet /norestart。  
結果をログ VCRedist_Install.log に出力。  
ダウンロードは curl（Windows 10 以降標準搭載）を利用。  
  
特徴 
2005 → 2008 SP1 → 2010 SP1 → 2012 Update 4 → 2013 → 2015-2022 統合版（最新）  
x86 / x64 両方インストール  
curl を使用してダウンロード（Windows 10/11 標準対応）  
結果をログに記録  
   
⚠ 注意点:  
再頒布パッケージは既にインストール済みの場合、処理は自動でスキップされます。   
古い VC++ は順にインストールする必要があります（後のバージョンでは完全互換ではないため）。  

```
@echo off
setlocal enabledelayedexpansion

:: ============================================
:: Visual C++ Redistributable 一括インストール
:: 2005 ～ 2022 全バージョン
:: ============================================

set LOGFILE=%~dp0VCRedist_Install.log
echo ================================ >> "%LOGFILE%"
echo 開始 %date% %time% >> "%LOGFILE%"
echo ================================ >> "%LOGFILE%"

set TEMPDIR=%TEMP%\VCRedist_%RANDOM%
mkdir "%TEMPDIR%"

:: ====== ダウンロードURLリスト ======
:: 2005
set URLS[1]=https://download.microsoft.com/download/1/4/6/14695FCA-22D0-4630-8A74-389A0C5F6B7E/vcredist_x86.EXE
set URLS[2]=https://download.microsoft.com/download/1/4/6/14695FCA-22D0-4630-8A74-389A0C5F6B7E/vcredist_x64.EXE

:: 2008 SP1
set URLS[3]=https://download.microsoft.com/download/2/9/5/295A6C04-C438-45A8-B90F-9356B3963C5C/vcredist_x86.exe
set URLS[4]=https://download.microsoft.com/download/2/9/5/295A6C04-C438-45A8-B90F-9356B3963C5C/vcredist_x64.exe

:: 2010 SP1
set URLS[5]=https://download.microsoft.com/download/1/6/5/165255E3-8F4F-4E9D-8C8C-4E7D64D5C798/vcredist_x86.exe
set URLS[6]=https://download.microsoft.com/download/1/6/5/165255E3-8F4F-4E9D-8C8C-4E7D64D5C798/vcredist_x64.exe

:: 2012 Update 4
set URLS[7]=https://download.microsoft.com/download/1/6/B/16B06F60-3B20-4FF2-B699-5E9B7962F9AE/VSU_4/vcredist_x86.exe
set URLS[8]=https://download.microsoft.com/download/1/6/B/16B06F60-3B20-4FF2-B699-5E9B7962F9AE/VSU_4/vcredist_x64.exe

:: 2013
set URLS[9]=https://download.microsoft.com/download/9/3/F/93FCF1E7-E6A4-478B-96E7-D4B285925B00/vcredist_x86.exe
set URLS[10]=https://download.microsoft.com/download/9/3/F/93FCF1E7-E6A4-478B-96E7-D4B285925B00/vcredist_x64.exe

:: 2015-2022 最新 (統合版)
set URLS[11]=https://aka.ms/vs/17/release/vc_redist.x86.exe
set URLS[12]=https://aka.ms/vs/17/release/vc_redist.x64.exe

:: ====== ダウンロード & インストール処理 ======
for /L %%i in (1,1,12) do (
    set "URL=!URLS[%%i]!"
    if not "!URL!"=="" (
        echo.
        echo [%%i] ダウンロード中: !URL!
        echo [%%i] ダウンロード中: !URL! >> "%LOGFILE%"

        set FILE=%TEMPDIR%\vcredist_%%i.exe
        curl -L -o "!FILE!" "!URL!" >> "%LOGFILE%" 2>&1

        if exist "!FILE!" (
            echo [%%i] ダウンロード成功: !FILE!
            echo [%%i] インストール開始 >> "%LOGFILE%"
            "!FILE!" /quiet /norestart >> "%LOGFILE%" 2>&1
            if !errorlevel! equ 0 (
                echo [%%i] インストール成功 >> "%LOGFILE%"
            ) else (
                echo [%%i] インストール失敗 (エラーコード=!errorlevel!) >> "%LOGFILE%"
            )
        ) else (
            echo [%%i] ダウンロード失敗 >> "%LOGFILE%"
        )
    )
)

:: 後処理
rd /s /q "%TEMPDIR%"

echo ================================ >> "%LOGFILE%"
echo 完了 %date% %time% >> "%LOGFILE%"
echo ================================ >> "%LOGFILE%"

echo.
echo すべての処理が完了しました。ログは %LOGFILE% を確認してください。
pause
endlocal

```

