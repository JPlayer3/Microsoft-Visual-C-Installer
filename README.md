ChatGPTに作成してもらった。

使い方（要点）
・管理者として PowerShell を起動（「管理者として実行」）。
・下のスクリプト全文を Install-VCRedists.ps1 などとして保存。
・PowerShell で保存したスクリプトのあるフォルダに移動し、.\Install-VCRedists.ps1 を実行。

注意
・インストーラはサイレント動作（例：/install /quiet /norestart）で実行します。異なるインストーラは別のスイッチが必要な場合があるので、その場合は InstallArgs を個別に上書きしてください。
・実行には 管理者権限 が必要です。
・ネットワーク環境や URL の有効期限によりダウンロードに失敗する場合があります。その場合は公式 Microsoft のページから該当の再頒布パッケージの URL を取得して -$InstallList に追加してください。

スクリプト（そのままコピペして実行できます）
```
<#
.SYNOPSIS
  複数の Visual C++ 再頒布パッケージを一括ダウンロード & サイレントインストールし、ログを出力するスクリプト。

.NOTES
  実行は管理者権限で行ってください。
  必要に応じて $InstallList に URL とインストーラ引数を追加・編集してください。
#>

# --- 設定 ---
# ログファイル
$logFolder = "$env:ProgramData\VCRedistInstaller"
if (-not (Test-Path $logFolder)) { New-Item -Path $logFolder -ItemType Directory -Force | Out-Null }
$timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
$logFile = Join-Path $logFolder "VCRedistInstall_$timestamp.log"

# 一時ダウンロードフォルダ
$tempDir = Join-Path $env:TEMP "VCRedistInstaller_$timestamp"
New-Item -Path $tempDir -ItemType Directory -Force | Out-Null

# インストール対象のリスト
# 各アイテムは @{ Name='表示名'; Url='ダウンロード先'; InstallArgs='インストール時引数' }
# aka.ms リダイレクトは最新版の Visual C++ 再頒布パッケージ（2015-2022 等）へ飛びます
$InstallList = @(
    @{ Name='VC++ Redistributable x64 (Latest 2015-2022)'; Url='https://aka.ms/vs/17/release/vc_redist.x64.exe'; InstallArgs='/install /quiet /norestart' },
    @{ Name='VC++ Redistributable x86 (Latest 2015-2022)'; Url='https://aka.ms/vs/17/release/vc_redist.x86.exe'; InstallArgs='/install /quiet /norestart' }
    # 例: 古いバージョンを追加する場合はここに行を追加
    # @{ Name='VC++ 2013 x86'; Url='https://download.microsoft.com/.../vcredist_x86.exe'; InstallArgs='/quiet /norestart' }
)

# --- ログ関数 ---
function Log {
    param($Text, $Level='INFO')
    $time = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $line = "[$time] [$Level] $Text"
    $line | Tee-Object -FilePath $logFile -Append
    Write-Output $line
}

# --- 管理者権限チェック ---
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "このスクリプトは管理者権限で実行する必要があります。PowerShellを「管理者として実行」してから再試行してください。"
    exit 1
}

Log "開始: VCRedist 一括インストールスクリプト" "START"
Log "ログファイル: $logFile"

# --- ダウンロード関数（リトライ対応） ---
function Download-File {
    param($Url, $OutPath, $MaxAttempts = 3)
    $attempt = 0
    while ($attempt -lt $MaxAttempts) {
        $attempt++
        try {
            Log "ダウンロード試行 ($attempt/$MaxAttempts): $Url"
            Invoke-WebRequest -Uri $Url -OutFile $OutPath -UseBasicParsing -TimeoutSec 120
            if (Test-Path $OutPath -PathType Leaf) {
                Log "ダウンロード成功: $OutPath"
                return $true
            }
        } catch {
            Log "ダウンロード失敗 (attempt $attempt): $($_.Exception.Message)" "WARN"
            Start-Sleep -Seconds (5 * $attempt)
        }
    }
    Log "ダウンロード失敗（最大試行回数）: $Url" "ERROR"
    return $false
}

# --- インストール処理 ---
foreach ($item in $InstallList) {
    $name = $item.Name
    $url = $item.Url
    $args = $item.InstallArgs
    $safeName = ($name -replace '[^0-9A-Za-z\-_\.]', '_')
    $installerPath = Join-Path $tempDir ($safeName + "_" + [IO.Path]::GetFileName($url))

    Log "処理対象: $name"
    # ダウンロード
    $dlOk = Download-File -Url $url -OutPath $installerPath -MaxAttempts 3
    if (-not $dlOk) {
        Log "スキップ: ダウンロードできないため $name のインストールを中止します。" "ERROR"
        continue
    }

    # 実行許可チェック（ブロック解除）
    try {
        Unblock-File -Path $installerPath -ErrorAction SilentlyContinue
    } catch {
        # ignore
    }

    # インストール実行
    Log "インストール実行: $name"
    try {
        $startInfo = @{
            FilePath = $installerPath
            ArgumentList = $args
            Wait = $true
            NoNewWindow = $true
            RedirectStandardOutput = $true
            RedirectStandardError = $true
        }
        # Start-Process ではリダイレクトの捕捉が面倒なため、Process クラスを利用
        $proc = New-Object System.Diagnostics.Process
        $proc.StartInfo.FileName = $installerPath
        $proc.StartInfo.Arguments = $args
        $proc.StartInfo.UseShellExecute = $false
        $proc.StartInfo.RedirectStandardOutput = $true
        $proc.StartInfo.RedirectStandardError = $true
        $proc.StartInfo.CreateNoWindow = $true
        $proc.Start() | Out-Null

        $stdOut = $proc.StandardOutput.ReadToEnd()
        $stdErr = $proc.StandardError.ReadToEnd()
        $proc.WaitForExit()
        $exitCode = $proc.ExitCode

        if ($stdOut) { Log "標準出力: $stdOut" }
        if ($stdErr) { Log "標準エラー: $stdErr" }

        Log "終了コード ($name): $exitCode"
        if ($exitCode -eq 0) {
            Log "インストール成功: $name"
        } else {
            Log "インストール失敗（非0終了）: $name" "ERROR"
        }
    } catch {
        Log "インストール実行中に例外発生: $($_.Exception.Message)" "ERROR"
    }
}

# --- 後片付け ---
try {
    Log "インストール処理完了。ダウンロード一時フォルダ: $tempDir"
    # 一時フォルダを残す/削除するか選べるように
    $keepTemp = $false
    if (-not $keepTemp) {
        Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        Log "一時フォルダを削除しました。"
    } else {
        Log "一時フォルダを保持します: $tempDir"
    }
} catch {
    Log "後片付けでエラー: $($_.Exception.Message)" "WARN"
}

Log "完了: VCRedist 一括インストールスクリプト 終了" "END"
```
補足（トラブルシューティング）
・インストールが失敗する場合はログ ($logFile) をまず確認してください。標準出力／標準エラーを記録しています。
・特定のインストーラだけ引数が違う場合は、$InstallList のその要素の InstallArgs を該当インストーラに合わせて修正してください（例：古い vcredist は /q /norestart など）。
・オフライン環境の場合は、あらかじめインストーラをダウンロードしてローカルパスを Url の代わりに指定できます（file:// 形式でもOK）。
・サイレントインストール後に再起動が必要な場合は、/norestart を外すか、手動で再起動してください。
