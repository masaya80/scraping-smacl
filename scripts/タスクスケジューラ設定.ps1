# ========================================
# SMCL 納品リスト処理システム
# Windowsタスクスケジューラ設定スクリプト
# 毎朝10:00に自動実行するタスクを作成
# ========================================

# 管理者権限チェック
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "このスクリプトは管理者権限で実行する必要があります。" -ForegroundColor Red
    Write-Host "PowerShellを管理者として実行してから再度お試しください。" -ForegroundColor Yellow
    pause
    exit 1
}

# スクリプトの場所を取得
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$BatchFile = Join-Path $ScriptPath "自動実行_毎朝10時.bat"

# バッチファイルの存在確認
if (-not (Test-Path $BatchFile)) {
    Write-Host "エラー: バッチファイルが見つかりません: $BatchFile" -ForegroundColor Red
    pause
    exit 1
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "SMCL 納品リスト処理システム" -ForegroundColor Cyan
Write-Host "タスクスケジューラ設定" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# タスク設定
$TaskName = "SMCL納品リスト処理_毎朝10時"
$TaskDescription = "SMCL納品リスト処理システムを毎朝10:00に自動実行"
$TaskPath = "\SMCL\"

Write-Host "設定内容:" -ForegroundColor Green
Write-Host "  タスク名: $TaskName" -ForegroundColor White
Write-Host "  実行時刻: 毎日 10:00" -ForegroundColor White
Write-Host "  実行ファイル: $BatchFile" -ForegroundColor White
Write-Host "  作業フォルダ: $ScriptPath" -ForegroundColor White
Write-Host ""

# 既存タスクの確認と削除
try {
    $ExistingTask = Get-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue
    if ($ExistingTask) {
        Write-Host "既存のタスクが見つかりました。削除しています..." -ForegroundColor Yellow
        Unregister-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -Confirm:$false
        Write-Host "既存タスクを削除しました。" -ForegroundColor Green
    }
} catch {
    # タスクが存在しない場合は何もしない
}

# タスクアクション（実行するコマンド）
$Action = New-ScheduledTaskAction -Execute $BatchFile -WorkingDirectory $ScriptPath

# タスクトリガー（毎日10:00に実行）
$Trigger = New-ScheduledTaskTrigger -Daily -At "10:00"

# タスク設定（システムアカウントで実行、ログオンしていなくても実行）
$Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RunOnlyIfNetworkAvailable

# タスクプリンシパル（最高権限で実行）
$Principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest

# タスクの作成
try {
    Write-Host "タスクスケジューラにタスクを登録しています..." -ForegroundColor Yellow
    
    Register-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -Action $Action -Trigger $Trigger -Settings $Settings -Principal $Principal -Description $TaskDescription
    
    Write-Host "✅ タスクの登録が完了しました！" -ForegroundColor Green
    Write-Host ""
    
    # 登録されたタスクの確認
    $RegisteredTask = Get-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath
    Write-Host "登録されたタスク情報:" -ForegroundColor Cyan
    Write-Host "  タスク名: $($RegisteredTask.TaskName)" -ForegroundColor White
    Write-Host "  状態: $($RegisteredTask.State)" -ForegroundColor White
    Write-Host "  次回実行予定: $(Get-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath | Get-ScheduledTaskInfo | Select-Object -ExpandProperty NextRunTime)" -ForegroundColor White
    
} catch {
    Write-Host "❌ タスクの登録に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
    pause
    exit 1
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "設定完了" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "📅 毎朝10:00に自動実行されます" -ForegroundColor Green
Write-Host "🔧 タスクスケジューラで確認・変更できます" -ForegroundColor Yellow
Write-Host "📊 実行ログは logs フォルダに保存されます" -ForegroundColor Yellow
Write-Host ""
Write-Host "手動でテスト実行する場合:" -ForegroundColor Cyan
Write-Host "  Start-ScheduledTask -TaskName '$TaskName' -TaskPath '$TaskPath'" -ForegroundColor White
Write-Host ""

# テスト実行の確認
$TestRun = Read-Host "今すぐテスト実行しますか？ (y/N)"
if ($TestRun -eq "y" -or $TestRun -eq "Y") {
    Write-Host ""
    Write-Host "テスト実行を開始します..." -ForegroundColor Yellow
    try {
        Start-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath
        Write-Host "✅ テスト実行を開始しました" -ForegroundColor Green
        Write-Host "実行状況はタスクスケジューラまたはログファイルで確認してください" -ForegroundColor Yellow
    } catch {
        Write-Host "❌ テスト実行の開始に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "設定が完了しました。何かキーを押すと終了します..." -ForegroundColor Green
pause
