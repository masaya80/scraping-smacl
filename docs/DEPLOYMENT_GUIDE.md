# SMCL 納品リスト処理システム デプロイガイド

## 📋 概要

SMCL 納品リスト処理システムを Windows 本番環境にデプロイし、毎朝 10:00 の自動実行を設定するための完全ガイドです。

## 🎯 デプロイ対象

- **システム名**: SMCL 納品リスト処理システム
- **実行環境**: Windows Server/Windows 10/11
- **実行スケジュール**: 毎朝 10:00 自動実行
- **データ保存先**: `\\sv001\Userdata\納品リスト処理\`

## 📋 事前準備チェックリスト

### ✅ システム要件

- [ ] Windows 10/11 または Windows Server 2016 以降
- [ ] Python 3.9 以降がインストール済み
- [ ] ネットワークドライブ `\\sv001\Userdata\` へのアクセス権限
- [ ] 管理者権限でのログイン可能
- [ ] インターネット接続（依存関係ダウンロード用）

### ✅ 必要な認証情報

- [ ] SMCL ログイン情報（企業コード、ユーザー ID、パスワード）
- [ ] LINE Bot アクセストークン
- [ ] LINE グループ ID
- [ ] Google Drive API 認証ファイル（credentials.json）

### ✅ ネットワーク環境

- [ ] `\\sv001\Userdata\納品リスト処理\` フォルダの作成権限
- [ ] SMCL ウェブサイトへのアクセス可能
- [ ] LINE API へのアクセス可能
- [ ] Google Drive API へのアクセス可能

## 🚀 デプロイ手順

### ステップ 1: システムファイルの配置

#### 1.1 デプロイディレクトリの作成

```cmd
mkdir C:\SMCL_System
cd C:\SMCL_System
```

#### 1.2 システムファイルのコピー

以下のファイルを `C:\SMCL_System\` にコピー：

```
C:\SMCL_System\
├── main.py                          # メインプログラム
├── requirements.txt                 # Python依存関係
├── 納品リスト処理バッチ.bat           # メインバッチファイル
├── 自動実行_毎朝10時.bat            # 自動実行用バッチ
├── タスクスケジューラ設定.ps1        # タスク設定スクリプト
├── services/                        # サービスモジュール
│   ├── __init__.py
│   ├── core/
│   │   ├── __init__.py
│   │   ├── config.py               # 設定ファイル
│   │   ├── logger.py
│   │   └── models.py
│   ├── data_processing/
│   ├── notification/
│   ├── scraping/
│   └── docs/
│       └── 角上魚類.xlsx            # マスタExcelファイル
└── credentials.json                 # Google Drive認証ファイル
```

### ステップ 2: 環境設定

#### 2.1 Python 環境の確認

```cmd
python --version
# または
py --version
```

**期待される出力**: `Python 3.9.x` 以降

#### 2.2 ネットワークドライブの確認

```cmd
dir \\sv001\Userdata\
```

#### 2.3 ネットワークフォルダの作成

```cmd
mkdir "\\sv001\Userdata\納品リスト処理"
mkdir "\\sv001\Userdata\納品リスト処理\logs"
```

### ステップ 3: 設定ファイルの編集

#### 3.1 SMCL 認証情報の設定

`services/core/config.py` を編集：

```python
# SMCL設定
self.target_url = "https://smclweb.cs-cxchange.net/smcl/view/lin/EDS001OLIN0000.aspx"
self.corp_code = "I26S"              # 企業コード
self.login_id = "0600200"            # ユーザーID
self.password = "toichi04"           # パスワード
```

#### 3.2 環境変数の設定（推奨）

```cmd
# LINE Bot設定
setx LINE_CHANNEL_ACCESS_TOKEN "your_line_access_token"
setx LINE_GROUP_ID "your_line_group_id"

# ネットワークパス設定（必要に応じて）
setx SMCL_NETWORK_PATH "\\sv001\Userdata\納品リスト処理"

# ログレベル設定
setx SMCL_LOG_LEVEL "INFO"
```

### ステップ 4: 依存関係のインストール

#### 4.1 仮想環境の作成とアクティベート

```cmd
cd C:\SMCL_System
py -m venv .venv
.venv\Scripts\activate.bat
```

#### 4.2 依存関係のインストール

```cmd
pip install --upgrade pip
pip install -r requirements.txt
```

### ステップ 5: 初期テスト

#### 5.1 設定確認テスト

```cmd
cd C:\SMCL_System
python -c "from services.core.config import Config; c=Config(); c.print_configuration_status()"
```

#### 5.2 手動実行テスト

```cmd
cd C:\SMCL_System
納品リスト処理バッチ.bat
```

**確認ポイント**:

- [ ] エラーなく実行完了
- [ ] ネットワークドライブへのファイル保存
- [ ] LINE 通知の送信
- [ ] ログファイルの生成

### ステップ 6: 自動実行の設定

#### 6.1 PowerShell での自動設定（推奨）

```powershell
# PowerShellを管理者として実行
cd C:\SMCL_System
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
.\タスクスケジューラ設定.ps1
```

#### 6.2 手動でのタスクスケジューラ設定

1. `Win + R` → `taskschd.msc`
2. 「基本タスクの作成」
3. 設定内容：
   - **名前**: `SMCL納品リスト処理_毎朝10時`
   - **説明**: `SMCL納品リスト処理システムの自動実行`
   - **トリガー**: 毎日 10:00
   - **操作**: `C:\SMCL_System\自動実行_毎朝10時.bat`
   - **開始場所**: `C:\SMCL_System`
   - **実行アカウント**: `SYSTEM`
   - **最高の特権で実行**: ✅

### ステップ 7: 本番テスト

#### 7.1 タスクスケジューラでのテスト実行

```powershell
# PowerShellで実行
Start-ScheduledTask -TaskName "SMCL納品リスト処理_毎朝10時" -TaskPath "\SMCL\"
```

#### 7.2 実行結果の確認

```cmd
# ログファイルの確認
type C:\SMCL_System\logs\batch_*.log

# ネットワークドライブの確認
dir "\\sv001\Userdata\納品リスト処理\%date:~0,4%%date:~5,2%%date:~8,2%"
```

## 📊 デプロイ後の確認項目

### ✅ 機能確認チェックリスト

- [ ] **スクレイピング**: SMCL サイトからのデータ取得
- [ ] **データ処理**: CSV からのデータ抽出
- [ ] **Excel 生成**: マスタファイルとの付け合わせ
- [ ] **PDF 生成**: 配車表・出庫依頼書の作成
- [ ] **画像変換**: PDF から画像への変換
- [ ] **LINE 通知**: 処理結果の通知送信
- [ ] **ファイル保存**: ネットワークドライブへの保存
- [ ] **ログ出力**: 実行ログの記録

### ✅ 自動実行確認

- [ ] タスクスケジューラでのタスク登録
- [ ] 次回実行時刻の設定確認
- [ ] テスト実行の成功
- [ ] ログファイルの自動生成

## 🔧 運用・メンテナンス

### 📅 日次確認項目

- [ ] 実行ログの確認（エラーの有無）
- [ ] 生成ファイルの確認
- [ ] LINE 通知の受信確認
- [ ] ネットワークドライブの容量確認

### 📊 週次確認項目

- [ ] ログファイルの整理（古いファイルの削除）
- [ ] システムリソースの確認
- [ ] バックアップファイルの確認

### 🔄 月次メンテナンス

- [ ] 依存関係の更新確認
- [ ] マスタファイルの更新
- [ ] 認証情報の有効期限確認
- [ ] システムログの分析

## 🚨 トラブルシューティング

### よくある問題と対処法

#### 1. ネットワークドライブアクセスエラー

**症状**: `ネットワークドライブ接続NG`

```cmd
# 対処法1: 手動でネットワークドライブをマップ
net use Z: \\sv001\Userdata /persistent:yes

# 対処法2: ローカル保存に切り替え
setx SMCL_DISABLE_NETWORK true
```

#### 2. Python 依存関係エラー

**症状**: `ModuleNotFoundError`

```cmd
# 対処法: 依存関係の再インストール
cd C:\SMCL_System
.venv\Scripts\activate.bat
pip install -r requirements.txt --force-reinstall
```

#### 3. タスクスケジューラ実行失敗

**症状**: タスクが「失敗」状態

```cmd
# 対処法1: バッチファイルの直接実行テスト
cd C:\SMCL_System
自動実行_毎朝10時.bat

# 対処法2: タスクの再作成
schtasks /delete /tn "SMCL納品リスト処理_毎朝10時" /f
# PowerShellで再設定
```

#### 4. SMCL 認証エラー

**症状**: ログイン失敗

- 認証情報の確認
- SMCL サイトの手動ログインテスト
- ネットワーク接続の確認

#### 5. LINE 通知エラー

**症状**: 通知が送信されない

```cmd
# 環境変数の確認
echo %LINE_CHANNEL_ACCESS_TOKEN%
echo %LINE_GROUP_ID%

# 手動テスト
python -c "from services.notification.line_bot import LineBotNotifier; LineBotNotifier().send_test_message()"
```

## 📋 デプロイ完了チェックシート

### 🎯 最終確認

- [ ] すべてのファイルが正しい場所に配置されている
- [ ] 環境変数が正しく設定されている
- [ ] 依存関係がすべてインストールされている
- [ ] 手動実行テストが成功している
- [ ] タスクスケジューラが正しく設定されている
- [ ] 自動実行テストが成功している
- [ ] ネットワークドライブにファイルが保存されている
- [ ] LINE 通知が正常に送信されている
- [ ] ログファイルが正常に生成されている

### 📞 サポート連絡先

デプロイ中に問題が発生した場合：

1. エラーログを収集
2. 実行環境情報を記録
3. 開発チームに連絡

### 📚 関連ドキュメント

- `WINDOWS_BATCH_SETUP.md`: Windows バッチ処理詳細ガイド
- `README.md`: システム概要
- `AUTHENTICATION_FLOW.md`: 認証設定詳細

---

## 🎉 デプロイ完了

このガイドに従ってデプロイが完了すると、システムは毎朝 10:00 に自動実行され、処理結果がネットワークドライブに保存され、LINE 通知が送信されます。

**次回実行予定**: 翌朝 10:00
**データ保存先**: `\\sv001\Userdata\納品リスト処理\YYYYMMDD\`
**ログ確認**: `C:\SMCL_System\logs\` および ネットワークドライブ
