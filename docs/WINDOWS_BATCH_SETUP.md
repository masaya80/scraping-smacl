# Windows バッチ処理セットアップガイド

## 概要

SMCL 納品リスト処理システムを Windows 環境で毎朝 10:00 に自動実行するためのセットアップガイドです。

## システム構成

### 📁 ファイル構成

```
scraping-smacl/
├── 納品リスト処理バッチ.bat          # メインバッチファイル
├── 自動実行_毎朝10時.bat            # タスクスケジューラ用バッチファイル
├── タスクスケジューラ設定.ps1        # タスクスケジューラ自動設定スクリプト
├── main.py                        # メインPythonスクリプト
├── requirements.txt               # Python依存関係
└── services/
    └── core/
        └── config.py              # 設定ファイル（ネットワークパス対応）
```

### 🗂️ データ保存構造

```
\\sv001\Userdata\納品リスト処理\      # ネットワークドライブ（メイン）
├── 20250817/                      # 日付別フォルダ（YYYYMMDD）
│   ├── downloads/                 # ダウンロードファイル
│   ├── *.xlsx                     # 生成されたExcelファイル
│   ├── *.pdf                      # 生成されたPDFファイル
│   └── *.jpg                      # 変換された画像ファイル
├── 20250818/
├── logs/
│   ├── 20250817/                  # 日付別ログフォルダ
│   └── batch_20250817_1000.log    # バッチ実行ログ
```

## セットアップ手順

### 1️⃣ 事前準備

#### Python 環境の確認

```cmd
py --version
```

Python 3.9 以上が必要です。

#### ネットワークドライブの確認

```cmd
dir \\sv001\Userdata\
```

ネットワークドライブにアクセスできることを確認してください。

### 2️⃣ 自動実行設定

#### PowerShell でタスクスケジューラを設定

1. PowerShell を**管理者として実行**
2. 以下のコマンドを実行：

```powershell
cd "C:\path\to\scraping-smacl"
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
.\タスクスケジューラ設定.ps1
```

#### 手動でタスクスケジューラを設定する場合

1. `Win + R` → `taskschd.msc` でタスクスケジューラを開く
2. 「基本タスクの作成」を選択
3. 以下の設定を入力：
   - **名前**: `SMCL納品リスト処理_毎朝10時`
   - **トリガー**: 毎日 10:00
   - **操作**: `自動実行_毎朝10時.bat`の完全パスを指定
   - **開始**: バッチファイルのあるフォルダを指定

### 3️⃣ 手動実行テスト

#### 通常実行（手動）

```cmd
cd "C:\path\to\scraping-smacl"
納品リスト処理バッチ.bat
```

#### 自動実行テスト

```cmd
cd "C:\path\to\scraping-smacl"
自動実行_毎朝10時.bat
```

## 設定オプション

### 環境変数での設定変更

#### ネットワークドライブパスの変更

```cmd
set SMCL_NETWORK_PATH=\\your-server\your-path\納品リスト処理
```

#### ネットワークドライブの無効化（ローカル保存）

```cmd
set SMCL_DISABLE_NETWORK=true
```

#### ネットワークドライブの強制有効化

```cmd
set SMCL_USE_NETWORK=true
```

### 設定ファイルでの変更

`services/core/config.py` の以下の部分を編集：

```python
# ネットワークドライブ設定
self.network_base_path = r'\\sv001\Userdata\納品リスト処理'
self.use_date_folders = True
self.date_folder_format = "%Y%m%d"  # YYYYMMDD形式
```

## ログとモニタリング

### 📊 ログファイルの場所

- **バッチ実行ログ**: `logs/batch_YYYYMMDD_HHMM.log`
- **アプリケーションログ**: `logs/YYYYMMDD/app.log`
- **ネットワークドライブログ**: `\\sv001\Userdata\納品リスト処理\logs\YYYYMMDD\`

### 🔍 ログの確認方法

```cmd
# 最新のバッチログを確認
type logs\batch_*.log | more

# 今日のアプリケーションログを確認
type logs\%date:~0,4%%date:~5,2%%date:~8,2%\app.log | more
```

### 📈 実行状況の確認

```powershell
# タスクの状態確認
Get-ScheduledTask -TaskName "SMCL納品リスト処理_毎朝10時"

# 最後の実行結果確認
Get-ScheduledTask -TaskName "SMCL納品リスト処理_毎朝10時" | Get-ScheduledTaskInfo
```

## トラブルシューティング

### ❌ よくあるエラーと対処法

#### 1. ネットワークドライブにアクセスできない

**症状**: `ネットワークドライブ接続NG` のメッセージ
**対処法**:

- ネットワーク接続を確認
- 認証情報を確認
- 環境変数 `SMCL_DISABLE_NETWORK=true` でローカル保存に切り替え

#### 2. Python 仮想環境の作成に失敗

**症状**: `仮想環境の作成に失敗しました`
**対処法**:

```cmd
# 手動で仮想環境を作成
py -m venv .venv
.venv\Scripts\activate.bat
pip install -r requirements.txt
```

#### 3. タスクスケジューラの実行に失敗

**症状**: タスクが「失敗」状態になる
**対処法**:

- バッチファイルのパスを絶対パスで指定
- 「最高の特権で実行する」にチェック
- 「ユーザーがログオンしているかどうかにかかわらず実行する」を選択

#### 4. 依存関係のインストールエラー

**症状**: `pip install` でエラーが発生
**対処法**:

```cmd
# Pythonとpipのアップグレード
py -m pip install --upgrade pip
py -m pip install --upgrade setuptools wheel

# 依存関係の再インストール
pip install -r requirements.txt --force-reinstall
```

### 🔧 デバッグモード

#### 詳細ログの有効化

環境変数を設定してデバッグモードで実行：

```cmd
set SMCL_DEBUG=true
set SMCL_LOG_LEVEL=DEBUG
納品リスト処理バッチ.bat
```

#### ヘッドレスモードの無効化（ブラウザ表示）

```cmd
set SMCL_HEADLESS=false
納品リスト処理バッチ.bat
```

## メンテナンス

### 🧹 ログファイルの自動削除

- バッチファイルは自動的に 7 日以上古いログを削除
- 10MB 以上のログファイルも自動削除

### 📦 バックアップ

重要なファイルの定期バックアップを推奨：

- `services/docs/角上魚類.xlsx` (マスタファイル)
- `services/core/config.py` (設定ファイル)
- `credentials.json` (認証ファイル)

### 🔄 システム更新

```cmd
# Gitでの更新（開発環境の場合）
git pull origin main

# 依存関係の更新
pip install -r requirements.txt --upgrade
```

## サポート

### 📞 問題が解決しない場合

1. ログファイルの内容を確認
2. エラーメッセージをコピー
3. 実行環境の情報を収集：
   - Windows バージョン
   - Python バージョン
   - ネットワーク環境

### 📚 関連ドキュメント

- `README.md`: システム全体の概要
- `AUTHENTICATION_FLOW.md`: 認証設定
- `requirements.txt`: 必要なライブラリ

