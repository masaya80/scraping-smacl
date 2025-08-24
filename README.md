# SMCL 納品リスト処理システム

角上魚類の納品リスト処理を自動化するシステムです。

## 🚀 クイックスタート

### 1) 手動実行（初回セットアップ込み）
以下を実行すると、仮想環境の作成・依存インストール・本体実行・ログ出力まで自動で行います。
```cmd
本番実行.bat
```

実行ログは `logs/run_*.log` に保存されます。最新ログの確認例:
```cmd
dir /b /od logs\run_*.log
```

### 2) 毎日 10:00 の自動実行を登録
Windows タスク スケジューラに登録します（ユーザー ログオン時に実行）。
```cmd
自動実行_登録.bat
```
タスク名は `SMCL_Delivery_Prod` です。補助コマンド:
```cmd
schtasks /Run /TN SMCL_Delivery_Prod
schtasks /Query /TN SMCL_Delivery_Prod /V /FO LIST
schtasks /Delete /TN SMCL_Delivery_Prod /F
```

### 3) UNC パス（ネットワークパス）から実行する場合
ネットワークパスから直接起動する場合は、UNC 対応バッチを利用してください。
```cmd
UNC対応で実行.bat
```

## 📁 主なファイル
```
scraping-smacl/
├── 本番実行.bat             # 本番実行（セットアップ＋ログ出力）
├── 自動実行_登録.bat       # タスクスケジューラ登録（毎日10:00）
├── UNC対応で実行.bat       # UNC 対応実行（手動用）
├── 通常実行.bat             # 通常実行（ローカルパス推奨）
├── main.py                  # メインプログラム
├── requirements.txt         # 依存定義
└── logs/                    # 実行ログ
```

## 📋 要件
- Windows 10/11
- Python 3.13 系（本番環境に既に導入済み想定）
- 必要に応じてネットワークアクセス権限

## 🔧 トラブルシューティング（簡易）
- 依存不足エラー（例: `ModuleNotFoundError`）
  - `run_prod_windows.bat` を再実行（依存が自動導入されます）
- UNC パス起因のカレントディレクトリエラー
  - `run_system_windows.bat` を使用、またはローカルディスクへコピーして実行
- 実行結果の確認
  - `logs/` の最新ログを確認

不明点があれば、最新のログファイルを添えて連絡してください。
