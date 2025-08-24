# Google Drive 認証フロー詳細ガイド

Google Drive API の認証について、ファイルの役割と認証フローを詳しく説明します。

## 📁 認証関連ファイルの役割

### credentials.json（手動ダウンロード）✅

- **役割**: Google Cloud Console からダウンロードする認証情報
- **取得先**: Google Cloud Console → 認証情報 → OAuth 2.0 クライアント ID
- **内容**: クライアント ID、クライアントシークレットなど
- **必要な操作**: **手動でダウンロード**して`credentials.json`として保存

### token.json（自動生成）🤖

- **役割**: 認証完了後にシステムが自動生成するアクセストークン
- **取得先**: **システムが自動生成**（ダウンロード不要）
- **内容**: アクセストークン、リフレッシュトークンなど
- **必要な操作**: **何もしない**（自動で作成される）

## 🔄 認証フロー詳細

### ステップ 1: credentials.json 準備

```bash
# Google Cloud Consoleからダウンロード
# → プロジェクトルートに「credentials.json」として配置
```

### ステップ 2: 初回認証実行

```bash
python check_config.py  # 設定確認
python main.py          # システム実行
```

### ステップ 3: 認証 URL 表示

```
認証が必要です。以下のURLにアクセスしてください:
https://accounts.google.com/o/oauth2/auth?client_id=123456789...
```

### ステップ 4: ブラウザ認証

1. 表示された URL をブラウザで開く
2. Google アカウントでログイン
3. アプリケーションの権限を**「許可」**

### ステップ 5: 認証コード取得

```
認証が完了しました。以下のコードをコピーしてください:
4/0AfgeXvqL-ABC123DEF456...
```

### ステップ 6: 認証コード設定

#### 方法 A: 環境変数（推奨）

```bash
export GOOGLE_AUTH_CODE="4/0AfgeXvqL-ABC123DEF456..."
python main.py
```

#### 方法 B: 手動入力

```bash
python main.py
# → 認証コードの入力を求められる
# → コピーした認証コードを貼り付け
```

### ステップ 7: token.json 自動生成 ✅

```
✅ Google Drive サービス初期化完了
📄 token.json が自動生成されました
```

## 🗂️ ファイル構成例

### 初期状態（credentials.json のみ）

```
scraping-smacl/
├── credentials.json     ← 手動配置
├── main.py
└── services/
```

### 認証完了後（token.json 自動生成）

```
scraping-smacl/
├── credentials.json     ← 手動配置
├── token.json          ← 自動生成 ✅
├── main.py
└── services/
```

## ⚙️ 認証方法の比較

| 方法         | 手動入力 | 環境変数 | config.py 直接設定 |
| ------------ | -------- | -------- | ------------------ |
| 手軽さ       | ⭐⭐     | ⭐⭐⭐   | ⭐⭐⭐             |
| 自動化       | ❌       | ✅       | ✅                 |
| セキュリティ | ⭐⭐     | ⭐⭐⭐   | ⭐                 |
| 推奨度       | 開発時   | 本番     | テスト             |

## 🔄 再認証が必要な場合

### いつ再認証が必要？

- token.json を削除した時
- 認証トークンの有効期限切れ
- 権限スコープを変更した時

### 再認証手順

```bash
# 1. 古いトークンを削除
rm token.json

# 2. システム再実行（認証フローが自動開始）
python main.py
```

## 🛡️ セキュリティ注意事項

### Git に含めてはいけないファイル

```bash
# .gitignore に追加
credentials.json  # クライアント認証情報
token.json       # アクセストークン
```

### 安全な管理方法

```bash
# 環境変数で管理（推奨）
export GOOGLE_DRIVE_CREDENTIALS="/secure/path/credentials.json"

# または専用ディレクトリ
mkdir ~/.smcl_config/
mv credentials.json ~/.smcl_config/
export GOOGLE_DRIVE_CREDENTIALS="$HOME/.smcl_config/credentials.json"
```

## 🔍 トラブルシューティング

### Q: token.json はどこからダウンロードしますか？

**A**: ダウンロードしません！システムが自動生成します。

### Q: 認証コードはどこで取得しますか？

**A**: ブラウザで認証 URL→「許可」→ 認証コード表示

### Q: 何度も認証が必要になります

**A**: token.json が削除されている可能性があります。.gitignore に追加してください。

### Q: 認証に失敗します

**A**:

1. credentials.json が正しく配置されているか確認
2. Google Drive API が有効化されているか確認
3. OAuth 同意画面が正しく設定されているか確認

## 📞 サポートコマンド

```bash
# 設定状況確認
python check_config.py

# Google Drive接続テスト
python -c "
from services.core.google_drive_uploader import GoogleDriveUploader
uploader = GoogleDriveUploader()
print('利用可能:', uploader.is_available())
print('接続テスト:', uploader.test_connection())
"
```
