# LINE Bot 設定ツール

このディレクトリには、SMCL 納品リスト処理システムの LINE Bot 設定に関するツールが含まれています。

## 📋 概要

LINE Bot を使用してシステムの処理結果をグループチャットに通知するための設定ツール群です。

## 🛠️ ツール一覧

### 1. `get_line_group_id.py`

**機能**: LINE グループ ID の取得とグループチャット設定の確認

```bash
python get_line_group_id.py
```

**用途**:

- LINE Bot のグループチャット設定状況を確認
- グループチャット接続テストの実行
- グループ設定手順の表示
- Webhook 用サーバーの作成

### 2. `check_line_settings.py`

**機能**: LINE Bot 設定の詳細確認

```bash
python check_line_settings.py
```

**用途**:

- 現在の LINE Bot 設定を詳細表示
- LINE Developers Console 設定の確認項目表示
- 設定手順のガイド表示

### 3. `simple_group_test.py`

**機能**: 既知のグループ ID でのテスト送信

```bash
python simple_group_test.py
```

**用途**:

- 複数の既知グループ ID でテスト送信
- 正しいグループ ID の特定
- 設定ファイルの自動更新

### 4. `webhook_group_id.py`

**機能**: Group ID 取得用 Webhook サーバー

```bash
python webhook_group_id.py
```

**用途**:

- LINE Webhook で Group ID を取得
- リアルタイムでの Webhook イベント監視
- Group ID と User ID の表示

### 5. `webhook_user_id.py`

**機能**: User ID 取得用 Webhook サーバー

```bash
python webhook_user_id.py
```

**用途**:

- 個人チャット用の User ID 取得
- Webhook イベントの監視

### 6. `simple_webhook_test.py`

**機能**: 簡単な Webhook テストサーバー

```bash
python simple_webhook_test.py
```

**用途**:

- Webhook 動作の基本テスト
- LINE Developers Console 設定確認

## 🚀 セットアップ手順

### 1. LINE Developers Console 設定

1. [LINE Developers Console](https://developers.line.biz/console/) にアクセス
2. プロバイダー「角上通知」を選択
3. Channel ID: `2007929536` のチャンネルを選択
4. Messaging API 設定:
   - `Allow bot to join group chats`: **ON**
   - `Use webhook`: **ON**（Group ID 確認時のみ）
   - `Auto-reply messages`: **OFF**
   - `Greeting messages`: **OFF**

### 2. LINE グループ作成

1. LINE アプリでグループを作成
2. グループ名を設定（例: 角上物流通知）
3. 必要なメンバーを追加

### 3. Bot をグループに追加

1. グループチャット画面を開く
2. 右上の設定アイコン → メンバー
3. 「招待」→「QR コードで招待」
4. LINE Developers の QR コードをスキャン
5. または Bot basic ID で検索して追加

### 4. Group ID 取得

#### 方法 A: Webhook を使用（推奨）

```bash
# 1. Webhookサーバーを起動
python webhook_group_id.py

# 2. 別ターミナルでngrokを起動
ngrok http 5000

# 3. ngrok URLをLINE Webhook URLに設定
# 4. グループでBotにメッセージ送信
# 5. Group IDが表示されるのでコピー
```

#### 方法 B: 既知 ID でテスト

```bash
python simple_group_test.py
```

### 5. 環境変数設定

取得した Group ID を環境変数に設定：

```bash
export LINE_GROUP_ID="取得したグループID"
export LINE_CHANNEL_ACCESS_TOKEN="チャンネルアクセストークン"
```

### 6. 設定確認

```bash
python check_line_settings.py
```

## 🔧 トラブルシューティング

### よくある問題

#### 1. Group ID が取得できない

- Webhook URL が正しく設定されているか確認
- ngrok が正常に動作しているか確認
- グループで Bot にメッセージを送信したか確認

#### 2. メッセージが送信されない

- Channel Access Token が正しいか確認
- Group ID が正しいか確認
- Bot がグループメンバーになっているか確認

#### 3. Webhook が反応しない

- LINE Developers Console で Webhook が有効になっているか確認
- ngrok URL が正しく設定されているか確認
- ファイアウォールでポートがブロックされていないか確認

## 📝 設定ファイル

設定は `services/core/config.py` で管理されています：

```python
# LINE Bot設定
self.line_channel_access_token = os.getenv('LINE_CHANNEL_ACCESS_TOKEN', 'default_token')
self.line_channel_secret = os.getenv('LINE_CHANNEL_SECRET', 'default_secret')
self.line_channel_id = os.getenv('LINE_CHANNEL_ID', '2007929536')
self.line_group_id = os.getenv('LINE_GROUP_ID', 'default_group_id')
```

## 💡 使用例

### 基本的な設定フロー

```bash
# 1. 設定状況確認
python check_line_settings.py

# 2. Group ID取得
python get_line_group_id.py

# 3. 環境変数設定
export LINE_GROUP_ID="Ca23b038e0208192c6efe1640d471e977"

# 4. テスト送信
python simple_group_test.py

# 5. 最終確認
python check_line_settings.py
```

## 📚 関連ドキュメント

- `../AUTHENTICATION_FLOW.md`: 認証フロー詳細
- `../README.md`: システム全体概要
- `../services/notification/line_bot.py`: LINE Bot 実装

## 🆘 サポート

問題が発生した場合は、以下の情報を収集してください：

1. エラーメッセージ
2. 実行環境（OS、Python バージョン）
3. LINE Developers Console 設定スクリーンショット
4. 環境変数設定状況

---

**注意**: このツール群は開発・設定用です。本番運用では環境変数による設定を使用してください。
