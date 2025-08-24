#!/bin/bash
# LINE Webhook Server起動スクリプト（署名検証対応）

echo "🚀 LINE Webhook サーバー起動中..."
echo "📝 Channel Secret設定中..."

# Channel Secretを環境変数に設定
export LINE_CHANNEL_SECRET="25d64f941d05535214a0462185672e91"

echo "✅ Channel Secret設定完了"
echo "🔄 サーバー起動中..."

# webhook_signature_fixed.pyを起動
cd /Users/masaya/Desktop/scraping-smacl/line_setup_tools
python webhook_signature_fixed.py
