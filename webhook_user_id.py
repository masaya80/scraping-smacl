#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LINE User ID 取得用 簡易Webhookサーバー
"""

from flask import Flask, request, jsonify
import json

app = Flask(__name__)

@app.route('/webhook', methods=['POST'])
def webhook():
    """Webhookエンドポイント"""
    try:
        data = request.get_json()
        
        # イベント情報を表示
        print("=" * 50)
        print("📨 Webhook受信:")
        print(json.dumps(data, indent=2, ensure_ascii=False))
        
        # User IDを抽出
        if 'events' in data:
            for event in data['events']:
                if 'source' in event and 'userId' in event['source']:
                    user_id = event['source']['userId']
                    print("=" * 50)
                    print(f"🎯 User ID発見: {user_id}")
                    print("=" * 50)
                    print("以下のコマンドで設定してください:")
                    print(f'export LINE_USER_ID="{user_id}"')
                    print("=" * 50)
        
        return jsonify({"status": "ok"})
        
    except Exception as e:
        print(f"エラー: {e}")
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    print("🚀 User ID取得用Webhookサーバー開始")
    print("📝 設定:")
    print("  1. ngrokを起動: ngrok http 5000")
    print("  2. ngrok URLをLINE DevelopersのWebhook URLに設定")
    print("  3. LINEでBotにメッセージ送信")
    print()
    app.run(debug=True, port=5000)
