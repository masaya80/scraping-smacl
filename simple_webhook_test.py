#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
シンプルなLINE Webhook テストサーバー
"""

from flask import Flask, request, jsonify
import json

app = Flask(__name__)

@app.route('/', methods=['GET'])
def index():
    """ルートエンドポイント"""
    return jsonify({
        "message": "LINE Webhook Test Server",
        "status": "running",
        "endpoints": ["/webhook"]
    })

@app.route('/webhook', methods=['POST', 'GET'])
def webhook():
    """Webhookエンドポイント"""
    try:
        if request.method == 'GET':
            return jsonify({
                "message": "Webhook endpoint working",
                "method": "GET"
            })
        
        # POSTリクエストの処理
        print("=" * 50)
        print("📨 Webhook受信!")
        
        # ヘッダー情報
        print("📥 Headers:")
        for key, value in request.headers.items():
            print(f"  {key}: {value}")
        
        # リクエストボディ
        data = request.get_json()
        print("📄 Body:")
        print(json.dumps(data, indent=2, ensure_ascii=False))
        
        # Group ID抽出
        if data and 'events' in data:
            for event in data['events']:
                if 'source' in event:
                    source = event['source']
                    if source.get('type') == 'group':
                        group_id = source.get('groupId')
                        user_id = source.get('userId')
                        
                        print("=" * 50)
                        print(f"🎯 Group ID: {group_id}")
                        print(f"👤 User ID: {user_id}")
                        print("=" * 50)
                        print("📋 設定用コマンド:")
                        print(f'export LINE_GROUP_ID="{group_id}"')
                        print("=" * 50)
        
        return jsonify({"status": "ok"})
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    print("🚀 Simple Webhook Test Server")
    print("📍 URL: http://localhost:5001")
    print("📍 Webhook: http://localhost:5001/webhook")
    app.run(debug=True, port=5001)
