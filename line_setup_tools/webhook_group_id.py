#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LINE Group ID 取得用 Webhookサーバー
"""

from flask import Flask, request, jsonify
import json
import hashlib
import hmac
import base64

app = Flask(__name__)

@app.route('/', methods=['GET'])
def index():
    """ルートエンドポイント（動作確認用）"""
    return jsonify({
        "message": "LINE Group ID取得用Webhookサーバー",
        "status": "running",
        "endpoint": "/webhook"
    })

@app.route('/webhook', methods=['POST', 'GET'])
def webhook():
    """Webhookエンドポイント"""
    try:
        # GETリクエストの場合（動作確認用）
        if request.method == 'GET':
            print("=" * 60)
            print("🔍 GET request received - webhook endpoint is working")
            print("=" * 60)
            return jsonify({
                "message": "Webhook endpoint is working",
                "method": "GET",
                "status": "ok"
            })
        
        # リクエストヘッダーをログ出力
        print("=" * 60)
        print("📥 リクエストヘッダー:")
        for key, value in request.headers.items():
            print(f"  {key}: {value}")
        
        # リクエストボディを取得
        body = request.get_data(as_text=True)
        print("=" * 60)
        print("📄 リクエストボディ (raw):")
        print(body)
        
        data = request.get_json()
        
        print("=" * 60)
        print("📨 Webhook受信:")
        print(json.dumps(data, indent=2, ensure_ascii=False))
        
        # Group ID と User ID を抽出
        if 'events' in data:
            for event in data['events']:
                if 'source' in event:
                    source = event['source']
                    source_type = source.get('type', 'unknown')
                    
                    print("=" * 60)
                    print(f"📍 送信元タイプ: {source_type}")
                    
                    if source_type == 'group':
                        group_id = source.get('groupId')
                        user_id = source.get('userId')
                        
                        print(f"👥 Group ID: {group_id}")
                        print(f"👤 User ID: {user_id}")
                        print("=" * 60)
                        print("🎯 グループ送信用設定コマンド:")
                        print(f'export LINE_GROUP_ID="{group_id}"')
                        print("=" * 60)
                        
                    elif source_type == 'user':
                        user_id = source.get('userId')
                        print(f"👤 User ID: {user_id}")
                        print("=" * 60)
                        print("⚠️ このシステムはグループチャット専用です")
                        print("グループでメッセージを送信してください")
                        print("=" * 60)
        
        # 正常レスポンスを返す
        response = jsonify({"status": "ok"})
        response.status_code = 200
        print("=" * 60)
        print("✅ 200 OKレスポンスを返しました")
        print("=" * 60)
        return response
        
    except Exception as e:
        print("=" * 60)
        print(f"❌ エラー発生: {e}")
        print(f"❌ エラータイプ: {type(e).__name__}")
        import traceback
        print(f"❌ スタックトレース:")
        print(traceback.format_exc())
        print("=" * 60)
        
        # エラー時も200を返す（LINEプラットフォーム用）
        error_response = jsonify({"status": "error", "message": str(e)})
        error_response.status_code = 200
        return error_response

if __name__ == "__main__":
    print("🚀 Group ID取得用Webhookサーバー開始")
    print("📝 使用方法:")
    print("  1. 別ターミナルで: ngrok http 5000")
    print("  2. ngrok URLをLINE Webhook URLに設定")
    print("  3. LINEグループでBotにメッセージ送信")
    print("  4. Group IDが表示されるのでコピー")
    print()
    app.run(debug=True, port=5000)
