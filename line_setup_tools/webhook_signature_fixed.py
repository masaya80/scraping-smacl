#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LINE Webhook 署名検証対応版 Group ID取得サーバー
"""

from flask import Flask, request, jsonify, abort
import json
import hashlib
import hmac
import base64
import os

app = Flask(__name__)

# LINE Channel Secretを環境変数から取得
LINE_CHANNEL_SECRET = os.getenv('LINE_CHANNEL_SECRET', '25d64f941d05535214a0462185672e91')

def validate_signature(body: bytes, signature: str) -> bool:
    """LINE Platform署名検証"""
    if not signature:
        return False
    
    try:
        # 署名を生成
        expected_signature = base64.b64encode(
            hmac.new(
                LINE_CHANNEL_SECRET.encode('utf-8'),
                body,
                hashlib.sha256
            ).digest()
        ).decode()
        
        return hmac.compare_digest(signature, expected_signature)
    except Exception as e:
        print(f"❌ 署名検証エラー: {e}")
        return False

@app.route('/', methods=['GET'])
def index():
    """ルートエンドポイント（動作確認用）"""
    return jsonify({
        "message": "LINE Group ID取得用Webhookサーバー（署名検証対応）",
        "status": "running",
        "endpoint": "/webhook",
        "channel_secret_configured": bool(LINE_CHANNEL_SECRET and LINE_CHANNEL_SECRET != 'your_channel_secret_here')
    })

@app.route('/webhook', methods=['POST', 'GET'])
def webhook():
    """Webhookエンドポイント（署名検証対応）"""
    try:
        # GETリクエストの場合（動作確認用）
        if request.method == 'GET':
            print("=" * 60)
            print("🔍 GET request - webhook endpoint is working")
            print(f"🔑 Channel Secret設定状況: {'✅ 設定済み' if LINE_CHANNEL_SECRET and LINE_CHANNEL_SECRET != 'your_channel_secret_here' else '❌ 未設定'}")
            print("=" * 60)
            return jsonify({
                "message": "Webhook endpoint is working",
                "method": "GET",
                "signature_verification": "enabled"
            })
        
        # POSTリクエストの場合
        print("=" * 60)
        print("📨 LINE Webhook受信")
        print("=" * 60)
        
        # リクエストボディを取得
        body = request.get_data()
        
        # 署名検証
        signature = request.headers.get('X-Line-Signature')
        print(f"📝 X-Line-Signature: {signature}")
        print(f"🔑 Channel Secret: {LINE_CHANNEL_SECRET[:20]}...")
        
        if not validate_signature(body, signature):
            print("❌ 署名検証失敗!")
            print("   Channel Secretが正しく設定されているか確認してください")
            abort(403)  # 署名検証失敗の場合は403を返す
        
        print("✅ 署名検証成功!")
        
        # リクエストヘッダーをログ出力
        print("📥 リクエストヘッダー:")
        for key, value in request.headers.items():
            if key.lower() != 'x-line-signature':  # 署名は既に表示済み
                print(f"  {key}: {value}")
        
        # JSONデータを解析
        try:
            data = json.loads(body.decode('utf-8'))
        except json.JSONDecodeError as e:
            print(f"❌ JSON解析エラー: {e}")
            return jsonify({"status": "error", "message": "Invalid JSON"}), 400
        
        print("=" * 60)
        print("📄 Webhook データ:")
        print(json.dumps(data, indent=2, ensure_ascii=False))
        
        # Group ID と User ID を抽出
        if 'events' in data:
            for i, event in enumerate(data['events']):
                print(f"\n📌 イベント {i+1}:")
                print(f"   タイプ: {event.get('type', 'unknown')}")
                
                if 'source' in event:
                    source = event['source']
                    source_type = source.get('type', 'unknown')
                    
                    print(f"   送信元: {source_type}")
                    
                    if source_type == 'group':
                        group_id = source.get('groupId')
                        user_id = source.get('userId')
                        
                        print("=" * 60)
                        print("🎯 グループチャット情報取得成功!")
                        print(f"👥 Group ID: {group_id}")
                        print(f"👤 User ID: {user_id}")
                        print("=" * 60)
                        print("📋 設定コマンド（コピーしてください）:")
                        print(f'export LINE_GROUP_ID="{group_id}"')
                        print("=" * 60)
                        
                        # ファイルにも保存
                        config_file = "line_group_config.txt"
                        with open(config_file, "w", encoding="utf-8") as f:
                            f.write(f"LINE_GROUP_ID={group_id}\n")
                            f.write(f"LINE_USER_ID={user_id}\n")
                        print(f"💾 設定を {config_file} に保存しました")
                        
                    elif source_type == 'user':
                        user_id = source.get('userId')
                        print("=" * 60)
                        print("📱 個人チャット情報:")
                        print(f"👤 User ID: {user_id}")
                        print("⚠️ このシステムはグループチャット専用です")
                        print("LINEグループでBotにメッセージを送信してください")
                        print("=" * 60)
                        
                    else:
                        print(f"⚠️ 未対応の送信元タイプ: {source_type}")
                
                # メッセージ内容も表示
                if event.get('type') == 'message':
                    message = event.get('message', {})
                    message_text = message.get('text', '')
                    print(f"   メッセージ: {message_text}")
        
        # 正常レスポンスを返す
        response = jsonify({"status": "success"})
        response.status_code = 200
        print("=" * 60)
        print("✅ 200 OK レスポンスを返しました")
        print("=" * 60)
        return response
        
    except Exception as e:
        print("=" * 60)
        print(f"❌ 予期しないエラー: {e}")
        print(f"❌ エラータイプ: {type(e).__name__}")
        import traceback
        print("❌ スタックトレース:")
        print(traceback.format_exc())
        print("=" * 60)
        
        # エラー時も200を返す（LINEプラットフォーム要件）
        error_response = jsonify({"status": "error", "message": str(e)})
        error_response.status_code = 200
        return error_response

if __name__ == "__main__":
    print("🚀 LINE Webhook サーバー（署名検証対応版）")
    print("=" * 60)
    print("📝 設定確認:")
    if LINE_CHANNEL_SECRET and LINE_CHANNEL_SECRET != 'your_channel_secret_here':
        print("✅ Channel Secret: 設定済み")
    else:
        print("❌ Channel Secret: 未設定")
        print("   以下のコマンドで設定してください:")
        print("   export LINE_CHANNEL_SECRET='あなたのChannelSecret'")
    
    print("\n📝 使用方法:")
    print("1. Channel Secretを環境変数に設定:")
    print("   export LINE_CHANNEL_SECRET='799dfc63eb76d45f6443f4be49833f47'")
    print("2. ngrokを起動: ngrok http 5001")
    print("3. ngrok URLをLINE Webhook URLに設定")
    print("4. LINEグループでBotにメッセージ送信")
    print("5. Group IDが表示され、設定ファイルに保存されます")
    print("=" * 60)
    
    app.run(debug=True, port=5001, host='0.0.0.0')
