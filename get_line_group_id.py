#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LINE Group ID 取得用スクリプト
"""

import sys
from pathlib import Path
from datetime import datetime

# プロジェクトルートをPythonパスに追加
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from services.core.config import Config
from services.notification.line_bot import LineBotNotifier


def check_line_configuration():
    """LINE設定状況を確認"""
    print("📱 LINE Bot グループ設定確認")
    print("=" * 50)
    
    config = Config()
    
    print("🔧 現在の設定:")
    print(f"  Channel ID: {config.line_channel_id}")
    print(f"  Channel Access Token: {'設定済み' if config.line_channel_access_token != 'dummy' else '❌未設定'}")
    print(f"  Group ID: {'設定済み' if config.line_group_id != 'dummy' else '❌未設定'}")
    
    print()
    return config


def test_line_group_connection(config):
    """LINEグループ接続テスト"""
    print("🔗 LINEグループ接続テスト")
    print("-" * 30)
    
    if not config.is_line_configured():
        print("❌ LINE設定が不完全です")
        print()
        print("⚠️ 不足している設定:")
        if config.line_channel_access_token == 'dummy':
            print("  - Channel Access Token")
        if config.line_group_id == 'dummy':
            print("  - Group ID")
        return False
    
    try:
        line_bot = LineBotNotifier()
        
        if not line_bot.enabled:
            print("❌ LINE Bot が無効化されています")
            return False
        
        # 接続テスト
        print("📤 グループチャットにテストメッセージ送信中...")
        
        test_message = f"""🧪 LINE Bot グループ接続テスト

📅 送信日時: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
📱 送信先: グループチャット
🤖 Bot名: 角上通知

✅ このメッセージが表示されていれば設定完了です！"""
        
        success = line_bot.send_message(test_message)
        
        if success:
            print("✅ グループチャット接続テスト成功！")
            print("📱 LINEグループでメッセージを確認してください")
            return True
        else:
            print("❌ グループチャット接続テスト失敗")
            return False
            
    except Exception as e:
        print(f"❌ エラー: {e}")
        return False


def show_group_setup_instructions():
    """グループ設定手順を表示"""
    print("📋 LINEグループ設定手順")
    print("=" * 50)
    print()
    
    print("【手順1】LINE Developers でグループチャット機能を有効化")
    print("  1. https://developers.line.biz/console/ にアクセス")
    print("  2. プロバイダー「角上通知」→ Channel を選択")
    print("  3. Messaging API → Bot settings")
    print("  4. 'Allow bot to join group chats' を ON に設定")
    print()
    
    print("【手順2】LINE グループを作成")
    print("  1. LINEアプリでグループを作成")
    print("  2. グループ名を設定（例: 角上物流通知）")
    print("  3. 必要なメンバーを追加")
    print()
    
    print("【手順3】BOTをグループに追加")
    print("  1. グループチャット画面を開く")
    print("  2. 右上の設定アイコン → メンバー")
    print("  3. 「招待」→ 「QRコードで招待」")
    print("  4. LINE Developers の QRコードをスキャン")
    print("  5. または Bot basic ID で検索して追加")
    print()
    
    print("【手順4】Group ID を取得")
    print("  1. Webhook を有効化（一時的）")
    print("  2. グループでBotにメッセージ送信")
    print("  3. Webhookログから Group ID を確認")
    print()
    
    print("【手順5】環境変数を設定")
    print("  export LINE_GROUP_ID=\"取得したグループID\"")
    print()


def create_webhook_for_group_id():
    """Group ID取得用Webhookサーバーを作成"""
    print("🌐 Group ID取得用 Webhookサーバー")
    print("-" * 40)
    
    webhook_code = '''#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LINE Group ID 取得用 Webhookサーバー
"""

from flask import Flask, request, jsonify
import json

app = Flask(__name__)

@app.route('/webhook', methods=['POST'])
def webhook():
    """Webhookエンドポイント"""
    try:
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
        
        return jsonify({"status": "ok"})
        
    except Exception as e:
        print(f"❌ エラー: {e}")
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    print("🚀 Group ID取得用Webhookサーバー開始")
    print("📝 使用方法:")
    print("  1. 別ターミナルで: ngrok http 5000")
    print("  2. ngrok URLをLINE Webhook URLに設定")
    print("  3. LINEグループでBotにメッセージ送信")
    print("  4. Group IDが表示されるのでコピー")
    print()
    app.run(debug=True, port=5000)
'''
    
    webhook_file = Path("webhook_group_id.py")
    with open(webhook_file, 'w', encoding='utf-8') as f:
        f.write(webhook_code)
    
    print(f"✅ Group ID取得用Webhookサーバーを作成: {webhook_file}")
    print()
    print("🚀 使用方法:")
    print("  1. pip install flask ngrok")
    print("  2. 別ターミナル: ngrok http 5000")
    print("  3. python webhook_group_id.py")
    print("  4. ngrok URLをLINE Webhook URLに設定")
    print("  5. LINEグループでBotにメッセージ送信")
    print("  6. Group IDをコピーして環境変数に設定")
    print()


def show_group_setup_guide():
    """グループ設定ガイドを表示"""
    print("🔄 グループチャット設定")
    print("-" * 30)
    
    print("環境変数で設定:")
    print("  export LINE_GROUP_ID=\"取得したグループID\"")
    print()
    
    print("または config.py で直接設定:")
    print("  line_group_id = \"取得したグループID\"")
    print()


def add_current_time_method():
    """現在時刻取得メソッドをconfigに追加"""
    config_addition = '''
    def _get_current_time(self) -> str:
        """現在時刻を取得"""
        from datetime import datetime
        return datetime.now().strftime('%Y-%m-%d %H:%M:%S')
'''
    
    # config.py に追加するコードを表示
    print("💡 config.py に追加推奨:")
    print(config_addition)


def main():
    """メイン関数"""
    print("🔧 LINE Bot グループ送信設定ツール")
    print("=" * 60)
    print()
    
    # 設定確認
    config = check_line_configuration()
    
    # グループ接続テスト
    if test_line_group_connection(config):
        print()
        print("🎉 LINEグループ設定完了！")
        print("PDF画像送信機能をグループで使用できます")
        return
    else:
        print("📋 グループ設定が必要です")
        print()
        show_group_setup_guide()
    
    print()
    
    # セットアップ手順
    show_group_setup_instructions()
    
    # Webhookサーバー作成
    answer = input("📡 Group ID取得用Webhookサーバーを作成しますか？ (y/N): ").lower()
    if answer in ['y', 'yes']:
        create_webhook_for_group_id()
    
    print()
    print("💡 設定完了後に再度このスクリプトを実行してください:")
    print("   python get_line_group_id.py")


if __name__ == "__main__":
    main()
