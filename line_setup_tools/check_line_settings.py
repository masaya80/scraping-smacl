#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LINE Bot 設定確認スクリプト
"""

import sys
from pathlib import Path

# プロジェクトルートをPythonパスに追加
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from services.core.config import Config


def main():
    """LINE Bot設定の詳細確認"""
    print("🔧 LINE Bot 設定詳細確認")
    print("=" * 50)
    
    config = Config()
    
    print("📱 現在の設定:")
    print(f"  Channel ID: {config.line_channel_id}")
    print(f"  Channel Secret: {config.line_channel_secret}")
    print(f"  Channel Access Token: {config.line_channel_access_token[:50]}...")
    print(f"  Group ID: {config.line_group_id}")
    print()
    
    print("🔍 確認すべき項目:")
    print("1. LINE Developers Console 設定:")
    print("   https://developers.line.biz/console/")
    print("   → プロバイダー: 角上通知")
    print("   → Channel ID: 2007929536")
    print()
    
    print("2. Messaging API 設定:")
    print("   ✅ Allow bot to join group chats: ON")
    print("   ✅ Use webhook: ON（Group ID確認時のみ）")
    print("   ✅ Auto-reply messages: OFF")
    print("   ✅ Greeting messages: OFF")
    print()
    
    print("3. グループ設定:")
    print("   ✅ LINEグループを作成済み")
    print("   ✅ Bot をグループに追加済み")
    print("   ✅ Bot がグループメンバーとして表示される")
    print()
    
    print("4. Group ID 確認方法:")
    print("   A. グループでBotにメッセージ送信")
    print("   B. LINE Developers → Messaging API → Webhook")
    print("   C. Webhook events ログを確認")
    print("   D. 'source.groupId' の値をコピー")
    print()
    
    print("💡 次のステップ:")
    print("1. 上記設定を確認")
    print("2. 正しいGroup IDを取得")
    print("3. export LINE_GROUP_ID=\"正しいID\"")
    print("4. python simple_group_test.py で再テスト")


if __name__ == "__main__":
    main()
