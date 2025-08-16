#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LINE Group ID 簡易テストスクリプト（ngrok不要）
"""

import sys
from pathlib import Path

# プロジェクトルートをPythonパスに追加
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from services.core.config import Config
from services.notification.line_bot import LineBotNotifier


def test_known_group_ids():
    """既知のGroup IDパターンをテスト"""
    print("🧪 LINE Group ID テストツール")
    print("=" * 50)
    
    config = Config()
    
    # 現在設定されているGroup ID
    current_group_id = config.line_group_id
    print(f"現在のGroup ID: {current_group_id}")
    print()
    
    # Group IDのパターンを確認
    print("🔍 Group IDの特徴:")
    print(f"  長さ: {len(current_group_id)} 文字")
    print(f"  先頭文字: {current_group_id[0] if current_group_id else 'なし'}")
    print()
    
    # 一般的なLINE Group IDパターン
    common_patterns = [
        current_group_id,  # 現在の設定値
        f"U{current_group_id[1:]}",  # U始まり（個人ID）
        f"R{current_group_id[1:]}",  # R始まり（Room ID）
    ]
    
    print("📋 テスト対象ID:")
    for i, group_id in enumerate(common_patterns, 1):
        if group_id and group_id != 'dummy':
            print(f"  {i}. {group_id}")
    
    print()
    
    # 各IDでテスト送信
    for i, test_id in enumerate(common_patterns, 1):
        if not test_id or test_id == 'dummy':
            continue
            
        print(f"🔄 パターン{i}をテスト: {test_id[:10]}...")
        
        try:
            # 一時的にGroup IDを変更
            original_id = config.line_group_id
            config.line_group_id = test_id
            
            # LINE Bot初期化
            line_bot = LineBotNotifier()
            
            if line_bot.enabled:
                # テストメッセージ送信
                test_message = f"🧪 Group IDテスト {i}\nID: {test_id[:15]}..."
                success = line_bot.send_message(test_message)
                
                if success:
                    print(f"  ✅ 成功! このIDが正しいGroup IDです")
                    print(f"  🎯 正しいGroup ID: {test_id}")
                    
                    # config.pyに正しいIDを設定
                    update_config_with_correct_id(test_id)
                    return test_id
                else:
                    print(f"  ❌ 失敗")
            else:
                print(f"  ❌ LINE Bot設定エラー")
            
            # 元のIDに戻す
            config.line_group_id = original_id
            
        except Exception as e:
            print(f"  ❌ エラー: {e}")
    
    print()
    print("⚠️  すべてのパターンが失敗しました")
    print("💡 手動でGroup IDを確認する必要があります")
    return None


def update_config_with_correct_id(correct_id):
    """正しいGroup IDでconfig.pyを更新"""
    try:
        print(f"\n📝 config.pyを正しいGroup IDで更新中...")
        print(f"新しいGroup ID: {correct_id}")
        
        # 環境変数設定コマンドを表示
        print("\n✅ 以下のコマンドで環境変数を設定してください:")
        print(f"export LINE_GROUP_ID=\"{correct_id}\"")
        
    except Exception as e:
        print(f"❌ config.py更新エラー: {e}")


def manual_group_id_input():
    """手動でGroup IDを入力"""
    print("\n🔧 手動Group ID設定")
    print("-" * 30)
    
    print("LINE Developers Console でGroup IDを確認:")
    print("1. https://developers.line.biz/console/ にアクセス")
    print("2. Channel選択 → Messaging API → Webhook")
    print("3. Webhookイベントログを確認")
    print("4. 'source.groupId' の値をコピー")
    print()
    
    group_id = input("Group IDを入力してください: ").strip()
    
    if group_id and len(group_id) > 10:
        print(f"\n🧪 入力されたGroup IDをテスト: {group_id}")
        
        try:
            config = Config()
            config.line_group_id = group_id
            
            line_bot = LineBotNotifier()
            if line_bot.enabled:
                success = line_bot.send_message(f"🧪 手動入力Group IDテスト\n{group_id}")
                
                if success:
                    print("✅ 成功! 正しいGroup IDです")
                    update_config_with_correct_id(group_id)
                    return group_id
                else:
                    print("❌ 送信失敗")
            
        except Exception as e:
            print(f"❌ エラー: {e}")
    
    print("❌ 無効なGroup IDです")
    return None


def show_line_developers_guide():
    """LINE Developers での確認方法を表示"""
    print("\n📋 LINE Developers でGroup ID確認方法")
    print("=" * 50)
    print()
    
    print("【方法1】Webhook イベントログで確認")
    print("1. https://developers.line.biz/console/ にアクセス")
    print("2. プロバイダー「角上通知」→ Channel「2007929536」選択")
    print("3. Messaging API → Webhook")
    print("4. 'Use webhook' を ON に設定")
    print("5. Webhook URL に一時的なURLを設定（例: https://example.com/webhook）")
    print("6. LINEグループでBotにメッセージ送信")
    print("7. Webhookログで 'source.groupId' を確認")
    print()
    
    print("【方法2】Bot をグループに追加して確認")
    print("1. LINEアプリでグループ作成")
    print("2. グループにBotを追加（QRコードまたは Bot ID検索）")
    print("3. 'Allow bot to join group chats' が ON であることを確認")
    print("4. グループでBotにメッセージ送信")
    print("5. Webhook ログで Group ID を確認")
    print()
    
    print("【確認すべき設定】")
    print("✅ Allow bot to join group chats: ON")
    print("✅ Use webhook: ON（Group ID確認時のみ）")
    print("✅ Auto-reply messages: OFF")
    print("✅ Bot がグループメンバーに追加済み")
    print()


def main():
    """メイン関数"""
    print("🔧 ngrok不要 Group ID確認ツール")
    print("=" * 60)
    print()
    
    # 自動テスト
    correct_id = test_known_group_ids()
    
    if correct_id:
        print("\n🎉 Group ID設定完了！")
        print("以下のコマンドでメインシステムを実行してください:")
        print("python main.py")
        return
    
    # 手動入力
    print("\n" + "="*50)
    answer = input("手動でGroup IDを入力しますか？ (y/N): ").lower()
    
    if answer in ['y', 'yes']:
        correct_id = manual_group_id_input()
        if correct_id:
            print("\n🎉 手動設定完了！")
            return
    
    # 確認方法ガイド
    show_line_developers_guide()
    
    print("\n💡 Group ID確認後、以下のコマンドで設定してください:")
    print("export LINE_GROUP_ID=\"確認したGroup ID\"")
    print("python simple_group_test.py")


if __name__ == "__main__":
    main()
