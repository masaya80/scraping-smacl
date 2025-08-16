#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Google Drive フォルダ設定スクリプト
"""

import sys
import os
from pathlib import Path

# プロジェクトルートをPythonパスに追加
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from services.core.config import Config
from services.core.google_drive_uploader import GoogleDriveUploader


def set_folder_config():
    """指定フォルダを設定"""
    
    folder_url = "https://drive.google.com/drive/u/5/folders/1WIYtbtWBFynj4A7xsje4Se4vxajLVwRt"
    folder_id = "1WIYtbtWBFynj4A7xsje4Se4vxajLVwRt"
    
    print("🔧 Google Drive フォルダ設定")
    print("=" * 50)
    print(f"📁 フォルダURL: {folder_url}")
    print(f"🆔 フォルダID: {folder_id}")
    print()
    
    # 現在の設定を確認
    config = Config()
    print("📋 現在の設定:")
    print(f"  フォルダID: {config.google_drive_folder_id}")
    print()
    
    # 設定方法を表示
    print("⚙️ 設定方法:")
    print()
    
    print("【方法1】環境変数で設定（推奨）:")
    print(f'export GOOGLE_DRIVE_FOLDER_ID="{folder_id}"')
    print()
    
    print("【方法2】config.py で直接設定（既に設定済み）:")
    print("services/core/config.py で設定済みです")
    print()
    
    print("【方法3】.env ファイルで設定:")
    env_content = f"""# Google Drive設定
GOOGLE_DRIVE_FOLDER_ID={folder_id}

# LINE Bot設定（必要に応じて）
LINE_CHANNEL_ACCESS_TOKEN=your_channel_access_token
LINE_USER_ID=your_user_id

# 認証コード（初回認証時）
GOOGLE_AUTH_CODE=your_auth_code
"""
    
    env_file = Path(".env")
    if not env_file.exists():
        with open(env_file, 'w', encoding='utf-8') as f:
            f.write(env_content)
        print(f"✅ .env ファイルを作成しました: {env_file}")
    else:
        print(f"⚠️  .env ファイルが既に存在します: {env_file}")
    
    print()


def test_folder_access():
    """フォルダアクセステスト"""
    print("🧪 フォルダアクセステスト")
    print("-" * 30)
    
    try:
        # Google Drive接続テスト
        uploader = GoogleDriveUploader()
        
        if not uploader.is_available():
            print("❌ Google Driveが利用できません")
            print("💡 credentials.json を設定してください")
            return False
        
        # 接続テスト
        if uploader.test_connection():
            print("✅ Google Drive接続成功")
            
            # ストレージ情報取得
            storage_info = uploader.get_storage_info()
            if storage_info.get('available'):
                print(f"👤 ユーザー: {storage_info.get('user_email')}")
                print(f"💾 使用容量: {storage_info.get('used_space_gb')} / {storage_info.get('total_space_gb')} GB")
                
            # フォルダ情報
            config = Config()
            folder_id = config.google_drive_folder_id
            if folder_id:
                print(f"📁 設定済みフォルダID: {folder_id}")
                print("📋 フォルダアクセス権限は実際のアップロード時に確認されます")
            else:
                print("⚠️  フォルダIDが設定されていません")
            
            return True
        else:
            print("❌ Google Drive接続失敗")
            return False
            
    except Exception as e:
        print(f"❌ テストエラー: {e}")
        return False


def create_test_upload():
    """テストアップロードを実行"""
    print("\n🚀 テストアップロード")
    print("-" * 20)
    
    try:
        from PIL import Image
        import tempfile
        
        # テスト画像を作成
        with tempfile.NamedTemporaryFile(suffix='.jpg', delete=False) as temp_file:
            img = Image.new('RGB', (300, 200), color='lightblue')
            img.save(temp_file, 'JPEG')
            temp_path = Path(temp_file.name)
        
        print(f"📄 テスト画像作成: {temp_path.name}")
        
        # アップロードテスト
        uploader = GoogleDriveUploader()
        public_url = uploader.upload_image_to_temporary_folder(temp_path)
        
        if public_url:
            print("✅ テストアップロード成功")
            print(f"🔗 公開URL: {public_url}")
            
            # フォルダURLを表示
            folder_url = "https://drive.google.com/drive/u/5/folders/1WIYtbtWBFynj4A7xsje4Se4vxajLVwRt"
            print(f"📁 アップロード先フォルダ: {folder_url}")
            
        else:
            print("❌ テストアップロード失敗")
        
        # テストファイルを削除
        temp_path.unlink()
        print("🗑️ テストファイルを削除")
        
        return public_url is not None
        
    except ImportError:
        print("❌ PIL (Pillow) ライブラリが必要です")
        print("pip install Pillow")
        return False
    except Exception as e:
        print(f"❌ テストアップロードエラー: {e}")
        return False


def main():
    """メイン関数"""
    print("🎯 Google Drive 特定フォルダ設定\n")
    
    # フォルダ設定
    set_folder_config()
    
    # アクセステスト
    if test_folder_access():
        print()
        
        # テストアップロード（オプション）
        answer = input("📤 テストアップロードを実行しますか？ (y/N): ").lower()
        if answer in ['y', 'yes']:
            create_test_upload()
    
    print()
    print("=" * 50)
    print("🎉 設定完了！")
    print()
    print("📝 次のステップ:")
    print("  1. python check_config.py     # 設定確認")
    print("  2. python main.py             # メインシステム実行")
    print()
    print("💡 注意事項:")
    print("  - 指定フォルダへの書き込み権限が必要です")
    print("  - 認証時に適切な権限を許可してください")
    print()


if __name__ == "__main__":
    main()
