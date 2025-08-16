#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Google Drive アップロードモジュール
"""

import os
import json
import time
from pathlib import Path
from typing import Optional, Union, Dict, Any
import tempfile

from .logger import Logger


class GoogleDriveUploader:
    """Google Drive アップロード クラス"""
    
    def __init__(self):
        self.logger = Logger(__name__)
        
        # Google Drive API の設定
        self.drive_available = False
        self.service = None
        
        try:
            from googleapiclient.discovery import build
            from googleapiclient.http import MediaFileUpload
            from google.oauth2.credentials import Credentials
            from google_auth_oauthlib.flow import Flow
            from google.auth.transport.requests import Request
            
            self.build = build
            self.MediaFileUpload = MediaFileUpload
            self.Credentials = Credentials
            self.Flow = Flow
            self.Request = Request
            
            self.drive_available = True
            self.logger.info("Google Drive API ライブラリが利用可能です")
            
        except ImportError as e:
            self.logger.warning(f"Google Drive API ライブラリが見つかりません: {e}")
        
        # 設定取得
        from .config import Config
        config = Config()
        drive_config = config.get_google_drive_config()
        
        # Google Drive設定
        self.credentials_file = drive_config['credentials_file']
        self.token_file = drive_config['token_file']
        self.folder_id = drive_config['folder_id']
        self.auth_code = drive_config['auth_code']
        
        # 認証スコープ
        self.scopes = ['https://www.googleapis.com/auth/drive.file']
        
        # 初期化時に認証を試行
        self._init_drive_service()
    
    def _init_drive_service(self):
        """Google Drive サービスを初期化"""
        if not self.drive_available:
            return
        
        try:
            creds = None
            
            # 認証情報ファイルが存在するかチェック
            if not os.path.exists(self.credentials_file):
                self.logger.warning(f"認証情報ファイルが見つかりません: {self.credentials_file}")
                self.logger.info("Google Cloud Consoleで認証情報を作成し、credentials.jsonとして保存してください")
                return
            
            # 認証情報ファイルの種類を判定
            try:
                with open(self.credentials_file, 'r') as f:
                    cred_data = json.load(f)
                
                if cred_data.get('type') == 'service_account':
                    # Service Account認証
                    self.logger.info("Service Account認証を使用します")
                    from google.oauth2 import service_account
                    creds = service_account.Credentials.from_service_account_file(
                        self.credentials_file, scopes=self.scopes
                    )
                else:
                    # OAuth2認証（既存のフロー）
                    self.logger.info("OAuth2認証を使用します")
                    
                    # 既存のトークンファイルを確認
                    if os.path.exists(self.token_file):
                        creds = self.Credentials.from_authorized_user_file(self.token_file, self.scopes)
                    
                    # 認証情報が無効な場合は再認証
                    if not creds or not creds.valid:
                        if creds and creds.expired and creds.refresh_token:
                            creds.refresh(self.Request())
                        else:
                            # 初回認証または再認証が必要
                            flow = self.Flow.from_client_secrets_file(self.credentials_file, self.scopes)
                            flow.redirect_uri = 'urn:ietf:wg:oauth:2.0:oob'  # デスクトップアプリ用
                            
                            # 認証URLを表示（手動認証）
                            auth_url, _ = flow.authorization_url(prompt='consent')
                            self.logger.info(f"認証が必要です。以下のURLにアクセスしてください:")
                            self.logger.info(auth_url)
                            self.logger.info("認証コードを入力してください（環境変数 GOOGLE_AUTH_CODE で設定可能）:")
                            
                            auth_code = self.auth_code
                            if not auth_code:
                                self.logger.warning("認証コードが設定されていません。手動認証が必要です")
                                self.logger.info("環境変数 GOOGLE_AUTH_CODE に認証コードを設定するか、")
                                self.logger.info("config.py の google_auth_code を設定してください")
                                return
                            
                            flow.fetch_token(code=auth_code)
                            creds = flow.credentials
                        
                        # トークンを保存
                        with open(self.token_file, 'w') as token:
                            token.write(creds.to_json())
                
            except Exception as e:
                self.logger.error(f"認証情報ファイル読み込みエラー: {str(e)}")
                return
            
            # Drive サービスを構築
            self.service = self.build('drive', 'v3', credentials=creds)
            self.logger.info("Google Drive サービス初期化完了")
            
        except Exception as e:
            self.logger.error(f"Google Drive サービス初期化エラー: {str(e)}")
            self.service = None
    
    def upload_image(self, image_path: Union[str, Path], filename: Optional[str] = None) -> Optional[str]:
        """
        画像をGoogle Driveにアップロードして公開URLを取得
        
        Args:
            image_path: 画像ファイルのパス
            filename: アップロード時のファイル名（指定しない場合は元ファイル名を使用）
        
        Returns:
            公開URL（失敗時はNone）
        """
        if not self.is_available():
            self.logger.error("Google Drive サービスが利用できません")
            return None
        
        try:
            image_path = Path(image_path)
            if not image_path.exists():
                self.logger.error(f"画像ファイルが見つかりません: {image_path}")
                return None
            
            # ファイル名を決定
            if filename is None:
                filename = image_path.name
            
            # タイムスタンプを付けてユニークにする
            timestamp = int(time.time())
            unique_filename = f"{timestamp}_{filename}"
            
            self.logger.info(f"Google Driveにアップロード開始: {unique_filename}")
            
            # ファイルのメタデータ
            file_metadata = {
                'name': unique_filename,
                'parents': [self.folder_id] if self.folder_id else []
            }
            
            # ファイルのMIMEタイプを決定
            mime_type = self._get_mime_type(image_path)
            
            # ファイルをアップロード
            media = self.MediaFileUpload(str(image_path), mimetype=mime_type)
            
            file = self.service.files().create(
                body=file_metadata,
                media_body=media,
                fields='id'
            ).execute()
            
            file_id = file.get('id')
            
            if not file_id:
                self.logger.error("ファイルIDの取得に失敗")
                return None
            
            # ファイルを公開設定にする
            self.service.permissions().create(
                fileId=file_id,
                body={'role': 'reader', 'type': 'anyone'}
            ).execute()
            
            # 公開URLを生成
            # Google Driveの直接画像URLの形式
            public_url = f"https://drive.google.com/uc?export=view&id={file_id}"
            
            self.logger.info(f"Google Driveアップロード完了: {public_url}")
            return public_url
            
        except Exception as e:
            self.logger.error(f"Google Driveアップロードエラー: {str(e)}")
            return None
    
    def upload_image_to_temporary_folder(self, image_path: Union[str, Path]) -> Optional[str]:
        """
        画像を一時フォルダにアップロード（自動削除用）
        
        Args:
            image_path: 画像ファイルのパス
        
        Returns:
            公開URL（失敗時はNone）
        """
        if not self.is_available():
            return None
        
        try:
            # 一時フォルダIDを取得または作成
            temp_folder_id = self._get_or_create_temp_folder()
            if not temp_folder_id:
                # 一時フォルダが作れない場合は通常アップロード
                return self.upload_image(image_path)
            
            # 元のフォルダIDを一時的に変更
            original_folder_id = self.folder_id
            self.folder_id = temp_folder_id
            
            try:
                result = self.upload_image(image_path)
                return result
            finally:
                # フォルダIDを元に戻す
                self.folder_id = original_folder_id
                
        except Exception as e:
            self.logger.error(f"一時フォルダアップロードエラー: {str(e)}")
            return None
    
    def cleanup_old_files(self, max_age_hours: int = 24):
        """古いファイルをクリーンアップ"""
        if not self.is_available():
            return
        
        try:
            # 一時フォルダのファイル一覧を取得
            temp_folder_id = self._get_temp_folder_id()
            if not temp_folder_id:
                return
            
            # 古いファイルを検索
            cutoff_time = time.time() - (max_age_hours * 3600)
            cutoff_datetime = time.strftime('%Y-%m-%dT%H:%M:%S', time.gmtime(cutoff_time))
            
            query = f"'{temp_folder_id}' in parents and createdTime < '{cutoff_datetime}'"
            
            results = self.service.files().list(
                q=query,
                fields="files(id, name, createdTime)"
            ).execute()
            
            files = results.get('files', [])
            
            # 古いファイルを削除
            deleted_count = 0
            for file in files:
                try:
                    self.service.files().delete(fileId=file['id']).execute()
                    self.logger.info(f"古いファイルを削除: {file['name']}")
                    deleted_count += 1
                except Exception as e:
                    self.logger.warning(f"ファイル削除エラー: {file['name']} - {str(e)}")
            
            if deleted_count > 0:
                self.logger.info(f"クリーンアップ完了: {deleted_count}個のファイルを削除")
            
        except Exception as e:
            self.logger.error(f"クリーンアップエラー: {str(e)}")
    
    def _get_or_create_temp_folder(self) -> Optional[str]:
        """一時フォルダを取得または作成"""
        try:
            folder_name = "LINE_Bot_Temp_Images"
            
            # 既存フォルダを検索
            query = f"name='{folder_name}' and mimeType='application/vnd.google-apps.folder'"
            if self.folder_id:
                query += f" and '{self.folder_id}' in parents"
            
            results = self.service.files().list(q=query, fields="files(id, name)").execute()
            folders = results.get('files', [])
            
            if folders:
                return folders[0]['id']
            
            # フォルダが存在しない場合は作成
            folder_metadata = {
                'name': folder_name,
                'mimeType': 'application/vnd.google-apps.folder',
                'parents': [self.folder_id] if self.folder_id else []
            }
            
            folder = self.service.files().create(body=folder_metadata, fields='id').execute()
            folder_id = folder.get('id')
            
            self.logger.info(f"一時フォルダを作成: {folder_name} (ID: {folder_id})")
            return folder_id
            
        except Exception as e:
            self.logger.error(f"一時フォルダ作成エラー: {str(e)}")
            return None
    
    def _get_temp_folder_id(self) -> Optional[str]:
        """一時フォルダIDを取得"""
        try:
            folder_name = "LINE_Bot_Temp_Images"
            query = f"name='{folder_name}' and mimeType='application/vnd.google-apps.folder'"
            
            results = self.service.files().list(q=query, fields="files(id)").execute()
            folders = results.get('files', [])
            
            return folders[0]['id'] if folders else None
            
        except Exception:
            return None
    
    def _get_mime_type(self, file_path: Path) -> str:
        """ファイルのMIMEタイプを取得"""
        ext = file_path.suffix.lower()
        if ext == '.jpg' or ext == '.jpeg':
            return 'image/jpeg'
        elif ext == '.png':
            return 'image/png'
        elif ext == '.gif':
            return 'image/gif'
        elif ext == '.webp':
            return 'image/webp'
        else:
            return 'image/jpeg'  # デフォルト
    
    def is_available(self) -> bool:
        """Google Drive が利用可能かどうかを確認"""
        return self.drive_available and self.service is not None
    
    def test_connection(self) -> bool:
        """Google Drive 接続テスト"""
        if not self.is_available():
            return False
        
        try:
            # 簡単なAPI呼び出しでテスト
            about = self.service.about().get(fields="user").execute()
            user_email = about.get('user', {}).get('emailAddress', '不明')
            
            self.logger.info(f"Google Drive接続テスト成功: {user_email}")
            return True
            
        except Exception as e:
            self.logger.error(f"Google Drive接続テストエラー: {str(e)}")
            return False
    
    def get_storage_info(self) -> Dict[str, Any]:
        """ストレージ情報を取得"""
        if not self.is_available():
            return {"available": False}
        
        try:
            about = self.service.about().get(fields="storageQuota,user").execute()
            
            storage_quota = about.get('storageQuota', {})
            user = about.get('user', {})
            
            # バイトを GB に変換
            def bytes_to_gb(bytes_value):
                if bytes_value:
                    return round(int(bytes_value) / (1024**3), 2)
                return 0
            
            return {
                "available": True,
                "user_email": user.get('emailAddress', '不明'),
                "total_space_gb": bytes_to_gb(storage_quota.get('limit')),
                "used_space_gb": bytes_to_gb(storage_quota.get('usage')),
                "folder_id": self.folder_id if self.folder_id else "ルート"
            }
            
        except Exception as e:
            self.logger.error(f"ストレージ情報取得エラー: {str(e)}")
            return {"available": False, "error": str(e)}
    
    def create_setup_instructions(self) -> str:
        """セットアップ手順を返す"""
        return """
Google Drive API セットアップ手順:

1. Google Cloud Console にアクセス
   https://console.cloud.google.com/

2. 新しいプロジェクトを作成（または既存プロジェクトを選択）

3. Google Drive API を有効化
   - APIとサービス > ライブラリ
   - "Google Drive API" を検索して有効化

4. 認証情報を作成
   - APIとサービス > 認証情報
   - "認証情報を作成" > "OAuthクライアントID"
   - アプリケーションタイプ: "デスクトップアプリケーション"
   - JSONファイルをダウンロードして "credentials.json" として保存

5. 環境変数を設定（オプション）:
   export GOOGLE_DRIVE_CREDENTIALS="credentials.json"
   export GOOGLE_DRIVE_TOKEN="token.json"
   export GOOGLE_DRIVE_FOLDER_ID="特定フォルダのID（オプション）"

6. 初回実行時に認証URLが表示されるので、ブラウザでアクセスして認証
"""
    
    def upload_pdf_to_temporary_folder(self, pdf_path: Union[str, Path]) -> Optional[str]:
        """
        PDFファイルを一時フォルダにアップロード
        
        Args:
            pdf_path: PDFファイルのパス
        
        Returns:
            公開URL（失敗時はNone）
        """
        if not self.is_available():
            return None
        
        try:
            pdf_path = Path(pdf_path)
            if not pdf_path.exists():
                self.logger.error(f"PDFファイルが見つかりません: {pdf_path}")
                return None
            
            # 一時フォルダIDを取得または作成
            temp_folder_id = self._get_temp_folder_id()
            if not temp_folder_id:
                temp_folder_id = self._create_temp_folder()
                if not temp_folder_id:
                    return None
            
            # ファイル名を生成（タイムスタンプ付き）
            timestamp = int(time.time())
            unique_filename = f"{timestamp}_{pdf_path.name}"
            
            self.logger.info(f"Google DriveにPDFアップロード開始: {unique_filename}")
            
            # ファイルメタデータ
            file_metadata = {
                'name': unique_filename,
                'parents': [temp_folder_id]
            }
            
            # PDFファイルをアップロード
            media = self.MediaFileUpload(
                str(pdf_path),
                mimetype='application/pdf',
                resumable=True
            )
            
            file = self.service.files().create(
                body=file_metadata,
                media_body=media,
                fields='id'
            ).execute()
            
            file_id = file.get('id')
            
            # ファイルを公開設定
            permission = {
                'type': 'anyone',
                'role': 'reader'
            }
            self.service.permissions().create(
                fileId=file_id,
                body=permission
            ).execute()
            
            # 公開URLを生成
            public_url = f"https://drive.google.com/uc?export=download&id={file_id}"
            
            self.logger.info(f"Google Drive PDFアップロード完了: {public_url}")
            return public_url
            
        except Exception as e:
            self.logger.error(f"Google Drive PDFアップロードエラー: {str(e)}")
            return None
