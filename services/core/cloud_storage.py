#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Google Drive 画像アップロードモジュール
"""

import os
from pathlib import Path
from typing import Optional, Union

from .logger import Logger
from .google_drive_uploader import GoogleDriveUploader


class CloudStorageUploader:
    """Google Drive 画像アップロード クラス"""
    
    def __init__(self):
        self.logger = Logger(__name__)
        
        # Google Drive設定のみ
        self.google_drive = GoogleDriveUploader()
        
        if self.google_drive.is_available():
            self.logger.info("Google Drive が利用可能です")
        else:
            self.logger.warning("Google Drive が利用できません。画像送信はテキスト情報のみになります")
        

    
    def upload_image(self, image_path: Union[str, Path], filename: Optional[str] = None) -> Optional[str]:
        """
        画像をGoogle Driveにアップロード
        
        Args:
            image_path: 画像ファイルのパス
            filename: アップロード時のファイル名
        
        Returns:
            公開URL（失敗時はNone）
        """
        try:
            if self.google_drive.is_available():
                self.logger.info("Google Driveにアップロード中...")
                url = self.google_drive.upload_image_to_temporary_folder(image_path)
                if url:
                    return url
                else:
                    self.logger.error("Google Driveアップロードに失敗しました")
                    return None
            else:
                self.logger.error("Google Driveが利用できません")
                return None
            
        except Exception as e:
            self.logger.error(f"画像アップロードエラー: {str(e)}")
            return None
    
    def upload_pdf(self, pdf_path: Union[str, Path], filename: Optional[str] = None) -> Optional[str]:
        """
        PDFをGoogle Driveにアップロード
        
        Args:
            pdf_path: PDFファイルのパス
            filename: アップロード時のファイル名
        
        Returns:
            公開URL（失敗時はNone）
        """
        try:
            if self.google_drive.is_available():
                self.logger.info("Google DriveにPDFアップロード中...")
                url = self.google_drive.upload_pdf_to_temporary_folder(pdf_path)
                if url:
                    return url
                else:
                    self.logger.error("Google Drive PDFアップロードに失敗しました")
                    return None
            else:
                self.logger.error("Google Driveが利用できません")
                return None
            
        except Exception as e:
            self.logger.error(f"PDFアップロードエラー: {str(e)}")
            return None
    
    def is_available(self) -> bool:
        """Google Drive が利用可能かどうかを確認"""
        return self.google_drive.is_available()
    
    def test_connection(self) -> bool:
        """Google Drive接続テスト"""
        try:
            if self.google_drive.test_connection():
                self.logger.info("✅ Google Drive接続テスト成功")
                return True
            else:
                self.logger.warning("❌ Google Drive接続テスト失敗")
                return False
                
        except Exception as e:
            self.logger.error(f"Google Drive接続テストエラー: {str(e)}")
            return False
    
    def get_storage_info(self):
        """Google Drive ストレージ情報を取得"""
        return self.google_drive.get_storage_info()
    
    def cleanup_old_files(self, max_age_hours: int = 24):
        """古いファイルをクリーンアップ"""
        if self.google_drive.is_available():
            self.google_drive.cleanup_old_files(max_age_hours)
