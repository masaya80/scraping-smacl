#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
設定管理モジュール
"""

import os
from pathlib import Path
from dataclasses import dataclass
from typing import Optional


@dataclass
class Config:
    """システム設定クラス"""
    
    def __init__(self):
        # テスト用設定
        self.test_mode = True

        # プロジェクトルートディレクトリ
        self.project_root = Path(__file__).parent.parent
        
        # SMCL設定
        self.target_url = "https://smclweb.cs-cxchange.net/smcl/view/lin/EDS001OLIN0000.aspx"
        self.corp_code = "I26S"
        self.login_id = "0600200"
        self.password = "toichi04"
        
        # ディレクトリ設定
        self.download_dir = self.project_root / "downloads"
        self.output_dir = self.project_root / "output"
        self.logs_dir = self.project_root / "logs"
        
        # マスタファイル設定
        self.master_excel_path = self.project_root / "docs" / "角上魚類.xlsx"
        
        # Selenium設定
        self.headless_mode = False
        self.selenium_timeout = 15
        
        # PDF抽出設定
        self.pdf_extract_mode = "pdfplumber"  # "pdfplumber" or "pymupdf"
        
        # Excel設定
        self.excel_template_dir = self.project_root / "templates"
        self.dispatch_template = self.excel_template_dir / "配車表_テンプレート.xlsx"
        self.shipping_template = self.excel_template_dir / "出庫依頼_テンプレート.xlsx"
        
        # LINE Bot設定
        # self.line_channel_secret = os.getenv('LINE_CHANNEL_SECRET', '25d64f941d05535214a0462185672e91')
        # self.line_channel_id = os.getenv('LINE_CHANNEL_ID', '2007929494')
        self.line_channel_access_token = os.getenv('LINE_CHANNEL_ACCESS_TOKEN', 'IVEs7s1v5o2XdR9dCT/Q101NpmsF0FvK0MxyJ9pNTqKZAYRq7mTmlghp/2cTo+vhWA5S280hrT7LZQ9HEFZ5Eu6dDtKgKywE+pUnE6zW/5Z5N148JR9IOHYA4pS8aeN4GGgrwbFNrnWqcL2eoaIhuAdB04t89/1O/w1cDnyilFU=')
        self.line_channel_secret = os.getenv('LINE_CHANNEL_SECRET', '799dfc63eb76d45f6443f4be49833f47')
        self.line_channel_id = os.getenv('LINE_CHANNEL_ID', '2007929536')
        self.line_group_id = os.getenv('LINE_GROUP_ID', 'Ca23b038e0208192c6efe1640d471e977')

        
        # Google Drive API設定
        self.google_drive_credentials_file = os.getenv('GOOGLE_DRIVE_CREDENTIALS', 
                                                      str(Path(__file__).parent.parent / 'credentials.json'))
        self.google_drive_token_file = os.getenv('GOOGLE_DRIVE_TOKEN', 
                                                str(self.project_root / 'token.json'))
        self.google_drive_folder_id = os.getenv('GOOGLE_DRIVE_FOLDER_ID', '')  # 特定フォルダID（空の場合はルートフォルダ）
        self.google_auth_code = os.getenv('GOOGLE_AUTH_CODE', '')  # 自動認証用（オプション）
        
        # PDF送信設定
        self.pdf_send_as_files = os.getenv('PDF_SEND_AS_FILES', 'false').lower() == 'true'  # PDFファイルとして送信するか
        
        # ログ設定
        self.log_level = "INFO"
        self.log_format = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
        self.log_max_bytes = 10 * 1024 * 1024  # 10MB
        self.log_backup_count = 5
        
        # その他設定
        self.max_retry_count = 3
        self.retry_delay = 2.0
        
        # 環境設定
        self.enable_excel_app = self._should_enable_excel_app()
        
        # PDF出力設定
        self.pdf_output_method = self._get_pdf_output_method()
        
        # ディレクトリを作成
        self._ensure_directories()
        
    def _should_enable_excel_app(self) -> bool:
        """Excel アプリケーションを使用するかどうかを判定（macOSでは無効化）"""
        import platform
        
        # 環境変数で強制的に有効/無効にできる
        force_enable = os.getenv('SMCL_FORCE_EXCEL_APP', '').lower() in ('true', '1', 'yes')
        force_disable = os.getenv('SMCL_DISABLE_EXCEL_APP', '').lower() in ('true', '1', 'yes')
        
        if force_enable:
            return True
        if force_disable:
            return False
        
        system = platform.system().lower()
        
        # macOSでは権限問題を回避するためExcelアプリを使用しない
        if system == "darwin":  # macOS
            return False
        # WindowsやLinuxでは使用可能
        elif system == "windows":
            return True
        else:
            return False
    
    def _get_pdf_output_method(self) -> str:
        """PDF出力方法を決定"""
        import platform
        
        # 環境変数で指定可能
        method = os.getenv('SMCL_PDF_METHOD', '').lower()
        if method in ['xlwings', 'weasyprint', 'html', 'native', 'libreoffice']:
            return method
        
        # デフォルト設定
        system = platform.system().lower()
        if system == "darwin":  # macOS
            return "enhanced"  # macOSでは改良版代替方法を使用
        elif system == "windows":
            return "xlwings"  # WindowsではExcelアプリ使用
        else:
            return "enhanced"  # Linuxなどでは改良版代替方法を使用
    
    def _ensure_directories(self):
        """必要なディレクトリを作成"""
        directories = [
            self.download_dir,
            self.output_dir,
            self.logs_dir,
            self.excel_template_dir
        ]
        
        for directory in directories:
            directory.mkdir(exist_ok=True)
    
    def get_output_filename(self, file_type: str, destination: str = None, timestamp: bool = True) -> str:
        """出力ファイル名を生成"""
        from datetime import datetime
        
        base_name = file_type
        if destination:
            base_name += f"_{destination}"
        
        if timestamp:
            timestamp_str = datetime.now().strftime('%Y%m%d_%H%M%S')
            base_name += f"_{timestamp_str}"
        
        return base_name + ".xlsx"
    
    def is_line_configured(self) -> bool:
        """LINE Bot設定が完了しているかチェック"""
        has_token = bool(self.line_channel_access_token and 
                        self.line_channel_access_token != 'dummy')
        has_group = bool(self.line_group_id and self.line_group_id != 'dummy')
        
        return has_token and has_group
    
    def get_line_target_id(self) -> str:
        """LINE送信先グループIDを取得"""
        return self.line_group_id
    
    def is_google_drive_configured(self) -> bool:
        """Google Drive設定が完了しているかチェック"""
        credentials_path = Path(self.google_drive_credentials_file)
        return credentials_path.exists()
    
    def get_google_drive_config(self) -> dict:
        """Google Drive設定を取得"""
        return {
            'credentials_file': self.google_drive_credentials_file,
            'token_file': self.google_drive_token_file,
            'folder_id': self.google_drive_folder_id,
            'auth_code': self.google_auth_code
        }
    
    def get_pdf_send_as_files(self) -> bool:
        """PDFをファイルとして送信するか取得"""
        return self.pdf_send_as_files
    
    def validate_master_file(self) -> bool:
        """マスタファイルの存在をチェック"""
        return self.master_excel_path.exists()
    
    @property
    def chrome_options(self) -> list:
        """Chrome オプションを取得"""
        options = [
            '--no-sandbox',
            '--disable-dev-shm-usage',
            '--disable-gpu',
            '--window-size=1920,1080'
        ]
        
        if self.headless_mode:
            options.append('--headless')
        
        return options
    
    def print_configuration_status(self):
        """設定状況を表示"""
        print("=== システム設定状況 ===")
        print()
        
        # LINE Bot設定
        print("📱 LINE Bot設定:")
        if self.is_line_configured():
            print("  ✅ グループチャット設定完了")
            print(f"  👥 Group ID: {self.line_group_id[:15]}...")
        else:
            print("  ❌ 未設定")
            print("  💡 環境変数 LINE_CHANNEL_ACCESS_TOKEN と LINE_GROUP_ID を設定してください")
        print()
        
        # Google Drive設定
        print("📦 Google Drive設定:")
        if self.is_google_drive_configured():
            print("  ✅ credentials.json 見つかりました")
            print(f"  📄 認証ファイル: {self.google_drive_credentials_file}")
            if self.google_drive_folder_id:
                print(f"  📁 専用フォルダ: {self.google_drive_folder_id}")
            else:
                print("  📁 ルートフォルダに保存")
        else:
            print("  ❌ credentials.json が見つかりません")
            print("  💡 GOOGLE_DRIVE_SETUP.md を参照してセットアップしてください")
        print()
        
        # マスタファイル設定
        print("📊 マスタファイル設定:")
        if self.validate_master_file():
            print("  ✅ マスタファイル見つかりました")
            print(f"  📄 ファイル: {self.master_excel_path}")
        else:
            print("  ❌ マスタファイルが見つかりません")
            print(f"  💡 {self.master_excel_path} を配置してください")
        print()
        
        # システム環境
        print("⚙️  システム環境:")
        print(f"  🖥️  Excel アプリ: {'有効' if self.enable_excel_app else '無効'}")
        print(f"  📄 PDF出力方法: {self.pdf_output_method}")
        print(f"  🕹️  ヘッドレスモード: {'有効' if self.headless_mode else '無効'}")
        print()
    
    def get_setup_commands(self) -> list:
        """セットアップコマンドの一覧を取得"""
        commands = []
        
        if not self.is_line_configured():
            commands.append("# LINE Bot設定")
            commands.append("export LINE_CHANNEL_ACCESS_TOKEN='your_channel_access_token'")
            commands.append("export LINE_GROUP_ID='your_group_id'")
            commands.append("")
        
        if not self.is_google_drive_configured():
            commands.append("# Google Drive設定")
            commands.append("# 1. GOOGLE_DRIVE_SETUP.md を参照して credentials.json を取得")
            commands.append("# 2. プロジェクトルートに配置")
            commands.append("# 3. オプション: 専用フォルダを使用")
            commands.append("# export GOOGLE_DRIVE_FOLDER_ID='folder_id'")
            commands.append("")
        
        if not self.validate_master_file():
            commands.append("# マスタファイル設定")
            commands.append(f"# {self.master_excel_path} にマスタExcelファイルを配置")
            commands.append("")
        
        return commands 