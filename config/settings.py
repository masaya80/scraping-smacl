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
        self.master_dir = self.project_root / "master"
        
        # マスタファイル設定
        self.master_excel_path = self.project_root / "docs" / "角上魚類　配送依頼・出庫依頼.xlsx"
        
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
        self.line_channel_access_token = 'dummy'
        self.line_user_id = 'dummy'
        
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
        if method in ['xlwings', 'weasyprint', 'html']:
            return method
        
        # デフォルト設定
        system = platform.system().lower()
        if system == "darwin":  # macOS
            return "weasyprint"  # macOSではweasyprint推奨
        elif system == "windows":
            return "xlwings"  # WindowsではExcelアプリ使用
        else:
            return "weasyprint"  # Linuxなどではweasyprint
    
    def _ensure_directories(self):
        """必要なディレクトリを作成"""
        directories = [
            self.download_dir,
            self.output_dir,
            self.logs_dir,
            self.master_dir,
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
        return bool(self.line_channel_access_token and self.line_user_id)
    
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