#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
設定管理モジュール
"""

import os
import platform
from pathlib import Path
from dataclasses import dataclass
from typing import Optional
from datetime import datetime


@dataclass
class Config:
    """システム設定クラス"""
    
    def __init__(self):
        # テスト用設定（この値を変更して環境を切り替え）
        self.test_mode = True
        
        # プロジェクトルートディレクトリ
        self.project_root = Path(__file__).parent.parent
        
        # SMCL設定
        self.target_url = "https://smclweb.cs-cxchange.net/smcl/view/lin/EDS001OLIN0000.aspx"
        self.corp_code = "I26S"
        self.login_id = "0600200"
        self.password = "toichi04"
        
        # test_modeに基づいて設定を初期化
        self._configure_based_on_mode()
        
        # 共通設定
        self._configure_common_settings()
        
        # ディレクトリを作成
        self._ensure_directories()
    
    def _configure_based_on_mode(self):
        """test_modeに基づいて設定を構成"""
        if self.test_mode:
            self._configure_for_test_mode()
        else:
            self._configure_for_production_mode()
    
    def _configure_for_test_mode(self):
        """開発・テスト用設定"""
        # ディレクトリ設定（ローカル）
        self.download_dir = self.project_root / "downloads"
        self.output_dir = self.project_root / "output"
        self.logs_dir = self.project_root / "logs"
        
        # ネットワークドライブ設定（無効）
        self.use_network_storage = False
        self.network_base_path = None
        
        # マスタファイル設定（ローカル）
        self.master_excel_path = self.project_root / "docs" / "角上魚類.xlsx"
        
        # Selenium設定（ブラウザ表示）
        self.headless_mode = False
        
        # LINE Bot設定（テスト用）
        self._configure_test_line_settings()
        
        # ログ設定（デバッグレベル）
        self.log_level = "DEBUG"
        
        print("🧪 テストモード設定を適用しました")
        print("   - ローカルファイル使用")
        print("   - ブラウザ表示あり")
        print("   - 確定処理: 個別設定")
    
    def _configure_for_production_mode(self):
        """本番用設定"""
        # Windowsネットワークドライブ設定
        self.network_base_path = self._get_network_base_path()
        self.use_network_storage = self._should_use_network_storage()
        
        # ディレクトリ設定（ネットワーク or ローカル）
        if self.use_network_storage:
            # ネットワークドライブ使用
            network_path = Path(self.network_base_path)
            self.download_dir = network_path / "downloads"
            self.output_dir = network_path
            self.logs_dir = network_path / "logs"
            # マスタファイル設定（ネットワーク）
            self.master_excel_path = network_path / "角上魚類.xlsx"
        else:
            # ローカルフォルダにフォールバック
            self.download_dir = self.project_root / "downloads"
            self.output_dir = self.project_root / "output"
            self.logs_dir = self.project_root / "logs"
            self.master_excel_path = self.project_root / "docs" / "角上魚類.xlsx"
        
        # Selenium設定（ヘッドレス）
        self.headless_mode = True
        
        # LINE Bot設定（本番用）
        self._configure_production_line_settings()
        
        # ログ設定（本番レベル）
        self.log_level = "INFO"
        
        print("🚀 本番モード設定を適用しました")
        print(f"   - ネットワークストレージ: {'有効' if self.use_network_storage else '無効（ローカル使用）'}")
        print("   - ヘッドレスモード")
        print("   - 確定処理: 個別設定")
    
    def _configure_test_line_settings(self):
        """テスト用LINE設定"""
        # 既存のテスト用設定
        self.line_channel_access_token = os.getenv('LINE_CHANNEL_ACCESS_TOKEN', 
            'IVEs7s1v5o2XdR9dCT/Q101NpmsF0FvK0MxyJ9pNTqKZAYRq7mTmlghp/2cTo+vhWA5S280hrT7LZQ9HEFZ5Eu6dDtKgKywE+pUnE6zW/5Z5N148JR9IOHYA4pS8aeN4GGgrwbFNrnWqcL2eoaIhuAdB04t89/1O/w1cDnyilFU=')
        self.line_channel_secret = os.getenv('LINE_CHANNEL_SECRET', '799dfc63eb76d45f6443f4be49833f47')
        self.line_channel_id = os.getenv('LINE_CHANNEL_ID', '2007929536')
        self.line_group_id = os.getenv('LINE_GROUP_ID', 'Ca23b038e0208192c6efe1640d471e977')
    
    def _configure_production_line_settings(self):
        """本番用LINE設定"""
        # 新しい本番用設定（環境変数から取得、デフォルト値は本番用）
        self.line_channel_access_token = os.getenv('LINE_CHANNEL_ACCESS_TOKEN', 
            'CoRdOqrzF2MLqc/WoFIgak/6aGmFMrP1jUTUoAMiEEXBOQ4XP7BZevrsnx7rDNV2i3tdMrxOvm1AIqygfA5fr+eKVajvndjTcxtOOnI4NOTQfJKvjLJ7fXxmcpn6G07mJ5BVNcU2c6uP68H9lJDeswdB04t89/1O/w1cDnyilFU=')
        self.line_channel_secret = os.getenv('LINE_CHANNEL_SECRET', '25d64f941d05535214a0462185672e91')
        self.line_channel_id = os.getenv('LINE_CHANNEL_ID', '2007929494')
        # self.line_group_id = os.getenv('LINE_GROUP_ID', 'Cd9d27348df3850582d0d31ea93e5afd7') #本番用
        self.line_group_id = os.getenv('LINE_GROUP_ID', 'Cc3557d7321fe7785abc86d8e2e6de168') #テスト用
    
    def _configure_common_settings(self):
        """共通設定（test_modeに関係なく適用）"""
        # ==============================================
        # 🚨 個別制御設定（現状の運用に合わせた設定）
        # ==============================================
        
        # 確定処理制御（現状：本番でも確定処理なし）
        self.enable_confirmation_process = self._get_confirmation_process_setting()
        
        # 検索条件制御（現状：本番もテストと同じ）
        self.enable_production_search_conditions = self._get_search_conditions_setting()
        
        print(f"⚙️ 個別制御設定:")
        print(f"   - 確定処理: {'有効' if self.enable_confirmation_process else '無効'}")
        print(f"   - 本番用検索条件: {'有効' if self.enable_production_search_conditions else '無効（テスト同様）'}")
        
        # ==============================================
        # 通常の共通設定
        # ==============================================
        
        # 日付別フォルダ設定
        self.use_date_folders = True
        self.date_folder_format = "%Y%m%d"  # YYYYMMDD形式
        
        # Selenium設定
        self.selenium_timeout = 15
        
        # PDF抽出設定
        self.pdf_extract_mode = "pdfplumber"  # "pdfplumber" or "pymupdf"
        
        # Excel設定
        self.excel_template_dir = self.project_root / "templates"
        self.dispatch_template = self.excel_template_dir / "配車表_テンプレート.xlsx"
        self.shipping_template = self.excel_template_dir / "出庫依頼_テンプレート.xlsx"
        
        # Google Drive API設定
        self.enable_google_drive = os.getenv('ENABLE_GOOGLE_DRIVE', 'false').lower() == 'true'
        self.google_drive_credentials_file = os.getenv('GOOGLE_DRIVE_CREDENTIALS', 
                                                      str(Path(__file__).parent.parent / 'credentials.json'))
        self.google_drive_token_file = os.getenv('GOOGLE_DRIVE_TOKEN', 
                                                str(self.project_root / 'token.json'))
        self.google_drive_folder_id = os.getenv('GOOGLE_DRIVE_FOLDER_ID', '')
        self.google_auth_code = os.getenv('GOOGLE_AUTH_CODE', '')
        
        # ログ設定
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
    
    def _get_confirmation_process_setting(self) -> bool:
        """確定処理の設定を取得"""
        # 環境変数での個別制御も可能
        env_setting = os.getenv('SMCL_ENABLE_CONFIRMATION', '').lower()
        if env_setting in ('true', '1', 'yes'):
            return True
        elif env_setting in ('false', '0', 'no'):
            return False
        
        # ==============================================
        # 🚨 現状の運用設定（手動で変更）
        # ==============================================
        # 現状：テスト・本番ともに確定処理なし
        CURRENT_CONFIRMATION_SETTING = False
        
        # 将来的に本番で確定処理を有効にする場合：
        # return not self.test_mode  # 本番のみ有効
        
        return CURRENT_CONFIRMATION_SETTING
    
    def _get_search_conditions_setting(self) -> bool:
        """検索条件設定を取得"""
        # 環境変数での個別制御も可能
        env_setting = os.getenv('SMCL_ENABLE_PRODUCTION_SEARCH', '').lower()
        if env_setting in ('true', '1', 'yes'):
            return True
        elif env_setting in ('false', '0', 'no'):
            return False
        
        # ==============================================
        # 🚨 現状の運用設定（手動で変更）
        # ==============================================
        # 現状：本番もテストと同じ検索条件
        CURRENT_SEARCH_SETTING = False
        
        # 将来的に本番で専用検索条件を使う場合：
        # return not self.test_mode  # 本番のみ専用条件
        
        return CURRENT_SEARCH_SETTING
        
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
    
    def is_google_drive_enabled(self) -> bool:
        """Google Drive機能が有効かチェック"""
        return self.enable_google_drive
    
    def is_google_drive_configured(self) -> bool:
        """Google Drive設定が完了しているかチェック"""
        if not self.enable_google_drive:
            return False
        credentials_path = Path(self.google_drive_credentials_file)
        return credentials_path.exists()
    
    def get_google_drive_config(self) -> dict:
        """Google Drive設定を取得"""
        return {
            'enabled': self.enable_google_drive,
            'credentials_file': self.google_drive_credentials_file,
            'token_file': self.google_drive_token_file,
            'folder_id': self.google_drive_folder_id,
            'auth_code': self.google_auth_code
        }

    
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

    def _get_network_base_path(self) -> str:
        """ネットワークドライブのベースパスを取得"""
        # 環境変数で設定可能
        network_path = os.getenv('SMCL_NETWORK_PATH', r'\\sv001\Userdata\納品リスト処理')
        return network_path
    
    def _should_use_network_storage(self) -> bool:
        """ネットワークストレージを使用するかどうかを判定"""
        # 環境変数で強制的に有効/無効にできる
        force_enable = os.getenv('SMCL_USE_NETWORK', '').lower() in ('true', '1', 'yes')
        force_disable = os.getenv('SMCL_DISABLE_NETWORK', '').lower() in ('true', '1', 'yes')
        
        if force_enable:
            return True
        if force_disable:
            return False
        
        # Windowsの場合のみデフォルトで有効
        system = platform.system().lower()
        if system == "windows":
            # ネットワークパスの存在確認
            try:
                network_path = Path(self.network_base_path)
                return network_path.exists()
            except Exception:
                return False
        
        return False
    
    def get_dated_output_dir(self, date_str: str = None) -> Path:
        """日付別の出力ディレクトリを取得"""
        if date_str is None:
            date_str = datetime.now().strftime(self.date_folder_format)
        
        if self.use_network_storage:
            # ネットワークドライブの日付別フォルダ
            base_path = Path(self.network_base_path)
            dated_dir = base_path / date_str
        else:
            # ローカルの日付別フォルダ
            if self.use_date_folders:
                dated_dir = self.output_dir / date_str
            else:
                dated_dir = self.output_dir
        
        # フォルダが存在しない場合は作成
        try:
            dated_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            # ネットワークドライブにアクセスできない場合はローカルフォルダにフォールバック
            if self.use_network_storage:
                print(f"ネットワークドライブにアクセスできません: {e}")
                print("ローカルフォルダを使用します")
                dated_dir = self.output_dir / date_str
                dated_dir.mkdir(parents=True, exist_ok=True)
            else:
                raise
        
        return dated_dir
    
    def get_dated_download_dir(self, date_str: str = None) -> Path:
        """日付別のダウンロードディレクトリを取得"""
        if date_str is None:
            date_str = datetime.now().strftime(self.date_folder_format)
        
        if self.use_network_storage:
            # ネットワークドライブの日付別フォルダ
            base_path = Path(self.network_base_path)
            dated_dir = base_path / "downloads" / date_str
        else:
            # ローカルの日付別フォルダ
            if self.use_date_folders:
                dated_dir = self.download_dir / date_str
            else:
                dated_dir = self.download_dir
        
        # フォルダが存在しない場合は作成
        try:
            dated_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            # ネットワークドライブにアクセスできない場合はローカルフォルダにフォールバック
            if self.use_network_storage:
                print(f"ネットワークドライブにアクセスできません: {e}")
                print("ローカルフォルダを使用します")
                dated_dir = self.download_dir / date_str
                dated_dir.mkdir(parents=True, exist_ok=True)
            else:
                raise
        
        return dated_dir
    
    def get_dated_logs_dir(self, date_str: str = None) -> Path:
        """日付別のログディレクトリを取得"""
        if date_str is None:
            date_str = datetime.now().strftime(self.date_folder_format)
        
        if self.use_network_storage:
            # ネットワークドライブの日付別フォルダ
            base_path = Path(self.network_base_path)
            dated_dir = base_path / "logs" / date_str
        else:
            # ローカルの日付別フォルダ
            if self.use_date_folders:
                dated_dir = self.logs_dir / date_str
            else:
                dated_dir = self.logs_dir
        
        # フォルダが存在しない場合は作成
        try:
            dated_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            # ネットワークドライブにアクセスできない場合はローカルフォルダにフォールバック
            if self.use_network_storage:
                print(f"ネットワークドライブにアクセスできません: {e}")
                print("ローカルフォルダを使用します")
                dated_dir = self.logs_dir / date_str
                dated_dir.mkdir(parents=True, exist_ok=True)
            else:
                raise
        
        return dated_dir
    
    def get_network_status(self) -> dict:
        """ネットワークドライブの状態を取得"""
        status = {
            "network_path": self.network_base_path,
            "use_network_storage": self.use_network_storage,
            "accessible": False,
            "error": None
        }
        
        if self.use_network_storage:
            try:
                network_path = Path(self.network_base_path)
                status["accessible"] = network_path.exists()
                if not status["accessible"]:
                    status["error"] = "パスが存在しません"
            except Exception as e:
                status["error"] = str(e)
        
        return status
    
    def print_configuration_status(self):
        """現在の設定状況を詳細表示"""
        print("=" * 60)
        print("🔧 SMCL システム設定状況")
        print("=" * 60)
        
        # 基本モード情報
        mode_icon = "🧪" if self.test_mode else "🚀"
        mode_name = "テストモード" if self.test_mode else "本番モード"
        print(f"{mode_icon} 動作モード: {mode_name}")
        print()
        
        # スクレイピング設定
        print("📱 スクレイピング設定:")
        print(f"   - ヘッドレスモード: {'有効' if self.headless_mode else '無効（ブラウザ表示）'}")
        print(f"   - 確定処理: {'有効' if self.enable_confirmation_process else '🚨 無効（現在の運用）'}")
        print(f"   - 検索日付: 📅 今日の日付を常に入力")
        print(f"   - 本番用検索条件: {'有効' if self.enable_production_search_conditions else '🚨 無効（テスト同様）'}")
        print(f"   - タイムアウト: {self.selenium_timeout}秒")
        print()
        
        # ファイル保存設定
        print("📁 ファイル保存設定:")
        if self.use_network_storage:
            network_status = self.get_network_status()
            status_icon = "✅" if network_status["accessible"] else "❌"
            print(f"   {status_icon} ネットワークストレージ: 有効")
            print(f"      パス: {self.network_base_path}")
            if not network_status["accessible"]:
                print(f"      エラー: {network_status.get('error', '不明')}")
        else:
            print("   📂 ローカルストレージ: 有効")
        
        print(f"   - ダウンロード: {self.download_dir}")
        print(f"   - 出力: {self.output_dir}")
        print(f"   - ログ: {self.logs_dir}")
        print(f"   - マスタファイル: {self.master_excel_path}")
        print()
        
        # LINE設定
        print("💬 LINE Bot 設定:")
        has_token = bool(self.line_channel_access_token and 
                        self.line_channel_access_token not in ['dummy', 'NEW_PRODUCTION_TOKEN_HERE'])
        has_secret = bool(self.line_channel_secret and 
                         self.line_channel_secret not in ['dummy', 'NEW_PRODUCTION_SECRET_HERE'])
        has_group = bool(self.line_group_id and self.line_group_id != 'dummy')
        
        token_icon = "✅" if has_token else "❌"
        secret_icon = "✅" if has_secret else "❌"
        group_icon = "✅" if has_group else "❌"
        
        print(f"   {token_icon} アクセストークン: {'設定済み' if has_token else '未設定'}")
        print(f"   {secret_icon} チャンネルシークレット: {'設定済み' if has_secret else '未設定'}")
        print(f"   {group_icon} グループID: {'設定済み' if has_group else '未設定'}")
        
        if has_token:
            token_preview = self.line_channel_access_token[:20] + "..." if len(self.line_channel_access_token) > 20 else self.line_channel_access_token
            print(f"      トークン（一部）: {token_preview}")
        
        line_configured = has_token and has_secret and has_group
        config_icon = "✅" if line_configured else "⚠️"
        print(f"   {config_icon} 設定状況: {'完了' if line_configured else '不完全'}")
        print()
        
        # Google Drive設定
        print("☁️ Google Drive 設定:")
        print(f"   - 機能: {'有効' if self.enable_google_drive else '無効'}")
        if self.enable_google_drive:
            credentials_exists = Path(self.google_drive_credentials_file).exists()
            cred_icon = "✅" if credentials_exists else "❌"
            print(f"   {cred_icon} 認証ファイル: {'存在' if credentials_exists else '不在'}")
            print(f"      パス: {self.google_drive_credentials_file}")
        print()
        
        # ログ設定
        print("📝 ログ設定:")
        print(f"   - レベル: {self.log_level}")
        print(f"   - 最大サイズ: {self.log_max_bytes // (1024*1024)}MB")
        print(f"   - バックアップ数: {self.log_backup_count}")
        print()
        
        # 設定切り替え方法
        print("🔄 設定切り替え方法:")
        print("【基本モード切り替え】")
        if self.test_mode:
            print("   本番モードに切り替えるには:")
            print("   services/core/config.py の21行目を以下に変更:")
            print("   self.test_mode = False")
        else:
            print("   テストモードに切り替えるには:")
            print("   services/core/config.py の21行目を以下に変更:")
            print("   self.test_mode = True")
        
        print("\n【個別制御設定】")
        print("   確定処理を有効にするには:")
        print("   services/core/config.py の _get_confirmation_process_setting() で:")
        print("   CURRENT_CONFIRMATION_SETTING = True に変更")
        
        print("\n   本番用検索条件を有効にするには:")
        print("   services/core/config.py の _get_search_conditions_setting() で:")
        print("   CURRENT_SEARCH_SETTING = True に変更")
        
        print("\n   【環境変数での制御も可能】")
        print("   export SMCL_ENABLE_CONFIRMATION=true  # 確定処理有効")
        print("   export SMCL_ENABLE_PRODUCTION_SEARCH=true  # 本番用検索条件有効")
        print()
        
        # 警告・エラーチェック
        warnings = []
        errors = []
        
        # マスタファイルチェック
        if not self.master_excel_path.exists():
            errors.append(f"マスタファイルが見つかりません: {self.master_excel_path}")
        
        # LINE設定チェック
        if not line_configured:
            warnings.append("LINE Bot設定が不完全です")
        
        # ネットワークドライブチェック（本番モードのみ）
        if not self.test_mode and self.use_network_storage:
            network_status = self.get_network_status()
            if not network_status["accessible"]:
                errors.append(f"ネットワークドライブにアクセスできません: {network_status.get('error', '不明')}")
        
        # 警告・エラー表示
        if warnings:
            print("⚠️ 警告:")
            for warning in warnings:
                print(f"   - {warning}")
            print()
        
        if errors:
            print("❌ エラー:")
            for error in errors:
                print(f"   - {error}")
            print()
        
        if not warnings and not errors:
            print("✅ すべての設定が正常です")
            print()
        
        print("=" * 60)
    
    def validate_configuration(self) -> dict:
        """設定の検証を行い、結果を辞書で返す"""
        result = {
            "valid": True,
            "warnings": [],
            "errors": [],
            "mode": "test" if self.test_mode else "production"
        }
        
        # マスタファイルチェック
        if not self.master_excel_path.exists():
            result["errors"].append(f"マスタファイルが見つかりません: {self.master_excel_path}")
            result["valid"] = False
        
        # LINE設定チェック
        line_configured = self.is_line_configured()
        if not line_configured:
            result["warnings"].append("LINE Bot設定が不完全です")
        
        # ネットワークドライブチェック（本番モードのみ）
        if not self.test_mode and self.use_network_storage:
            network_status = self.get_network_status()
            if not network_status["accessible"]:
                result["errors"].append(f"ネットワークドライブにアクセスできません: {network_status.get('error', '不明')}")
                result["valid"] = False
        
        return result