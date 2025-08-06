#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ログ機能モジュール
"""

import logging
import logging.handlers
from pathlib import Path
from datetime import datetime
from typing import Optional


class Logger:
    """ログ管理クラス"""
    
    _loggers = {}
    _configured = False
    
    def __init__(self, name: str):
        self.name = name
        self.logger = self._get_logger(name)
    
    @classmethod
    def configure(cls, log_dir: Path, log_level: str = "INFO", 
                 log_format: str = None, max_bytes: int = 10*1024*1024,
                 backup_count: int = 5):
        """ログ設定を初期化"""
        if cls._configured:
            return
            
        # ログディレクトリを作成
        log_dir.mkdir(exist_ok=True)
        
        # ログファイルパス
        log_file = log_dir / f"smcl_system_{datetime.now().strftime('%Y%m%d')}.log"
        
        # フォーマット設定
        if not log_format:
            log_format = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
        
        formatter = logging.Formatter(log_format)
        
        # ルートロガーの設定
        root_logger = logging.getLogger()
        root_logger.setLevel(getattr(logging, log_level.upper()))
        
        # コンソールハンドラー
        console_handler = logging.StreamHandler()
        console_handler.setLevel(getattr(logging, log_level.upper()))
        console_handler.setFormatter(formatter)
        root_logger.addHandler(console_handler)
        
        # ファイルハンドラー（ローテーション対応）
        file_handler = logging.handlers.RotatingFileHandler(
            log_file, maxBytes=max_bytes, backupCount=backup_count,
            encoding='utf-8'
        )
        file_handler.setLevel(getattr(logging, log_level.upper()))
        file_handler.setFormatter(formatter)
        root_logger.addHandler(file_handler)
        
        cls._configured = True
    
    def _get_logger(self, name: str) -> logging.Logger:
        """ロガーインスタンスを取得"""
        if name not in self._loggers:
            # 初回設定時にログ設定を自動で行う
            if not self._configured:
                from config.settings import Config
                config = Config()
                self.configure(
                    log_dir=config.logs_dir,
                    log_level=config.log_level,
                    log_format=config.log_format,
                    max_bytes=config.log_max_bytes,
                    backup_count=config.log_backup_count
                )
            
            self._loggers[name] = logging.getLogger(name)
        
        return self._loggers[name]
    
    def debug(self, message: str, *args, **kwargs):
        """デバッグレベルのログを出力"""
        self.logger.debug(message, *args, **kwargs)
    
    def info(self, message: str, *args, **kwargs):
        """情報レベルのログを出力"""
        self.logger.info(message, *args, **kwargs)
    
    def warning(self, message: str, *args, **kwargs):
        """警告レベルのログを出力"""
        self.logger.warning(message, *args, **kwargs)
    
    def error(self, message: str, *args, **kwargs):
        """エラーレベルのログを出力"""
        self.logger.error(message, *args, **kwargs)
    
    def critical(self, message: str, *args, **kwargs):
        """クリティカルレベルのログを出力"""
        self.logger.critical(message, *args, **kwargs)
    
    def exception(self, exception: Exception, message: str = None):
        """例外情報を含むログを出力"""
        if message:
            self.logger.exception(f"{message}: {str(exception)}")
        else:
            self.logger.exception(f"例外が発生しました: {str(exception)}")
    
    def log_phase_start(self, phase: str, description: str = None):
        """フェーズ開始ログ"""
        message = f"=== {phase} 開始 ==="
        if description:
            message += f" ({description})"
        self.info(message)
    
    def log_phase_end(self, phase: str, success: bool = True, duration: float = None):
        """フェーズ終了ログ"""
        status = "成功" if success else "失敗"
        message = f"=== {phase} {status} ==="
        if duration is not None:
            message += f" (処理時間: {duration:.1f}秒)"
        
        if success:
            self.info(message)
        else:
            self.error(message)
    
    def log_operation(self, operation: str, details: dict = None):
        """操作ログ"""
        message = f"操作: {operation}"
        if details:
            detail_str = ", ".join([f"{k}={v}" for k, v in details.items()])
            message += f" ({detail_str})"
        self.info(message)
    
    def log_file_operation(self, operation: str, file_path: Path, success: bool = True):
        """ファイル操作ログ"""
        status = "成功" if success else "失敗"
        message = f"ファイル{operation}: {file_path} ({status})"
        
        if success:
            self.info(message)
        else:
            self.error(message)
    
    def log_data_processing(self, operation: str, count: int, details: dict = None):
        """データ処理ログ"""
        message = f"データ{operation}: {count}件"
        if details:
            detail_str = ", ".join([f"{k}={v}" for k, v in details.items()])
            message += f" ({detail_str})"
        self.info(message) 