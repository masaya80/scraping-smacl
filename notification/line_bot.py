#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LINE Bot 通知モジュール
"""

import requests
import json
from typing import Dict, List, Optional, Any
from datetime import datetime

from utils.logger import Logger
from data.excel_processor import ValidationError


class LineBotNotifier:
    """LINE Bot 通知クラス"""
    
    def __init__(self):
        self.logger = Logger(__name__)
        
        # 設定から取得
        from config.settings import Config
        config = Config()
        
        self.channel_access_token = config.line_channel_access_token
        self.user_id = config.line_user_id
        self.api_url = "https://api.line.me/v2/bot/message/push"
        
        # 設定確認
        if not self.channel_access_token or not self.user_id:
            self.logger.warning("LINE Bot設定が不完全です。通知機能は無効になります。")
            self.enabled = False
        else:
            self.enabled = True
            self.logger.info("LINE Bot通知機能が有効です")
    
    def send_message(self, message: str) -> bool:
        """LINE メッセージを送信"""
        if not self.enabled:
            self.logger.warning("LINE Bot が無効のため、メッセージ送信をスキップしました")
            return False
        
        try:
            headers = {
                'Authorization': f'Bearer {self.channel_access_token}',
                'Content-Type': 'application/json'
            }
            
            data = {
                'to': self.user_id,
                'messages': [
                    {
                        'type': 'text',
                        'text': message
                    }
                ]
            }
            
            response = requests.post(
                self.api_url,
                headers=headers,
                data=json.dumps(data),
                timeout=10
            )
            
            if response.status_code == 200:
                self.logger.info("LINE メッセージ送信成功")
                return True
            else:
                self.logger.error(f"LINE メッセージ送信失敗: {response.status_code} - {response.text}")
                return False
                
        except requests.exceptions.RequestException as e:
            self.logger.error(f"LINE API リクエストエラー: {str(e)}")
            return False
        except Exception as e:
            self.logger.error(f"LINE メッセージ送信エラー: {str(e)}")
            return False
    
    def send_process_summary(self, summary: Dict[str, Any]) -> bool:
        """処理サマリーを送信"""
        try:
            self.logger.info("処理サマリー通知送信")
            
            # サマリーメッセージを構築
            message = self._build_summary_message(summary)
            
            return self.send_message(message)
            
        except Exception as e:
            self.logger.error(f"処理サマリー送信エラー: {str(e)}")
            return False
    
    def _build_summary_message(self, summary: Dict[str, Any]) -> str:
        """サマリーメッセージを構築"""
        try:
            lines = [
                "📊 SMCL 納品リスト処理完了",
                "",
                f"🕐 処理日時: {summary.get('処理日時', '不明')}",
                "",
                "📈 処理結果:",
                f"  ✅ 正常データ: {summary.get('正常データ件数', 0)}件",
                f"  ❌ エラーデータ: {summary.get('エラーデータ件数', 0)}件",
                f"  📄 生成ファイル: {summary.get('生成ファイル数', 0)}個"
            ]
            
            # エラーがある場合は警告を追加
            error_count = summary.get('エラーデータ件数', 0)
            if error_count > 0:
                lines.extend([
                    "",
                    "⚠️ エラーが発生しています。",
                    "詳細はエラーレポートを確認してください。"
                ])
            else:
                lines.extend([
                    "",
                    "✨ すべて正常に処理されました！"
                ])
            
            return "\n".join(lines)
            
        except Exception as e:
            self.logger.error(f"サマリーメッセージ構築エラー: {str(e)}")
            return "処理完了通知でエラーが発生しました"
    
    def send_error_details(self, errors: List[ValidationError]) -> bool:
        """エラー詳細を送信"""
        try:
            self.logger.info("エラー詳細通知送信")
            
            if not errors:
                return True
            
            # エラー詳細メッセージを構築
            message = self._build_error_message(errors)
            
            return self.send_message(message)
            
        except Exception as e:
            self.logger.error(f"エラー詳細送信エラー: {str(e)}")
            return False
    
    def _build_error_message(self, errors: List[ValidationError]) -> str:
        """エラーメッセージを構築"""
        try:
            lines = [
                "🚨 エラー詳細レポート",
                "",
                f"エラー総数: {len(errors)}件"
            ]
            
            # エラーを種別ごとに集計
            error_summary = self._summarize_errors(errors)
            
            lines.append("")
            lines.append("📋 エラー種別:")
            
            for error_type, count in error_summary.items():
                lines.append(f"  • {error_type}: {count}件")
            
            # 上位エラーの詳細を表示（最大5件）
            lines.append("")
            lines.append("🔍 主要エラー:")
            
            for i, error in enumerate(errors[:5]):
                lines.append(f"  {i+1}. {error.item_name}")
                lines.append(f"     {error.description}")
                if i < 4 and i < len(errors) - 1:
                    lines.append("")
            
            if len(errors) > 5:
                lines.append(f"     ... 他 {len(errors) - 5}件")
            
            lines.extend([
                "",
                "📄 詳細はエラーレポートExcelを確認してください。"
            ])
            
            return "\n".join(lines)
            
        except Exception as e:
            self.logger.error(f"エラーメッセージ構築エラー: {str(e)}")
            return "エラー詳細の構築中にエラーが発生しました"
    
    def _summarize_errors(self, errors: List[ValidationError]) -> Dict[str, int]:
        """エラーを種別ごとに集計"""
        summary = {}
        
        for error in errors:
            error_type = error.error_type
            if error_type in summary:
                summary[error_type] += 1
            else:
                summary[error_type] = 1
        
        # 件数順にソート
        sorted_summary = dict(sorted(summary.items(), key=lambda x: x[1], reverse=True))
        
        return sorted_summary
    
    def send_start_notification(self) -> bool:
        """処理開始通知を送信"""
        try:
            self.logger.info("処理開始通知送信")
            
            message = (
                "🚀 SMCL 納品リスト処理開始\n"
                "\n"
                f"開始時刻: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                "\n"
                "処理内容:\n"
                "1. スマクラからの納品リストダウンロード\n"
                "2. PDFデータ抽出\n"
                "3. マスタとの付け合わせ\n"
                "4. 配車表・出庫依頼書作成\n"
                "\n"
                "完了までしばらくお待ちください..."
            )
            
            return self.send_message(message)
            
        except Exception as e:
            self.logger.error(f"処理開始通知エラー: {str(e)}")
            return False
    
    def send_phase_notification(self, phase: str, status: str = "開始") -> bool:
        """フェーズ別通知を送信"""
        try:
            self.logger.info(f"フェーズ通知送信: {phase} {status}")
            
            emoji_map = {
                "スクレイピング": "🕷️",
                "データ抽出": "📄",
                "マスタ付け合わせ": "🔍",
                "Excel生成": "📊",
                "通知": "📢"
            }
            
            emoji = emoji_map.get(phase, "⚙️")
            
            message = f"{emoji} {phase} {status}\n{datetime.now().strftime('%H:%M:%S')}"
            
            return self.send_message(message)
            
        except Exception as e:
            self.logger.error(f"フェーズ通知エラー: {str(e)}")
            return False
    
    def send_emergency_notification(self, error_message: str) -> bool:
        """緊急エラー通知を送信"""
        try:
            self.logger.info("緊急エラー通知送信")
            
            message = (
                "🆘 緊急エラー発生\n"
                "\n"
                f"発生時刻: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                "\n"
                f"エラー内容:\n{error_message}\n"
                "\n"
                "至急確認してください。"
            )
            
            return self.send_message(message)
            
        except Exception as e:
            self.logger.error(f"緊急エラー通知エラー: {str(e)}")
            return False
    
    def send_file_generation_notification(self, files: List[str]) -> bool:
        """ファイル生成通知を送信"""
        try:
            self.logger.info("ファイル生成通知送信")
            
            if not files:
                return True
            
            lines = [
                "📁 ファイル生成完了",
                "",
                f"生成ファイル数: {len(files)}個",
                ""
            ]
            
            for i, file_path in enumerate(files[:10], 1):  # 最大10個まで表示
                file_name = file_path.split('/')[-1] if '/' in file_path else file_path
                lines.append(f"{i}. {file_name}")
            
            if len(files) > 10:
                lines.append(f"... 他 {len(files) - 10}個")
            
            message = "\n".join(lines)
            
            return self.send_message(message)
            
        except Exception as e:
            self.logger.error(f"ファイル生成通知エラー: {str(e)}")
            return False
    
    def test_connection(self) -> bool:
        """LINE Bot接続テスト"""
        try:
            self.logger.info("LINE Bot接続テスト")
            
            test_message = (
                "🧪 LINE Bot接続テスト\n"
                "\n"
                f"テスト時刻: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                "\n"
                "このメッセージが届いていれば、\n"
                "LINE Bot設定は正常です。"
            )
            
            result = self.send_message(test_message)
            
            if result:
                self.logger.info("LINE Bot接続テスト成功")
            else:
                self.logger.error("LINE Bot接続テスト失敗")
            
            return result
            
        except Exception as e:
            self.logger.error(f"LINE Bot接続テストエラー: {str(e)}")
            return False 