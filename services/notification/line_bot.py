#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LINE Bot 通知モジュール
"""

import requests
import json
import os
import tempfile
from typing import Dict, List, Optional, Any, Union
from datetime import datetime
from pathlib import Path

from ..core.logger import Logger
from ..data_processing.excel_processor import ValidationError
from ..core.pdf_to_image import PDFToImageConverter
from ..core.cloud_storage import CloudStorageUploader


class LineBotNotifier:
    """LINE Bot 通知クラス"""
    
    def __init__(self):
        self.logger = Logger(__name__)
        
        # 設定から取得
        from ..core.config import Config
        config = Config()
        
        self.channel_access_token = config.line_channel_access_token
        self.group_id = config.get_line_target_id()  # Group ID専用
        self.api_url = "https://api.line.me/v2/bot/message/push"
        self.content_api_url = "https://api-data.line.me/v2/bot/message/content"
        
        # PDF変換クラスを初期化
        self.pdf_converter = PDFToImageConverter()
        
        # クラウドストレージクラスを初期化
        self.cloud_uploader = CloudStorageUploader()
        
        # 設定確認
        if not self.channel_access_token or not self.group_id or self.group_id == 'dummy':
            self.logger.warning("LINE Bot設定が不完全です。通知機能は無効になります。")
            self.enabled = False
        else:
            self.enabled = True
            self.logger.info("LINE Bot通知機能が有効です（グループチャット専用）")
    
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
                'to': self.group_id,
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
                f"  📊 生成Excel: {summary.get('生成Excelファイル数', 0)}個",
                f"  📄 生成PDF: {summary.get('生成PDFファイル数', 0)}個",
                f"  📥 ダウンロードPDF: {summary.get('ダウンロードPDFファイル数', 0)}個",
                f"  📋 総PDFファイル: {summary.get('総PDFファイル数', 0)}個",
                f"  🖼️ 生成画像: {summary.get('総画像数', 0)}枚",
                f"  ✅ 変換成功: {summary.get('画像変換成功数', 0)}ファイル",
                f"  ❌ 変換失敗: {summary.get('画像変換失敗数', 0)}ファイル"
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
    
    def send_image_from_url(self, original_content_url: str, preview_image_url: str, alt_text: str = "画像") -> bool:
        """画像URLからLINE画像メッセージを送信"""
        if not self.enabled:
            self.logger.warning("LINE Bot が無効のため、画像送信をスキップしました")
            return False
        
        try:
            headers = {
                'Authorization': f'Bearer {self.channel_access_token}',
                'Content-Type': 'application/json'
            }
            
            data = {
                'to': self.group_id,
                'messages': [
                    {
                        'type': 'image',
                        'originalContentUrl': original_content_url,
                        'previewImageUrl': preview_image_url,
                        'altText': alt_text
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
                self.logger.info("LINE 画像送信成功")
                return True
            else:
                self.logger.error(f"LINE 画像送信失敗: {response.status_code} - {response.text}")
                return False
                
        except Exception as e:
            self.logger.error(f"LINE 画像送信エラー: {str(e)}")
            return False
    
    def send_pdf_as_images(
        self, 
        pdf_path: str, 
        title: str = "PDF画像", 
        max_pages: int = 3,
        upload_to_cloud: bool = True,
        save_to_output: bool = True
    ) -> bool:
        """
        PDFファイルを画像に変換してLINEで送信
        
        Args:
            pdf_path: PDFファイルのパス
            title: 送信する画像のタイトル
            max_pages: 最大送信ページ数
            upload_to_cloud: クラウドストレージにアップロードするか
            save_to_output: outputディレクトリに画像を保存するか
        
        Returns:
            送信成功可否
        """
        if not self.enabled:
            self.logger.warning("LINE Bot が無効のため、PDF画像送信をスキップしました")
            return False
        
        if not self.pdf_converter.is_available():
            self.logger.error("PDF変換機能が利用できません")
            return False
        
        try:
            pdf_path = Path(pdf_path)
            if not pdf_path.exists():
                self.logger.error(f"PDFファイルが見つかりません: {pdf_path}")
                return False
            
            self.logger.info(f"PDF画像送信開始: {pdf_path.name}")
            
            # 一時ディレクトリを作成
            with tempfile.TemporaryDirectory() as temp_dir:
                temp_dir_path = Path(temp_dir)
                
                # PDFを画像に変換（LINE送信に最適化された設定）
                settings = self.pdf_converter.get_optimal_settings_for_line()
                settings['max_pages'] = max_pages
                
                image_paths = self.pdf_converter.convert_pdf_to_images(
                    pdf_path=pdf_path,
                    output_dir=temp_dir_path,
                    **settings
                )
                
                if not image_paths:
                    self.logger.error("PDF画像変換に失敗しました")
                    return False
                
                # outputディレクトリに画像を保存
                saved_image_paths = []
                if save_to_output:
                    saved_image_paths = self._save_images_to_output(image_paths, pdf_path)
                
                # 各画像を送信（メッセージなしでシンプルに）
                success_count = 0
                for i, image_path in enumerate(image_paths, 1):
                    try:
                        if upload_to_cloud:
                            # クラウドストレージにアップロードして送信
                            success = self._send_image_via_cloud_storage(
                                image_path, 
                                f"{title} - ページ {i}"
                            )
                        else:
                            # Base64で直接送信（非推奨：サイズ制限あり）
                            success = self._send_image_as_base64(
                                image_path,
                                f"{title} - ページ {i}"
                            )
                        
                        if success:
                            success_count += 1
                            self.logger.info(f"ページ {i} 送信成功")
                        else:
                            self.logger.warning(f"ページ {i} 送信失敗")
                            
                    except Exception as e:
                        self.logger.error(f"ページ {i} 送信エラー: {str(e)}")
                
                # 保存された画像パスをログ出力
                if saved_image_paths:
                    self.logger.info(f"画像をoutputディレクトリに保存しました: {len(saved_image_paths)}枚")
                    for saved_path in saved_image_paths:
                        self.logger.info(f"  - {saved_path}")
                
                return success_count > 0
                
        except Exception as e:
            self.logger.error(f"PDF画像送信エラー: {str(e)}")
            return False
    
    def _save_images_to_output(self, image_paths: List[Path], original_pdf_path: Path) -> List[Path]:
        """
        画像ファイルをoutputディレクトリに保存
        
        Args:
            image_paths: 一時ディレクトリ内の画像ファイルパスのリスト
            original_pdf_path: 元のPDFファイルのパス
        
        Returns:
            保存された画像ファイルのパスリスト
        """
        try:
            # 設定からoutputディレクトリを取得
            from ..core.config import Config
            config = Config()
            output_dir = Path(config.output_dir)
            
            # outputディレクトリが存在しない場合は作成
            output_dir.mkdir(parents=True, exist_ok=True)
            
            saved_paths = []
            base_name = original_pdf_path.stem  # 拡張子なしのファイル名
            
            for i, image_path in enumerate(image_paths, 1):
                try:
                    # 保存先ファイル名を生成
                    image_extension = image_path.suffix  # .jpg, .png など
                    output_filename = f"{base_name}_page_{i:02d}{image_extension}"
                    output_path = output_dir / output_filename
                    
                    # ファイルをコピー
                    import shutil
                    shutil.copy2(image_path, output_path)
                    
                    saved_paths.append(output_path)
                    
                except Exception as e:
                    self.logger.error(f"画像ファイル保存エラー ({image_path.name}): {str(e)}")
                    continue
            
            return saved_paths
            
        except Exception as e:
            self.logger.error(f"画像保存処理エラー: {str(e)}")
            return []
    
    def _send_image_via_cloud_storage(self, image_path: Path, alt_text: str) -> bool:
        """
        画像をクラウドストレージにアップロードしてURLで送信
        
        Args:
            image_path: 画像ファイルのパス
            alt_text: 代替テキスト
        
        Returns:
            送信成功可否
        """
        try:
            # クラウドストレージに画像をアップロード
            public_url = self.cloud_uploader.upload_image(image_path)
            
            if public_url:
                # URLが取得できた場合は画像メッセージとして送信
                self.logger.info(f"クラウドURL取得成功: {public_url}")
                return self.send_image_from_url(
                    original_content_url=public_url,
                    preview_image_url=public_url,  # プレビューも同じURLを使用
                    alt_text=alt_text
                )
            else:
                # アップロードに失敗した場合はテキスト情報を送信
                self.logger.warning("クラウドアップロードに失敗、テキスト情報で送信")
                return self._send_image_as_base64(image_path, alt_text)
            
        except Exception as e:
            self.logger.error(f"クラウドストレージ経由画像送信エラー: {str(e)}")
            # エラーの場合もテキスト情報で送信
            return self._send_image_as_base64(image_path, alt_text)
    
    def _send_image_as_base64(self, image_path: Path, alt_text: str) -> bool:
        """
        画像をBase64データとして送信（テキストメッセージとして）
        
        Note: LINE Bot APIでは直接Base64画像は送信できないため、
              画像の内容をテキストで説明する形にフォールバック
        
        Args:
            image_path: 画像ファイルのパス
            alt_text: 代替テキスト
        
        Returns:
            送信成功可否
        """
        try:
            # 画像ファイルの基本情報を送信
            file_size = image_path.stat().st_size
            size_mb = file_size / (1024 * 1024)
            
            message = (
                f"🖼️ {alt_text}\n"
                f"ファイル名: {image_path.name}\n"
                f"サイズ: {size_mb:.2f} MB\n"
                f"📁 ローカルに保存されました"
            )
            
            return self.send_message(message)
            
        except Exception as e:
            self.logger.error(f"Base64画像送信エラー: {str(e)}")
            return False
    
    def send_pdf_summary_with_images(self, pdf_files: List[Path], max_files: int = 3) -> bool:
        """
        複数のPDFファイルのサマリーと代表画像を送信
        
        Args:
            pdf_files: PDFファイルのリスト
            max_files: 最大処理ファイル数
        
        Returns:
            送信成功可否
        """
        if not self.enabled:
            self.logger.warning("LINE Bot が無効のため、PDFサマリー送信をスキップしました")
            return False
        
        try:
            if not pdf_files:
                self.send_message("📄 生成されたPDFファイルはありません")
                return True
            
            # サマリーメッセージを送信
            summary_lines = [
                "📄 生成PDF一覧",
                "",
                f"総ファイル数: {len(pdf_files)}",
                ""
            ]
            
            # ファイル一覧を追加
            for i, pdf_file in enumerate(pdf_files[:10], 1):  # 最大10件表示
                file_size = pdf_file.stat().st_size / (1024 * 1024)  # MB
                summary_lines.append(f"{i}. {pdf_file.name} ({file_size:.1f} MB)")
            
            if len(pdf_files) > 10:
                summary_lines.append(f"... 他 {len(pdf_files) - 10} ファイル")
            
            self.send_message("\n".join(summary_lines))
            
            # 主要なPDFファイルの画像を送信
            success_count = 0
            for i, pdf_file in enumerate(pdf_files[:max_files], 1):
                try:
                    self.logger.info(f"PDF画像送信: {pdf_file.name}")
                    success = self.send_pdf_as_images(
                        str(pdf_file),
                        f"生成ファイル {i}",
                        max_pages=2  # 各PDFの最初の2ページ
                    )
                    
                    if success:
                        success_count += 1
                        
                except Exception as e:
                    self.logger.error(f"PDF {pdf_file.name} の画像送信エラー: {str(e)}")
            
            # 送信結果を報告
            if success_count > 0:
                result_message = f"✅ {success_count}/{min(len(pdf_files), max_files)} ファイルの画像送信完了"
            else:
                result_message = "⚠️ PDF画像の送信に失敗しました"
            
            self.send_message(result_message)
            return success_count > 0
            
        except Exception as e:
            self.logger.error(f"PDFサマリー送信エラー: {str(e)}")
            return False
    
    def send_pdf_as_file(self, pdf_path: Union[str, Path], title: str = "PDFファイル") -> bool:
        """
        PDFファイルをドキュメントとしてLINEで送信
        
        Args:
            pdf_path: PDFファイルのパス
            title: 送信するPDFのタイトル
        
        Returns:
            送信成功可否
        """
        if not self.enabled:
            self.logger.warning("LINE Bot が無効のため、PDF送信をスキップしました")
            return False
        
        try:
            pdf_path = Path(pdf_path)
            if not pdf_path.exists():
                self.logger.error(f"PDFファイルが見つかりません: {pdf_path}")
                return False
            
            self.logger.info(f"PDF送信開始: {pdf_path.name}")
            
            # Google DriveにPDFをアップロード
            pdf_url = self.cloud_uploader.upload_pdf(pdf_path)
            
            if pdf_url:
                # ドキュメントメッセージとして送信
                return self._send_document_message(pdf_url, title, pdf_path.name)
            else:
                # Google Driveアップロードに失敗した場合はファイル情報を送信
                self.logger.warning("PDF アップロードに失敗、ファイル情報で送信")
                return self._send_pdf_info_message(pdf_path, title)
                
        except Exception as e:
            self.logger.error(f"PDF送信エラー: {str(e)}")
            return False
    
    def _send_document_message(self, document_url: str, title: str, filename: str) -> bool:
        """ドキュメントメッセージを送信"""
        try:
            headers = {
                'Authorization': f'Bearer {self.channel_access_token}',
                'Content-Type': 'application/json'
            }
            
            # Flex Message形式でPDFダウンロードリンクを作成
            flex_content = {
                "type": "bubble",
                "body": {
                    "type": "box",
                    "layout": "vertical",
                    "contents": [
                        {
                            "type": "text",
                            "text": "📄 PDFファイル",
                            "weight": "bold",
                            "size": "lg",
                            "color": "#333333"
                        },
                        {
                            "type": "text",
                            "text": title,
                            "size": "md",
                            "color": "#555555",
                            "margin": "sm"
                        },
                        {
                            "type": "text",
                            "text": filename,
                            "size": "sm",
                            "color": "#888888",
                            "margin": "sm"
                        }
                    ]
                },
                "footer": {
                    "type": "box",
                    "layout": "vertical",
                    "contents": [
                        {
                            "type": "button",
                            "action": {
                                "type": "uri",
                                "label": "PDFを開く",
                                "uri": document_url
                            },
                            "style": "primary",
                            "color": "#1DB446"
                        }
                    ]
                }
            }
            
            data = {
                'to': self.group_id,
                'messages': [
                    {
                        'type': 'flex',
                        'altText': f'{title} - {filename}',
                        'contents': flex_content
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
                self.logger.info("LINE PDF送信成功")
                return True
            else:
                self.logger.error(f"LINE PDF送信失敗: {response.status_code} - {response.text}")
                return False
                
        except Exception as e:
            self.logger.error(f"ドキュメントメッセージ送信エラー: {str(e)}")
            return False
    
    def _send_pdf_info_message(self, pdf_path: Path, title: str) -> bool:
        """PDFファイル情報をテキストメッセージとして送信"""
        try:
            file_size = pdf_path.stat().st_size
            size_mb = file_size / (1024 * 1024)
            
            message = (
                f"📄 {title}\n"
                f"ファイル名: {pdf_path.name}\n"
                f"サイズ: {size_mb:.2f} MB\n"
                f"📁 ローカルに保存されました"
            )
            
            return self.send_message(message)
            
        except Exception as e:
            self.logger.error(f"PDF情報メッセージ送信エラー: {str(e)}")
            return False
    
    def send_pdf_summary_with_files(self, pdf_files: List[Path], max_files: int = 3, send_as_files: bool = False) -> bool:
        """
        複数のPDFファイルのサマリーとファイル送信
        
        Args:
            pdf_files: PDFファイルのリスト
            max_files: 最大処理ファイル数
            send_as_files: Trueの場合PDFファイルとして送信、Falseの場合画像として送信
        
        Returns:
            送信成功可否
        """
        if not self.enabled:
            self.logger.warning("LINE Bot が無効のため、PDFサマリー送信をスキップしました")
            return False
        
        try:
            if not pdf_files:
                self.send_message("📄 PDFファイルはありません")
                return True
            
            # ファイルを種類別に分類
            generated_files = []
            downloaded_files = []
            
            for pdf_file in pdf_files:
                # ファイル名に「納品リスト」が含まれている場合はダウンロードファイル
                if "納品リスト" in pdf_file.name:
                    downloaded_files.append(pdf_file)
                else:
                    generated_files.append(pdf_file)
            
            # サマリーメッセージを送信
            summary_lines = [
                "📄 PDF一覧",
                "",
                f"総ファイル数: {len(pdf_files)}",
                f"  📄 生成PDF: {len(generated_files)}個",
                f"  📥 ダウンロードPDF: {len(downloaded_files)}個",
                f"送信形式: {'PDFファイル' if send_as_files else 'PDF画像'}",
                ""
            ]
            
            # ファイル一覧を追加
            file_count = 0
            if generated_files:
                summary_lines.append("📄 生成PDFファイル:")
                for pdf_file in generated_files[:5]:  # 最大5件表示
                    file_count += 1
                    file_size = pdf_file.stat().st_size / (1024 * 1024)  # MB
                    summary_lines.append(f"  {file_count}. {pdf_file.name} ({file_size:.1f} MB)")
                
                if len(generated_files) > 5:
                    summary_lines.append(f"     ... 他 {len(generated_files) - 5} ファイル")
                summary_lines.append("")
            
            if downloaded_files:
                summary_lines.append("📥 ダウンロードPDFファイル:")
                for pdf_file in downloaded_files[:5]:  # 最大5件表示
                    file_count += 1
                    file_size = pdf_file.stat().st_size / (1024 * 1024)  # MB
                    summary_lines.append(f"  {file_count}. {pdf_file.name} ({file_size:.1f} MB)")
                
                if len(downloaded_files) > 5:
                    summary_lines.append(f"     ... 他 {len(downloaded_files) - 5} ファイル")
            
            self.send_message("\n".join(summary_lines))
            
            # PDFファイルを送信
            success_count = 0
            for i, pdf_file in enumerate(pdf_files[:max_files], 1):
                try:
                    # ファイルタイプを判定
                    file_type = "ダウンロードファイル" if "納品リスト" in pdf_file.name else "生成ファイル"
                    
                    if send_as_files:
                        self.logger.info(f"PDF送信: {pdf_file.name}")
                        success = self.send_pdf_as_file(
                            str(pdf_file),
                            f"{file_type} {i}"
                        )
                    else:
                        self.logger.info(f"PDF画像送信: {pdf_file.name}")
                        success = self.send_pdf_as_images(
                            str(pdf_file),
                            f"{file_type} {i}",
                            max_pages=2  # 各PDFの最初の2ページ
                        )
                    
                    if success:
                        success_count += 1
                        
                except Exception as e:
                    self.logger.error(f"PDF {pdf_file.name} の送信エラー: {str(e)}")
            
            return success_count > 0
            
        except Exception as e:
            self.logger.error(f"PDFサマリー送信エラー: {str(e)}")
            return False
    
    def send_converted_images(self, converted_images: Dict[str, List[Path]]) -> bool:
        """
        変換済み画像を送信（疎結合版）
        
        Args:
            converted_images: PDFファイル名をキーとした画像パスリストの辞書
        
        Returns:
            送信成功可否
        """
        if not self.enabled:
            self.logger.warning("LINE Bot が無効のため、画像送信をスキップしました")
            return False
        
        try:
            if not converted_images:
                self.send_message("📄 変換された画像はありません")
                return True
            
            # 画像統計を計算
            total_pdfs = len(converted_images)
            total_images = sum(len(images) for images in converted_images.values())
            successful_pdfs = sum(1 for images in converted_images.values() if images)
            
            # サマリーメッセージを送信
            summary_lines = [
                "📄 PDF画像一覧",
                "",
                f"変換PDFファイル数: {total_pdfs}",
                f"生成画像数: {total_images}",
                f"変換成功: {successful_pdfs}ファイル",
                ""
            ]
            
            # ファイル別詳細を追加
            for pdf_name, image_paths in converted_images.items():
                if image_paths:
                    file_type = "📥 納品リスト" if "納品リスト" in pdf_name else "📄 生成ファイル"
                    summary_lines.append(f"{file_type}: {pdf_name} ({len(image_paths)}枚)")
            
            self.send_message("\n".join(summary_lines))
            
            # 各画像を送信
            success_count = 0
            total_sent = 0
            
            for pdf_name, image_paths in converted_images.items():
                if not image_paths:
                    continue
                
                for i, image_path in enumerate(image_paths, 1):
                    try:
                        # 画像ファイルが存在するか確認
                        if not image_path.exists():
                            self.logger.warning(f"画像ファイルが見つかりません: {image_path}")
                            continue
                        
                        # 画像を直接送信
                        success = self._send_image_via_cloud_storage(
                            image_path,
                            f"{pdf_name} - ページ {i}"
                        )
                        
                        if success:
                            success_count += 1
                            self.logger.info(f"画像送信成功: {image_path.name}")
                        else:
                            self.logger.warning(f"画像送信失敗: {image_path.name}")
                        
                        total_sent += 1
                        
                    except Exception as e:
                        self.logger.error(f"画像送信エラー ({image_path.name}): {str(e)}")
                        total_sent += 1
            
            self.logger.info(f"画像送信完了: {success_count}/{total_sent} 成功")
            return success_count > 0
            
        except Exception as e:
            self.logger.error(f"変換済み画像送信エラー: {str(e)}")
            return False
    
    def send_integrated_completion_notification(
        self,
        summary: Dict[str, Any],
        error_data: List = None,
        converted_images: Dict[str, List[Path]] = None,
        max_images: int = 3,
        send_all_delivery_lists: bool = True
    ) -> bool:
        """
        統合された処理完了通知を送信（ビジネス向け）
        
        Args:
            summary: 処理結果のサマリー
            error_data: エラーデータのリスト
            converted_images: 変換済み画像の辞書
            max_images: 最大送信画像数
        
        Returns:
            送信成功可否
        """
        if not self.enabled:
            self.logger.warning("LINE Bot が無効のため、統合通知をスキップしました")
            return False
        
        try:
            self.logger.info("統合完了通知の送信開始")
            
            # メイン通知メッセージを送信
            main_message = self._build_integrated_message(summary, error_data)
            self.send_message(main_message)
            
            # 重要なPDF画像のみを送信（ビジネス側に必要最小限）
            if converted_images:
                self._send_business_relevant_images(converted_images, max_images, send_all_delivery_lists)
            
            self.logger.info("統合完了通知の送信完了")
            return True
            
        except Exception as e:
            self.logger.error(f"統合完了通知送信エラー: {str(e)}")
            return False
    
    def _build_integrated_message(self, summary: Dict[str, Any], error_data: List = None) -> str:
        """統合されたビジネス向けメッセージを構築"""
        try:
            lines = [
                "✅ SMCL 納品リスト処理完了",
                "",
                f"🕐 処理日時: {summary.get('処理日時', '不明')}",
                ""
            ]
            
            # エラーの有無のみを表示（詳細な数値は削除）
            error_count = summary.get('エラーデータ件数', 0)
            
            if error_count == 0:
                lines.extend([
                    "🎉 すべて正常に処理されました！"
                ])
            else:
                lines.extend([
                    "⚠️ 一部エラーがあります。詳細確認が必要です。"
                ])
            
            # 画像がある場合のみ画像表示の案内を追加
            total_images = summary.get('総画像数', 0)
            
            return "\n".join(lines)
            
        except Exception as e:
            self.logger.error(f"統合メッセージ構築エラー: {str(e)}")
            return "処理完了通知でエラーが発生しました"
    
    def _send_business_relevant_images(self, converted_images: Dict[str, List[Path]], max_images: int, send_all_delivery_lists: bool = True) -> bool:
        """
        ビジネス側に関連する重要な画像のみを送信
        
        Args:
            converted_images: 変換済み画像の辞書
            max_images: 最大送信画像数
            send_all_delivery_lists: 全ての納品リスト画像を送信するか
        
        Returns:
            送信成功可否
        """
        try:
            if not converted_images:
                return True
            
            # ビジネス側に重要なPDFファイルを優先順位付けして選択
            priority_files = []
            
            # 1. 出庫依頼書（最も重要）
            for pdf_name, image_paths in converted_images.items():
                if "出庫依頼" in pdf_name and image_paths:
                    priority_files.extend([(pdf_name, path, 1) for path in image_paths[:2]])  # 最大2ページ
            
            # 2. 配車表・アリスト
            for pdf_name, image_paths in converted_images.items():
                if ("アリスト" in pdf_name or "配車" in pdf_name or "LT" in pdf_name) and image_paths:
                    priority_files.extend([(pdf_name, path, 2) for path in image_paths[:1]])  # 最大1ページ
                    
            # 3. 納品リスト（参考用）
            for pdf_name, image_paths in converted_images.items():
                if "納品リスト" in pdf_name and image_paths:
                    if send_all_delivery_lists:
                        # 全ての納品リスト画像を送信
                        priority_files.extend([(pdf_name, path, 3) for path in image_paths])
                        self.logger.info(f"納品リスト全画像追加: {pdf_name} ({len(image_paths)}枚)")
                    else:
                        # 従来通り最大1ページ
                        priority_files.extend([(pdf_name, path, 3) for path in image_paths[:1]])  # 最大1ページ
            
            # 優先度でソートし、最大画像数まで制限
            priority_files.sort(key=lambda x: x[2])
            
            if send_all_delivery_lists:
                # 全ての納品リストを送信する場合
                # 優先度1,2（出庫依頼書、配車表）を優先選択
                high_priority_files = [f for f in priority_files if f[2] <= 2]
                delivery_list_files = [f for f in priority_files if f[2] == 3]  # 納品リスト
                
                # 高優先度ファイルを制限内で選択
                selected_high_priority = high_priority_files[:max_images]
                
                # 残りの枠で納品リストを追加（制限なし）
                selected_files = selected_high_priority + delivery_list_files
                
                self.logger.info(f"高優先度画像: {len(selected_high_priority)}枚, 納品リスト: {len(delivery_list_files)}枚")
            else:
                # 従来通りの制限
                selected_files = priority_files[:max_images]
            
            if not selected_files:
                self.logger.info("送信する重要画像が見つかりませんでした")
                return True
            
            # 選択された重要画像を送信
            success_count = 0
            for pdf_name, image_path, priority in selected_files:
                try:
                    # 画像が存在するか確認
                    if not image_path.exists():
                        self.logger.warning(f"画像ファイルが見つかりません: {image_path}")
                        continue
                    
                    # ファイルタイプに応じたタイトルを設定
                    if "出庫依頼" in pdf_name:
                        title = "📄 出庫依頼書"
                    elif any(keyword in pdf_name for keyword in ["アリスト", "配車", "LT"]):
                        title = "🚛 配車表"
                    elif "納品リスト" in pdf_name:
                        title = "📋 納品リスト"
                    else:
                        title = "📄 生成ファイル"
                    
                    # 画像を送信
                    success = self._send_image_via_cloud_storage(
                        image_path,
                        title
                    )
                    
                    if success:
                        success_count += 1
                        self.logger.info(f"重要画像送信成功: {image_path.name}")
                    else:
                        self.logger.warning(f"重要画像送信失敗: {image_path.name}")
                        
                except Exception as e:
                    self.logger.error(f"重要画像送信エラー ({image_path.name}): {str(e)}")
            
            self.logger.info(f"重要画像送信完了: {success_count}/{len(selected_files)} 成功")
            return success_count > 0
            
        except Exception as e:
            self.logger.error(f"ビジネス関連画像送信エラー: {str(e)}")
            return False