#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF画像変換処理モジュール
"""

from pathlib import Path
from typing import List, Dict, Optional
from datetime import datetime

from ..core.logger import Logger
from ..core.pdf_to_image import PDFToImageConverter


class PDFImageProcessor:
    """PDF画像変換処理クラス"""
    
    def __init__(self, config=None):
        self.logger = Logger(__name__)
        self.config = config
        self.pdf_converter = PDFToImageConverter()
        
    def process_all_pdfs(self, pdf_files: List[Path], output_dir: Path) -> Dict[str, List[Path]]:
        """
        複数のPDFファイルを画像に変換
        
        Args:
            pdf_files: 変換対象のPDFファイルリスト
            output_dir: 画像出力ディレクトリ
        
        Returns:
            PDFファイル名をキーとした画像パスリストの辞書
        """
        try:
            self.logger.info("PDF画像変換処理開始")
            
            if not self.pdf_converter.is_available():
                self.logger.error("PDF変換機能が利用できません")
                return {}
            
            # 出力ディレクトリを作成
            output_dir.mkdir(parents=True, exist_ok=True)
            
            converted_images = {}
            total_images = 0
            
            for pdf_file in pdf_files:
                try:
                    self.logger.info(f"PDF変換開始: {pdf_file.name}")
                    
                    # PDFを画像に変換
                    image_paths = self._convert_single_pdf(pdf_file, output_dir)
                    
                    if image_paths:
                        converted_images[pdf_file.name] = image_paths
                        total_images += len(image_paths)
                        self.logger.info(f"PDF変換完了: {pdf_file.name} -> {len(image_paths)}枚")
                    else:
                        self.logger.warning(f"PDF変換失敗: {pdf_file.name}")
                        converted_images[pdf_file.name] = []
                        
                except Exception as e:
                    self.logger.error(f"PDF変換エラー ({pdf_file.name}): {str(e)}")
                    converted_images[pdf_file.name] = []
            
            self.logger.info(f"PDF画像変換処理完了: {len(pdf_files)}ファイル -> {total_images}枚の画像")
            return converted_images
            
        except Exception as e:
            self.logger.error(f"PDF画像変換処理エラー: {str(e)}")
            return {}
    
    def _convert_single_pdf(self, pdf_file: Path, output_dir: Path) -> List[Path]:
        """
        単一PDFファイルを画像に変換
        
        Args:
            pdf_file: PDFファイルのパス
            output_dir: 画像出力ディレクトリ
        
        Returns:
            生成された画像ファイルのパスリスト
        """
        try:
            if not pdf_file.exists():
                self.logger.error(f"PDFファイルが見つかりません: {pdf_file}")
                return []
            
            # LINE送信に最適化された設定を取得
            settings = self.pdf_converter.get_optimal_settings_for_line()
            
            # ファイル種別に応じて設定を調整
            if "納品リスト" in pdf_file.name:
                settings['max_pages'] = 3  # 納品リストは最大3ページ
            else:
                settings['max_pages'] = 2  # 出庫依頼書・配車表は最大2ページ
            
            # PDFを画像に変換
            image_paths = self.pdf_converter.convert_pdf_to_images(
                pdf_path=pdf_file,
                output_dir=output_dir,
                **settings
            )
            
            return image_paths
            
        except Exception as e:
            self.logger.error(f"単一PDF変換エラー ({pdf_file.name}): {str(e)}")
            return []
    
    def get_image_summary(self, converted_images: Dict[str, List[Path]]) -> Dict[str, int]:
        """
        変換結果のサマリーを取得
        
        Args:
            converted_images: 変換結果の辞書
        
        Returns:
            サマリー情報の辞書
        """
        try:
            total_pdfs = len(converted_images)
            total_images = sum(len(images) for images in converted_images.values())
            successful_pdfs = sum(1 for images in converted_images.values() if images)
            failed_pdfs = total_pdfs - successful_pdfs
            
            # ファイル種別別の集計
            generated_pdfs = 0
            downloaded_pdfs = 0
            generated_images = 0
            downloaded_images = 0
            
            for pdf_name, images in converted_images.items():
                if "納品リスト" in pdf_name:
                    downloaded_pdfs += 1
                    downloaded_images += len(images)
                else:
                    generated_pdfs += 1
                    generated_images += len(images)
            
            summary = {
                "総PDFファイル数": total_pdfs,
                "総画像数": total_images,
                "成功PDFファイル数": successful_pdfs,
                "失敗PDFファイル数": failed_pdfs,
                "生成PDFファイル数": generated_pdfs,
                "生成PDF画像数": generated_images,
                "ダウンロードPDFファイル数": downloaded_pdfs,
                "ダウンロードPDF画像数": downloaded_images
            }
            
            return summary
            
        except Exception as e:
            self.logger.error(f"サマリー作成エラー: {str(e)}")
            return {}
    
    def cleanup_old_images(self, output_dir: Path, days_to_keep: int = 7) -> int:
        """
        古い画像ファイルをクリーンアップ
        
        Args:
            output_dir: 画像ディレクトリ
            days_to_keep: 保持日数
        
        Returns:
            削除したファイル数
        """
        try:
            if not output_dir.exists():
                return 0
            
            from datetime import timedelta
            cutoff_date = datetime.now() - timedelta(days=days_to_keep)
            
            deleted_count = 0
            for image_file in output_dir.glob("*.jpg"):
                try:
                    # ファイルの更新日時を確認
                    file_mtime = datetime.fromtimestamp(image_file.stat().st_mtime)
                    
                    if file_mtime < cutoff_date:
                        image_file.unlink()
                        deleted_count += 1
                        self.logger.debug(f"古い画像ファイルを削除: {image_file.name}")
                        
                except Exception as e:
                    self.logger.warning(f"ファイル削除エラー ({image_file.name}): {str(e)}")
                    continue
            
            if deleted_count > 0:
                self.logger.info(f"古い画像ファイルを削除しました: {deleted_count}個")
            
            return deleted_count
            
        except Exception as e:
            self.logger.error(f"画像クリーンアップエラー: {str(e)}")
            return 0
