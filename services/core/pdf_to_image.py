#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF to Image 変換モジュール
"""

import io
import base64
from pathlib import Path
from typing import List, Optional, Union
from PIL import Image

from .logger import Logger


class PDFToImageConverter:
    """PDF から画像への変換クラス"""
    
    def __init__(self):
        self.logger = Logger(__name__)
        
        # 必要なライブラリの動的インポート
        try:
            from pdf2image import convert_from_path
            self.convert_from_path = convert_from_path
            self.pdf2image_available = True
            self.logger.info("pdf2image ライブラリが利用可能です")
        except ImportError:
            self.pdf2image_available = False
            self.logger.warning("pdf2image ライブラリが見つかりません。PDFから画像への変換は無効です")
    
    def convert_pdf_to_images(
        self, 
        pdf_path: Union[str, Path], 
        output_dir: Optional[Union[str, Path]] = None,
        dpi: int = 200,
        format: str = 'JPEG',
        quality: int = 85,
        max_pages: int = 10
    ) -> List[Path]:
        """
        PDFファイルを画像ファイルに変換
        
        Args:
            pdf_path: PDFファイルのパス
            output_dir: 出力ディレクトリ（指定しない場合は同じディレクトリ）
            dpi: 解像度
            format: 画像フォーマット（JPEG, PNG）
            quality: JPEG品質（1-100）
            max_pages: 最大変換ページ数
        
        Returns:
            生成された画像ファイルのパスリスト
        """
        if not self.pdf2image_available:
            self.logger.error("pdf2image ライブラリが利用できません")
            return []
        
        try:
            pdf_path = Path(pdf_path)
            if not pdf_path.exists():
                self.logger.error(f"PDFファイルが見つかりません: {pdf_path}")
                return []
            
            # 出力ディレクトリの設定
            if output_dir is None:
                output_dir = pdf_path.parent
            else:
                output_dir = Path(output_dir)
                output_dir.mkdir(parents=True, exist_ok=True)
            
            self.logger.info(f"PDF変換開始: {pdf_path.name}")
            
            # PDFを画像に変換
            images = self.convert_from_path(
                pdf_path,
                dpi=dpi,
                first_page=1,
                last_page=max_pages
            )
            
            if not images:
                self.logger.warning(f"PDFから画像を抽出できませんでした: {pdf_path}")
                return []
            
            # 画像ファイルを保存
            image_paths = []
            base_name = pdf_path.stem
            
            for i, image in enumerate(images, 1):
                # ファイル名を生成
                if format.upper() == 'JPEG':
                    ext = '.jpg'
                elif format.upper() == 'PNG':
                    ext = '.png'
                else:
                    ext = '.jpg'  # デフォルト
                
                image_filename = f"{base_name}_page_{i:02d}{ext}"
                image_path = output_dir / image_filename
                
                # 画像を保存
                if format.upper() == 'JPEG':
                    # JPEGの場合は品質設定を適用
                    image.save(image_path, format, quality=quality, optimize=True)
                else:
                    image.save(image_path, format)
                
                image_paths.append(image_path)
                self.logger.info(f"画像生成完了: {image_filename}")
            
            self.logger.info(f"PDF変換完了: {len(image_paths)}枚の画像を生成")
            return image_paths
            
        except Exception as e:
            self.logger.error(f"PDF変換エラー: {str(e)}")
            return []
    
    def convert_pdf_to_base64_images(
        self, 
        pdf_path: Union[str, Path], 
        dpi: int = 150,
        format: str = 'JPEG',
        quality: int = 80,
        max_pages: int = 5
    ) -> List[str]:
        """
        PDFをBase64エンコードされた画像データに変換
        
        Args:
            pdf_path: PDFファイルのパス
            dpi: 解像度
            format: 画像フォーマット
            quality: JPEG品質
            max_pages: 最大変換ページ数
        
        Returns:
            Base64エンコードされた画像データのリスト
        """
        if not self.pdf2image_available:
            self.logger.error("pdf2image ライブラリが利用できません")
            return []
        
        try:
            pdf_path = Path(pdf_path)
            if not pdf_path.exists():
                self.logger.error(f"PDFファイルが見つかりません: {pdf_path}")
                return []
            
            self.logger.info(f"PDFをBase64画像に変換開始: {pdf_path.name}")
            
            # PDFを画像に変換
            images = self.convert_from_path(
                pdf_path,
                dpi=dpi,
                first_page=1,
                last_page=max_pages
            )
            
            if not images:
                self.logger.warning(f"PDFから画像を抽出できませんでした: {pdf_path}")
                return []
            
            # 画像をBase64に変換
            base64_images = []
            
            for i, image in enumerate(images, 1):
                try:
                    # メモリ上で画像をエンコード
                    buffer = io.BytesIO()
                    
                    if format.upper() == 'JPEG':
                        image.save(buffer, format='JPEG', quality=quality, optimize=True)
                    else:
                        image.save(buffer, format=format)
                    
                    # Base64エンコード
                    buffer.seek(0)
                    image_base64 = base64.b64encode(buffer.getvalue()).decode('utf-8')
                    base64_images.append(image_base64)
                    
                    self.logger.info(f"ページ{i}をBase64に変換完了")
                    
                except Exception as e:
                    self.logger.error(f"ページ{i}のBase64変換エラー: {str(e)}")
                    continue
            
            self.logger.info(f"Base64変換完了: {len(base64_images)}枚の画像データを生成")
            return base64_images
            
        except Exception as e:
            self.logger.error(f"PDF Base64変換エラー: {str(e)}")
            return []
    
    def get_pdf_page_count(self, pdf_path: Union[str, Path]) -> int:
        """
        PDFのページ数を取得
        
        Args:
            pdf_path: PDFファイルのパス
        
        Returns:
            ページ数（エラーの場合は0）
        """
        if not self.pdf2image_available:
            return 0
        
        try:
            import fitz  # PyMuPDF
            doc = fitz.open(pdf_path)
            page_count = len(doc)
            doc.close()
            return page_count
        except ImportError:
            # PyMuPDFがない場合はpdf2imageで確認
            try:
                pdf_path = Path(pdf_path)
                if not pdf_path.exists():
                    return 0
                
                # 最初のページのみ変換してページ数を推定
                images = self.convert_from_path(pdf_path, dpi=50, last_page=1)
                if images:
                    # 実際のページ数を取得するには全変換が必要だが、ここでは概算
                    return 1  # 少なくとも1ページはある
                return 0
            except Exception:
                return 0
        except Exception as e:
            self.logger.error(f"PDFページ数取得エラー: {str(e)}")
            return 0
    
    def is_available(self) -> bool:
        """変換機能が利用可能かどうかを確認"""
        return self.pdf2image_available
    
    def get_optimal_settings_for_line(self) -> dict:
        """LINE送信に最適化された設定を取得"""
        return {
            'dpi': 150,  # 適度な解像度
            'format': 'JPEG',  # ファイルサイズ重視
            'quality': 80,  # バランスの取れた品質
            'max_pages': 3  # LINE送信を考慮したページ数制限
        }
