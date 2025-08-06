#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CSV データ抽出モジュール
"""

import pandas as pd
from pathlib import Path
from typing import List, Dict, Optional
from dataclasses import dataclass
from datetime import datetime
import glob

from utils.logger import Logger
from data.pdf_extractor import DeliveryDocument, DeliveryItem


@dataclass
class OrderItem:
    """受注アイテムのデータクラス（内部使用用）"""
    item_code: str = ""
    quantity: int = 0
    extracted_date: str = ""
    source_file: str = ""


class CSVExtractor:
    """CSV データ抽出クラス"""
    
    def __init__(self):
        self.logger = Logger(__name__)
    
    def extract_order_data(self, csv_path: Path) -> List[DeliveryDocument]:
        """CSVファイルから受注データを抽出"""
        try:
            self.logger.info(f"CSV抽出開始: {csv_path}")
            
            if not csv_path.exists():
                self.logger.error(f"CSVファイルが見つかりません: {csv_path}")
                return []
            
            # CSVファイルを読み込み（CP932エンコーディング）
            df = pd.read_csv(csv_path, encoding='cp932')
            
            # 必要な列が存在するかチェック
            required_columns = ['商品コード（発注用）', '発注数量（バラ）']
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                self.logger.error(f"必要な列が見つかりません: {missing_columns}")
                return []
            
            # データを抽出してDeliveryDocumentに変換
            file_date = self._extract_date_from_filename(csv_path)
            
            # 商品コードと数量の組み合わせでグループ化して集計
            grouped = df.groupby('商品コード（発注用）')['発注数量（バラ）'].sum().reset_index()
            
            # DeliveryDocumentとDeliveryItemを作成
            document = DeliveryDocument()
            document.document_id = f"CSV_{file_date}_{csv_path.stem}"
            document.document_date = file_date
            document.supplier = "角上魚類"
            document.items = []
            
            for _, row in grouped.iterrows():
                item_code = str(row['商品コード（発注用）']).strip()
                quantity = int(row['発注数量（バラ）'])
                
                # 空の商品コードやゼロ数量はスキップ
                if not item_code or item_code == 'nan' or quantity <= 0:
                    continue
                
                # DeliveryItemを作成
                delivery_item = DeliveryItem()
                delivery_item.item_code = item_code
                delivery_item.item_name = f"商品コード_{item_code}"
                delivery_item.quantity = quantity
                delivery_item.unit = "個"
                delivery_item.delivery_date = file_date
                delivery_item.notes = f"CSV抽出元: {csv_path.name}"
                
                document.items.append(delivery_item)
                self.logger.debug(f"商品コード: {item_code}, 数量: {quantity}")
            
            self.logger.info(f"CSV抽出完了: {len(document.items)} 品目（ユニーク）を抽出")
            return [document] if document.items else []
            
        except Exception as e:
            self.logger.error(f"CSV抽出エラー: {str(e)}")
            return []
    
    def _extract_date_from_filename(self, csv_path: Path) -> str:
        """ファイル名から日付を抽出"""
        try:
            filename = csv_path.stem
            
            # 日付パターンを検索（例: 20250806）
            import re
            date_pattern = r'(\d{8})'
            date_match = re.search(date_pattern, filename)
            
            if date_match:
                date_str = date_match.group(1)
                # 日付をフォーマット
                formatted_date = f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:8]}"
                return formatted_date
            else:
                # ファイルの更新日時を使用
                return datetime.fromtimestamp(csv_path.stat().st_mtime).strftime('%Y-%m-%d')
                
        except Exception as e:
            self.logger.error(f"日付抽出エラー: {str(e)}")
            return datetime.now().strftime('%Y-%m-%d')
    
    def find_latest_csv_file(self, directory: Path, pattern: str = "受注伝票_*.csv") -> Optional[Path]:
        """最新のCSVファイルを検索"""
        try:
            self.logger.info(f"CSVファイル検索開始: {directory}")
            
            # パターンに一致するファイルを検索
            csv_files = list(directory.glob(pattern))
            
            if not csv_files:
                self.logger.warning(f"パターンに一致するCSVファイルが見つかりません: {pattern}")
                return None
            
            # ファイル名の日付部分で並び替え（新しい順）
            csv_files.sort(key=lambda x: x.stem, reverse=True)
            
            latest_file = csv_files[0]
            self.logger.info(f"最新のCSVファイルを発見: {latest_file}")
            
            return latest_file
            
        except Exception as e:
            self.logger.error(f"CSVファイル検索エラー: {str(e)}")
            return None
    
    def find_today_csv_file(self, directory: Path) -> Optional[Path]:
        """今日の日付のCSVファイルを検索"""
        try:
            today_str = datetime.now().strftime('%Y%m%d')
            pattern = f"受注伝票_{today_str}*.csv"
            
            csv_files = list(directory.glob(pattern))
            
            if not csv_files:
                self.logger.warning(f"今日の日付のCSVファイルが見つかりません: {pattern}")
                return None
            
            # 時刻が最新のファイルを選択
            csv_files.sort(key=lambda x: x.stem, reverse=True)
            
            latest_file = csv_files[0]
            self.logger.info(f"今日の日付のCSVファイルを発見: {latest_file}")
            
            return latest_file
            
        except Exception as e:
            self.logger.error(f"今日のCSVファイル検索エラー: {str(e)}")
            return None
    
    def extract_multiple_csv_files(self, directory: Path, pattern: str = "受注伝票_*.csv") -> List[DeliveryDocument]:
        """複数のCSVファイルからデータを抽出"""
        try:
            self.logger.info(f"複数CSV抽出開始: {directory}")
            
            csv_files = list(directory.glob(pattern))
            if not csv_files:
                self.logger.warning(f"CSVファイルが見つかりません: {directory}")
                return []
            
            all_documents = []
            for csv_file in csv_files:
                documents = self.extract_order_data(csv_file)
                all_documents.extend(documents)
            
            self.logger.info(f"複数CSV抽出完了: {len(all_documents)} 件の受注データ")
            return all_documents
            
        except Exception as e:
            self.logger.error(f"複数CSV抽出エラー: {str(e)}")
            return []