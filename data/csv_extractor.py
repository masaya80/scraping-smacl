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

try:
    import chardet
except ImportError:
    chardet = None


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
            
            # CSVファイルを読み込み（エンコーディング問題があるため、複数の方法を試行）
            df = self._read_csv_with_fallback(csv_path)
            
            if df is None or df.empty:
                self.logger.error("CSVファイルを読み込めませんでした")
                return []
            
            # 列インデックスで必要な列を特定（列名が文字化けしている場合に対応）
            product_code_col_idx = self._find_column_by_index_or_name(df, 108, ['商品コード（発注用）'])
            quantity_col_idx = self._find_column_by_index_or_name(df, 146, ['発注数量（バラ）'])
            warehouse_col_idx = self._find_column_by_index_or_name(df, 2, ['倉庫名', 'C列'])  # C列（index 2）を倉庫名として使用
            
            if product_code_col_idx is None or quantity_col_idx is None:
                self.logger.error(f"必要な列が見つかりません。商品コード列: {product_code_col_idx}, 数量列: {quantity_col_idx}")
                return []
            
            # データを抽出してDeliveryDocumentに変換
            file_date = self._extract_date_from_filename(csv_path)
            
            # 商品コードと数量の列を取得
            product_code_col = df.columns[product_code_col_idx]
            quantity_col = df.columns[quantity_col_idx]
            warehouse_col = df.columns[warehouse_col_idx] if warehouse_col_idx is not None else None
            
            self.logger.info(f"使用する列: 商品コード={product_code_col} (列{product_code_col_idx}), 数量={quantity_col} (列{quantity_col_idx}), 倉庫={warehouse_col} (列{warehouse_col_idx})")
            
            # 商品コード、倉庫名、数量の組み合わせでグループ化して集計
            group_cols = [product_code_col]
            if warehouse_col is not None:
                group_cols.append(warehouse_col)
            
            grouped = df.groupby(group_cols)[quantity_col].sum().reset_index()
            
            # DeliveryDocumentとDeliveryItemを作成
            document = DeliveryDocument()
            document.document_id = f"CSV_{file_date}_{csv_path.stem}"
            document.document_date = file_date
            document.supplier = "角上魚類"
            document.items = []
            
            for _, row in grouped.iterrows():
                item_code = str(row[product_code_col]).strip()
                quantity = int(row[quantity_col])
                warehouse = str(row[warehouse_col]).strip() if warehouse_col is not None and warehouse_col in row else ""
                
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
                delivery_item.warehouse = warehouse  # 倉庫名を設定
                delivery_item.notes = f"CSV抽出元: {csv_path.name}"
                
                document.items.append(delivery_item)
                self.logger.debug(f"商品コード: {item_code}, 数量: {quantity}, 倉庫: {warehouse}")
            
            self.logger.info(f"CSV抽出完了: {len(document.items)} 品目（ユニーク）を抽出")
            return [document] if document.items else []
            
        except Exception as e:
            self.logger.error(f"CSV抽出エラー: {str(e)}")
            return []
    
    def _detect_encoding(self, csv_path: Path) -> str:
        """CSVファイルのエンコーディングを検出"""
        try:
            # 直接一般的なエンコーディングを試行（より確実）
            return self._try_common_encodings(csv_path)
                
        except Exception as e:
            self.logger.warning(f"エンコーディング検出エラー: {str(e)}")
            return 'cp932'
    
    def _try_common_encodings(self, csv_path: Path) -> str:
        """一般的なエンコーディングを試行"""
        # 日本語CSVファイルでよく使われるエンコーディングを優先
        encodings = ['cp932', 'shift_jis', 'utf-8', 'utf-8-sig', 'iso-2022-jp']
        
        for encoding in encodings:
            try:
                # pandasで実際に読み込みテストを行う
                test_df = pd.read_csv(csv_path, encoding=encoding, nrows=1)
                
                # 日本語の列名があることを確認（文字化けしていないか）
                columns_str = str(test_df.columns.tolist())
                self.logger.debug(f"エンコーディング {encoding} での列名例: {columns_str[:100]}...")
                
                if '商品' in columns_str or '数量' in columns_str:
                    self.logger.info(f"エンコーディング試行成功（列名確認済み）: {encoding}")
                    return encoding
                elif '�' in columns_str:
                    # 文字化けがある場合は次を試行
                    self.logger.debug(f"エンコーディング {encoding} で文字化けを検出")
                    continue
                else:
                    # 文字化けがなく読み込めた場合は採用
                    self.logger.info(f"エンコーディング試行成功: {encoding}")
                    return encoding
                
            except (UnicodeDecodeError, UnicodeError, pd.errors.EmptyDataError):
                continue
            except Exception as e:
                self.logger.debug(f"エンコーディング試行エラー ({encoding}): {str(e)}")
                continue
        
        # すべて失敗した場合はcp932をデフォルトとして返す
        self.logger.warning("適切なエンコーディングが見つかりませんでした。cp932を使用します。")
        return 'cp932'
    
    def _read_csv_with_fallback(self, csv_path: Path):
        """複数の方法でCSVファイルを読み込み"""
        # 複数の読み込み設定を試行
        configs = [
            {'encoding': 'cp932', 'engine': 'c', 'low_memory': False},
            {'encoding': 'shift_jis', 'engine': 'c', 'low_memory': False},
            {'encoding': 'cp932', 'on_bad_lines': 'skip', 'engine': 'python'},
            {'encoding': 'shift_jis', 'on_bad_lines': 'skip', 'engine': 'python'},
            {'encoding': 'cp932', 'engine': 'python'},
            {'encoding': 'shift_jis', 'engine': 'python'},
            {'encoding': 'latin1', 'engine': 'c', 'low_memory': False},
            {'encoding': 'utf-8', 'engine': 'c', 'low_memory': False},
            {'encoding': 'latin1', 'on_bad_lines': 'skip', 'engine': 'python'},
            {'encoding': 'utf-8', 'on_bad_lines': 'skip', 'engine': 'python'},
        ]
        
        for config in configs:
            try:
                self.logger.debug(f"CSV読み込み試行: {config}")
                # CSVファイルを読み込み（エンジンがpythonの場合はlow_memoryを除外）
                read_params = config.copy()
                if config.get('engine') == 'python' and 'low_memory' in read_params:
                    read_params.pop('low_memory')
                
                df = pd.read_csv(csv_path, **read_params)
                
                self.logger.debug(f"読み込み結果: 列数={len(df.columns)}, 行数={len(df)}")
                
                # 最低限の行数と列数をチェック
                if len(df.columns) >= 150 and len(df) > 10:
                    self.logger.info(f"CSV読み込み成功: {config} (列数: {len(df.columns)}, 行数: {len(df)})")
                    return df
                elif len(df.columns) >= 150:
                    self.logger.warning(f"CSV読み込み（行数少ない）: {config} (列数: {len(df.columns)}, 行数: {len(df)})")
                    # 行数が少ないが、列数が正しい場合は使用
                    return df
                else:
                    self.logger.warning(f"CSV読み込み（条件不適合）: {config} (列数: {len(df.columns)}, 行数: {len(df)})")
                    
            except Exception as e:
                self.logger.warning(f"CSV読み込みエラー ({config}): {str(e)}")
                continue
        
        self.logger.error("すべての設定でCSV読み込みに失敗しました")
        return None
    
    def _find_column_by_index_or_name(self, df, expected_index: int, possible_names: list):
        """列インデックスまたは列名で列を特定"""
        try:
            # まず期待されるインデックスの列が存在するかチェック
            if expected_index < len(df.columns):
                self.logger.info(f"列{expected_index}を使用: {df.columns[expected_index]}")
                return expected_index
            
            # 列名で検索
            for name in possible_names:
                if name in df.columns:
                    col_idx = df.columns.get_loc(name)
                    self.logger.info(f"列名「{name}」で発見: 列{col_idx}")
                    return col_idx
            
            # 部分一致で検索
            for i, col_name in enumerate(df.columns):
                for name in possible_names:
                    if name in str(col_name):
                        self.logger.info(f"部分一致で発見「{name}」: 列{i} ({col_name})")
                        return i
            
            return None
            
        except Exception as e:
            self.logger.error(f"列検索エラー: {str(e)}")
            return None
    
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
        """最新のCSVファイルを検索（`directory` 直下と `directory/unconfirmed` の両方を探索）"""
        try:
            self.logger.info(f"CSVファイル検索開始: {directory}")
            
            # 直下と unconfirmed サブディレクトリを探索
            candidates = []
            candidates.extend(list(directory.glob(pattern)))
            unconfirmed_dir = directory / "unconfirmed"
            if unconfirmed_dir.exists():
                candidates.extend(list(unconfirmed_dir.glob(pattern)))
            
            if not candidates:
                self.logger.warning(f"パターンに一致するCSVファイルが見つかりません: {pattern}")
                return None
            
            # 更新時刻の新しい順で選択
            candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
            latest_file = candidates[0]
            self.logger.info(f"最新のCSVファイルを発見: {latest_file}")
            return latest_file
            
        except Exception as e:
            self.logger.error(f"CSVファイル検索エラー: {str(e)}")
            return None
    
    def find_today_csv_file(self, directory: Path) -> Optional[Path]:
        """今日の日付のCSVファイルを検索（直下と `unconfirmed` を探索）"""
        try:
            today_str = datetime.now().strftime('%Y%m%d')
            pattern = f"受注伝票_{today_str}*.csv"
            
            candidates = []
            candidates.extend(list(directory.glob(pattern)))
            unconfirmed_dir = directory / "unconfirmed"
            if unconfirmed_dir.exists():
                candidates.extend(list(unconfirmed_dir.glob(pattern)))
            
            if not candidates:
                self.logger.warning(f"今日の日付のCSVファイルが見つかりません: {pattern}")
                return None
            
            # 更新時刻の新しい順で選択
            candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
            latest_file = candidates[0]
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