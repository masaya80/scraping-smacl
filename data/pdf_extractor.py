#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF データ抽出モジュール
"""

import re
from pathlib import Path
from typing import List, Dict, Optional, Any
from dataclasses import dataclass

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

from utils.logger import Logger


@dataclass
class DeliveryItem:
    """納品アイテムのデータクラス"""
    item_code: str = ""
    item_name: str = ""
    quantity: int = 0
    unit: str = ""
    unit_price: float = 0.0
    total_price: float = 0.0
    delivery_date: str = ""
    destination: str = ""
    supplier: str = ""
    warehouse: str = ""  # 倉庫名を追加
    notes: str = ""


@dataclass
class DeliveryDocument:
    """納品書類のデータクラス"""
    document_id: str = ""
    document_date: str = ""
    delivery_destination: str = ""
    supplier: str = ""
    total_amount: float = 0.0
    items: List[DeliveryItem] = None
    
    def __post_init__(self):
        if self.items is None:
            self.items = []


class PDFExtractor:
    """PDF データ抽出クラス"""
    
    def __init__(self, extract_mode: str = "pdfplumber"):
        self.extract_mode = extract_mode
        self.logger = Logger(__name__)
        
        # ライブラリの確認
        if extract_mode == "pdfplumber" and not pdfplumber:
            self.logger.warning("pdfplumber がインストールされていません。PyMuPDF を使用します。")
            self.extract_mode = "pymupdf"
        
        if extract_mode == "pymupdf" and not fitz:
            self.logger.warning("PyMuPDF がインストールされていません。pdfplumber を使用します。")
            self.extract_mode = "pdfplumber"
        
        if not pdfplumber and not fitz:
            raise ImportError("PDF処理ライブラリがインストールされていません。pdfplumber または PyMuPDF をインストールしてください。")
    
    def extract_delivery_data(self, pdf_path: Path) -> List[DeliveryDocument]:
        """PDFファイルから納品データを抽出"""
        try:
            self.logger.info(f"PDF抽出開始: {pdf_path}")
            
            if not pdf_path.exists():
                self.logger.error(f"PDFファイルが見つかりません: {pdf_path}")
                return []
            
            # PDFファイル名から基本情報を抽出
            file_info = self._extract_file_info(pdf_path)
            
            # PDF内容を抽出
            if self.extract_mode == "pdfplumber":
                text_content = self._extract_with_pdfplumber(pdf_path)
            else:
                text_content = self._extract_with_pymupdf(pdf_path)
            
            if not text_content:
                self.logger.warning(f"PDFからテキストが抽出できませんでした: {pdf_path}")
                return []
            
            # 抽出したテキストから納品データを解析
            documents = self._parse_delivery_text(text_content, file_info)
            
            self.logger.info(f"PDF抽出完了: {len(documents)} 件の納品書類を抽出")
            return documents
            
        except Exception as e:
            self.logger.error(f"PDF抽出エラー: {str(e)}")
            return []
    
    def _extract_file_info(self, pdf_path: Path) -> Dict[str, str]:
        """ファイル名から基本情報を抽出"""
        filename = pdf_path.stem
        info = {"filename": filename}
        
        # 日時パターンを抽出（例: 20250718074638）
        date_pattern = r'(\d{8})(\d{6})'
        date_match = re.search(date_pattern, filename)
        if date_match:
            date_str = date_match.group(1)
            time_str = date_match.group(2)
            # 日付をフォーマット
            formatted_date = f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:8]}"
            formatted_time = f"{time_str[:2]}:{time_str[2:4]}:{time_str[4:6]}"
            info["extracted_date"] = formatted_date
            info["extracted_time"] = formatted_time
        
        # 文書種類を判定
        if "納品リストD" in filename:
            info["document_type"] = "納品リストD"
        elif "集計" in filename:
            info["document_type"] = "集計"
        
        return info
    
    def _extract_with_pdfplumber(self, pdf_path: Path) -> str:
        """pdfplumber を使用してテキストを抽出"""
        try:
            with pdfplumber.open(pdf_path) as pdf:
                text_content = ""
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text_content += page_text + "\n"
                
                # テーブルデータも抽出を試行
                tables = []
                for page in pdf.pages:
                    page_tables = page.extract_tables()
                    if page_tables:
                        tables.extend(page_tables)
                
                # テーブルデータをテキストに変換
                if tables:
                    text_content += "\n--- テーブルデータ ---\n"
                    for table in tables:
                        for row in table:
                            if row:
                                row_text = "\t".join([str(cell) if cell else "" for cell in row])
                                text_content += row_text + "\n"
                
                return text_content
                
        except Exception as e:
            self.logger.error(f"pdfplumber での抽出エラー: {str(e)}")
            return ""
    
    def _extract_with_pymupdf(self, pdf_path: Path) -> str:
        """PyMuPDF (fitz) を使用してテキストを抽出"""
        try:
            doc = fitz.open(pdf_path)
            text_content = ""
            
            for page_num in range(len(doc)):
                page = doc.load_page(page_num)
                page_text = page.get_text()
                text_content += page_text + "\n"
            
            doc.close()
            return text_content
            
        except Exception as e:
            self.logger.error(f"PyMuPDF での抽出エラー: {str(e)}")
            return ""
    
    def _parse_delivery_text(self, text_content: str, file_info: Dict[str, str]) -> List[DeliveryDocument]:
        """抽出したテキストから納品データを解析"""
        try:
            documents = []
            
            # 基本的な納品書類情報を作成
            document = DeliveryDocument()
            document.document_id = file_info.get("filename", "")
            document.document_date = file_info.get("extracted_date", "")
            
            # テキストから各種情報を抽出
            self._extract_supplier_info(text_content, document)
            self._extract_destination_info(text_content, document)
            self._extract_items(text_content, document)
            self._calculate_totals(document)
            
            if document.items:
                documents.append(document)
                self.logger.info(f"納品書類を解析: {len(document.items)} 品目")
            else:
                self.logger.warning("品目が見つかりませんでした")
            
            return documents
            
        except Exception as e:
            self.logger.error(f"テキスト解析エラー: {str(e)}")
            return []
    
    def _extract_supplier_info(self, text: str, document: DeliveryDocument):
        """供給者情報を抽出"""
        try:
            # 供給者名のパターンを検索
            supplier_patterns = [
                r'仕入先[：:]\s*(.+)',
                r'供給者[：:]\s*(.+)',
                r'販売者[：:]\s*(.+)',
            ]
            
            for pattern in supplier_patterns:
                match = re.search(pattern, text)
                if match:
                    document.supplier = match.group(1).strip()
                    break
            
            if not document.supplier:
                # デフォルト設定
                document.supplier = "角上魚類"
                
        except Exception as e:
            self.logger.error(f"供給者情報抽出エラー: {str(e)}")
    
    def _extract_destination_info(self, text: str, document: DeliveryDocument):
        """納品先情報を抽出"""
        try:
            # 納品先のパターンを検索
            destination_patterns = [
                r'納品先[：:]\s*(.+)',
                r'お客様[：:]\s*(.+)',
                r'配送先[：:]\s*(.+)',
            ]
            
            for pattern in destination_patterns:
                match = re.search(pattern, text)
                if match:
                    document.delivery_destination = match.group(1).strip()
                    break
            
        except Exception as e:
            self.logger.error(f"納品先情報抽出エラー: {str(e)}")
    
    def _extract_items(self, text: str, document: DeliveryDocument):
        """品目情報を抽出"""
        try:
            # まず商品コードを抽出して品目を作成
            item_codes = self._extract_item_codes(text)
            
            if item_codes:
                # 商品コードが見つかった場合、各コードに対してアイテムを作成
                for code in item_codes:
                    item = DeliveryItem()
                    item.item_code = code
                    # 商品名は後でマスタから取得
                    item.item_name = f"商品コード_{code}"
                    item.quantity = 1  # デフォルト数量
                    document.items.append(item)
                    self.logger.info(f"商品コード抽出: {code}")
            else:
                # 従来の方法でテーブル形式のデータを探す
                self._extract_items_from_table(text, document)
                
                # 品目が見つからない場合、別のパターンで検索
                if not document.items:
                    self._extract_items_alternative(text, document)
                
        except Exception as e:
            self.logger.error(f"品目情報抽出エラー: {str(e)}")
    
    def _extract_item_codes(self, text: str) -> List[str]:
        """PDFから商品コードを抽出"""
        try:
            import re
            
            # 商品コードのパターンを検索
            code_patterns = [
                r'商品コード[：:\s]*(\d+)',  # 「商品コード ： 134490」形式
                r'アイテムコード[：:\s]*(\d+)',
                r'品番[：:\s]*(\d+)',
            ]
            
            extracted_codes = []
            
            for pattern in code_patterns:
                matches = re.findall(pattern, text, re.IGNORECASE)
                if matches:
                    extracted_codes.extend(matches)
                    self.logger.info(f"商品コードパターン「{pattern}」で発見: {matches}")
            
            # 重複を除去
            unique_codes = list(set(extracted_codes))
            self.logger.info(f"抽出された商品コード: {unique_codes}")
            
            return unique_codes
            
        except Exception as e:
            self.logger.error(f"商品コード抽出エラー: {str(e)}")
            return []
    
    def _extract_items_from_table(self, text: str, document: DeliveryDocument):
        """テーブル形式から品目情報を抽出（従来の方法）"""
        try:
            lines = text.split('\n')
            in_table = False
            
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                
                # テーブルヘッダーを検出
                if any(header in line for header in ['商品名', '品名', '数量', '単価', '金額']):
                    in_table = True
                    continue
                
                if in_table:
                    # 品目データを解析
                    item = self._parse_item_line(line)
                    if item and item.item_name:
                        document.items.append(item)
                        
                    # テーブル終了を検出
                    if any(end_word in line for end_word in ['合計', '小計', '総額', '---']):
                        in_table = False
                        
        except Exception as e:
            self.logger.error(f"テーブル形式品目抽出エラー: {str(e)}")
    
    def _parse_item_line(self, line: str) -> Optional[DeliveryItem]:
        """1行の品目データを解析"""
        try:
            # タブ区切りまたは空白区切りでデータを分割
            parts = re.split(r'\t+|\s{2,}', line)
            parts = [part.strip() for part in parts if part.strip()]
            
            if len(parts) < 2:
                return None
            
            item = DeliveryItem()
            
            # データの構造を推測して割り当て
            if len(parts) >= 4:
                item.item_name = parts[0]
                item.quantity = self._parse_number(parts[1], int, 0)
                item.unit = parts[2] if len(parts) > 2 else ""
                item.unit_price = self._parse_number(parts[3], float, 0.0)
                
                if len(parts) >= 5:
                    item.total_price = self._parse_number(parts[4], float, 0.0)
                else:
                    item.total_price = item.quantity * item.unit_price
            
            elif len(parts) >= 2:
                item.item_name = parts[0]
                # 数量または金額として解釈
                number_value = self._parse_number(parts[1], float, 0.0)
                if number_value > 1000:  # 金額として判定
                    item.total_price = number_value
                else:  # 数量として判定
                    item.quantity = int(number_value)
            
            return item if item.item_name else None
            
        except Exception as e:
            self.logger.error(f"品目行解析エラー: {str(e)}")
            return None
    
    def _extract_items_alternative(self, text: str, document: DeliveryDocument):
        """代替方式で品目を抽出"""
        try:
            # 商品名らしきパターンを探す
            item_patterns = [
                r'(\w+(?:\s+\w+)*)\s+(\d+)\s*(個|本|kg|g)?\s*(\d+(?:,\d{3})*)',
                r'(.+?)\s+(\d+)\s+(\d+(?:,\d{3})*)'
            ]
            
            for pattern in item_patterns:
                matches = re.finditer(pattern, text)
                for match in matches:
                    item = DeliveryItem()
                    item.item_name = match.group(1).strip()
                    item.quantity = int(match.group(2))
                    
                    if len(match.groups()) >= 4:
                        item.unit = match.group(3) or ""
                        price_str = match.group(4).replace(',', '')
                        item.total_price = float(price_str)
                    else:
                        price_str = match.group(3).replace(',', '')
                        item.total_price = float(price_str)
                    
                    if item.quantity > 0:
                        item.unit_price = item.total_price / item.quantity
                    
                    document.items.append(item)
                
                if document.items:
                    break
                    
        except Exception as e:
            self.logger.error(f"代替品目抽出エラー: {str(e)}")
    
    def _parse_number(self, text: str, number_type: type, default: Any) -> Any:
        """数値文字列を解析"""
        try:
            # カンマ、全角数字、空白を処理
            cleaned = re.sub(r'[,，\s]', '', text)
            cleaned = cleaned.translate(str.maketrans('０１２３４５６７８９', '0123456789'))
            
            if number_type == int:
                return int(float(cleaned))
            else:
                return number_type(cleaned)
                
        except (ValueError, TypeError):
            return default
    
    def _calculate_totals(self, document: DeliveryDocument):
        """合計金額を計算"""
        try:
            total = sum(item.total_price for item in document.items)
            document.total_amount = total
            
        except Exception as e:
            self.logger.error(f"合計計算エラー: {str(e)}")
    
    def extract_multiple_files(self, pdf_directory: Path) -> List[DeliveryDocument]:
        """複数のPDFファイルからデータを抽出"""
        try:
            self.logger.info(f"複数PDF抽出開始: {pdf_directory}")
            
            pdf_files = list(pdf_directory.glob("*.pdf"))
            if not pdf_files:
                self.logger.warning(f"PDFファイルが見つかりません: {pdf_directory}")
                return []
            
            all_documents = []
            for pdf_file in pdf_files:
                documents = self.extract_delivery_data(pdf_file)
                all_documents.extend(documents)
            
            self.logger.info(f"複数PDF抽出完了: {len(all_documents)} 件の納品書類")
            return all_documents
            
        except Exception as e:
            self.logger.error(f"複数PDF抽出エラー: {str(e)}")
            return [] 