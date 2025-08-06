#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel 処理モジュール
"""

import pandas as pd
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Tuple, Optional, Any
from dataclasses import dataclass

try:
    import openpyxl
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
except ImportError:
    openpyxl = None

from utils.logger import Logger
from data.pdf_extractor import DeliveryDocument, DeliveryItem


@dataclass
class MasterItem:
    """マスタアイテムのデータクラス"""
    item_code: str = ""
    item_name: str = ""
    supplier: str = ""
    category: str = ""
    unit_price: float = 0.0
    delivery_destinations: List[str] = None
    notes: str = ""
    
    def __post_init__(self):
        if self.delivery_destinations is None:
            self.delivery_destinations = []


@dataclass
class ValidationError:
    """バリデーションエラーのデータクラス"""
    error_type: str = ""
    item_name: str = ""
    expected_value: str = ""
    actual_value: str = ""
    description: str = ""
    document_id: str = ""


class ExcelProcessor:
    """Excel 処理クラス"""
    
    def __init__(self):
        self.logger = Logger(__name__)
        self.master_data: Dict[str, MasterItem] = {}
        
        if not openpyxl:
            raise ImportError("openpyxl がインストールされていません。pip install openpyxl を実行してください。")
    
    def load_master_data(self, master_file_path: Path) -> bool:
        """マスタExcelファイルの「登録商品マスター」シートを読み込み"""
        try:
            self.logger.info(f"マスタファイル読み込み開始: {master_file_path}")
            
            if not master_file_path.exists():
                self.logger.error(f"マスタファイルが見つかりません: {master_file_path}")
                return False
            
            # 「登録商品マスター」シートを読み込み
            df = pd.read_excel(master_file_path, sheet_name='登録商品マスター')
            
            # カラム名を正規化
            df.columns = [str(col).strip() for col in df.columns]
            self.logger.info(f"マスタシートのカラム: {list(df.columns)}")
            
            # A列（角上魚類商品コード）をキーとして辞書に変換
            for _, row in df.iterrows():
                item = MasterItem()
                
                # A列の商品コード（数値型）
                item.item_code = str(int(row.iloc[0])) if pd.notna(row.iloc[0]) else ""
                
                # 他の列の情報も取得
                item.item_name = self._get_value_from_row(row, ['商品名', '品名'])
                item.supplier = "角上魚類"  # デフォルト
                
                # その他の情報
                if '担当者' in df.columns:
                    item.notes = self._get_value_from_row(row, ['担当者'])
                
                # 商品コードをキーとして保存（A列の値）
                if item.item_code:
                    self.master_data[item.item_code] = item
                    self.logger.debug(f"マスタ登録: コード={item.item_code}, 商品名={item.item_name}")
            
            self.logger.info(f"マスタデータ読み込み完了: {len(self.master_data)} 件")
            return True
            
        except Exception as e:
            self.logger.error(f"マスタファイル読み込みエラー: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
    
    def _get_value_from_row(self, row: pd.Series, possible_columns: List[str]) -> str:
        """行から指定された可能なカラム名のいずれかの値を取得"""
        for col in possible_columns:
            if col in row.index:
                value = str(row[col]).strip()
                if value and value != 'nan':
                    return value
        return ""
    
    def _get_numeric_value_from_row(self, row: pd.Series, possible_columns: List[str]) -> float:
        """行から数値データを取得"""
        for col in possible_columns:
            if col in row.index:
                try:
                    value = pd.to_numeric(row[col], errors='coerce')
                    if pd.notna(value):
                        return float(value)
                except:
                    continue
        return 0.0
    
    def validate_with_master(self, documents: List[DeliveryDocument], master_file_path: Path) -> Tuple[List[DeliveryDocument], List[ValidationError]]:
        """納品データをマスタと付け合わせて検証"""
        try:
            self.logger.info("マスタ付け合わせ処理開始")
            
            # マスタデータ読み込み
            if not self.master_data:
                if not self.load_master_data(master_file_path):
                    return documents, []
            
            validated_documents = []
            all_errors = []
            
            for document in documents:
                validated_items = []
                document_errors = []
                
                for item in document.items:
                    # マスタデータと照合
                    validation_result = self._validate_item_with_master(item, document.document_id)
                    
                    if validation_result["is_valid"]:
                        # マスタ情報で補完
                        enhanced_item = self._enhance_item_with_master(item, validation_result["master_item"])
                        validated_items.append(enhanced_item)
                    else:
                        # エラー情報を記録
                        document_errors.extend(validation_result["errors"])
                        # エラーがあってもアイテムは含める（エラー情報付き）
                        item.notes = f"エラー: {', '.join([err.description for err in validation_result['errors']])}"
                        validated_items.append(item)
                
                # 検証済み文書を作成
                validated_document = document
                validated_document.items = validated_items
                validated_documents.append(validated_document)
                all_errors.extend(document_errors)
            
            self.logger.info(f"マスタ付け合わせ完了: 正常 {len(validated_documents)} 件, エラー {len(all_errors)} 件")
            return validated_documents, all_errors
            
        except Exception as e:
            self.logger.error(f"マスタ付け合わせエラー: {str(e)}")
            return documents, []
    
    def _validate_item_with_master(self, item: DeliveryItem, document_id: str) -> Dict[str, Any]:
        """個別アイテムのマスタ検証（商品コードベース）"""
        result = {
            "is_valid": True,
            "errors": [],
            "master_item": None
        }
        
        # 商品コードでマスタデータから検索
        master_item = self._find_master_item_by_code(item.item_code)
        
        if not master_item:
            # 商品コードが見つからない
            error = ValidationError(
                error_type="商品未登録",
                item_name=item.item_name,
                expected_value="マスタに存在する商品コード",
                actual_value=item.item_code,
                description=f"商品コード「{item.item_code}」はマスタに登録されていません",
                document_id=document_id
            )
            result["errors"].append(error)
            result["is_valid"] = False
            return result
        
        result["master_item"] = master_item
        self.logger.info(f"マスタ突合成功: コード={item.item_code}, 商品名={master_item.item_name}")
        
        # 単価チェック（もしPDFに単価情報があれば）
        if master_item.unit_price > 0 and item.unit_price > 0:
            price_diff_ratio = abs(item.unit_price - master_item.unit_price) / master_item.unit_price
            if price_diff_ratio > 0.1:  # 10%以上の差異
                error = ValidationError(
                    error_type="単価差異",
                    item_name=master_item.item_name,
                    expected_value=str(master_item.unit_price),
                    actual_value=str(item.unit_price),
                    description=f"単価が標準価格と {price_diff_ratio:.1%} 差異があります",
                    document_id=document_id
                )
                result["errors"].append(error)
        
        return result
    
    def _find_master_item_by_code(self, item_code: str) -> Optional[MasterItem]:
        """商品コードでマスタデータから商品を検索"""
        if item_code in self.master_data:
            return self.master_data[item_code]
        return None
    
    def _find_master_item(self, item_name: str) -> Optional[MasterItem]:
        """マスタデータから商品を検索（商品名ベース・従来互換）"""
        # 商品名での検索（従来の機能を維持）
        for master_item in self.master_data.values():
            if master_item.item_name == item_name:
                return master_item
        
        # 部分一致で検索
        for master_item in self.master_data.values():
            if item_name in master_item.item_name or master_item.item_name in item_name:
                return master_item
        
        return None
    
    def _enhance_item_with_master(self, item: DeliveryItem, master_item: MasterItem) -> DeliveryItem:
        """マスタ情報でアイテムを補強"""
        enhanced_item = item
        
        # マスタから正しい商品名を設定
        if master_item.item_name:
            enhanced_item.item_name = master_item.item_name
            self.logger.info(f"商品名補完: {item.item_code} -> {master_item.item_name}")
        
        # 不足情報を補完
        if not enhanced_item.item_code and master_item.item_code:
            enhanced_item.item_code = master_item.item_code
        
        if not enhanced_item.supplier and master_item.supplier:
            enhanced_item.supplier = master_item.supplier
        
        return enhanced_item
    
    def write_errors_to_master_sheet(self, errors: List[ValidationError], master_file_path: Path) -> bool:
        """エラーデータをマスタExcelの「エラーリスト」シートに書き込み"""
        try:
            self.logger.info("エラーリストシートへの書き込み開始")
            
            if not errors:
                self.logger.info("エラーデータがありません")
                return True
            
            if not master_file_path.exists():
                self.logger.error(f"マスタファイルが見つかりません: {master_file_path}")
                return False
            
            # Excelファイルを読み込み
            wb = load_workbook(master_file_path)
            
            # エラーリストシートを取得または作成
            if "エラーリスト" in wb.sheetnames:
                ws = wb["エラーリスト"]
                # 既存データをクリア（ヘッダー以外）
                if ws.max_row > 1:
                    ws.delete_rows(2, ws.max_row)
            else:
                ws = wb.create_sheet("エラーリスト")
            
            # ヘッダー作成
            headers = ["発生日時", "エラー種別", "商品コード", "商品名", "説明", "文書ID"]
            for col, header in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col, value=header)
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
            
            # エラーデータを書き込み
            for row, error in enumerate(errors, 2):
                ws.cell(row=row, column=1, value=datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
                ws.cell(row=row, column=2, value=error.error_type)
                ws.cell(row=row, column=3, value=error.actual_value)  # 商品コード
                ws.cell(row=row, column=4, value=error.item_name)
                ws.cell(row=row, column=5, value=error.description)
                ws.cell(row=row, column=6, value=error.document_id)
            
            # 列幅調整
            for column in ws.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                ws.column_dimensions[column_letter].width = adjusted_width
            
            # ファイル保存
            wb.save(master_file_path)
            wb.close()
            
            self.logger.info(f"エラーリストシート更新完了: {len(errors)}件のエラーを記載")
            return True
            
        except Exception as e:
            self.logger.error(f"エラーリストシート書き込みエラー: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
    
    def create_error_excel(self, errors: List[ValidationError]) -> Optional[Path]:
        """エラーデータをExcelファイルに出力（バックアップ用）"""
        try:
            self.logger.info("エラーExcel作成開始")
            
            if not errors:
                self.logger.info("エラーデータがありません")
                return None
            
            # 出力ディレクトリを取得
            from config.settings import Config
            config = Config()
            
            # ファイル名生成
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            error_file = config.output_dir / f"エラーデータ_{timestamp}.xlsx"
            
            # ワークブック作成
            wb = Workbook()
            ws = wb.active
            ws.title = "エラーデータ"
            
            # ヘッダー作成
            headers = ["エラー種別", "商品名", "期待値", "実際値", "説明", "文書ID", "発生日時"]
            for col, header in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col, value=header)
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")
            
            # データ行作成
            for row, error in enumerate(errors, 2):
                ws.cell(row=row, column=1, value=error.error_type)
                ws.cell(row=row, column=2, value=error.item_name)
                ws.cell(row=row, column=3, value=error.expected_value)
                ws.cell(row=row, column=4, value=error.actual_value)
                ws.cell(row=row, column=5, value=error.description)
                ws.cell(row=row, column=6, value=error.document_id)
                ws.cell(row=row, column=7, value=datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            
            # 列幅調整
            for column in ws.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                ws.column_dimensions[column_letter].width = adjusted_width
            
            # ファイル保存
            wb.save(error_file)
            
            self.logger.info(f"エラーExcel作成完了: {error_file}")
            return error_file
            
        except Exception as e:
            self.logger.error(f"エラーExcel作成エラー: {str(e)}")
            return None
    
    def group_by_destination(self, documents: List[DeliveryDocument]) -> Dict[str, List[DeliveryDocument]]:
        """納品先ごとにデータをグループ化"""
        try:
            self.logger.info("納品先別グループ化開始")
            
            grouped = {}
            
            for document in documents:
                destination = document.delivery_destination or "不明"
                
                if destination not in grouped:
                    grouped[destination] = []
                
                grouped[destination].append(document)
            
            self.logger.info(f"グループ化完了: {len(grouped)} 納品先")
            return grouped
            
        except Exception as e:
            self.logger.error(f"グループ化エラー: {str(e)}")
            return {}
    
    def create_dispatch_table(self, destination: str, documents: List[DeliveryDocument]) -> Optional[Path]:
        """配車表を作成"""
        try:
            self.logger.info(f"配車表作成開始: {destination}")
            
            from config.settings import Config
            config = Config()
            
            # ファイル名生成
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            safe_destination = self._sanitize_filename(destination)
            dispatch_file = config.output_dir / f"配車表_{safe_destination}_{timestamp}.xlsx"
            
            # ワークブック作成
            wb = Workbook()
            ws = wb.active
            ws.title = "配車表"
            
            # タイトル
            ws.merge_cells('A1:G1')
            title_cell = ws['A1']
            title_cell.value = f"配車表 - {destination}"
            title_cell.font = Font(size=16, bold=True)
            title_cell.alignment = Alignment(horizontal='center')
            
            # ヘッダー
            headers = ["配送日", "商品名", "数量", "単位", "配送先", "備考", "担当者"]
            for col, header in enumerate(headers, 1):
                cell = ws.cell(row=3, column=col, value=header)
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="E6F3FF", end_color="E6F3FF", fill_type="solid")
            
            # データ行
            row = 4
            for document in documents:
                for item in document.items:
                    ws.cell(row=row, column=1, value=document.document_date)
                    ws.cell(row=row, column=2, value=item.item_name)
                    ws.cell(row=row, column=3, value=item.quantity)
                    ws.cell(row=row, column=4, value=item.unit)
                    ws.cell(row=row, column=5, value=destination)
                    ws.cell(row=row, column=6, value=item.notes)
                    ws.cell(row=row, column=7, value="")  # 担当者は空欄
                    row += 1
            
            # 罫線とフォーマット
            self._apply_table_formatting(ws, 3, row - 1, len(headers))
            
            # 列幅調整
            self._adjust_column_widths(ws)
            
            # ファイル保存
            wb.save(dispatch_file)
            
            self.logger.info(f"配車表作成完了: {dispatch_file}")
            return dispatch_file
            
        except Exception as e:
            self.logger.error(f"配車表作成エラー: {str(e)}")
            return None
    
    def create_shipping_request(self, destination: str, documents: List[DeliveryDocument]) -> Optional[Path]:
        """出庫依頼書を作成"""
        try:
            self.logger.info(f"出庫依頼書作成開始: {destination}")
            
            from config.settings import Config
            config = Config()
            
            # ファイル名生成
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            safe_destination = self._sanitize_filename(destination)
            shipping_file = config.output_dir / f"出庫依頼_{safe_destination}_{timestamp}.xlsx"
            
            # ワークブック作成
            wb = Workbook()
            ws = wb.active
            ws.title = "出庫依頼"
            
            # タイトル
            ws.merge_cells('A1:F1')
            title_cell = ws['A1']
            title_cell.value = f"出庫依頼書 - {destination}"
            title_cell.font = Font(size=16, bold=True)
            title_cell.alignment = Alignment(horizontal='center')
            
            # 基本情報
            ws.cell(row=2, column=1, value="依頼日:")
            ws.cell(row=2, column=2, value=datetime.now().strftime('%Y-%m-%d'))
            ws.cell(row=2, column=4, value="納品先:")
            ws.cell(row=2, column=5, value=destination)
            
            # ヘッダー
            headers = ["商品コード", "商品名", "数量", "単位", "単価", "金額"]
            for col, header in enumerate(headers, 1):
                cell = ws.cell(row=4, column=col, value=header)
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="FFE6CC", end_color="FFE6CC", fill_type="solid")
            
            # データ行と合計計算
            row = 5
            total_amount = 0
            
            for document in documents:
                for item in document.items:
                    ws.cell(row=row, column=1, value=item.item_code)
                    ws.cell(row=row, column=2, value=item.item_name)
                    ws.cell(row=row, column=3, value=item.quantity)
                    ws.cell(row=row, column=4, value=item.unit)
                    ws.cell(row=row, column=5, value=item.unit_price)
                    ws.cell(row=row, column=6, value=item.total_price)
                    total_amount += item.total_price
                    row += 1
            
            # 合計行
            ws.cell(row=row, column=5, value="合計:")
            ws.cell(row=row, column=5).font = Font(bold=True)
            ws.cell(row=row, column=6, value=total_amount)
            ws.cell(row=row, column=6).font = Font(bold=True)
            
            # 罫線とフォーマット
            self._apply_table_formatting(ws, 4, row, len(headers))
            
            # 列幅調整
            self._adjust_column_widths(ws)
            
            # ファイル保存
            wb.save(shipping_file)
            
            self.logger.info(f"出庫依頼書作成完了: {shipping_file}")
            return shipping_file
            
        except Exception as e:
            self.logger.error(f"出庫依頼書作成エラー: {str(e)}")
            return None
    
    def _sanitize_filename(self, filename: str) -> str:
        """ファイル名に使用できない文字を除去"""
        import re
        return re.sub(r'[<>:"/\\|?*]', '_', filename)
    
    def _apply_table_formatting(self, ws, start_row: int, end_row: int, num_cols: int):
        """テーブルに罫線とフォーマットを適用"""
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        for row in range(start_row, end_row + 1):
            for col in range(1, num_cols + 1):
                ws.cell(row=row, column=col).border = thin_border
    
    def _adjust_column_widths(self, ws):
        """列幅を自動調整"""
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column_letter].width = adjusted_width 