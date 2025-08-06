#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SMCL 納品リスト処理システム メインコントローラー

システムフロー:
1. スマクラにログイン
2. 納品リストをダウンロード  
3. 納品データを抽出
4. 納品リストをマスタExcelと付け合わせ
5. 納品先に合わせて、配車表・出庫依頼のExcelを作成
6. web上から遅れるFax送信(まだなくてもいい)
7. LineBotで通知
"""

import sys
import os
from datetime import datetime
from pathlib import Path

# プロジェクトルートをPythonパスに追加
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from config.settings import Config
from utils.logger import Logger
from scraping.smcl_scraper import SMCLScraper
from data.pdf_extractor import PDFExtractor
from data.csv_extractor import CSVExtractor
from data.excel_processor import ExcelProcessor
from notification.line_bot import LineBotNotifier


class DeliveryListProcessor:
    """納品リスト処理システムのメインクラス"""
    
    def __init__(self):
        self.config = Config()
        self.logger = Logger(__name__)
        self.scraper = None
        self.pdf_extractor = PDFExtractor()
        self.csv_extractor = CSVExtractor()
        self.excel_processor = ExcelProcessor()
        self.line_notifier = LineBotNotifier()
        
    def run(self):
        """メイン処理を実行"""
        try:
            self.logger.info("=== 納品リスト処理システム開始 ===")
            start_time = datetime.now()
            
            # フェーズ1: スマクラログインと納品リストダウンロード
            # if not self._phase1_scraping():
            #     self.logger.error("フェーズ1: スクレイピング処理が失敗しました")
            #     return False
                
            # フェーズ2: CSVデータ抽出
            extracted_data = self._phase2_data_extraction()
            if not extracted_data:
                self.logger.error("フェーズ2: データ抽出が失敗しました")
                return False
            print(extracted_data)
                
            # # フェーズ3: マスタExcelとの付け合わせ
            validated_data, error_data = self._phase3_master_validation(extracted_data)
            
            # # フェーズ4: 配車表・出庫依頼Excel作成
            # if not self._phase4_excel_generation(validated_data):
            #     self.logger.error("フェーズ4: Excel生成が失敗しました")
            #     return False
                
            # # フェーズ5: LineBot通知
            # self._phase5_notification(validated_data, error_data)
            
            # 処理完了
            end_time = datetime.now()
            duration = end_time - start_time
            self.logger.info(f"=== 処理完了 (処理時間: {duration.total_seconds():.1f}秒) ===")
            
            return True
            
        except Exception as e:
            self.logger.error(f"メイン処理でエラーが発生しました: {str(e)}")
            self.logger.exception(e)
            return False
            
        finally:
            self._cleanup()
    
    def _phase1_scraping(self):
        """フェーズ1: スマクラログインと納品リストダウンロード"""
        try:
            self.logger.info("フェーズ1: スクレイピング処理開始")
            
            self.scraper = SMCLScraper(
                download_dir=self.config.download_dir,
                headless=self.config.headless_mode
            )
            
            # スマクラにログインして納品リストをダウンロード
            return self.scraper.download_delivery_lists()
            
        except Exception as e:
            self.logger.error(f"フェーズ1でエラー: {str(e)}")
            return False
    
    def _phase2_data_extraction(self):
        """フェーズ2: CSVデータ抽出"""
        try:
            self.logger.info("フェーズ2: CSVデータ抽出処理開始")
            
            # 今日の日付のCSVファイルを取得
            csv_file = self.csv_extractor.find_today_csv_file(Path(self.config.download_dir))
            
            if not csv_file:
                # 今日のファイルがない場合は最新のファイルを取得
                self.logger.warning("今日の日付のCSVファイルが見つかりません。最新のファイルを使用します。")
                csv_file = self.csv_extractor.find_latest_csv_file(Path(self.config.download_dir))
            
            if not csv_file:
                self.logger.error("CSVファイルが見つかりませんでした")
                return None
            
            # CSVからデータを抽出
            documents = self.csv_extractor.extract_order_data(csv_file)
            
            # アイテム数をカウント
            total_items = sum(len(doc.items) for doc in documents)
            
            self.logger.info(f"合計 {total_items} 件のデータを抽出しました")
            return documents
            
        except Exception as e:
            self.logger.error(f"フェーズ2でエラー: {str(e)}")
            return None
    
    def _phase3_master_validation(self, extracted_data):
        """フェーズ3: マスタExcelとの付け合わせ"""
        try:
            self.logger.info("フェーズ3: マスタ付け合わせ処理開始")
            
            validated_data, error_data = self.excel_processor.validate_with_master(
                extracted_data, self.config.master_excel_path
            )
            
            # エラーデータがある場合の処理
            if error_data:
                # エラーリストシートに書き込み
                success = self.excel_processor.write_errors_to_master_sheet(
                    error_data, self.config.master_excel_path
                )
                
                if success:
                    self.logger.warning(f"エラーデータを「エラーリスト」シートに記載しました: {len(error_data)}件")
                else:
                    # エラーリストシートへの書き込みが失敗した場合、バックアップファイルを作成
                    error_file = self.excel_processor.create_error_excel(error_data)
                    self.logger.warning(f"エラーデータをバックアップファイルに出力しました: {error_file}")
            else:
                self.logger.info("エラーはありませんでした")
                
            return validated_data, error_data
            
        except Exception as e:
            self.logger.error(f"フェーズ3でエラー: {str(e)}")
            return [], []
    
    def _phase4_excel_generation(self, validated_data):
        """フェーズ4: 配車表・出庫依頼Excel作成"""
        try:
            self.logger.info("フェーズ4: Excel生成処理開始")
            
            # 納品先ごとにグループ化
            grouped_data = self.excel_processor.group_by_destination(validated_data)
            
            generated_files = []
            for destination, data in grouped_data.items():
                # 配車表作成
                dispatch_file = self.excel_processor.create_dispatch_table(destination, data)
                if dispatch_file:
                    generated_files.append(dispatch_file)
                    
                # 出庫依頼作成
                shipping_file = self.excel_processor.create_shipping_request(destination, data)
                if shipping_file:
                    generated_files.append(shipping_file)
            
            self.logger.info(f"{len(generated_files)} 個のExcelファイルを生成しました")
            return True
            
        except Exception as e:
            self.logger.error(f"フェーズ4でエラー: {str(e)}")
            return False
    
    def _phase5_notification(self, validated_data, error_data):
        """フェーズ5: LineBot通知"""
        try:
            self.logger.info("フェーズ5: 通知処理開始")
            
            # 処理結果のサマリーを作成
            summary = {
                "処理日時": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                "正常データ件数": len(validated_data),
                "エラーデータ件数": len(error_data),
                "生成ファイル数": len(self._get_generated_excel_files())
            }
            
            # LINE通知を送信
            self.line_notifier.send_process_summary(summary)
            
            # エラーがある場合は詳細通知
            if error_data:
                self.line_notifier.send_error_details(error_data)
                
        except Exception as e:
            self.logger.error(f"フェーズ5でエラー: {str(e)}")
    
    def _get_downloaded_pdf_files(self):
        """ダウンロードされたPDFファイルのリストを取得"""
        download_dir = Path(self.config.download_dir)
        return list(download_dir.glob("*.pdf"))
    
    def _get_generated_excel_files(self):
        """生成されたExcelファイルのリストを取得"""
        output_dir = Path(self.config.output_dir)
        return list(output_dir.glob("*.xlsx"))
    
    def _cleanup(self):
        """リソースのクリーンアップ"""
        if self.scraper:
            self.scraper.cleanup()


def main():
    """メイン関数"""
    processor = DeliveryListProcessor()
    success = processor.run()
    
    if success:
        print("✅ 処理が正常に完了しました")
        sys.exit(0)
    else:
        print("❌ 処理が失敗しました")
        sys.exit(1)


if __name__ == "__main__":
    main() 