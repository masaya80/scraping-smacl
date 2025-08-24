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

from services.core.config import Config
from services.core.logger import Logger
from services.scraping.smcl_scraper import SMCLScraper
from services.data_processing.csv_extractor import CSVExtractor
from services.data_processing.excel_processor import ExcelProcessor
from services.data_processing.pdf_image_processor import PDFImageProcessor
from services.notification.line_bot import LineBotNotifier


class DeliveryListProcessor:
    """納品リスト処理システムのメインクラス"""
    
    def __init__(self):
        self.config = Config()
        self.logger = Logger(__name__)
        self.scraper = None
        self.csv_extractor = CSVExtractor()
        self.excel_processor = ExcelProcessor(self.config)
        self.pdf_image_processor = PDFImageProcessor(self.config)
        self.line_notifier = LineBotNotifier()
        
        # 日付別ディレクトリを設定
        self.today_str = datetime.now().strftime(self.config.date_folder_format)
        self.dated_output_dir = self.config.get_dated_output_dir(self.today_str)
        self.dated_download_dir = self.config.get_dated_download_dir(self.today_str)
        self.dated_logs_dir = self.config.get_dated_logs_dir(self.today_str)
        
    def run(self):
        """メイン処理を実行"""
        try:
            self.logger.info("=== 納品リスト処理システム開始 ===")
            self.logger.info(f"処理日: {self.today_str}")
            self.logger.info(f"出力先: {self.dated_output_dir}")
            
            # ネットワークドライブの状態確認
            network_status = self.config.get_network_status()
            if network_status["use_network_storage"]:
                if network_status["accessible"]:
                    self.logger.info(f"ネットワークドライブ接続OK: {network_status['network_path']}")
                else:
                    self.logger.warning(f"ネットワークドライブ接続NG: {network_status.get('error', '不明なエラー')}")
                    self.logger.info("ローカルフォルダを使用します")
            
            start_time = datetime.now()
            
            # # フェーズ1: スマクラログインと納品リストダウンロード
            if not self._phase1_scraping():
                 self.logger.error("フェーズ1: スクレイピング処理が失敗しました")
                 return False
            
            # データなしの場合の専用処理
            if hasattr(self.scraper, 'no_data_found') and self.scraper.no_data_found:
                self.logger.info("📭 該当するデータがありません - 専用通知を送信します")
                self._send_no_data_notification()
                return True  # 処理は成功（データがないだけ）
                
            # フェーズ2: CSVデータ抽出
            extracted_data = self._phase2_data_extraction()
            if not extracted_data:
                self.logger.error("フェーズ2: データ抽出が失敗しました")
                return False
            
            # # フェーズ3: マスタExcelとの付け合わせ
            validated_data, error_data = self._phase3_master_validation(extracted_data)
            
            # # フェーズ4: 配車表・出庫依頼Excel作成
            if not self._phase4_excel_generation(validated_data):
                self.logger.error("フェーズ4: Excel生成が失敗しました")
                return False
                
            # # フェーズ5: PDF画像変換
            converted_images = self._phase5_pdf_image_conversion()
            if not converted_images:
                self.logger.warning("フェーズ4.5: PDF画像変換で画像が生成されませんでした")
                
            # # フェーズ6: LineBot通知
            self._phase6_notification(validated_data, error_data, converted_images)
            
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
                download_dir=self.dated_download_dir,
                headless=self.config.headless_mode,
                config=self.config
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
            csv_file = self.csv_extractor.find_today_csv_file(self.dated_download_dir)
        
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
        """フェーズ4: 倉庫別注文処理 - 既存Excelシートに数量を挿入"""
        try:
            self.logger.info("フェーズ4: 倉庫別注文処理開始")
            
            # 要求された仕様に従って既存マスタExcelに数量を挿入
            success = self.excel_processor.process_warehouse_orders(
                validated_data, self.config.master_excel_path
            )
            
            if success:
                self.logger.info("倉庫別注文処理が正常に完了しました")
            else:
                self.logger.error("倉庫別注文処理が失敗しました")
            
            return success
            
        except Exception as e:
            self.logger.error(f"フェーズ4でエラー: {str(e)}")
            return False
    
    def _phase5_pdf_image_conversion(self):
        """フェーズ5: PDF画像変換"""
        try:
            self.logger.info("フェーズ5: PDF画像変換処理開始")
            
            # 今日のPDFファイルを取得
            generated_pdf_files = self._get_generated_pdf_files()
            downloaded_pdf_files = self._get_downloaded_pdf_files()
            all_pdf_files = downloaded_pdf_files + generated_pdf_files
            
            if not all_pdf_files:
                self.logger.warning("変換対象のPDFファイルがありません")
                return {}
            
            self.logger.info(f"PDF画像変換対象: {len(all_pdf_files)}ファイル")
            for pdf_file in all_pdf_files:
                self.logger.info(f"  - {pdf_file.name}")
            
            # 画像出力ディレクトリを設定
            output_dir = self.dated_output_dir
            
            # PDF画像変換を実行
            converted_images = self.pdf_image_processor.process_all_pdfs(all_pdf_files, output_dir)
            
            # 変換結果のサマリーを取得
            summary = self.pdf_image_processor.get_image_summary(converted_images)
            
            self.logger.info("フェーズ5: PDF画像変換処理完了")
            self.logger.info(f"  変換結果: {summary.get('総PDFファイル数', 0)}ファイル -> {summary.get('総画像数', 0)}枚")
            self.logger.info(f"  成功: {summary.get('成功PDFファイル数', 0)}ファイル, 失敗: {summary.get('失敗PDFファイル数', 0)}ファイル")
            
            # 古い画像ファイルをクリーンアップ
            deleted_count = self.pdf_image_processor.cleanup_old_images(output_dir, days_to_keep=7)
            if deleted_count > 0:
                self.logger.info(f"古い画像ファイルを削除: {deleted_count}個")
            
            return converted_images
            
        except Exception as e:
            self.logger.error(f"フェーズ4.5でエラー: {str(e)}")
            return {}
    
    def _phase6_notification(self, validated_data, error_data, converted_images):
        """フェーズ6: LineBot通知（統合版）"""
        try:
            self.logger.info("フェーズ6: 統合通知処理開始")
            
            # 生成されたファイルを取得
            excel_files = self._get_generated_excel_files()
            generated_pdf_files = self._get_generated_pdf_files()
            downloaded_pdf_files = self._get_downloaded_pdf_files()
            
            # 画像変換結果のサマリーを取得
            image_summary = self.pdf_image_processor.get_image_summary(converted_images)
            
            # 処理結果のサマリーを作成
            summary = {
                "処理日時": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                "正常データ件数": len(validated_data),
                "エラーデータ件数": len(error_data),
                "生成Excelファイル数": len(excel_files),
                "生成PDFファイル数": len(generated_pdf_files),
                "ダウンロードPDFファイル数": len(downloaded_pdf_files),
                "総PDFファイル数": len(generated_pdf_files) + len(downloaded_pdf_files),
                "総画像数": image_summary.get("総画像数", 0),
                "画像変換成功数": image_summary.get("成功PDFファイル数", 0),
                "画像変換失敗数": image_summary.get("失敗PDFファイル数", 0)
            }
            
            # 統合されたビジネス向け通知を送信
            self.line_notifier.send_integrated_completion_notification(
                summary=summary,
                error_data=error_data,
                converted_images=converted_images,
                max_images=5,  # 重要な画像（出庫依頼書・配車表）を最大5枚まで送信
                send_all_delivery_lists=True  # 全ての納品リスト画像を送信
            )
            
            # 重要なエラーがある場合のみ詳細通知を追加送信
            if error_data and len(error_data) > 5:
                self.logger.info("多数のエラーが発生したため詳細通知を送信")
                self.line_notifier.send_error_details(error_data)
                
        except Exception as e:
            self.logger.error(f"フェーズ6でエラー: {str(e)}")
    
    def _get_downloaded_pdf_files(self):
        """今日ダウンロードされたPDFファイルのリストを取得"""
        # 日付別ダウンロードディレクトリから取得
        today_files = []
        for pdf_file in self.dated_download_dir.glob("*.pdf"):
            today_files.append(pdf_file)
        
        # ファイル名で並び替え（最新順）
        today_files.sort(key=lambda x: x.name, reverse=True)
        
        self.logger.info(f"今日ダウンロードされたPDFファイル: {len(today_files)}個")
        return today_files
    
    def _get_generated_excel_files(self):
        """今日生成されたExcelファイルのリストを取得"""
        # 日付別出力ディレクトリから取得
        today_files = []
        for excel_file in self.dated_output_dir.glob("*.xlsx"):
            today_files.append(excel_file)
        
        # ファイル名で並び替え（最新順）
        today_files.sort(key=lambda x: x.name, reverse=True)
        
        self.logger.info(f"今日生成されたExcelファイル: {len(today_files)}個")
        return today_files
    
    def _get_generated_pdf_files(self):
        """今日生成されたPDFファイルのリストを取得（最新の出庫依頼書と配車表のみ）"""
        # 日付別出力ディレクトリから取得
        today_files = []
        for pdf_file in self.dated_output_dir.glob("*.pdf"):
            today_files.append(pdf_file)
        
        # ファイル名で並び替え（最新順）
        today_files.sort(key=lambda x: x.name, reverse=True)
        
        # 出庫依頼書と配車表の最新ファイルのみを取得
        filtered_files = []
        found_types = set()
        
        for pdf_file in today_files:
            file_type = None
            if "出庫依頼" in pdf_file.name:
                file_type = "出庫依頼"
            elif any(keyword in pdf_file.name for keyword in ["アリスト", "配車", "LT"]):
                file_type = "配車表"
            
            if file_type and file_type not in found_types:
                filtered_files.append(pdf_file)
                found_types.add(file_type)
                
            # 2種類とも見つかったら終了
            if len(found_types) >= 2:
                break
        
        self.logger.info(f"今日生成されたPDFファイル（フィルタ後）: {len(filtered_files)}個")
        for pdf_file in filtered_files:
            self.logger.info(f"  - {pdf_file.name}")
        
        return filtered_files
    
    def _send_no_data_notification(self):
        """データなし時の専用LINE通知"""
        try:
            self.logger.info("📭 データなし通知処理開始")
            
            from datetime import datetime
            from services.notification.line_bot import LineBotNotifier
            
            # LineBotNotifier初期化
            line_bot = LineBotNotifier()
            
            # 今日の日付取得
            today_str = datetime.today().strftime('%Y年%m月%d日')
            
            # 専用メッセージ作成
            message = f"📭 受注データ確認結果\n\n{today_str}の受注データが見つかりませんでした。"
            
            # LINE通知送信
            self.logger.info(f"LINE通知送信: {message}")
            line_bot.send_message(message)
            
            self.logger.info("✅ データなし通知を送信しました")
            
        except Exception as e:
            self.logger.error(f"データなし通知送信エラー: {str(e)}")
            self.logger.exception(e)
    
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