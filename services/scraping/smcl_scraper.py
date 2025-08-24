#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SMCL スクレイピングモジュール
"""

import time
import platform
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime

from ..core.logger import Logger


class SMCLScraper:
    """SMCL Webサイトのスクレイピングクラス"""
    
    def __init__(self, download_dir: Path, headless: bool = True, config=None):
        self.download_dir = Path(download_dir)
        self.headless = headless
        self.config = config
        self.driver = None
        self.logger = Logger(__name__)
        
        # ログイン情報（設定から取得することも可能）
        self.target_url = "https://smclweb.cs-cxchange.net/smcl/view/lin/EDS001OLIN0000.aspx"
        self.corp_code = "I26S"
        self.login_id = "0600200"
        self.password = "toichi04"
        
        # 確定処理制御フラグ（configから取得）
        if config and hasattr(config, 'enable_confirmation_process'):
            self.enable_confirmation_process = config.enable_confirmation_process
        else:
            self.enable_confirmation_process = True  # デフォルトは有効
        
        # データなし状態の管理
        self.no_data_found = False
        
        if not self.enable_confirmation_process:
            self.logger.info("⚠️  確定処理は無効になっています（テストモード）")
        
        # ダウンロードディレクトリを作成
        self.download_dir.mkdir(exist_ok=True)
    
    def setup_driver(self):
        """Chromeドライバーを設定"""
        try:
            chrome_options = Options()
            
            # 基本オプション
            chrome_options.add_argument('--no-sandbox')
            chrome_options.add_argument('--disable-dev-shm-usage')
            chrome_options.add_argument('--disable-gpu')
            chrome_options.add_argument('--window-size=1920,1080')
            
            if self.headless:
                chrome_options.add_argument('--headless')
            
            # Windows用の追加設定
            if platform.system() == 'Windows':
                chrome_options.add_argument('--disable-software-rasterizer')
            
            # ダウンロード設定
            download_dir_str = str(self.download_dir.absolute())
            prefs = {
                "download.default_directory": download_dir_str,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True
            }
            chrome_options.add_experimental_option("prefs", prefs)
            
            # ChromeDriverManagerを使用
            self.logger.info(f"システム: {platform.system()} ({platform.machine()})")
            driver_path = ChromeDriverManager().install()
            
            if platform.system() == 'Windows':
                driver_path = driver_path.replace('/', '\\')
            
            service = Service(driver_path)
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
            
            # ヘッドレスモードでのダウンロード設定
            self.driver.execute_cdp_cmd("Page.setDownloadBehavior", {
                "behavior": "allow",
                "downloadPath": download_dir_str
            })
            
            self.logger.info("ChromeDriver の初期化が完了しました")
            return True
            
        except Exception as e:
            self.logger.error(f"ChromeDriver の初期化に失敗しました: {str(e)}")
            
            # 代替方法を試行
            try:
                self.driver = webdriver.Chrome(options=chrome_options)
                self.driver.execute_cdp_cmd("Page.setDownloadBehavior", {
                    "behavior": "allow",
                    "downloadPath": download_dir_str
                })
                self.logger.info("代替方法でChromeDriverを初期化しました")
                return True
            except Exception as e2:
                self.logger.error(f"代替方法も失敗しました: {str(e2)}")
                return False
    
    def access_site(self):
        """SMCLサイトにアクセス"""
        try:
            self.logger.info(f"SMCLサイトにアクセス: {self.target_url}")
            self.driver.get(self.target_url)
            
            # ページ読み込み完了を待つ
            WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            
            WebDriverWait(self.driver, 10).until(
                lambda driver: driver.execute_script("return document.readyState") == "complete"
            )
            
            self.logger.info(f"現在のURL: {self.driver.current_url}")
            self.logger.info(f"ページタイトル: {self.driver.title}")
            
            # 再ログインボタンの確認とクリック
            self._handle_relogin_button()
            
            return True
            
        except TimeoutException:
            self.logger.error("ページの読み込みがタイムアウトしました")
            return False
        except Exception as e:
            self.logger.error(f"サイトアクセスエラー: {str(e)}")
            return False
    
    def _handle_relogin_button(self):
        """再ログインボタンの処理"""
        try:
            # ID指定での検索を優先
            relogin_button = self.driver.find_element(By.ID, "LogoutLinkButton")
            self.logger.info("再ログインボタンが見つかりました（ID指定）")
            
            self.driver.execute_script("arguments[0].scrollIntoView(true);", relogin_button)
            time.sleep(1)
            relogin_button.click()
            self.logger.info("再ログインボタンをクリックしました")
            
            time.sleep(3)
            WebDriverWait(self.driver, 10).until(
                lambda driver: driver.execute_script("return document.readyState") == "complete"
            )
            
        except NoSuchElementException:
            try:
                # テキスト検索を試行
                relogin_button = self.driver.find_element(By.XPATH, "//a[contains(text(), '再ログイン')]")
                self.logger.info("再ログインボタンが見つかりました（テキスト検索）")
                
                self.driver.execute_script("arguments[0].scrollIntoView(true);", relogin_button)
                time.sleep(1)
                relogin_button.click()
                self.logger.info("再ログインボタンをクリックしました")
                
                time.sleep(3)
                
            except NoSuchElementException:
                self.logger.info("再ログインボタンが見つかりませんでした")
    
    def login(self):
        """SMCLサイトにログイン"""
        try:
            self.logger.info("ログイン処理開始")
            
            # 企業コード入力
            corp_field = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "FormView2_CorpCdTextBox"))
            )
            corp_field.clear()
            corp_field.send_keys(self.corp_code)
            self.logger.info(f"企業コード入力: {self.corp_code}")
            
            # ログインID入力
            login_field = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "FormView1_LoginIdTextBox"))
            )
            login_field.clear()
            login_field.send_keys(self.login_id)
            self.logger.info(f"ログインID入力: {self.login_id}")
            
            # パスワード入力
            password_field = self.driver.find_element(By.ID, "FormView1_LoginPwTextBox")
            password_field.clear()
            password_field.send_keys(self.password)
            self.logger.info("パスワード入力完了")
            
            # ログインボタンクリック
            login_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.ID, "FormView1_btnLogin"))
            )
            
            self.driver.execute_script("arguments[0].scrollIntoView(true);", login_button)
            time.sleep(1)
            login_button.click()
            self.logger.info("ログインボタンをクリックしました")
            
            # ログイン結果確認
            time.sleep(3)
            current_url = self.driver.current_url
            
            # エラーメッセージチェック
            try:
                error_elements = self.driver.find_elements(By.XPATH, "//*[contains(text(), 'エラー') or contains(text(), '失敗') or contains(text(), '無効')]")
                if error_elements:
                    self.logger.error("ログインエラーメッセージが検出されました")
                    for elem in error_elements:
                        self.logger.error(f"エラー: {elem.text}")
                    return False
            except:
                pass
            
            self.logger.info("ログインが完了しました")
            return True
            
        except TimeoutException:
            self.logger.error("ログインフィールドの検索がタイムアウトしました")
            return False
        except Exception as e:
            self.logger.error(f"ログインエラー: {str(e)}")
            return False
    
    def select_user_and_navigate(self):
        """ユーザーを選択してアプリトップに遷移"""
        try:
            self.logger.info("ユーザー選択処理開始")
            
            WebDriverWait(self.driver, 10).until(
                lambda driver: driver.execute_script("return document.readyState") == "complete"
            )
            
            # ユーザー選択リンクをクリック
            user_link = WebDriverWait(self.driver, 15).until(
                EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_GridView1_ctl02_Label21"))
            )
            
            self.logger.info(f"ユーザー選択リンク: {user_link.text}")
            
            self.driver.execute_script("arguments[0].scrollIntoView(true);", user_link)
            time.sleep(1)
            user_link.click()
            self.logger.info("ユーザー選択リンクをクリックしました")
            
            # 遷移完了を待つ
            time.sleep(3)
            WebDriverWait(self.driver, 15).until(
                lambda driver: driver.execute_script("return document.readyState") == "complete"
            )
            
            self.logger.info("アプリトップへの遷移が完了しました")
            return True
            
        except TimeoutException:
            self.logger.error("ユーザー選択リンクの検索がタイムアウトしました")
            return False
        except Exception as e:
            self.logger.error(f"ユーザー選択エラー: {str(e)}")
            return False
    
    def _set_search_date(self):
        """検索日付を今日の日付に設定"""
        try:
            from datetime import datetime, timedelta
            
            # 今日の日付を取得（yyyy/mm/dd形式）
            today_str = datetime.now().strftime('%Y/%m/%d')
            # 昨日の日付を取得
            # today_str = (datetime.now() - timedelta(days=2)).strftime('%Y/%m/%d')
            
            self.logger.info(f"検索日付設定開始: {today_str}")
            
            # 日付入力欄を見つける
            date_input = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_FormView1_ImpDateFromTextBox"))
            )
            
            # 既存の値をクリアして今日の日付を入力
            date_input.clear()
            date_input.send_keys(today_str)
            
            self.logger.info(f"検索日付を設定しました: {today_str}")
            
            # 少し待つ（JavaScript処理のため）
            time.sleep(1)
            
        except TimeoutException:
            self.logger.warning("日付入力欄が見つかりませんでした")
        except Exception as e:
            self.logger.error(f"日付設定エラー: {str(e)}")
    
    def _check_no_data_message(self):
        """「該当するデータがありません。」メッセージをチェック"""
        try:
            # メッセージエリアを検索
            message_element = self.driver.find_element(
                By.ID, "ctl00_messageArea_RepeatMessage_ctl00_messageLabel"
            )
            
            message_text = message_element.text.strip()
            self.logger.info(f"メッセージエリアの内容: '{message_text}'")
            
            if "該当するデータがありません" in message_text:
                self.logger.warning("🔍 該当するデータがありません - メッセージを検出しました")
                return True
            
            return False
            
        except NoSuchElementException:
            # メッセージエリアが見つからない（通常の検索結果がある場合）
            self.logger.debug("メッセージエリアが見つかりませんでした（データありの状態）")
            return False
        except Exception as e:
            self.logger.error(f"データなしメッセージチェックエラー: {str(e)}")
            return False
    
    def navigate_to_order_list_and_search(self):
        """受注一覧に遷移して検索"""
        try:
            self.logger.info("受注一覧遷移処理開始")
            
            # 受注一覧タブをクリック
            tab_link = WebDriverWait(self.driver, 15).until(
                EC.element_to_be_clickable((By.ID, "ctl00_tab3link"))
            )
            
            self.driver.execute_script("arguments[0].scrollIntoView(true);", tab_link)
            time.sleep(1)
            tab_link.click()
            self.logger.info("受注一覧タブをクリックしました")
            
            # ページ読み込み完了を待つ
            time.sleep(3)
            WebDriverWait(self.driver, 15).until(
                lambda driver: driver.execute_script("return document.readyState") == "complete"
            )
            
            # =======================================================
            # 🚨 検索条件設定（現状：test_modeに関係なく同じ条件）
            # =======================================================
            
            # 今日の日付を入力（常に実行）
            self._set_search_date()
            
            # セレクトボックスの設定（現在は個別制御で無効化）
            search_conditions_enabled = (self.config and 
                                       hasattr(self.config, 'enable_production_search_conditions') and 
                                       self.config.enable_production_search_conditions)
            
            if search_conditions_enabled:
                # 将来的に本番用検索条件を有効にする場合
                select_element = WebDriverWait(self.driver, 15).until(
                    EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_FormView1_DeciDropDownList"))
                )

                select = Select(select_element)
                current_option = select.first_selected_option
                self.logger.info(f"現在の選択: {current_option.text}")
                
                select.select_by_value("0")  # 参照済
                selected_option = select.first_selected_option
                self.logger.info(f"変更後の選択: {selected_option.text}")
                
                time.sleep(1)
            else:
                self.logger.info("🚨 本番用検索条件は無効（テスト同様）- セレクトボックス変更をスキップ")
            
            # 検索ボタンをクリック
            search_button = WebDriverWait(self.driver, 15).until(
                EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_FormView1_Button1"))
            )
            
            self.driver.execute_script("arguments[0].scrollIntoView(true);", search_button)
            time.sleep(1)
            search_button.click()
            self.logger.info("検索ボタンをクリックしました")
            
            # 検索結果読み込み完了を待つ
            time.sleep(5)
            WebDriverWait(self.driver, 20).until(
                lambda driver: driver.execute_script("return document.readyState") == "complete"
            )
            
            # 検索結果確認
            rows = self.driver.find_elements(By.XPATH, "//table//tr[position()>1]")
            self.logger.info(f"検索結果データ行数: {len(rows)}")
            
            # 「該当するデータがありません。」メッセージをチェック
            if self._check_no_data_message():
                self.logger.warning("🔍 該当するデータが見つかりませんでした")
                self.no_data_found = True
                return True  # エラーではないので True を返す
            else:
                self.no_data_found = False

            return True
            
        except TimeoutException:
            self.logger.error("受注一覧遷移でタイムアウトしました")
            return False
        except Exception as e:
            self.logger.error(f"受注一覧遷移エラー: {str(e)}")
            return False
    
    def process_order_details_and_download(self, target_link_id=None):
        """受注伝票詳細処理とダウンロード"""
        try:
            # 対象リンクIDを決定（デフォルトは最初のリンク）
            link_id = target_link_id or "ctl00_ContentPlaceHolder1_GridView1_ctl02_ImpDateLinkButton"
            self.logger.info(f"受注伝票詳細処理開始: {link_id}")
            
            # 詳細ページに遷移
            detail_link = WebDriverWait(self.driver, 15).until(
                EC.element_to_be_clickable((By.ID, link_id))
            )
            
            self.driver.execute_script("arguments[0].scrollIntoView(true);", detail_link)
            time.sleep(1)
            detail_link.click()
            self.logger.info("受注伝票詳細リンクをクリックしました")
            
            # ページ読み込み完了を待つ
            time.sleep(3)
            WebDriverWait(self.driver, 15).until(
                lambda driver: driver.execute_script("return document.readyState") == "complete"
            )
            
            # 印刷種類を「納品リストD（集計）」に変更
            print_kind_select = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_FormView2_PrintKindDropDownList"))
            )
            
            select = Select(print_kind_select)
            current_option = select.first_selected_option
            self.logger.info(f"現在の印刷種類: {current_option.text}")
            
            select.select_by_value("0400")  # 納品リストD（集計）
            selected_option = select.first_selected_option
            self.logger.info(f"変更後の印刷種類: {selected_option.text}")
            
            time.sleep(1)
            
            # ダウンロードボタンをクリック（ダウンロード）
            print_button = WebDriverWait(self.driver, 15).until(
                EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_FormView2_PrintButton"))
            )
            
            self.driver.execute_script("arguments[0].scrollIntoView(true);", print_button)
            time.sleep(1)
            print_button.click()
            self.logger.info("印刷ボタンをクリックしました（ダウンロード開始）")
            
            # ダウンロード完了を待つ
            time.sleep(5)
            
            # 確定処理（フラグによって制御）
            if self.enable_confirmation_process:
                self.logger.info("確定処理を実行します")
                
                # チェックボックスをクリック
                checkbox = WebDriverWait(self.driver, 15).until(
                    EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_baseCheckbox1"))
                )
                
                self.driver.execute_script("arguments[0].scrollIntoView(true);", checkbox)
                time.sleep(1)
                checkbox.click()
                self.logger.info("チェックボックスをクリックしました")

                # 確定ボタンをクリック
                confirm_button = WebDriverWait(self.driver, 15).until(
                    EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_DecideButton"))
                )
                confirm_button.click()
                self.logger.info("確定ボタンをクリックしました")
            else:
                self.logger.info("🧪 テストモードのため確定処理をスキップしました")
                self.logger.info("   - チェックボックスクリック: スキップ")
                self.logger.info("   - 確定ボタンクリック: スキップ")
            
            time.sleep(2)
            return True
            
        except TimeoutException:
            self.logger.error("受注伝票詳細処理でタイムアウトしました")
            return False
        except Exception as e:
            self.logger.error(f"受注伝票詳細処理エラー: {str(e)}")
            return False
    
    def check_no_data_message(self):
        """「該当するデータがありません」メッセージをチェック"""
        try:
            message_element = self.driver.find_element(By.ID, "ctl00_messageArea_RepeatMessage_ctl00_messageLabel")
            message_text = message_element.text.strip()
            
            if "該当するデータがありません" in message_text:
                self.logger.info("「該当するデータがありません」メッセージを確認")
                return True
            else:
                self.logger.info("まだ処理すべきデータがあります")
                return False
                
        except NoSuchElementException:
            self.logger.info("メッセージ要素が見つかりません（データがある可能性）")
            return False
        except Exception as e:
            self.logger.error(f"メッセージ確認エラー: {str(e)}")
            return False
    
    def download_delivery_lists(self):
        """納品リストをダウンロード（メイン処理）"""
        try:
            self.logger.log_phase_start("納品リストダウンロード")
            
            # ドライバー設定
            if not self.setup_driver():
                return False
            
            # サイトアクセス
            if not self.access_site():
                return False
            
            # ログイン
            if not self.login():
                return False
            
            # ユーザー選択
            if not self.select_user_and_navigate():
                return False
            
            # すべての受注伝票を処理
            attempt = 0
            max_attempts = 50  # 安全のための最大試行回数
            processed_orders = set()  # テストモード用：処理済み受注ID管理
            test_mode = self.config and getattr(self.config, 'test_mode', False)
            no_new_data_count = 0  # テストモード用：新しいデータがない回数をカウント
            
            if test_mode:
                self.logger.info("🧪 テストモード: 確定処理なしで全件処理します")
            
            while attempt < max_attempts:
                attempt += 1
                self.logger.info(f"処理ラウンド {attempt}")
                
                # 受注一覧検索
                if not self.navigate_to_order_list_and_search():
                    self.logger.error("受注一覧検索に失敗")
                    break
                
                # データ存在確認
                if self.check_no_data_message():
                    self.logger.info("すべての納品リストのダウンロードが完了")
                    return True
                
                if attempt == 1:
                    # CSVダウンロード
                    self.download_csv()
                
                # 詳細処理とダウンロード
                if test_mode:
                    # テストモードでは全ての利用可能なリンクを順次処理
                    available_links = self._get_available_order_links()
                    if not available_links:
                        self.logger.info("🏁 テストモード: 利用可能な受注リンクがないため処理終了")
                        return True
                    
                    processed_any = False
                    for link_id in available_links:
                        if link_id not in processed_orders:
                            self.logger.info(f"📋 新規処理開始: {link_id}")
                            
                            if self.process_order_details_and_download(target_link_id=link_id):
                                processed_orders.add(link_id)
                                processed_any = True
                                no_new_data_count = 0
                                self.logger.info(f"✅ 処理完了: {link_id}")
                                
                                # 受注一覧に戻る
                                self._navigate_back_to_order_list()
                                break  # 1件処理したら次のループへ
                            else:
                                self.logger.error(f"❌ 処理失敗: {link_id}")
                                break
                        else:
                            self.logger.debug(f"⏭️ スキップ (処理済み): {link_id}")
                    
                    if not processed_any:
                        no_new_data_count += 1
                        self.logger.info(f"🔄 新規データなし (重複検出 {no_new_data_count}回目)")
                        
                        if no_new_data_count >= 3:
                            self.logger.info("🏁 テストモード: 新しいデータがないため処理終了")
                            return True
                else:
                    # 本番モードでは従来通り
                    if not self.process_order_details_and_download():
                        self.logger.error("詳細処理に失敗")
                        break
                
                # テストモードでの処理件数制限（安全策）
                if test_mode and len(processed_orders) >= 10:
                    self.logger.info(f"🧪 テストモード: 最大処理件数({len(processed_orders)}件)に達したため終了")
                    return True
                
                time.sleep(3)
            
            if attempt >= max_attempts:
                self.logger.warning(f"最大試行回数({max_attempts})に達したため処理を終了")
            
            return False
            
        except Exception as e:
            self.logger.error(f"納品リストダウンロードでエラー: {str(e)}")
            return False
        finally:
            self.cleanup()
    
    # 受注一覧CSVダウンロード
    def download_csv(self):
        """CSVダウンロード"""
        try:
            self.logger.info("CSVダウンロード開始")
            download_button = WebDriverWait(self.driver, 15).until(
                EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_DownloadButton"))
            )
            download_button.click()
            self.logger.info("CSVダウンロードボタンをクリックしました")
            time.sleep(5)
            return True
        except Exception as e:
            self.logger.error(f"CSVダウンロードエラー: {str(e)}")
            return False
    
    def _get_available_order_links(self):
        """現在利用可能な受注リンクのIDリストを取得"""
        try:
            available_links = []
            
            # GridView内の全ての受注リンクを検索
            # パターン: ctl00_ContentPlaceHolder1_GridView1_ctl0X_ImpDateLinkButton
            for i in range(2, 20):  # ctl02からctl19まで（安全のため広めに検索）
                link_id = f"ctl00_ContentPlaceHolder1_GridView1_ctl0{i}_ImpDateLinkButton"
                
                try:
                    element = self.driver.find_element(By.ID, link_id)
                    if element.is_displayed() and element.is_enabled():
                        available_links.append(link_id)
                        self.logger.debug(f"利用可能なリンク発見: {link_id}")
                except NoSuchElementException:
                    # 要素が見つからない場合は終了
                    break
                except Exception as e:
                    self.logger.debug(f"リンクチェックエラー {link_id}: {e}")
                    continue
            
            self.logger.info(f"利用可能な受注リンク: {len(available_links)}件")
            for link_id in available_links:
                self.logger.debug(f"  - {link_id}")
            
            return available_links
            
        except Exception as e:
            self.logger.error(f"受注リンク取得エラー: {str(e)}")
            return []
    
    def _navigate_back_to_order_list(self):
        """詳細画面から受注一覧画面に戻る"""
        try:
            self.logger.info("📋 受注一覧画面に戻ります")
            
            # まずブラウザの戻るボタンを試す
            self.driver.back()
            time.sleep(2)
            
            # さらに一覧画面に戻る必要がある場合
            current_url = self.driver.current_url
            if "EDS001VORD" not in current_url:
                self.driver.back()
                time.sleep(2)
            
            # ページ読み込み完了を待つ
            WebDriverWait(self.driver, 15).until(
                lambda driver: driver.execute_script("return document.readyState") == "complete"
            )
            
            self.logger.info("✅ 受注一覧画面に戻りました")
            return True
            
        except Exception as e:
            self.logger.error(f"受注一覧画面への遷移エラー: {str(e)}")
            return False

    def _get_current_order_id(self):
        """現在処理対象の受注IDを取得（廃止予定）"""
        try:
            # GridView の最初の行から日付と受注情報を取得
            date_link = self.driver.find_element(
                By.ID, "ctl00_ContentPlaceHolder1_GridView1_ctl02_ImpDateLinkButton"
            )
            
            # 日付テキストを取得（例：2025/01/17）
            date_text = date_link.text.strip()
            
            # 同じ行から他の情報も取得（受注番号等があれば）
            try:
                # 受注番号セルを探す（GridViewの同じ行）
                row = date_link.find_element(By.XPATH, "./../../..")  # tr要素に遡る
                cells = row.find_elements(By.TAG_NAME, "td")
                
                # セル情報を組み合わせて一意なIDを作成
                cell_texts = []
                for i, cell in enumerate(cells[:5]):  # 最初の5セルまで
                    cell_text = cell.text.strip()
                    if cell_text and len(cell_text) < 50:  # 長すぎるテキストは除外
                        cell_texts.append(f"{i}:{cell_text}")
                
                order_id = f"{date_text}|{','.join(cell_texts)}"
                self.logger.debug(f"受注ID生成: {order_id}")
                return order_id
                
            except Exception as detail_error:
                # 詳細情報取得に失敗した場合は日付のみを使用
                self.logger.debug(f"詳細情報取得失敗、日付のみ使用: {detail_error}")
                return date_text
            
        except NoSuchElementException:
            self.logger.warning("処理対象の受注IDを取得できませんでした")
            return None
        except Exception as e:
            self.logger.error(f"受注ID取得エラー: {str(e)}")
            return None

    def cleanup(self):
        """リソースのクリーンアップ"""
        if self.driver:
            try:
                time.sleep(2)
                self.driver.quit()
                self.logger.info("ブラウザを閉じました")
            except Exception as e:
                self.logger.error(f"ブラウザクリーンアップエラー: {str(e)}")
            finally:
                self.driver = None 