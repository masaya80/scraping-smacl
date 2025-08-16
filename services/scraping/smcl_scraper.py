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
            
            # セレクトボックスを「参照済」に変更
            if not (self.config and getattr(self.config, 'test_mode', False)):
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

            return True
            
        except TimeoutException:
            self.logger.error("受注一覧遷移でタイムアウトしました")
            return False
        except Exception as e:
            self.logger.error(f"受注一覧遷移エラー: {str(e)}")
            return False
    
    def process_order_details_and_download(self):
        """受注伝票詳細処理とダウンロード"""
        try:
            self.logger.info("受注伝票詳細処理開始")
            
            # 詳細ページに遷移
            detail_link = WebDriverWait(self.driver, 15).until(
                EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_GridView1_ctl02_ImpDateLinkButton"))
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
            test_mode = self.config and getattr(self.config, 'test_mode', False)
            
            while True:
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
                if not self.process_order_details_and_download():
                    self.logger.error("詳細処理に失敗")
                    break
                
                # test_modeがtrueの場合は1回で終了
                if test_mode:
                    self.logger.info("テストモードのため1回で処理を終了します")
                    return True
                
                time.sleep(3)
            
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