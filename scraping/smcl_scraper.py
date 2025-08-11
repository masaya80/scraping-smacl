#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SMCL スクレイピングモジュール
"""

import os  # noqa: F401
import time
import platform
from pathlib import Path
from datetime import datetime  # noqa: F401
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager

from utils.logger import Logger


class SMCLScraper:
    """SMCL Webサイトのスクレイピングクラス"""
    
    def __init__(self, download_dir: Path, headless: bool = True):
        self.download_dir = Path(download_dir)
        self.headless = headless
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
            
            # エラーメッセージチェック
            try:
                error_elements = self.driver.find_elements(By.XPATH, "//*[contains(text(), 'エラー') or contains(text(), '失敗') or contains(text(), '無効')]")
                if error_elements:
                    self.logger.error("ログインエラーメッセージが検出されました")
                    for elem in error_elements:
                        self.logger.error(f"エラー: {elem.text}")
                    return False
            except Exception:
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
            select_element = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_FormView1_RefDropDownList"))
            )
            
            select = Select(select_element)
            current_option = select.first_selected_option
            self.logger.info(f"現在の選択: {current_option.text}")
            
            select.select_by_value("1")  # 参照済
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
                EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_FormView2_DownloadButton"))
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
            max_attempts = 10
            attempt = 0
            
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
                
                # 詳細処理とダウンロード
                if not self.process_order_details_and_download():
                    self.logger.error("詳細処理に失敗")
                    break
                
                time.sleep(3)
            
            if attempt >= max_attempts:
                self.logger.warning(f"最大試行回数({max_attempts})に達しました")
            
            return False
            
        except Exception as e:
            self.logger.error(f"納品リストダウンロードでエラー: {str(e)}")
            return False
        finally:
            self.cleanup()
    
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

    # -------------------------------------------------------------
    # 未確定CSVダウンロード関連の新規メソッド
    # -------------------------------------------------------------
    def navigate_to_order_shipping_list(self) -> bool:
        """受注・出荷一覧に遷移（直接URL）"""
        try:
            target_url = "https://smclweb.cs-cxchange.net/smcl/view/ord/EDS001VORD0000.aspx"
            self.logger.info(f"受注・出荷一覧に遷移: {target_url}")
            self.driver.get(target_url)

            WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            WebDriverWait(self.driver, 15).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
            return True
        except Exception as e:
            self.logger.error(f"受注・出荷一覧への遷移エラー: {str(e)}")
            return False

    def _find_first(self, xpaths: list, timeout: int = 10):
        """複数XPathのいずれかで最初に見つかった要素を返す"""
        last_error = None
        for xp in xpaths:
            try:
                return WebDriverWait(self.driver, timeout).until(
                    EC.presence_of_element_located((By.XPATH, xp))
                )
            except Exception as e:
                last_error = e
                continue
        if last_error:
            raise last_error
        return None

    def set_confirmation_status_unconfirmed(self) -> bool:
        """検索条件の『確定状況』を『未確定』系に設定（見つかった場合のみ）"""
        try:
            # 確定状況セレクトをラベル近傍や名称で探索
            select_xpaths = [
                "//label[contains(., '確定状況')]/following::select[1]",
                "//label[contains(., '確定')]/following::select[1]",
                "//select[contains(@id, 'Confirm') or contains(@id, 'Kakutei') or contains(@name, 'Confirm') or contains(@name, 'Kakutei')]",
            ]
            select_elem = self._find_first(select_xpaths, timeout=8)

            select = Select(select_elem)
            # optionをテキスト部分一致で選択
            options = select.options
            chosen = False
            for opt in options:
                text = (opt.text or "").strip()
                if any(key in text for key in ["未確定", "未確定あり", "未確定のみ"]):
                    select.select_by_visible_text(text)
                    chosen = True
                    self.logger.info(f"確定状況を選択: {text}")
                    break

            if not chosen:
                self.logger.warning("確定状況の『未確定』系オプションが見つかりませんでした（スキップ）")
            return True
        except Exception as e:
            self.logger.warning(f"確定状況設定の警告: {str(e)}（続行します）")
            return False

    def click_search_on_order_shipping(self) -> bool:
        """受注・出荷一覧の検索ボタンをクリック"""
        try:
            btn_xpaths = [
                "//input[@type='submit' and contains(@value, '検索')]",
                "//button[contains(., '検索')]",
                "//a[contains(., '検索')]",
            ]
            btn = self._find_first(btn_xpaths, timeout=8)
            self.driver.execute_script("arguments[0].scrollIntoView(true);", btn)
            time.sleep(0.5)
            btn.click()

            # 検索結果の読み込み待ち
            WebDriverWait(self.driver, 20).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
            time.sleep(1.0)
            self.logger.info("検索ボタンをクリックしました")
            return True
        except Exception as e:
            self.logger.error(f"検索クリックエラー: {str(e)}")
            return False

    def _snapshot_existing_csv(self) -> set:
        """現時点のダウンロード済みCSVファイル名のスナップショットを取得"""
        try:
            existing = set(p.name for p in Path(self.download_dir).glob("*.csv"))
            return existing
        except Exception:
            return set()

    def _wait_for_new_csv(self, before: set, timeout: int = 30) -> Path | None:
        """新規CSVダウンロード完了を待機してパスを返す"""
        end = time.time() + timeout
        last_size = None
        last_path = None
        while time.time() < end:
            candidates = [p for p in Path(self.download_dir).glob("受注伝票_*.csv") if p.name not in before]
            if candidates:
                # 最新のものを選ぶ
                candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
                path = candidates[0]
                size = path.stat().st_size
                if last_path and path == last_path and last_size == size:
                    # 前回と同じサイズ → ダウンロード完了とみなす
                    return path
                last_path, last_size = path, size
            time.sleep(0.5)
        return None

    def _try_click_with_xpaths(self, xpaths: list, scroll: bool = True) -> bool:
        """与えたXPathリストのいずれかを見つけてクリック（最初に成功したらTrue）"""
        for xp in xpaths:
            try:
                elem = WebDriverWait(self.driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, xp))
                )
                if scroll:
                    self.driver.execute_script("arguments[0].scrollIntoView(true);", elem)
                    time.sleep(0.2)
                elem.click()
                return True
            except Exception:
                continue
        return False

    def _try_click_with_js_keyword_search(self, keywords: list) -> bool:
        """JSでa/button/inputを走査し、innerText/valueにキーワードを含む要素をクリック"""
        try:
            js = """
            const kws = arguments[0];
            const nodes = Array.from(document.querySelectorAll('a,button,input[type="button"],input[type="submit"]'));
            function txt(el){ return (el.innerText||'') + ' ' + (el.value||''); }
            for (const el of nodes){
              const t = txt(el);
              for (const kw of kws){
                if (t.includes(kw)){
                  el.scrollIntoView({behavior:'instant', block:'center'});
                  el.click();
                  return true;
                }
              }
            }
            return false;
            """
            kws = keywords
            return bool(self.driver.execute_script(js, kws))
        except Exception:
            return False

    def click_download_all_details(self, dest_subdir: str = "unconfirmed") -> bool:
        """全明細ダウンロードをクリックし、新規CSVを待機して保存位置をログ出力

        dest_subdir: ダウンロード後に移動するサブディレクトリ名（例: 'unconfirmed', 'today'）
        """
        try:
            before = self._snapshot_existing_csv()

            # ボタン探索を複数戦略でポーリング
            keywords = [
                "全明細ダウンロード", "明細ダウンロード", "全件ダウンロード", "一括ダウンロード",
                "CSVダウンロード", "CSV出力", "ダウンロード"
            ]
            xpath_sets = [
                [
                    "//a[contains(., '全明細') and contains(., 'ダウンロード')]",
                    "//a[contains(., 'ダウンロード') and contains(., '明細')]",
                    "//button[contains(., '全明細') and contains(., 'ダウンロード')]",
                    "//button[contains(., 'ダウンロード')]",
                    "//input[(contains(@value,'全明細') or contains(@value,'ダウンロード')) and (@type='button' or @type='submit')]",
                ],
                [
                    "//a[contains(., 'CSV')]",
                    "//button[contains(., 'CSV')]",
                    "//input[contains(@value,'CSV')]",
                ],
            ]

            clicked = False
            end = time.time() + 30
            while time.time() < end and not clicked:
                # 1) XPathで直接クリック
                for xpaths in xpath_sets:
                    if self._try_click_with_xpaths(xpaths):
                        clicked = True
                        break
                # 2) JSでキーワード走査
                if not clicked and self._try_click_with_js_keyword_search(keywords):
                    clicked = True

                if not clicked:
                    time.sleep(1.0)

            if not clicked:
                self.logger.error("全明細ダウンロードボタンが見つかりませんでした")
                return False

            self.logger.info("全明細ダウンロードをクリックしました")

            csv_path = self._wait_for_new_csv(before, timeout=60)
            if not csv_path:
                self.logger.error("CSVのダウンロードを確認できませんでした")
                return False

            # 指定サブディレクトリに移動（存在しなければ作成）
            dest_dir = Path(self.download_dir) / dest_subdir
            dest_dir.mkdir(exist_ok=True)
            dest = dest_dir / csv_path.name
            try:
                csv_path.replace(dest)
                self.logger.info(f"CSVを保存({dest_subdir}): {dest}")
            except Exception:
                # 同名がある場合はそのまま
                self.logger.info(f"CSVを検出: {csv_path}")
            return True
        except Exception as e:
            self.logger.error(f"全明細ダウンロード処理エラー: {str(e)}")
            return False

    def download_unconfirmed_csv_via_ui(self) -> bool:
        """（UIフィルタ使用）未確定のみを検索してCSVをダウンロード

        注: 画面に『未確定』フィルタが無い場合は不発のため、通常は
            download_unconfirmed_csv()（当日=未確定運用）を利用してください。
        """
        try:
            self.logger.log_phase_start("未確定CSVダウンロード（UIフィルタ）")

            if not self.setup_driver():
                return False

            if not self.access_site():
                return False

            if not self.login():
                return False

            if not self.select_user_and_navigate():
                return False

            if not self.navigate_to_order_shipping_list():
                return False

            # 確定状況を『未確定』に（見つからない場合は続行）
            self.set_confirmation_status_unconfirmed()

            if not self.click_search_on_order_shipping():
                return False

            if not self.click_download_all_details():
                return False

            self.logger.info("未確定CSVのダウンロードが完了しました（UIフィルタ）")
            return True
        except Exception as e:
            self.logger.error(f"未確定CSVダウンロード（UIフィルタ）でエラー: {str(e)}")
            return False
        finally:
            self.cleanup()

    # -------------------------------------------------------------
    # 今日分CSVダウンロード（未確定≒当日という運用定義）
    # -------------------------------------------------------------
    def set_received_date_today(self) -> bool:
        """受信日(From/To)を当日に設定（見つかった入力へセット）"""
        try:
            today = datetime.now().strftime('%Y/%m/%d')

            # From候補
            from_xpaths = [
                "//label[contains(., '受信') and contains(., '日')]/following::input[1]",
                "//input[contains(@id, 'From') and (@type='text' or @type='date')]",
                "//input[contains(@name, 'From') and (@type='text' or @type='date')]",
                "(//input[@type='date'])[1]",
            ]
            # To候補
            to_xpaths = [
                "//label[contains(., '受信') and contains(., '日')]/following::input[2]",
                "//input[contains(@id, 'To') and (@type='text' or @type='date')]",
                "//input[contains(@name, 'To') and (@type='text' or @type='date')]",
                "(//input[@type='date'])[2]",
            ]

            def _set_input(xpaths: list) -> bool:
                try:
                    elem = self._find_first(xpaths, timeout=6)
                    self.driver.execute_script("arguments[0].scrollIntoView(true);", elem)
                    time.sleep(0.2)
                    elem.clear()
                    elem.send_keys(today)
                    return True
                except Exception:
                    return False

            ok_from = _set_input(from_xpaths)
            ok_to = _set_input(to_xpaths)

            if not ok_from and not ok_to:
                self.logger.warning("受信日フィルタの入力欄が見つかりませんでした（スキップ）")
                return False

            self.logger.info(f"受信日フィルタを当日に設定: {today}")
            return True
        except Exception as e:
            self.logger.warning(f"受信日設定の警告: {str(e)}（続行します）")
            return False

    def download_unconfirmed_csv(self) -> bool:
        """未確定CSVをダウンロード（未確定≒当日という運用定義）"""
        try:
            self.logger.log_phase_start("未確定CSVダウンロード")

            if not self.setup_driver():
                return False

            if not self.access_site():
                return False

            if not self.login():
                return False

            if not self.select_user_and_navigate():
                return False

            if not self.navigate_to_order_shipping_list():
                return False

            # 受信日を当日に（未確定≒当日の運用）
            self.set_received_date_today()

            if not self.click_search_on_order_shipping():
                return False

            if not self.click_download_all_details(dest_subdir="unconfirmed"):
                # フォールバック: ページ送りしながらダウンロードを試みる
                self.logger.warning("全明細DLが見つからないため、ページ送りDLへフォールバックします（未確定）")
                if not self._download_unconfirmed_csv_paginated():
                    return False

            self.logger.info("未確定CSVのダウンロードが完了しました")
            return True
        except Exception as e:
            self.logger.error(f"未確定CSVダウンロードでエラー: {str(e)}")
            return False
        finally:
            self.cleanup()

    # -------------------------------------------------------------
    # ページネーション対応フォールバック
    # -------------------------------------------------------------
    def _click_next_page(self) -> bool:
        """ページャの『次へ』等をクリック。成功でTrue"""
        xpaths = [
            "//a[contains(., '次へ')]",
            "//a[contains(., '次')]",
            "//a[normalize-space(text())='>']",
            "//button[contains(., '次')]",
        ]
        if self._try_click_with_xpaths(xpaths):
            WebDriverWait(self.driver, 10).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
            time.sleep(0.5)
            return True
        # JS走査
        return self._try_click_with_js_keyword_search(["次", ">", "Next"]) or False

    def _download_unconfirmed_csv_paginated(self, max_pages: int = 20) -> bool:
        """ページ送りしながら各ページで未確定CSVダウンロードを試行"""
        success_count = 0
        for _ in range(max_pages):
            if self.click_download_all_details(dest_subdir="unconfirmed"):
                success_count += 1
            # 次ページが無ければ終了
            if not self._click_next_page():
                break
        if success_count == 0:
            self.logger.error("どのページでもCSVダウンロードに失敗しました")
            return False
        self.logger.info(f"ページ送りDL成功（未確定）: {success_count}ファイル")
        return True