import time
import os
import platform
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime
import sys

# 対象URL
TARGET_URL = "https://smclweb.cs-cxchange.net/smcl/view/lin/EDS001OLIN0000.aspx"
    
 # ログイン情報
CORP_CODE = "I26S"
LOGIN_ID = "0600200"
PASSWORD = "toichi04"

"""
ドライバーの設定
"""
def setup_driver(download_dir=None):
    """Chromeドライバーをヘッドレスモードで設定"""
    chrome_options = Options()
    
    # ヘッドレスモードを有効化
    # chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--window-size=1920,1080')
    
    # Windows用の追加設定
    if platform.system() == 'Windows':
        chrome_options.add_argument('--disable-software-rasterizer')
    
    # ダウンロードディレクトリの設定
    if download_dir:
        # Windowsのパスを正規化
        download_dir = os.path.abspath(download_dir)
        prefs = {
            "download.default_directory": download_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        }
        chrome_options.add_experimental_option("prefs", prefs)
    
    try:
        # ChromeDriverManagerを使用してドライバーを自動的に管理
        print(f"システム: {platform.system()} ({platform.machine()})")
        
        # プラットフォームに応じてChromeDriverをインストール
        driver_path = ChromeDriverManager().install()
        
        # Windows環境での特別な処理
        if platform.system() == 'Windows':
            # Windowsではパスの区切り文字を確認
            driver_path = driver_path.replace('/', os.sep)
        
        service = Service(driver_path)
        driver = webdriver.Chrome(service=service, options=chrome_options)
        
        # ヘッドレスモードでのダウンロードを有効化
        if download_dir:
            driver.execute_cdp_cmd("Page.setDownloadBehavior", {
                "behavior": "allow",
                "downloadPath": download_dir
            })
        
        return driver
    
    except Exception as e:
        print(f"ChromeDriverの初期化エラー: {str(e)}")
        print("\n代替方法を試しています...")
        
        # 代替方法: システムにインストールされているChromeDriverを使用
        try:
            driver = webdriver.Chrome(options=chrome_options)
            
            # ヘッドレスモードでのダウンロードを有効化
            if download_dir:
                driver.execute_cdp_cmd("Page.setDownloadBehavior", {
                    "behavior": "allow",
                    "downloadPath": download_dir
                })
            
            return driver
        except Exception as e2:
            print(f"代替方法も失敗しました: {str(e2)}")
            raise

def login_smcl(driver, corp_code, login_id, password):
    """SMCLサイトにログイン"""
    try:
        print("ログイン情報を入力しています...")
        
        # 企業コードフィールドを探して入力
        print("企業コードを入力しています...")
        corp_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "FormView2_CorpCdTextBox"))
        )
        corp_field.clear()
        corp_field.send_keys(corp_code)
        print(f"  企業コード「{corp_code}」を入力しました")
        
        # ログインIDフィールドを探して入力
        print("ログインIDを入力しています...")
        login_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "FormView1_LoginIdTextBox"))
        )
        login_field.clear()
        login_field.send_keys(login_id)
        print(f"  ログインID「{login_id}」を入力しました")
        
        # パスワードフィールドを探して入力
        print("パスワードを入力しています...")
        password_field = driver.find_element(By.ID, "FormView1_LoginPwTextBox")
        password_field.clear()
        password_field.send_keys(password)
        print("  パスワードを入力しました")
        
        # ログインボタンをクリック
        print("ログインボタンをクリックしています...")
        login_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "FormView1_btnLogin"))
        )
        
        # ボタンが表示されているか確認してスクロール
        driver.execute_script("arguments[0].scrollIntoView(true);", login_button)
        time.sleep(1)
        
        login_button.click()
        print("  ログインボタンをクリックしました")
        
        # ログイン後のページ読み込みを待つ
        print("ログイン処理を待っています...")
        time.sleep(3)
        
        # ページの変化を確認（URLの変更やページタイトルの変更など）
        current_url = driver.current_url
        page_title = driver.title
        print(f"  ログイン後のURL: {current_url}")
        print(f"  ログイン後のページタイトル: {page_title}")
        
        # ログイン成功の判定（エラーメッセージがないかチェック）
        try:
            # エラーメッセージ要素があるかチェック
            error_elements = driver.find_elements(By.XPATH, "//*[contains(text(), 'エラー') or contains(text(), '失敗') or contains(text(), '無効')]")
            if error_elements:
                print("❌ ログインエラーメッセージが検出されました:")
                for elem in error_elements:
                    print(f"  - {elem.text}")
                return False
        except:
            pass
        
        print("✅ ログインが完了しました")
        return True
        
    except TimeoutException:
        print("❌ ログインフィールドの検索がタイムアウトしました")
        return False
    except Exception as e:
        print(f"❌ ログインエラー: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

"""
サイトアクセス
"""
def access_smcl_site(driver, target_url):
    """SMCLサイトにアクセス"""
    try:
        # 指定されたURLにアクセス
        print(f"SMCLサイトにアクセスしています: {target_url}")
        driver.get(target_url)
        
        # ページの読み込みを待つ
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )
        
        # JavaScriptの実行完了を待つ
        WebDriverWait(driver, 10).until(
            lambda driver: driver.execute_script("return document.readyState") == "complete"
        )
        
        print(f"現在のURL: {driver.current_url}")
        print(f"ページタイトル: {driver.title}")
        
        # ページの基本情報を取得
        print("\n=== ページ情報 ===")
        
        # 再ログインボタンの存在を確認してクリック
        try:
            # ID指定での検索を優先
            relogin_button = driver.find_element(By.ID, "LogoutLinkButton")
            print("✓ 再ログインボタンが見つかりました（ID指定）")
            print(f"  テキスト: {relogin_button.text}")
            
            # 再ログインボタンをクリック
            print("再ログインボタンをクリックしています...")
            driver.execute_script("arguments[0].scrollIntoView(true);", relogin_button)
            time.sleep(1)
            relogin_button.click()
            print("  再ログインボタンをクリックしました")
            
            # ページの読み込みを待つ
            time.sleep(3)
            WebDriverWait(driver, 10).until(
                lambda driver: driver.execute_script("return document.readyState") == "complete"
            )
            
            print(f"  再ログイン後のURL: {driver.current_url}")
            print(f"  再ログイン後のページタイトル: {driver.title}")
            
        except NoSuchElementException:
            # ID指定で見つからない場合はテキスト検索を試す
            try:
                relogin_button = driver.find_element(By.XPATH, "//a[contains(text(), '再ログイン')]")
                print("✓ 再ログインボタンが見つかりました（テキスト検索）")
                print(f"  テキスト: {relogin_button.text}")
                
                # 再ログインボタンをクリック
                print("再ログインボタンをクリックしています...")
                driver.execute_script("arguments[0].scrollIntoView(true);", relogin_button)
                time.sleep(1)
                relogin_button.click()
                print("  再ログインボタンをクリックしました")
                
                # ページの読み込みを待つ
                time.sleep(3)
                WebDriverWait(driver, 10).until(
                    lambda driver: driver.execute_script("return document.readyState") == "complete"
                )
                
                print(f"  再ログイン後のURL: {driver.current_url}")
                print(f"  再ログイン後のページタイトル: {driver.title}")
                
            except NoSuchElementException:
                print("✗ 再ログインボタンが見つかりませんでした")
        
        # ログインフォームの存在を確認（再ログイン後に再確認）
        try:
            # 少し待ってからログインフォームを確認
            WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "FormView2_CorpCdTextBox"))
            )
            
            corp_field = driver.find_element(By.ID, "FormView2_CorpCdTextBox")
            login_field = driver.find_element(By.ID, "FormView1_LoginIdTextBox")
            password_field = driver.find_element(By.ID, "FormView1_LoginPwTextBox")
            login_button = driver.find_element(By.ID, "FormView1_btnLogin")
            
            print("✓ ログインフォームが見つかりました")
            print(f"  企業コードフィールド: {corp_field.get_attribute('type')}")
            print(f"  ログインIDフィールド: {login_field.get_attribute('type')}")
            print(f"  パスワードフィールド: {password_field.get_attribute('type')}")
            print(f"  ログインボタン: {login_button.get_attribute('value')}")
            
        except (NoSuchElementException, TimeoutException):
            print("✗ ログインフォームが見つかりませんでした")
        
        # ページ内の主要な要素を確認
        try:
            # 画像要素を確認
            images = driver.find_elements(By.TAG_NAME, "img")
            print(f"✓ 画像要素: {len(images)}個")
            
            # テーブル要素を確認
            tables = driver.find_elements(By.TAG_NAME, "table")
            print(f"✓ テーブル要素: {len(tables)}個")
            
            # リンク要素を確認
            links = driver.find_elements(By.TAG_NAME, "a")
            print(f"✓ リンク要素: {len(links)}個")
            
        except Exception as e:
            print(f"ページ要素の確認中にエラー: {str(e)}")
        
        return True
        
    except TimeoutException:
        print("❌ ページの読み込みがタイムアウトしました")
        return False
    except Exception as e:
        print(f"❌ サイトアクセスエラー: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def select_user_and_navigate(driver):
    """ユーザーを選択してアプリトップに遷移"""
    try:
        print("ユーザー選択処理を開始しています...")
        
        # ページの読み込み完了を待つ
        WebDriverWait(driver, 10).until(
            lambda driver: driver.execute_script("return document.readyState") == "complete"
        )
        
        print(f"現在のURL: {driver.current_url}")
        print(f"現在のページタイトル: {driver.title}")
        
        # 指定されたIDのaタグを探してクリック
        print("ユーザー選択リンクを探しています...")
        user_link = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_GridView1_ctl02_Label21"))
        )
        
        print("✓ ユーザー選択リンクが見つかりました")
        print(f"  リンクテキスト: {user_link.text}")
        print(f"  リンク要素: {user_link.tag_name}")
        
        # リンクが表示されているか確認してスクロール
        driver.execute_script("arguments[0].scrollIntoView(true);", user_link)
        time.sleep(1)
        
        # リンクをクリック
        print("ユーザー選択リンクをクリックしています...")
        user_link.click()
        print("  ユーザー選択リンクをクリックしました")
        
        # ページ遷移の完了を待つ
        print("アプリトップへの遷移を待っています...")
        time.sleep(3)
        
        # ページの読み込み完了を待つ
        WebDriverWait(driver, 15).until(
            lambda driver: driver.execute_script("return document.readyState") == "complete"
        )
        
        # 遷移後のページ情報を表示
        print(f"遷移後のURL: {driver.current_url}")
        print(f"遷移後のページタイトル: {driver.title}")
        
        # アプリトップページかどうかを簡単にチェック
        try:
            # アプリトップページの特徴的な要素があるかチェック（例：メニューなど）
            page_source = driver.page_source
            if "アプリ" in page_source or "メニュー" in page_source or "ホーム" in page_source:
                print("✓ アプリトップページに正常に遷移したと思われます")
            else:
                print("? アプリトップページかどうかの確認ができませんでした")
        except Exception as e:
            print(f"ページ内容の確認中にエラー: {str(e)}")
        
        print("✅ ユーザー選択とアプリトップへの遷移が完了しました")
        return True
        
    except TimeoutException:
        print("❌ ユーザー選択リンクの検索がタイムアウトしました")
        print("利用可能な要素を確認してみます...")
        
        try:
            # GridView内の要素を確認
            gridview = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_GridView1")
            links = gridview.find_elements(By.TAG_NAME, "a")
            print(f"GridView内のリンク数: {len(links)}")
            
            for i, link in enumerate(links[:5]):  # 最初の5個まで表示
                try:
                    print(f"  リンク{i+1}: ID={link.get_attribute('id')}, テキスト={link.text}")
                except:
                    print(f"  リンク{i+1}: 情報取得失敗")
                    
        except NoSuchElementException:
            print("GridViewが見つかりませんでした")
        
        return False
        
    except Exception as e:
        print(f"❌ ユーザー選択エラー: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def navigate_to_order_list_and_search(driver):
    """受注一覧に遷移して該当伝票を検索"""
    try:
        print("受注一覧への遷移処理を開始しています...")
        
        # ページの読み込み完了を待つ
        WebDriverWait(driver, 10).until(
            lambda driver: driver.execute_script("return document.readyState") == "complete"
        )
        
        print(f"現在のURL: {driver.current_url}")
        print(f"現在のページタイトル: {driver.title}")
        
        # ステップ1: id="ctl00_tab3link"のaタグをクリック
        print("1. 受注一覧タブをクリックしています...")
        tab_link = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.ID, "ctl00_tab3link"))
        )
        
        print("✓ 受注一覧タブが見つかりました")
        print(f"  タブテキスト: {tab_link.text}")
        
        # タブが表示されているか確認してスクロール
        driver.execute_script("arguments[0].scrollIntoView(true);", tab_link)
        time.sleep(1)
        
        # タブをクリック
        tab_link.click()
        print("  受注一覧タブをクリックしました")
        
        # ページ遷移の完了を待つ
        print("受注一覧ページの読み込みを待っています...")
        time.sleep(3)
        
        WebDriverWait(driver, 15).until(
            lambda driver: driver.execute_script("return document.readyState") == "complete"
        )
        
        print(f"遷移後のURL: {driver.current_url}")
        print(f"遷移後のページタイトル: {driver.title}")
        
        # ステップ2: セレクトボックスの値を「未参照」に変更
        print("\n2. セレクトボックスを「未参照」に変更しています...")
        select_element = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_FormView1_RefDropDownList"))
        )
        
        print("✓ セレクトボックスが見つかりました")
        
        # Selectオブジェクトを作成
        select = Select(select_element)
        
        # 現在の選択値を表示
        current_option = select.first_selected_option
        print(f"  現在の選択: {current_option.text} (value={current_option.get_attribute('value')})")
        
        # 「参照済」(value="1")を選択（テスト用）
        select.select_by_value("1")
        # select.select_by_index("2")  # インデックス指定の場合
        
        # 選択後の値を確認
        selected_option = select.first_selected_option
        print(f"  変更後の選択: {selected_option.text} (value={selected_option.get_attribute('value')})")
        
        time.sleep(1)
        
        # ステップ3: 検索ボタンを押下
        print("\n3. 検索ボタンを押下しています...")
        search_button = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_FormView1_Button1"))
        )
        
        print("✓ 検索ボタンが見つかりました")
        print(f"  ボタンテキスト: {search_button.get_attribute('value')}")
        
        # ボタンが表示されているか確認してスクロール
        driver.execute_script("arguments[0].scrollIntoView(true);", search_button)
        time.sleep(1)
        
        # 検索ボタンをクリック
        search_button.click()
        print("  検索ボタンをクリックしました")
        
        # 検索結果の読み込みを待つ
        print("検索結果の読み込みを待っています...")
        time.sleep(5)
        
        WebDriverWait(driver, 20).until(
            lambda driver: driver.execute_script("return document.readyState") == "complete"
        )
        
        # 検索結果の確認
        print(f"検索後のURL: {driver.current_url}")
        print(f"検索後のページタイトル: {driver.title}")
        
        # 検索結果があるかチェック
        try:
            # GridViewやテーブルなどの結果表示要素を確認
            result_tables = driver.find_elements(By.TAG_NAME, "table")
            print(f"✓ 検索結果テーブル数: {len(result_tables)}")
            
            # データ行があるかチェック
            rows = driver.find_elements(By.XPATH, "//table//tr[position()>1]")  # ヘッダー行以外
            print(f"✓ データ行数: {len(rows)}")
            
            if len(rows) > 0:
                print("✓ 検索結果にデータが見つかりました")
            else:
                print("? 検索結果にデータが見つかりませんでした")
                
        except Exception as e:
            print(f"検索結果の確認中にエラー: {str(e)}")
        
        print("✅ 受注一覧への遷移と検索が完了しました")
        return True
        
    except TimeoutException:
        print("❌ 要素の検索がタイムアウトしました")
        print("利用可能な要素を確認してみます...")
        
        try:
            # ページ内のリンクやボタンを確認
            links = driver.find_elements(By.TAG_NAME, "a")
            buttons = driver.find_elements(By.TAG_NAME, "input")
            selects = driver.find_elements(By.TAG_NAME, "select")
            
            print(f"ページ内のリンク数: {len(links)}")
            print(f"ページ内のボタン数: {len(buttons)}")
            print(f"ページ内のセレクト数: {len(selects)}")
            
            # tab3linkを探してみる
            for link in links[:10]:  # 最初の10個まで確認
                link_id = link.get_attribute('id')
                if link_id and 'tab' in link_id.lower():
                    print(f"  タブリンク候補: ID={link_id}, テキスト={link.text}")
                    
        except Exception as e:
            print(f"要素確認中にエラー: {str(e)}")
        
        return False
        
    except Exception as e:
        print(f"❌ 受注一覧遷移エラー: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def process_order_details_and_download(driver):
    """受注伝票の詳細処理とダウンロード"""
    try:
        print("受注伝票の詳細処理を開始しています...")
        
        # ページの読み込み完了を待つ
        WebDriverWait(driver, 10).until(
            lambda driver: driver.execute_script("return document.readyState") == "complete"
        )
        
        print(f"現在のURL: {driver.current_url}")
        print(f"現在のページタイトル: {driver.title}")
        
        # ステップ1: 受注伝票の詳細ページに遷移
        print("1. 受注伝票詳細ページに遷移しています...")
        detail_link = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_GridView1_ctl02_ImpDateLinkButton"))
        )
        
        print("✓ 受注伝票詳細リンクが見つかりました")
        print(f"  リンクテキスト: {detail_link.text}")
        
        # リンクが表示されているか確認してスクロール
        driver.execute_script("arguments[0].scrollIntoView(true);", detail_link)
        time.sleep(1)
        
        # リンクをクリック
        detail_link.click()
        print("  受注伝票詳細リンクをクリックしました")
        
        # ページ遷移の完了を待つ
        print("詳細ページの読み込みを待っています...")
        time.sleep(3)
        
        WebDriverWait(driver, 15).until(
            lambda driver: driver.execute_script("return document.readyState") == "complete"
        )
        
        print(f"詳細ページのURL: {driver.current_url}")
        print(f"詳細ページタイトル: {driver.title}")
        
        # ステップ2: 印刷種類を「納品リストD（集計）」に変更
        print("\n2. 印刷種類を「納品リストD（集計）」に変更しています...")
        print_kind_select = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_FormView2_PrintKindDropDownList"))
        )
        
        print("✓ 印刷種類セレクトボックスが見つかりました")
        
        # Selectオブジェクトを作成
        select = Select(print_kind_select)
        
        # 現在の選択値を表示
        current_option = select.first_selected_option
        print(f"  現在の選択: {current_option.text} (value={current_option.get_attribute('value')})")
        
        # 「納品リストD（集計）」(value="0500")を選択
        select.select_by_value("0400")
        
        # 選択後の値を確認
        selected_option = select.first_selected_option
        print(f"  変更後の選択: {selected_option.text} (value={selected_option.get_attribute('value')})")
        
        time.sleep(1)
        
        # ステップ3: 印刷ボタンをクリックしてダウンロード
        print("\n3. 印刷ボタンをクリックしてダウンロードしています...")
        print_button = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_FormView2_PrintButton"))
        )
        
        print("✓ 印刷ボタンが見つかりました")
        print(f"  ボタンテキスト: {print_button.get_attribute('value')}")
        
        # ボタンが表示されているか確認してスクロール
        driver.execute_script("arguments[0].scrollIntoView(true);", print_button)
        time.sleep(1)
        
        # 印刷ボタンをクリック
        print_button.click()
        print("  印刷ボタンをクリックしました")
        
        # ダウンロード処理の完了を待つ
        print("ダウンロード処理を待っています...")
        time.sleep(5)  # ダウンロード処理のため少し長めに待機
        
        print("=== ここまでダウンロード処理 ===")
        
        # ステップ4: チェックボックスをクリック
        print("\n4. チェックボックスをクリックしています...")
        checkbox = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_baseCheckbox1"))
        )
        
        print("✓ チェックボックスが見つかりました")
        
        # チェックボックスが表示されているか確認してスクロール
        driver.execute_script("arguments[0].scrollIntoView(true);", checkbox)
        time.sleep(1)
        
        # チェックボックスをクリック
        checkbox.click()
        print("  チェックボックスをクリックしました")
        
        # 少し待機
        time.sleep(2)
        
        print("✅ 受注伝票の詳細処理とダウンロードが完了しました")
        return True
        
    except TimeoutException:
        print("❌ 要素の検索がタイムアウトしました")
        return False
        
    except Exception as e:
        print(f"❌ 受注伝票処理エラー: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def check_no_data_message(driver):
    """「該当するデータがありません」メッセージをチェック"""
    try:
        message_element = driver.find_element(By.ID, "ctl00_messageArea_RepeatMessage_ctl00_messageLabel")
        message_text = message_element.text.strip()
        print(f"メッセージ確認: {message_text}")
        
        if "該当するデータがありません" in message_text:
            print("✓ 「該当するデータがありません」メッセージが見つかりました")
            return True
        else:
            print("✓ まだ処理すべきデータがあります")
            return False
            
    except NoSuchElementException:
        print("メッセージ要素が見つかりませんでした（データがある可能性があります）")
        return False
    except Exception as e:
        print(f"メッセージ確認エラー: {str(e)}")
        return False

def process_all_orders(driver):
    """すべての受注伝票を処理するループ"""
    try:
        print("すべての受注伝票の処理を開始します...")
        max_attempts = 10  # 無限ループを避けるため最大試行回数を設定
        attempt = 0
        
        while attempt < max_attempts:
            attempt += 1
            print(f"\n=== 処理ラウンド {attempt} ===")
            
            # 受注一覧の検索を実行
            print("受注一覧の検索を実行しています...")
            if not navigate_to_order_list_and_search(driver):
                print("❌ 受注一覧の検索に失敗しました")
                break
                
            # 「該当するデータがありません」メッセージをチェック
            if check_no_data_message(driver):
                print("✅ すべての受注伝票の処理が完了しました")
                return True
                
            # 受注伝票の詳細処理とダウンロード
            print("\n受注伝票の詳細処理を実行しています...")
            if not process_order_details_and_download(driver):
                print("❌ 受注伝票の詳細処理に失敗しました")
                break
                
            # 少し待機してから次のラウンドへ
            print("次の処理ラウンドまで待機しています...")
            time.sleep(3)
        
        if attempt >= max_attempts:
            print(f"⚠️  最大試行回数({max_attempts})に達しました")
            
        return False
        
    except Exception as e:
        print(f"❌ 受注伝票処理ループエラー: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

"""
メイン処理
"""
def main():
    # 開始時刻を記録
    start_time = datetime.now()
    print(f"\n=== SMCL サイトアクセス・ログインスクリプト ===")
    print(f"開始時刻: {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"実行環境: {platform.system()} {platform.release()}")
    
    # ダウンロードディレクトリの設定（将来的なダウンロード用）
    download_dir = os.path.join(os.getcwd(), 'downloads')
    download_dir = os.path.abspath(download_dir)
    
    if not os.path.exists(download_dir):
        os.makedirs(download_dir)
    
    print(f"ダウンロードディレクトリ: {download_dir}")
    
    # ドライバーの設定
    driver = None
    try:
        driver = setup_driver(download_dir)
        
        # SMCLサイトにアクセス
        if access_smcl_site(driver, TARGET_URL):
            print("\n✅ サイトアクセスが正常に完了しました")
            
            # ログイン実行
            print("\n=== ログイン処理開始 ===")
            if login_smcl(driver, CORP_CODE, LOGIN_ID, PASSWORD):
                print("\n✅ ログインが正常に完了しました")
                
                # ユーザー選択とアプリトップへの遷移
                print("\n=== ユーザー選択処理開始 ===")
                if select_user_and_navigate(driver):
                    print("\n✅ ユーザー選択とアプリトップへの遷移が正常に完了しました")
                    
                    # すべての受注伝票を処理
                    print("\n=== すべての受注伝票処理開始 ===")
                    if process_all_orders(driver):
                        print("\n✅ すべての受注伝票の処理が正常に完了しました")
                        
                        # 最終的なページ情報を表示
                        print(f"最終URL: {driver.current_url}")
                        print(f"最終ページタイトル: {driver.title}")
                        
                    else:
                        print("\n❌ 受注伝票の処理に失敗しました")
                    
                else:
                    print("\n❌ ユーザー選択に失敗しました")
                
            else:
                print("\n❌ ログインに失敗しました")
        else:
            print("\n❌ サイトアクセスに失敗しました")
            
    except Exception as e:
        print(f"\n❌ エラーが発生しました: {str(e)}")
        import traceback
        traceback.print_exc()
        
    finally:
        # ブラウザを閉じる
        if driver:
            time.sleep(2)
            driver.quit()
            print("\nブラウザを閉じました")
        
        # 処理時間を表示
        end_time = datetime.now()
        duration = end_time - start_time
        print(f"終了時刻: {end_time.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"処理時間: {duration.total_seconds():.1f}秒")

if __name__ == "__main__":
    main()
