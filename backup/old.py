import time
import os
import platform
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime
import sys

def setup_driver(download_dir=None):
    """Chromeドライバーをヘッドレスモードで設定"""
    chrome_options = Options()
    
    # ヘッドレスモードを有効化
    chrome_options.add_argument('--headless')
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

def login_rexass(driver, login_id, password):
    """rexass5.comにログイン"""
    try:
        # ログインページにアクセス
        print("ログインページにアクセスしています...")
        driver.get("https://www.rexass5.com/login/login/init.do")
        
        # ページの読み込みを待つ
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )
        
        # ログインIDフィールドを探す（正しい名前は 'logid'）
        print("ログインIDを入力しています...")
        login_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "logid"))
        )
        login_field.clear()
        login_field.send_keys(login_id)
        
        # パスワードフィールドを探す（正しい名前は 'pasid'）
        print("パスワードを入力しています...")
        password_field = driver.find_element(By.NAME, "pasid")
        password_field.clear()
        password_field.send_keys(password)
        
        # ログインボタンをクリック（クラス名 'action_login' を持つbuttonタグ）
        print("ログインボタンをクリックしています...")
        login_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CLASS_NAME, "action_login"))
        )
        login_button.click()
        
        # ログイン後のページ読み込みを待つ（URLの変更を待つ）
        print("ログイン処理を待っています...")
        WebDriverWait(driver, 10).until_not(
            EC.url_contains("login/login/init.do")
        )
        
        print("ログインに成功しました")
        return True
        
    except Exception as e:
        print(f"ログインエラー: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def navigate_and_download(driver, download_url, download_dir):
    """ダウンロードページにアクセスしてファイルをダウンロード"""
    try:
        # ダウンロードページにアクセス
        print(f"ダウンロードページにアクセスしています: {download_url}")
        driver.get(download_url)
        
        # ページの読み込みを待つ（より長めに設定）
        print("ページの読み込みを待っています...")
        time.sleep(3)
        
        # ページが完全に読み込まれるまで待つ
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )
        
        # JavaScriptの実行完了を待つ
        driver.execute_script("return document.readyState") == "complete"
        time.sleep(2)
        
        print(f"現在のURL: {driver.current_url}")
        
        # 「親子・各店ダウンロード」ボタンを探す（クラス名 'action_download'）
        print("「親子・各店ダウンロード」ボタンを探しています...")
        
        try:
            # 複数の方法でボタンを探す
            download_button = None
            
            # 方法1: クラス名で探す
            try:
                download_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.CLASS_NAME, "action_download"))
                )
            except TimeoutException:
                pass
            
            # 方法2: テキストを含むボタンを探す
            if not download_button:
                try:
                    download_button = driver.find_element(
                        By.XPATH, "//button[contains(text(), '親子') and contains(text(), 'ダウンロード')]"
                    )
                except NoSuchElementException:
                    pass
            
            if not download_button:
                raise Exception("ダウンロードボタンが見つかりません")
            
            print("ボタンを見つけました。クリックしています...")
            
            # ボタンが表示されているか確認してスクロール
            driver.execute_script("arguments[0].scrollIntoView(true);", download_button)
            time.sleep(1)
            
            # ボタンをクリック
            download_button.click()
            
            # ダウンロード完了を待つ
            print("ダウンロードを待っています...")
            
            # ダウンロード前のファイルリストを取得
            before_download = set(os.listdir(download_dir)) if os.path.exists(download_dir) else set()
            
            # ダウンロードが完了するまで待つ（最大60秒）
            download_complete = False
            for i in range(60):
                time.sleep(1)
                current_files = set(os.listdir(download_dir)) if os.path.exists(download_dir) else set()
                new_files = current_files - before_download
                
                # 新しいファイルが見つかり、一時ファイルでない場合
                if new_files:
                    actual_files = [f for f in new_files if not f.endswith(('.tmp', '.crdownload', '.part'))]
                    if actual_files:
                        download_complete = True
                        break
                
                if i % 5 == 0:
                    print(f"  ダウンロード待機中... ({i}秒経過)")
            
            if download_complete:
                # ダウンロードされたファイルの処理
                for new_file in actual_files:
                    original_path = os.path.join(download_dir, new_file)
                    
                    # 現在の日時を追加したファイル名を作成
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    base_name, ext = os.path.splitext(new_file)
                    new_name = f"{base_name}_{timestamp}{ext}"
                    new_path = os.path.join(download_dir, new_name)
                    
                    # ファイル名を変更
                    try:
                        os.rename(original_path, new_path)
                        print(f"ダウンロード完了: {new_name}")
                    except Exception as e:
                        print(f"ファイル名の変更に失敗: {str(e)}")
                        print(f"ダウンロード完了: {new_file}")
            else:
                print("ダウンロードがタイムアウトしました")
            
            return download_complete
            
        except Exception as e:
            print(f"ボタンの検索でエラー: {str(e)}")
            
            # デバッグ用：ページに存在するボタンを表示
            buttons = driver.find_elements(By.TAG_NAME, "button")
            print(f"\n現在のページに存在するボタン: {len(buttons)}個")
            for i, btn in enumerate(buttons[:5]):  # 最初の5個のみ表示
                print(f"  Button {i+1}: {btn.text}, class={btn.get_attribute('class')}")
            
            return False
            
    except Exception as e:
        print(f"ダウンロードエラー: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def main():
    # 開始時刻を記録
    start_time = datetime.now()
    print(f"\n=== Rexass5 自動ダウンロードスクリプト ===")
    print(f"開始時刻: {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"実行環境: {platform.system()} {platform.release()}")
    
    # ダウンロードディレクトリの設定（OSに依存しない方法）
    download_dir = os.path.join(os.getcwd(), 'downloads')
    download_dir = os.path.abspath(download_dir)  # 絶対パスに変換
    
    if not os.path.exists(download_dir):
        os.makedirs(download_dir)
    
    print(f"ダウンロードディレクトリ: {download_dir}")
    
    # ドライバーの設定
    driver = None
    try:
        driver = setup_driver(download_dir)
        
        # ログイン情報
        login_id = "EENFU5"
        password = "TPN6L4"
        
        # ログイン実行
        if login_rexass(driver, login_id, password):
            # ダウンロードページへ移動
            download_url = "https://www.rexass5.com/sanwa/OrderDown/?SESSION_AOP_SYSCD=001"
            
            if navigate_and_download(driver, download_url, download_dir):
                print("\n✅ ダウンロードが正常に完了しました")
                print(f"ファイルは {download_dir} に保存されています")
                
                # ダウンロードされたファイルのリストを表示
                print("\n--- ダウンロードされたファイル ---")
                for file in sorted(os.listdir(download_dir)):
                    if not file.startswith('.'):  # 隠しファイルを除外
                        file_path = os.path.join(download_dir, file)
                        if os.path.isfile(file_path):
                            file_size = os.path.getsize(file_path) / 1024  # KB単位
                            print(f"  📄 {file} ({file_size:.1f} KB)")
            else:
                print("\n❌ ダウンロードに失敗しました")
        else:
            print("\n❌ ログインに失敗しました")
            
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