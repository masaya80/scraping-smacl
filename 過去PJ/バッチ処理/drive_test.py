#!/usr/bin/env python
# -*- coding: utf-8 -*-

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import os
import re
import io
from datetime import datetime
import pandas as pd
import platform
from zoneinfo import ZoneInfo  # Python 3.9以降の場合

# --------------------------
# 認証設定
# --------------------------

# 使用するスコープ（Drive, スプレッドシート両方が必要な場合）
SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets"
]

tantosha_file = "担当者確認.txt"
output_path = r'\\sv001\Userdata\鮮魚相場表'
tantosha_file_path = os.path.join(output_path, tantosha_file)
date_file_path = os.path.join(output_path, 'date')

# サービスアカウントの認証情報ファイルのパス（適宜パスを変更してください）
SERVICE_ACCOUNT_FILE = "toichi-363dc57bbd54.json"

def authenticate():
    """
    サービスアカウントの認証情報から Google API の認証を行い、
    Drive API のサービスオブジェクトを返します。
    """
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    drive_service = build('drive', 'v3', credentials=creds)
    return drive_service

def get_output_path():
    return output_path
    """
    実行環境に応じたデスクトップのパスを返します。
    Windowsの場合はUSERPROFILE環境変数、macの場合はホームディレクトリ直下のDesktopを利用します。
    """
    system = platform.system()
    if system == "Windows":
        desktop_path = os.path.join(os.environ['USERPROFILE'], 'Desktop')
    else:
        # macOSやその他のUnix系システムではホームディレクトリのDesktopを利用
        desktop_path = os.path.join(os.path.expanduser("~"), 'Desktop')
    return desktop_path

# --------------------------
# 最新のテキストファイルをダウンロードする関数
# --------------------------
def copy_newest_txt_to_desktop(drive_service, folder_id):
    """
    指定フォルダ内のファイルのうち、名前が「yyMMDD_HHmm.txt」の形式に合致するものから
    最新のものをデスクトップにダウンロード（コピー）します。ただし、ファイル名の日付が今日以前の場合はコピーしません。
    """
    # 指定フォルダ内の全ファイルを取得
    results = drive_service.files().list(
        q=f"'{folder_id}' in parents",
        fields="files(id, name)"
    ).execute()
    files = results.get('files', [])

    # 「yyMMDD_HHmm.txt」に合致するか判定する正規表現パターン
    pattern = re.compile(r'^(\d{6}_\d{4})\.txt$')

    newest_file = None
    newest_datetime = None

    # 各ファイルについて、名前がパターンにマッチすれば日付を抽出して比較
    for file in files:
        filename = file.get('name', '')
        match = pattern.match(filename)
        if match:
            date_str = match.group(1)  # 例: "250301_1031"
            try:
                file_datetime = datetime.strptime(date_str, "%y%m%d_%H%M")
            except ValueError:
                continue  # 形式に問題があればスキップ
            if newest_datetime is None or file_datetime > newest_datetime:
                newest_datetime = file_datetime
                newest_file = file

    if newest_file is None:
        print(f"フォルダ(ID: {folder_id})内に対象となるテキストファイルが見つかりませんでした。")
        return

    # 日付チェック：ファイル名から抽出した日付が「今日」の日付であるかチェック（JST基準）
    now_jst = datetime.now(ZoneInfo("Asia/Tokyo")).date()
    if newest_datetime.date() < now_jst:
        print(f"最新ファイルの日時({newest_datetime.date()})が今日({now_jst})以前のため、コピーをスキップします。")
        return

    print(f"最新のファイル: {newest_file['name']} (ID: {newest_file['id']}) をコピーします。")

    # ダウンロード処理
    request = drive_service.files().get_media(fileId=newest_file['id'])
    desktop_path = date_file_path
    local_file_path = os.path.join(desktop_path, newest_file['name'])

    with io.FileIO(local_file_path, 'wb') as fh:
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()
            if status:
                print(f"Download {int(status.progress() * 100)}%")
    
    print(f"ファイルを {local_file_path} にコピーしました。")


# --------------------------
# Google Drive上の最新「最終更新日.txt」をダウンロードし、担当者情報を返す関数
# --------------------------
def process_latest_saishin_folder(drive_service, folder_id):
    """
    Google Drive 内または指定フォルダ内から「最終更新日」フォルダを検索し、その中の
    「最終更新日.txt」をShift_JISでダウンロードして読み込み、カンマ区切りの内容を
    各変数に割り当てます。
    返却する変数:
      - kakai: 係課
      - tantousha_code: 担当者コード
      - tantousha_name: 担当者名
      - uriage_date: 売上日
      - latest_update: 最終更新日
    """
    

    # フォルダ内から「最終更新日.txt」ファイルを検索
    file_results = drive_service.files().list(
        q=f"'{folder_id}' in parents and name = '最終更新日.txt'",
        fields="files(id, name)"
    ).execute()
    files = file_results.get('files', [])
    if not files:
        print("『最終更新日.txt』ファイルが見つかりませんでした。")
        return None
    latest_file = files[0]
    file_id = latest_file.get('id')
    print(f"見つかったファイル: {latest_file.get('name')} (ID: {file_id}) をダウンロードします。")

    # ファイルをダウンロードしてShift_JISでデコード
    request = drive_service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
        if status:
            print(f"Download {int(status.progress() * 100)}%")
    fh.seek(0)
    content_bytes = fh.read()
    content_str = content_bytes.decode('shift_jis')

    # カンマ区切りで内容を各変数に割り当て
    parts = content_str.split(',')
    if len(parts) < 5:
        print("ファイルの内容が期待される形式ではありません。")
        return None
    kakai = parts[0].strip()
    tantousha_code = parts[1].strip()
    tantousha_name = parts[2].strip()
    uriage_date = parts[3].strip()
    latest_update = parts[4].strip()

    print("読み込んだ内容:")
    print("係課:", kakai)
    print("担当者コード:", tantousha_code)
    print("担当者名:", tantousha_name)
    print("売上日:", uriage_date)
    print("最終更新日:", latest_update)

    return kakai, tantousha_code, tantousha_name, uriage_date, latest_update



def update_tantosha_confirmation(drive_service, folder_id):
    """
    売上日(new_uriage)が、今日の日付の場合は、現在の時刻が14:00未満なら更新を実施し、
    14:00以降なら更新しないようにします。
    また、売上日が未来の日付であれば更新します。
    なお、全て日本時刻（JST）での判定です。

    新規レコードを追加する際は、【売上日, 担当者（担当者氏名）, 最終更新日】でソートして保存します。
    """
    # ① 最新情報の取得（この部分は元の実装と同様）
    new_data = process_latest_saishin_folder(drive_service, folder_id)
    if new_data is None:
        print("新しい担当者情報の取得に失敗しました。")
        return
    new_kakai, new_code, new_name, new_uriage, new_latest_update = new_data

    print("売上日:", new_uriage)

    # 売上日のチェック（日本時刻で判定）
    try:
        # 売上日の文字列は "YYYY/MM/DD" 形式と仮定
        sales_date = datetime.strptime(new_uriage, "%Y/%m/%d").date()
    except Exception as e:
        print("売上日の変換に失敗しました:", e)
        return

    # 現在の日本時刻を取得
    now_jst = datetime.now(ZoneInfo("Asia/Tokyo"))
    today = now_jst.date()

    # 売上日が過去の場合は更新しない
    if sales_date < today:
        print("売上日が過去の日付です。担当者確認テキストへの書き込みは行いません。")
        return
    # 売上日が今日の場合は、現在時刻が14:00未満なら更新、それ以降なら更新しない
    elif sales_date == today:
        if now_jst.hour < 14:
            print("売上日が今日ですが、現在時刻が14:00未満なので更新します。")
        else:
            print("売上日が今日で、かつ現在時刻が14:00以降なので、担当者確認テキストへの書き込みは行いません。")
            return
    else:
        # 売上日が未来なら更新
        print("売上日が未来の日付なので更新します。")

    # ② 担当者確認ファイルの読み込み（固定幅フォーマット）
    widths = [2, 9, 10, 15, 25]
    columns = ['管轄部署コード', '担当者コード', '担当者氏名', '売上日', '最終更新日']
    try:
        df = pd.read_fwf(tantosha_file_path, widths=widths, names=columns, encoding='shift_jis')
    except Exception as e:
        print("担当者確認ファイルの読み込みに失敗しました:", e)
        return

    # 既存レコードの抽出などの処理
    df['担当者コード_stripped'] = df['担当者コード'].astype(str).str.strip()
    target_df = df[df['担当者コード_stripped'] == new_code]

    need_update = False
    if not target_df.empty:
        try:
            target_df['最終更新日_dt'] = pd.to_datetime(target_df['最終更新日'].str.strip(), format="%Y/%m/%d %H:%M:%S")
            current_latest = target_df['最終更新日_dt'].max()
            new_dt = datetime.strptime(new_latest_update, "%Y/%m/%d %H:%M:%S")
            if new_dt > current_latest:
                need_update = True
                print("更新があります。既存の最新最終更新日:", current_latest, "新しい最終更新日:", new_dt)
            else:
                print("更新はありません。既存の最新最終更新日が新しいです。")
        except Exception as e:
            print("日付の変換エラー:", e)
            need_update = True
    else:
        need_update = True
        print("該当の担当者コードのレコードは存在しないため、新規追加します。")

    if need_update:
        new_line = "{:<2}{:<8}{}{: <15}{:<25}".format(
            new_kakai,
            new_code,
            new_name.ljust(10, '　'),  # 担当者氏名は全角スペースで右側の余白埋め
            new_uriage,
            new_latest_update
        )

        print("追加する新規レコード:")
        print(new_line)
        new_row = {
            '管轄部署コード': new_kakai,
            '担当者コード': new_code,
            '担当者氏名': new_name,
            '売上日': new_uriage,
            '最終更新日': new_latest_update
        }
        new_row_df = pd.DataFrame([new_row])
        df = pd.concat([df, new_row_df], ignore_index=True)

        # 追加後、【売上日, 担当者（担当者氏名）, 最終更新日】の順でソートする
        try:
            df['売上日_dt'] = pd.to_datetime(df['売上日'].str.strip(), format="%Y/%m/%d", errors='coerce')
        except Exception as e:
            print("売上日の日付変換エラー:", e)
            df['売上日_dt'] = None

        try:
            df['最終更新日_dt'] = pd.to_datetime(df['最終更新日'].str.strip(), format="%Y/%m/%d %H:%M:%S", errors='coerce')
        except Exception as e:
            print("最終更新日の日付変換エラー:", e)
            df['最終更新日_dt'] = None

        df = df.sort_values(by=['売上日', '担当者氏名', '最終更新日'])

        # 補助列を一度削除する前に、重複の排除（例：担当者コードで重複している場合、最新のレコード（＝一番下の行）だけ残す）
        df = df.drop_duplicates(subset=['売上日', '担当者氏名', '最終更新日'], keep='last')

        # 排除後、改めてソート（必要に応じて再度並び替え）
        df = df.sort_values(by=['売上日', '担当者氏名', '最終更新日'])

        # 補助列は不要なので削除
        df = df.drop(columns=['売上日_dt', '最終更新日_dt'])
    else:
        print("更新は不要です。")
    print(df)

    # ソート後の DataFrame を担当者確認ファイルに出力（固定幅フォーマットで整形）
    desktop_path = get_output_path()
    output_file = os.path.join(desktop_path, tantosha_file)
    with open(output_file, "w", encoding="shift_jis") as f:
        for idx, row in df.iterrows():
            line = "{:<2}{:<8}{}{: <15}{:<25}".format(
                str(row['管轄部署コード']).strip(),
                str(row['担当者コード']).strip(),
                str(row['担当者氏名']).strip().ljust(10, '　'),
                str(row['売上日']).strip(),
                str(row['最終更新日']).strip()
            )
            f.write(line + "\n")
    print("更新後の担当者確認ファイルをデスクトップに保存しました:", output_file)


def finalize_confirmation_file():
    """
    担当者確認ファイルを再度読み込み、【売上日, 担当者氏名, 最終更新日】の順でソートした上で、
    担当者コードの重複を排除します。重複は、同じ担当者コードが存在する場合、最新のレコード（並び順上下にあるもの）だけを残します。
    最後に整形済みデータを元のファイルに上書き保存します。
    """
    # ファイルパスを取得（get_output_path() と tantosha_file は元の実装と共通のものを利用）
    file_path = os.path.join(get_output_path(), tantosha_file)
    
    # 固定幅フォーマット用の設定（読み込み用）
    widths = [2, 9, 10, 15, 25]
    columns = ['管轄部署コード', '担当者コード', '担当者氏名', '売上日', '最終更新日']

    # ファイル読み込み
    try:
        df = pd.read_fwf(file_path, widths=widths, names=columns, encoding='shift_jis')
    except Exception as e:
        print("担当者確認ファイルの読み込みに失敗しました:", e)
        return

    # 日付変換用の補助列を追加
    try:
        df['売上日_dt'] = pd.to_datetime(df['売上日'].str.strip(), format="%Y/%m/%d", errors='coerce')
    except Exception as e:
        print("売上日の日付変換エラー:", e)
        df['売上日_dt'] = None

    try:
        df['最終更新日_dt'] = pd.to_datetime(df['最終更新日'].str.strip(), format="%Y/%m/%d %H:%M:%S", errors='coerce')
    except Exception as e:
        print("最終更新日の日付変換エラー:", e)
        df['最終更新日_dt'] = None

    df = df.sort_values(by=['売上日', '担当者氏名', '最終更新日'])

    # 補助列を一度削除する前に、重複の排除（例：担当者コードで重複している場合、最新のレコード（＝一番下の行）だけ残す）
    df = df.drop_duplicates(subset=['売上日', '担当者氏名', '最終更新日'], keep='last')

    # 排除後、改めてソート（必要に応じて再度並び替え）
    df = df.sort_values(by=['売上日', '担当者氏名', '最終更新日'])
    
    # 補助列は不要なので削除
    df = df.drop(columns=['売上日_dt', '最終更新日_dt'])
    
    # 固定幅フォーマットに整形してファイルに上書き保存
    with open(file_path, "w", encoding="shift_jis") as f:
        for idx, row in df.iterrows():
            line = "{:<2}{:<8}{}{: <15}{:<25}".format(
                str(row['管轄部署コード']).strip(),
                str(row['担当者コード']).strip(),
                str(row['担当者氏名']).strip().ljust(10, '　'),
                str(row['売上日']).strip(),
                str(row['最終更新日']).strip()
            )
            f.write(line + "\n")
    print("担当者確認ファイルの最終的なソートおよび重複排除が完了しました:", file_path)



# --------------------------
# メイン処理
# --------------------------

def main():
    # 「相場表」フォルダのID（Google Drive 上の「相場表」フォルダIDに置き換えてください）
    root_folder_id = "1fOIj6Ffaao4q8zFnLa_SnHJbLHe1pPDG"

    # 認証処理とサービスオブジェクトの生成
    drive_service = authenticate()

    # 「相場表」フォルダ内の各個人フォルダ（サブフォルダ）を取得
    results = drive_service.files().list(
        q=f"'{root_folder_id}' in parents and mimeType = 'application/vnd.google-apps.folder'",
        fields="files(id, name)"
    ).execute()
    folders = results.get('files', [])

    if not folders:
        print("『相場表』フォルダ内にサブフォルダが見つかりませんでした。")
        return

    # 各個人フォルダ毎に最新のテキストファイルを取得してデスクトップにコピー
    for folder in folders:
        individual_name = folder.get('name')
        folder_id = folder.get('id')
        print(f"\nProcessing folder for: {individual_name} (ID: {folder_id})")
        copy_newest_txt_to_desktop(drive_service, folder_id)
        update_tantosha_confirmation(drive_service, folder_id)
        finalize_confirmation_file()

if __name__ == "__main__":
    main()
