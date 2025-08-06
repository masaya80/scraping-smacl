import pandas as pd
import os
from typing import List, Dict, Union
from datetime import datetime

def extracted_orders_from_csv(file_path: str = "docs/test4.xlsx", sheet_name: str = "商品マスタ") -> List[Dict[str, Union[str, int]]]:
    """
    受注伝票からExcelデータを読み込んで、商品番号と発注数量を配列として抽出する関数
    
    Args:
        file_path (str): 受注伝票ファイルのパス（Excelファイル）
        sheet_name (str): Excelシート名（デフォルト: "アリスト"）
        
    Returns:
        List[Dict]: [{productNumber: 商品番号, orderquantity: 発注数量}, ...]の形式
        
    Raises:
        FileNotFoundError: ファイルが見つからない場合
        ValueError: データの形式が正しくない場合
    """
    
    # ファイルの存在確認
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"ファイルが見つかりません: {file_path}")
    
    try:
        # Excelファイルを読み込み（指定されたシート）
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # データフレームの内容を確認（デバッグ用）
        print("📊 読み込んだデータの列名:")
        print(df.columns.tolist())
        print("\n📊 データの最初の5行:")
        print(df.head())
        
        # 空のDataFrameの場合の処理
        if df.empty or len(df.columns) == 0:
            print("⚠️  データが空です。別のシートを確認してください。")
            # 利用可能なシート名を表示
            excel_file = pd.ExcelFile(file_path)
            print(f"利用可能なシート: {excel_file.sheet_names}")
            return []
        
        # 必要な列を特定（柔軟な列名マッチング）
        product_number_col = None
        order_quantity_col = None
        
        # 商品番号の列を探す（商品コードも含む）
        for col in df.columns:
            col_lower = str(col).lower()
            if any(keyword in col_lower for keyword in ['魚上魚類商品コード']):
                product_number_col = col
                break
        
        # 発注数量の列を探す（受注数も含む）
        for col in df.columns:
            col_lower = str(col).lower()
            if any(keyword in col_lower for keyword in ['発注数量']):
                order_quantity_col = col
                break
        
        if product_number_col is None:
            print("⚠️  商品番号の列が見つかりません。")
            print("利用可能な列:", df.columns.tolist())
            raise ValueError("商品番号の列が見つかりません。列名に'魚上魚類商品コード'を含めてください。")
        
        if order_quantity_col is None:
            print("⚠️  発注数量の列が見つかりません。")
            print("利用可能な列:", df.columns.tolist())
            raise ValueError("発注数量の列が見つかりません。列名に '発注数量'のいずれかを含めてください。")
        
        print(f"\n✅ 使用する列:")
        print(f"   商品コード: {product_number_col}")
        print(f"   発注数量: {order_quantity_col}")
        
        # データを抽出して配列に変換
        extracted_data = []
        
        for index, row in df.iterrows():
            # 空の行はスキップ
            if pd.isna(row[product_number_col]) or pd.isna(row[order_quantity_col]):
                continue
            
            # データ型を適切に変換
            product_number = str(row[product_number_col]).strip()
            
            # 数量を整数に変換（小数点がある場合は丸める）
            try:
                order_quantity = int(float(row[order_quantity_col]))
            except (ValueError, TypeError):
                print(f"⚠️  行 {index + 1}: 数量 '{row[order_quantity_col]}' を数値に変換できませんでした。スキップします。")
                continue
            
            # 辞書形式で追加
            extracted_data.append({
                "productNumber": product_number,
                "orderQuantity": order_quantity
            })
        
        print(f"\n✅ 抽出完了: {len(extracted_data)}件のデータを抽出しました")
        
        return extracted_data
        
    except pd.errors.EmptyDataError:
        raise ValueError("ファイルが空です")
    except Exception as e:
        raise ValueError(f"ファイルの読み込み中にエラーが発生しました: {str(e)}")


def save_to_downloads(data: List[Dict[str, Union[str, int]]], filename_prefix: str = "extracted_orders") -> str:
    """
    抽出されたデータをdownloadsディレクトリにCSVファイルとして保存する関数
    
    Args:
        data (List[Dict]): 抽出されたデータ
        filename_prefix (str): ファイル名のプレフィックス
        
    Returns:
        str: 保存されたファイルのフルパス
    """
    
    # プロジェクト内のdownloadsディレクトリのパスを取得
    downloads_dir = "downloads"
    
    # ディレクトリが存在しない場合は作成
    os.makedirs(downloads_dir, exist_ok=True)
    
    # タイムスタンプ付きのファイル名を作成
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{filename_prefix}_{timestamp}.csv"
    file_path = os.path.join(downloads_dir, filename)
    
    # データをDataFrameに変換
    df = pd.DataFrame(data)
    
    # CSVファイルとして保存（UTF-8エンコーディング）
    df.to_csv(file_path, index=False, encoding='utf-8-sig')
    
    print(f"💾 ファイルを保存しました: {file_path}")
    print(f"📊 保存されたデータ: {len(data)}件")
    
    return file_path


def save_to_downloads_excel(data: List[Dict[str, Union[str, int]]], filename_prefix: str = "extracted_orders") -> str:
    """
    抽出されたデータをdownloadsディレクトリにExcelファイルとして保存する関数
    
    Args:
        data (List[Dict]): 抽出されたデータ
        filename_prefix (str): ファイル名のプレフィックス
        
    Returns:
        str: 保存されたファイルのフルパス
    """
    
    # プロジェクト内のdownloadsディレクトリのパスを取得
    downloads_dir = "downloads"
    
    # ディレクトリが存在しない場合は作成
    os.makedirs(downloads_dir, exist_ok=True)
    
    # タイムスタンプ付きのファイル名を作成
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{filename_prefix}_{timestamp}.xlsx"
    file_path = os.path.join(downloads_dir, filename)
    
    # データをDataFrameに変換
    df = pd.DataFrame(data)
    
    # Excelファイルとして保存
    df.to_excel(file_path, index=False, engine='openpyxl')
    
    print(f"💾 Excelファイルを保存しました: {file_path}")
    print(f"📊 保存されたデータ: {len(data)}件")
    
    return file_path


# 使用例とテスト
if __name__ == "__main__":
    try:
        # 関数を実行（test4.xlsxの商品マスタから読み込み）
        result = extracted_orders_from_csv("docs/test4.xlsx", "商品マスタ")
        
        if result:
            print("\n🎉 抽出結果:")
            for i, order in enumerate(result, 1):
                print(f"  {i}. {order}")
                
            print(f"\n📊 総件数: {len(result)}件")
            
            # Downloadsディレクトリに保存
            csv_path = save_to_downloads(result, "受注データ")
            excel_path = save_to_downloads_excel(result, "受注データ")
            
            print(f"\n✅ 処理完了!")
            print(f"📁 CSV: {csv_path}")
            print(f"📁 Excel: {excel_path}")
            print(f"📂 保存先: downloadsディレクトリ")
        else:
            print("\n⚠️  抽出できるデータがありませんでした。")
            
    except Exception as e:
        print(f"❌ エラーが発生しました: {e}")