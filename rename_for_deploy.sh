#!/bin/bash

# =============================================================================
# SMCL システム デプロイ用ファイル名変換スクリプト
# =============================================================================
# 
# このスクリプトは、日本語ファイル名を英語に変換してWindows環境での
# 文字化けを防ぎます。デプロイ前に実行してください。
#
# 使用方法:
#   chmod +x rename_for_deploy.sh
#   ./rename_for_deploy.sh
#
# =============================================================================

set -e  # エラー時に停止

# スクリプトのディレクトリに移動
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

echo "🔄 SMCL システム デプロイ用ファイル名変換を開始します..."
echo "📁 作業ディレクトリ: $SCRIPT_DIR"
echo ""

# ファイル名変換マッピング
declare -A RENAME_MAP=(
    # バッチファイル
    ["本番実行.bat"]="production_run.bat"
    ["通常実行.bat"]="normal_run.bat"
    ["自動実行_登録.bat"]="auto_schedule_register.bat"
    ["UNC対応で実行.bat"]="unc_compatible_run.bat"
    
    # スクリプトディレクトリ内
    ["scripts/システム状態確認.bat"]="scripts/system_status_check.bat"
    ["scripts/タスクスケジューラ設定.ps1"]="scripts/task_scheduler_setup.ps1"
    ["scripts/デプロイスクリプト.bat"]="scripts/deploy_script.bat"
    ["scripts/納品リスト処理バッチ_python版.bat"]="scripts/delivery_list_batch_python.bat"
    ["scripts/納品リスト処理バッチ.bat"]="scripts/delivery_list_batch.bat"
    ["scripts/自動実行_毎朝10時.bat"]="scripts/auto_run_daily_10am.bat"
    
    # ツールディレクトリ内
    ["tools/debug_batch.bat"]="tools/debug_batch.bat"
    ["tools/minimal_test.bat"]="tools/minimal_test.bat"
    ["tools/run_with_cmd.bat"]="tools/run_with_cmd.bat"
    ["tools/test_simple.bat"]="tools/test_simple.bat"
    
    # サービスドキュメント
    ["services/docs/角上魚類.xlsx"]="services/docs/kakujo_gyorui.xlsx"
    
    # 過去プロジェクト（削除対象だが念のため）
    ["過去PJ"]="past_projects"
    ["過去PJ/バッチ処理"]="past_projects/batch_processing"
    ["過去PJ/バッチ処理/相場表バッチ.bat"]="past_projects/batch_processing/market_price_batch.bat"
)

# ファイル名変換実行関数
rename_file() {
    local old_name="$1"
    local new_name="$2"
    
    if [[ -e "$old_name" ]]; then
        echo "  📝 変換: $old_name → $new_name"
        
        # 新しいファイル名のディレクトリが存在しない場合は作成
        local new_dir=$(dirname "$new_name")
        if [[ "$new_dir" != "." && ! -d "$new_dir" ]]; then
            mkdir -p "$new_dir"
        fi
        
        # ファイル/ディレクトリを移動
        mv "$old_name" "$new_name"
        return 0
    else
        echo "  ⚠️  スキップ: $old_name (見つかりません)"
        return 1
    fi
}

# メイン変換処理
total_renamed=0

echo "🔍 ファイル名変換を実行中..."
echo ""

for old_name in "${!RENAME_MAP[@]}"; do
    new_name="${RENAME_MAP[$old_name]}"
    echo "🔍 チェック中: $old_name"
    
    if rename_file "$old_name" "$new_name"; then
        ((total_renamed++))
    fi
done

echo ""
echo "✅ ファイル名変換完了!"
echo "📊 変換されたファイル・ディレクトリ数: $total_renamed"
echo ""

# 変換後のファイル存在確認
echo "🔍 重要ファイルの存在確認（変換後）:"
important_files_after=(
    "main.py"
    "requirements.txt"
    "README.md"
    "services"
    "production_run.bat"
    "normal_run.bat"
    "auto_schedule_register.bat"
    "unc_compatible_run.bat"
)

all_important_exist=true
for file in "${important_files_after[@]}"; do
    if [[ -e "$file" ]]; then
        echo "  ✅ $file"
    else
        echo "  ❌ $file (見つかりません!)"
        all_important_exist=false
    fi
done

echo ""
if $all_important_exist; then
    echo "🎉 すべての重要ファイルが確認できました。"
else
    echo "⚠️  一部の重要ファイルが見つかりません。確認してください。"
fi

echo ""
echo "📋 変換されたファイル名の対応表:"
echo "----------------------------------------"
for old_name in "${!RENAME_MAP[@]}"; do
    new_name="${RENAME_MAP[$old_name]}"
    if [[ -e "$new_name" ]]; then
        echo "  $old_name → $new_name"
    fi
done

echo ""
echo "📦 次のステップ:"
echo "  1. ./cleanup_for_deploy.sh でクリーンアップを実行"
echo "  2. ディレクトリを圧縮"
echo "  3. Google Drive にアップロード"
echo "  4. Windows PC でダウンロード・展開"
echo ""
echo "💡 Windows側では英語ファイル名で展開されるため、文字化けは発生しません。"
echo "🎯 ファイル名変換が正常に完了しました!"
