#!/bin/bash

# =============================================================================
# SMCL システム デプロイ用クリーンアップスクリプト
# =============================================================================
# 
# このスクリプトは、システムをデプロイする前に不要なファイルを削除します。
# 実行前に必ずバックアップを取ってください。
#
# 使用方法:
#   chmod +x cleanup_for_deploy.sh
#   ./cleanup_for_deploy.sh
#
# =============================================================================

set -e  # エラー時に停止

# スクリプトのディレクトリに移動
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

echo "🧹 SMCL システム デプロイ用クリーンアップを開始します..."
echo "📁 作業ディレクトリ: $SCRIPT_DIR"
echo ""

# 削除するファイル・ディレクトリのリスト
CLEANUP_TARGETS=(
    # Python キャッシュファイル
    "__pycache__"
    "services/__pycache__"
    "services/core/__pycache__"
    "services/data_processing/__pycache__"
    "services/notification/__pycache__"
    "services/scraping/__pycache__"
    
    # ログファイル・ディレクトリ
    "logs"
    "services/logs"
    
    # 出力ファイル・ディレクトリ
    "output"
    "services/output"
    
    # ダウンロードファイル
    "services/downloads"
    
    # 過去プロジェクト
    "過去PJ"
    
    # macOS システムファイル
    ".DS_Store"
    "*/.DS_Store"
    "._*"
    "**/._*"
    
    # 一時ファイル
    "*.tmp"
    "*.temp"
    
    # Python コンパイル済みファイル
    "*.pyc"
    "*.pyo"
    "*.pyd"
    
    # ログファイル
    "*.log"
    
    # 仮想環境（もしあれば）
    "venv"
    ".venv"
    "env"
    ".env"
    
    # IDE設定ファイル
    ".vscode"
    ".idea"
    
    # その他の不要ファイル
    "*.swp"
    "*.swo"
    "*~"
    "#*#"
    ".pytest_cache"
    "*.egg-info"
    
    # エディタの一時ファイル
    "*.sublime-workspace"
    "*.sublime-project"
    ".vscode/settings.json"
    ".idea/"
)

# 削除実行関数
cleanup_files() {
    local target="$1"
    local count=0
    
    # ファイルパターンの場合
    if [[ "$target" == *"*"* ]]; then
        # find を使用してパターンマッチングファイルを検索・削除
        while IFS= read -r -d '' file; do
            if [[ -e "$file" ]]; then
                echo "  🗑️  削除: $file" >&2
                rm -rf "$file"
                ((count++))
            fi
        done < <(find . -name "$target" -print0 2>/dev/null || true)
    else
        # 通常のファイル・ディレクトリの場合
        if [[ -e "$target" ]]; then
            echo "  🗑️  削除: $target" >&2
            rm -rf "$target"
            ((count++))
        fi
    fi
    
    echo $count
}

# メイン削除処理
total_deleted=0

for target in "${CLEANUP_TARGETS[@]}"; do
    echo "🔍 チェック中: $target"
    deleted=$(cleanup_files "$target")
    total_deleted=$((total_deleted + deleted))
done

echo ""
echo "✅ クリーンアップ完了!"
echo "📊 削除されたファイル・ディレクトリ数: $total_deleted"
echo ""

# 残存ファイルサイズの確認
if command -v du >/dev/null 2>&1; then
    echo "📏 クリーンアップ後のディレクトリサイズ:"
    du -sh . 2>/dev/null || echo "  サイズ計算に失敗しました"
    echo ""
fi

# 重要ファイルの存在確認
echo "🔍 重要ファイルの存在確認:"
important_files=(
    "main.py"
    "requirements.txt"
    "README.md"
    "services"
    "本番実行.bat"
    "通常実行.bat"
)

all_important_exist=true
for file in "${important_files[@]}"; do
    if [[ -e "$file" ]]; then
        echo "  ✅ $file"
    else
        echo "  ❌ $file (見つかりません!)"
        all_important_exist=false
    fi
done

echo ""
if $all_important_exist; then
    echo "🎉 すべての重要ファイルが確認できました。デプロイ準備完了です!"
else
    echo "⚠️  一部の重要ファイルが見つかりません。確認してください。"
    exit 1
fi

echo ""
echo "📦 次のステップ:"
echo "  1. このディレクトリを圧縮してください"
echo "  2. ファイル転送サービスでアップロードしてください"
echo "  3. デプロイ先でダウンロード・展開してください"
echo ""
echo "🎯 デプロイ用クリーンアップが正常に完了しました!"
