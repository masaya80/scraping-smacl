#!/bin/bash

# =============================================================================
# SMCL システム 統合デプロイ準備スクリプト
# =============================================================================
# 
# このスクリプトは以下の処理を順次実行します：
# 1. 日本語ファイル名を英語に変換（文字化け防止）
# 2. 不要ファイルのクリーンアップ
# 3. 圧縮ファイルの作成
#
# 使用方法:
#   chmod +x prepare_for_deploy.sh
#   ./prepare_for_deploy.sh
#
# =============================================================================

set -e  # エラー時に停止

# スクリプトのディレクトリに移動
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

# 現在の日時を取得（圧縮ファイル名用）
TIMESTAMP=$(date +"%Y%m%d_%H%M%S")
ARCHIVE_NAME="smcl_system_${TIMESTAMP}"

echo "🚀 SMCL システム デプロイ準備を開始します..."
echo "📁 作業ディレクトリ: $SCRIPT_DIR"
echo "📦 作成予定アーカイブ: ${ARCHIVE_NAME}.tar.gz"
echo ""

# ステップ1: ファイル名変換
echo "==================== ステップ 1: ファイル名変換 ===================="
if [[ -f "./rename_for_deploy.sh" ]]; then
    echo "📝 日本語ファイル名を英語に変換中..."
    ./rename_for_deploy.sh
    echo ""
else
    echo "⚠️  rename_for_deploy.sh が見つかりません。ファイル名変換をスキップします。"
    echo ""
fi

# ステップ2: クリーンアップ
echo "==================== ステップ 2: クリーンアップ ===================="
if [[ -f "./cleanup_for_deploy.sh" ]]; then
    echo "🧹 不要ファイルをクリーンアップ中..."
    ./cleanup_for_deploy.sh
    echo ""
else
    echo "⚠️  cleanup_for_deploy.sh が見つかりません。クリーンアップをスキップします。"
    echo ""
fi

# ステップ3: 圧縮
echo "==================== ステップ 3: アーカイブ作成 ===================="
echo "📦 システムファイルを圧縮中..."

# 圧縮対象から除外するファイル・ディレクトリ
EXCLUDE_PATTERNS=(
    "--exclude=*.tar.gz"
    "--exclude=*.zip"
    "--exclude=.git"
    "--exclude=.gitignore"
    "--exclude=*.DS_Store"
    "--exclude=__pycache__"
    "--exclude=*.pyc"
    "--exclude=*.log"
    "--exclude=logs"
    "--exclude=output"
    "--exclude=.venv"
    "--exclude=venv"
    "--exclude=node_modules"
    "--exclude=.pytest_cache"
    "--exclude=*.egg-info"
    "--exclude=past_projects"  # 変換後の過去プロジェクト名
    "--exclude=過去PJ"         # 変換前の過去プロジェクト名（念のため）
)

# tar.gz形式で圧縮
if tar czf "../${ARCHIVE_NAME}.tar.gz" "${EXCLUDE_PATTERNS[@]}" -C .. "$(basename "$SCRIPT_DIR")"; then
    echo "✅ アーカイブ作成完了: ../${ARCHIVE_NAME}.tar.gz"
    
    # ファイルサイズを表示
    if command -v ls >/dev/null 2>&1; then
        archive_size=$(ls -lh "../${ARCHIVE_NAME}.tar.gz" | awk '{print $5}')
        echo "📏 アーカイブサイズ: $archive_size"
    fi
else
    echo "❌ アーカイブ作成に失敗しました。"
    exit 1
fi

echo ""
echo "==================== デプロイ準備完了 ===================="
echo ""
echo "🎉 デプロイ準備が正常に完了しました！"
echo ""
echo "📋 次のステップ:"
echo "  1. ../${ARCHIVE_NAME}.tar.gz を Google Drive にアップロード"
echo "  2. Windows PC でダウンロード"
echo "  3. 7-Zip や WinRAR で展開"
echo "  4. production_run.bat を実行してシステム起動"
echo ""
echo "💡 重要な変更点:"
echo "  - 本番実行.bat → production_run.bat"
echo "  - 通常実行.bat → normal_run.bat"
echo "  - 自動実行_登録.bat → auto_schedule_register.bat"
echo "  - UNC対応で実行.bat → unc_compatible_run.bat"
echo ""
echo "🔧 Windows側での実行方法:"
echo "  production_run.bat をダブルクリックして実行"
echo ""
echo "✨ 文字化けは発生しません！"
