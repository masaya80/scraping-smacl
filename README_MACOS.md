# macOS環境での実行権限回避方法

## 問題の概要

macOS上でこのシステムを実行すると、Excel周りの処理でシステムから実行権限の確認ダイアログが表示される問題がありました。

## 解決策

このシステムでは自動的にOS環境を検出し、以下の対応を行います：

### macOS環境での自動対応

1. **Excel アプリケーションの使用無効化**
   - macOSでは `xlwings` を使用したExcel アプリケーション起動を自動的に無効化
   - 権限確認ダイアログが表示されなくなります

2. **代替PDF出力方法の使用**
   - `weasyprint` を使用した高品質PDF出力
   - Excel のレイアウトとセル結合を正確に再現
   - 権限確認なしで動作

### Windows環境での互換性

- Windows環境では従来通り `xlwings` を使用してExcel アプリケーション経由でPDF出力
- 最高品質のPDF出力を維持

## インストール

### 基本パッケージ
```bash
pip install -r requirements.txt
```

### macOS用追加パッケージ（推奨）
```bash
# PDF出力品質向上のため
pip install weasyprint
```

`weasyprint` のインストールで問題が発生する場合：

```bash
# macOSでweasprintの依存関係インストール
brew install pango libffi
pip install weasyprint
```

## 環境変数による手動制御

必要に応じて以下の環境変数で動作を制御できます：

### Excel アプリケーション制御
```bash
# 強制的にExcel アプリケーションを使用（macOSでも）
export SMCL_FORCE_EXCEL_APP=true

# 強制的にExcel アプリケーションを無効化（Windowsでも）
export SMCL_DISABLE_EXCEL_APP=true
```

### PDF出力方法制御
```bash
# PDF出力方法を指定
export SMCL_PDF_METHOD=weasyprint  # weasyprint使用
export SMCL_PDF_METHOD=xlwings     # Excel アプリ使用
export SMCL_PDF_METHOD=html        # HTMLファイルのみ生成
```

## トラブルシューティング

### PDF出力ができない場合

1. **weasyprint インストールエラー**
   ```bash
   # macOSの場合
   brew install cairo pango gdk-pixbuf libffi
   pip install --upgrade pip setuptools wheel
   pip install weasyprint
   ```

2. **HTMLファイルが生成される場合**
   - `weasyprint` がインストールされていない場合、HTMLファイルが代替生成されます
   - ブラウザでHTMLファイルを開き、印刷機能でPDF化可能

### 権限エラーが継続する場合

1. **環境変数で強制無効化**
   ```bash
   export SMCL_DISABLE_EXCEL_APP=true
   python main.py
   ```

2. **システム環境設定での許可**
   - システム環境設定 > セキュリティとプライバシー > プライバシー
   - "自動化" または "アクセシビリティ" でPythonやターミナルを許可

## 動作確認

システムが正しく動作しているかログで確認できます：

```bash
# ログでPDF出力方法を確認
tail -f logs/smcl_system_*.log | grep "PDF出力"
```

期待されるログメッセージ：
- macOS: `Excel アプリケーションの使用が無効化されています。代替方法を使用します`
- Windows: `xlwingsでPDF出力: シート名 -> ファイル名`

## 本番環境（Windows）での運用

- Windows環境では自動的に最適な設定（xlwings使用）が適用されます
- macOSでの開発・テスト時のみ代替手法が使用されます
- 最終的なPDF品質はWindows環境で確認してください
