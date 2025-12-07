# 業務データ自動処理・統合システム

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![Google Apps Script](https://img.shields.io/badge/GAS-Apps%20Script-green.svg)](https://developers.google.com/apps-script)

メール添付PDFの自動収集からExcelデータ化、複数データソースの統合まで、一連の業務プロセスを自動化するシステムです。

## 📋 プロジェクト概要

このプロジェクトは、以下の業務フローを自動化します：

1. **Gmail添付PDFの自動保存** (Google Apps Script)
2. **PDFからExcelへのデータ抽出** (VBA + Power Query)
3. **Excelデータのスプレッドシートへの集約** (Google Apps Script)
4. **複数データソースの統合と照合** (Python/Google Colab)

### 主な特徴

- ✅ メール受信からデータ統合まで完全自動化
- ✅ 重複データの自動検知と排除
- ✅ 複数データソースの管理番号による照合
- ✅ スタイル付き統合レポートの自動生成
- ✅ エラーハンドリングとログ記録

## 🔄 全体フロー

![Automation Workflow Diagram](./docs/images/workflow-diagram.png)

```
┌─────────────────────────────────────────────────────────────────┐
│                        業務データ自動処理フロー                  　　　　　　　　　　　　　　　　　　　　　　　　　 　　│
└─────────────────────────────────────────────────────────────────┘

① PDF自動保存 (GAS)
   Gmail → PDFファイル抽出 → Google Drive保存
          ↓
② PDF→Excelデータ化 (VBA + Power Query)
   PDF → Power Query → テキスト抽出 → Excel形式保存
          ↓
③ データ集約 (GAS)
   複数Excelファイル → CSV読み込み → スプレッドシート統合
          ↓
④ データ統合・照合 (Python)
   3つのデータソース → 管理番号で統合 → 照合用Excel出力
```

## 🔍 技術選定の背景

### なぜPower Queryを使用するのか？

ステップ②でPower Queryを採用した理由：

1. **GAS OCRの精度不足**
   - Google Apps ScriptのOCR機能では、テキスト認識精度が業務要件を満たさなかった
   - 特に数字や記号の誤認識が多発

2. **横向きPDFの対応**
   - GASでは横向き（ランドスケープ）PDFの右側のテキストが正しく取得できない問題があった
   - Power Queryは向きに関係なく正確にテキストを抽出可能

3. **高速処理**
   - Power QueryはExcelのネイティブ機能として最適化されており、処理が高速

### ⚠️ 重要な制限事項

**Power Queryは「テキスト埋め込み型PDF」のみ対応**

- ✅ **対応**: デジタル生成されたPDF（Word、Excel等から出力）
- ❌ **非対応**: スキャンした画像PDF、手書き文書をPDF化したもの

#### 対応判別方法

PDFファイルを開いて、テキストを選択できるか確認：
- テキスト選択可能 → ✅ このシステムで処理可能
- テキスト選択不可 → ❌ OCR処理が必要（別途対応が必要）

## 📁 ディレクトリ構造
```
business-automation-workflow/
├── README.md
├── LICENSE
├── .gitignore
├── step1_gmail_pdf_extraction/
│   ├── README.md
│   └── code.gs
├── step2_pdf_to_excel/
│   ├── README.md
│   ├── ExportPdfDataToExcel.bas
│   └── LIMITATIONS.md
├── step3_data_aggregation/
│   ├── README.md
│   └── transferExcelDataToSheet.gs
├── step4_data_integration/
│   ├── README.md
│   ├── integrate_data.ipynb
│   └── requirements.txt
└── docs/
    ├── image
```

## 🚀 クイックスタート

### 前提条件

- Google Workspace アカウント
- Microsoft Excel (Power Query対応版)
- Google Colab アカウント
- **処理対象のPDFがテキスト埋め込み型であること** ⚠️

### セットアップ

各ステップの詳細なセットアップ方法は、各ディレクトリのREADME.mdを参照してください：

1. [Step 1: Gmail PDF自動保存](./step1_gmail_pdf_extraction/)
2. [Step 2: PDF→Excelデータ化](./step2_pdf_to_excel/)
3. [Step 3: データ集約](./step3_data_aggregation/)
4. [Step 4: データ統合・照合](./step4_data_integration/)

## 🛠️ 技術スタック

- **Google Apps Script** - メール処理、ファイル操作
- **VBA** - PDF→Excelデータ変換の制御
- **Power Query** - PDFテキスト抽出（テキスト埋め込み型のみ対応）
- **Python 3.8+** - データ統合と分析
  - pandas
  - gspread
  - xlsxwriter
- **Google Colab** - 実行環境

## ⚠️ 重要な注意事項

### PDF処理の制限

- ⚠️ **テキスト埋め込み型PDFのみ対応**: 
  - このシステムはPower Queryを使用するため、画像PDFやスキャン文書は処理できません
  - PDFを開いてテキスト選択ができることを事前に確認してください

- 📄 **横向きPDF対応**:
  - Power Queryは横向き（ランドスケープ）PDFも正しく処理可能
  - GAS OCRでは横向きPDFの右側が欠落する問題がありましたが、Power Queryでは解決済み

## 📊 使用例

このシステムは以下のような業務に適用できます：

- 注文書・請求書の自動処理
- 修理依頼書の管理
- 定期レポートの集約
- 複数システム間のデータ照合

## 🤝 コントリビューション

改善提案やバグ報告は Issue でお願いします。

## 📝 ライセンス

MIT License - 詳細は [LICENSE](./LICENSE) ファイルを参照してください。

## 📧 お問い合わせ

質問や相談がある場合は Issue を作成してください。

---

**作成者**: [@yasumorishima](https://github.com/yasumorishima)  
**作成日**: 2025年12月  
**対象ユーザー**: 業務自動化を検討している方、データ処理の効率化を目指す方
