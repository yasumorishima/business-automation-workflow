# \# 業務データ自動処理・統合システム

# 

# メール添付PDFの自動収集からExcelデータ化、複数データソースの統合まで、一連の業務プロセスを自動化するシステムです。

# 

# \## 📋 プロジェクト概要

# 

# このプロジェクトは、以下の業務フローを自動化します：

# 

# 1\. \*\*Gmail添付PDFの自動保存\*\* (Google Apps Script)

# 2\. \*\*PDFからExcelへのデータ抽出\*\* (VBA + Power Query)

# 3\. \*\*Excelデータのスプレッドシートへの集約\*\* (Google Apps Script)

# 4\. \*\*複数データソースの統合と照合\*\* (Python/Google Colab)

# 

# \### 主な特徴

# 

# \- ✅ メール受信からデータ統合まで完全自動化

# \- ✅ 重複データの自動検知と排除

# \- ✅ 複数データソースの管理番号による照合

# \- ✅ スタイル付き統合レポートの自動生成

# \- ✅ エラーハンドリングとログ記録

# 

# \## 🔄 全体フロー

# ```

# ┌─────────────────────────────────────────────────────────────────┐

# │                        業務データ自動処理フロー                          │

# └─────────────────────────────────────────────────────────────────┘

# 

# ① PDF自動保存 (GAS)

# &nbsp;  Gmail → PDFファイル抽出 → Google Drive保存

# &nbsp;         ↓

# ② PDF→Excelデータ化 (VBA + Power Query)

# &nbsp;  PDF → Power Query → テキスト抽出 → Excel形式保存

# &nbsp;         ↓

# ③ データ集約 (GAS)

# &nbsp;  複数Excelファイル → CSV読み込み → スプレッドシート統合

# &nbsp;         ↓

# ④ データ統合・照合 (Python)

# &nbsp;  3つのデータソース → 管理番号で統合 → 照合用Excel出力

# ```

# 

# \## 🔍 技術選定の背景

# 

# \### なぜPower Queryを使用するのか？

# 

# ステップ②でPower Queryを採用した理由：

# 

# 1\. \*\*GAS OCRの精度不足\*\*

# &nbsp;  - Google Apps ScriptのOCR機能では、テキスト認識精度が業務要件を満たさなかった

# &nbsp;  - 特に数字や記号の誤認識が多発

# 

# 2\. \*\*横向きPDFの対応\*\*

# &nbsp;  - GASでは横向き（ランドスケープ）PDFの右側のテキストが正しく取得できない問題があった

# &nbsp;  - Power Queryは向きに関係なく正確にテキストを抽出可能

# 

# 3\. \*\*高速処理\*\*

# &nbsp;  - Power QueryはExcelのネイティブ機能として最適化されており、処理が高速

# 

# \### ⚠️ 重要な制限事項

# 

# \*\*Power Queryは「テキスト埋め込み型PDF」のみ対応\*\*

# 

# \- ✅ \*\*対応\*\*: デジタル生成されたPDF（Word、Excel等から出力）

# \- ❌ \*\*非対応\*\*: スキャンした画像PDF、手書き文書をPDF化したもの

# 

# \#### 対応判別方法

# 

# PDFファイルを開いて、テキストを選択できるか確認：

# \- テキスト選択可能 → ✅ このシステムで処理可能

# \- テキスト選択不可 → ❌ OCR処理が必要（別途対応が必要）

# 

# \#### 画像PDFの場合の対処法

# 

# もし画像PDFを処理する必要がある場合：

# 

# 1\. \*\*事前にOCR処理を実施\*\*

# &nbsp;  - Adobe Acrobat Pro のOCR機能

# &nbsp;  - Google Drive のプレビュー機能

# &nbsp;  - 専用OCRツール（ABBYY FineReader等）

# 

# 2\. \*\*代替ツールの検討\*\*

# &nbsp;  - Python + pytesseract

# &nbsp;  - Google Cloud Vision API

# &nbsp;  - AWS Textract

# 

# \## 📁 ディレクトリ構造

# ```

# automation-workflow/

# ├── README.md

# ├── step1\_gmail\_pdf\_extraction/

# │   ├── README.md

# │   └── code.gs

# ├── step2\_pdf\_to\_excel/

# │   ├── README.md

# │   ├── ExportPdfDataToExcel.bas

# │   └── LIMITATIONS.md  # PDF対応形式の詳細説明

# ├── step3\_data\_aggregation/

# │   ├── README.md

# │   └── transferExcelDataToSheet.gs

# ├── step4\_data\_integration/

# │   ├── README.md

# │   ├── integrate\_data.ipynb

# │   └── requirements.txt

# └── docs/

# &nbsp;   ├── setup\_guide.md

# &nbsp;   ├── flow\_diagram.png

# &nbsp;   └── pdf\_compatibility.md  # PDF互換性ガイド

# ```

# 

# \## 🚀 セットアップ

# 

# \### 前提条件

# 

# \- Google Workspace アカウント

# \- Microsoft Excel (Power Query対応版)

# \- Google Colab アカウント

# \- \*\*処理対象のPDFがテキスト埋め込み型であること\*\* ⚠️

# 

# \### ステップ1: Gmail PDF自動保存

# 

# 1\. Google Apps Scriptプロジェクトを作成

# 2\. `step1\_gmail\_pdf\_extraction/code.gs` をコピー

# 3\. 設定項目を編集：

# ```javascript

# &nbsp;  const CONFIG\_TARGET\_FROM\_EMAIL = "sender@example.com";

# &nbsp;  const CONFIG\_ATTACHMENT\_DRIVE\_FOLDER\_ID = "YOUR\_FOLDER\_ID";

# ```

# 4\. 時間主導型トリガーを設定（推奨: 1時間ごと）

# 

# \### ステップ2: PDF→Excelデータ化

# 

# \*\*⚠️ 事前確認: PDFがテキスト選択可能か確認してください\*\*

# 

# 1\. Excelで新しいブックを作成

# 2\. Power Queryエディタを開く（データ → データの取得）

# 3\. Google DriveのPDFフォルダに接続

# 4\. VBAエディタを開く（Alt + F11）

# 5\. `step2\_pdf\_to\_excel/ExportPdfDataToExcel.bas` をインポート

# 6\. 設定項目を編集：

# ```vba

# &nbsp;  Const OUTPUT\_FOLDER\_PATH As String = "C:\\Your\\Output\\Path\\"

# &nbsp;  Const TARGET\_SHEET\_NAME As String = "CombinedPDFData"

# ```

# 

# \#### Power Query設定のポイント

# 

# \- データソース: PDF

# \- 変換: テーブルに変換

# \- 区切り記号: タブ、改行を認識

# \- データ型: テキストのまま保持

# 

# \### ステップ3: データ集約

# 

# 1\. Google スプレッドシートを作成

# 2\. スクリプトエディタから `step3\_data\_aggregation/transferExcelDataToSheet.gs` をコピー

# 3\. 設定項目を編集：

# ```javascript

# &nbsp;  const CONFIG\_EXCEL\_DRIVE\_FOLDER\_ID = "YOUR\_FOLDER\_ID";

# &nbsp;  const CONFIG\_TARGET\_SPREADSHEET\_ID = "YOUR\_SPREADSHEET\_ID";

# ```

# 4\. カスタムメニューから手動実行またはトリガー設定

# 

# \### ステップ4: データ統合・照合

# 

# 1\. Google Colabで `step4\_data\_integration/integrate\_data.ipynb` を開く

# 2\. 必要なライブラリをインストール：

# ```python

# &nbsp;  !pip install xlsxwriter

# ```

# 3\. スプレッドシートURLを設定：

# ```python

# &nbsp;  SPREADSHEET\_URL = 'YOUR\_SPREADSHEET\_URL\_HERE'

# ```

# 4\. ファイルをアップロードして実行

# 

# \## 💡 使用方法

# 

# \### 日次運用

# 

# 1\. \*\*自動実行\*\*: ステップ①は時間トリガーで自動実行

# 2\. \*\*手動実行\*\*: ステップ②③を必要に応じて実行

# 3\. \*\*統合処理\*\*: ステップ④で月次/週次統合レポート生成

# 

# \### データの確認

# 

# \- \*\*PDF保存状況\*\*: Google Driveフォルダを確認

# \- \*\*Excel変換状況\*\*: 出力フォルダを確認

# \- \*\*スプレッドシート\*\*: Google スプレッドシートで明細確認

# \- \*\*統合データ\*\*: ダウンロードされたExcelファイルを確認

# 

# \## 🛠️ 技術スタック

# 

# \- \*\*Google Apps Script\*\* - メール処理、ファイル操作

# \- \*\*VBA\*\* - PDF→Excelデータ変換の制御

# \- \*\*Power Query\*\* - PDFテキスト抽出（テキスト埋め込み型のみ対応）

# \- \*\*Python\*\* - データ統合と分析

# &nbsp; - pandas

# &nbsp; - gspread

# &nbsp; - xlsxwriter

# \- \*\*Google Colab\*\* - 実行環境

# 

# \## ⚙️ カスタマイズ

# 

# \### ファイル命名規則の変更

# 

# `step1\_gmail\_pdf\_extraction/code.gs`:

# ```javascript

# const CONFIG\_ATTACHMENT\_PREFIX = "カスタム\_プレフィックス\_";

# ```

# 

# \### ヘッダー名の変更

# 

# `step3\_data\_aggregation/transferExcelDataToSheet.gs`:

# ```javascript

# const HEADERS = \['No.', 'ファイル名', 'カスタム列1', 'カスタム列2', ...];

# ```

# 

# \### 統合キーの変更

# 

# `step4\_data\_integration/integrate\_data.ipynb`:

# ```python

# df\_merged = pd.merge(df\_a, df\_b, on='カスタムキー', how='left')

# ```

# 

# \## 📊 出力サンプル

# 

# \### スプレッドシート（ステップ3）

# 

# | No. | ファイル名 | 発注日 | 発注番号 | 管理番号 | 得意先 | ... |

# |-----|-----------|--------|---------|---------|--------|-----|

# | 1   | file1.pdf | 2025-01-15 | ORD-001 | MNG-12345 | 顧客A | ... |

# 

# \### 統合Excel（ステップ4）

# 

# \- マルチレベルヘッダー（データソース別に色分け）

# \- 自動列幅調整

# \- ウィンドウ枠固定

# \- 日付形式の統一

# 

# \## ⚠️ 注意事項

# 

# \### PDF処理の制限

# 

# \- ⚠️ \*\*テキスト埋め込み型PDFのみ対応\*\*: 

# &nbsp; - このシステムはPower Queryを使用するため、画像PDFやスキャン文書は処理できません

# &nbsp; - PDFを開いてテキスト選択ができることを事前に確認してください

# &nbsp; - 画像PDFの場合は、事前にOCR処理を行うか、別のツールを検討してください

# 

# \- 📄 \*\*横向きPDF対応\*\*:

# &nbsp; - Power Queryは横向き（ランドスケープ）PDFも正しく処理可能

# &nbsp; - GAS OCRでは横向きPDFの右側が欠落する問題がありましたが、Power Queryでは解決済み

# 

# \### セキュリティ

# 

# \- ⚠️ \*\*機密情報の取り扱い\*\*: 

# &nbsp; - Google Drive のフォルダIDやスプレッドシートIDは環境変数で管理推奨

# &nbsp; - メールアドレスやファイル名に個人情報が含まれる場合は適切に保護

# 

# \### パフォーマンス

# 

# \- 📊 \*\*大量データ処理\*\*: 

# &nbsp; - ステップ②: Power Queryは1回の処理で50-100ファイル程度を推奨

# &nbsp; - ステップ③: 100ファイル以上の場合はバッチ処理を検討

# &nbsp; - ステップ④: メモリ使用量に注意（Colabの制限: 12GB）

# 

# \### エラー対応

# 

# \- 🔍 \*\*ログの確認\*\*: 

# &nbsp; - GAS: 実行ログを確認（表示 → ログ）

# &nbsp; - VBA: イミディエイトウィンドウを確認

# &nbsp; - Python: セル出力を確認

# 

# \## 🔧 トラブルシューティング

# 

# \### PDFからテキストが抽出できない

# 

# \*\*症状\*\*: Power Queryで処理しても空白や文字化けが発生

# 

# \*\*原因と対処法\*\*:

# 1\. \*\*画像PDFの可能性\*\* → PDFを開いてテキスト選択を試す

# &nbsp;  - 選択不可 → OCR処理を実施

# 2\. \*\*暗号化PDF\*\* → パスワード保護を解除

# 3\. \*\*破損PDF\*\* → 元ファイルを再取得

# 

# \### 横向きPDFで一部のテキストが欠ける

# 

# \*\*Power Query使用時\*\*: この問題は発生しません（GAS OCRの既知の問題を回避済み）

# 

# \## 🤝 貢献

# 

# 改善提案やバグ報告は Issue でお願いします。

# 

# \## 📝 ライセンス

# 

# MIT License - 自由に使用・改変できますが、自己責任でお願いします。

# 

# \## 📧 お問い合わせ

# 

# 質問や相談がある場合は Issue を作成してください。

# 

# ---

# 

# \*\*作成日\*\*: 2025年12月  

# \*\*バージョン\*\*: 1.0.0  

# \*\*対象ユーザー\*\*: 業務自動化を検討している方、データ処理の効率化を目指す方  

# \*\*重要な前提\*\*: テキスト埋め込み型PDFの処理に特化したシステムです

