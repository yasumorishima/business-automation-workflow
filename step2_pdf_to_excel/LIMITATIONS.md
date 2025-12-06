\# PDF処理の制限事項



\## テキスト埋め込み型PDFのみ対応



このシステムは、\*\*テキスト情報が埋め込まれたPDF\*\*のみを処理できます。



\### ✅ 対応するPDF



\- Microsoft Word から出力したPDF

\- Excel から出力したPDF

\- PowerPoint から出力したPDF

\- Webページを「PDFとして保存」したもの

\- プログラムで生成されたPDF（請求書システム等）



\### ❌ 対応しないPDF



\- スキャナーで読み取った画像PDF

\- 写真や画像をPDF化したもの

\- 手書き文書をスキャンしたPDF

\- FAXをPDF化したもの



\## 確認方法



\### 方法1: テキスト選択テスト



1\. PDFファイルを Adobe Reader や ブラウザで開く

2\. マウスでテキストを選択してみる

3\. 結果：

&nbsp;  - ✅ \*\*選択できる\*\* → このシステムで処理可能

&nbsp;  - ❌ \*\*選択できない\*\* → 画像PDFのため処理不可



\### 方法2: ファイルサイズの確認



\*\*目安\*\*:

\- テキスト埋め込み型: 通常 50KB - 500KB

\- 画像PDF: 通常 1MB 以上



※ あくまで目安です。確実な確認は方法1で行ってください。



\## 画像PDFを処理したい場合



以下の方法でテキスト埋め込み型PDFに変換してください：



\### オプション1: Adobe Acrobat Pro



1\. Adobe Acrobat Pro で PDFを開く

2\. 「ツール」→ 「テキスト認識」→ 「このファイル内」

3\. OCR処理後、別名で保存



\### オプション2: Google Drive



1\. Google Drive にPDFをアップロード

2\. 右クリック → 「アプリで開く」→ 「Google ドキュメント」

3\. 「ファイル」→ 「ダウンロード」→ 「PDFドキュメント」



\### オプション3: オンラインOCRツール



\- \[OCR.space](https://ocr.space/)

\- \[OnlineOCR.net](https://www.onlineocr.net/)



\### オプション4: Python + Tesseract（技術者向け）

```python

from pdf2image import convert\_from\_path

import pytesseract



\# PDFを画像に変換

images = convert\_from\_path('input.pdf')



\# OCR実行

text = pytesseract.image\_to\_string(images\[0], lang='jpn')

```



\## 技術的背景



\### なぜPower Queryを採用したのか



1\. \*\*GAS OCRの精度不足\*\*

&nbsp;  - Google Apps ScriptのOCR機能は画像PDFに対応しているが、精度が低い

&nbsp;  - 数字や記号の誤認識が多発

&nbsp;  - 業務データとして使用するには信頼性が不足



2\. \*\*横向きPDFの問題\*\*

&nbsp;  - GAS OCRでは横向き（ランドスケープ）PDFの右側が正しく読み取れない

&nbsp;  - Power Queryはテキストデータを直接読むため、向きに依存しない



3\. \*\*処理速度\*\*

&nbsp;  - OCR処理: 1ページあたり5-10秒

&nbsp;  - Power Query: 1ページあたり1秒未満



\### Power Query の仕組み



Power Query は PDFから\*\*既に埋め込まれているテキストデータ\*\*を抽出します。

```

テキスト埋め込み型PDF

&nbsp;   ├─ 表示用（レンダリング情報）

&nbsp;   └─ テキストデータ ← Power Queryはここを読む

```



画像PDFにはテキストデータが存在しないため、Power Queryでは処理できません。

```

画像PDF

&nbsp;   └─ 画像データのみ ← テキストデータが存在しない

```



\## まとめ



| 項目 | テキスト埋め込み型 | 画像PDF |

|------|------------------|---------|

| テキスト選択 | ✅ 可能 | ❌ 不可 |

| Power Query処理 | ✅ 対応 | ❌ 非対応 |

| 処理速度 | ✅ 高速 | - |

| 精度 | ✅ 100% | - |

| 事前処理 | 不要 | OCR必須 |



\*\*推奨\*\*: 送信元にテキスト埋め込み型PDFでの提供を依頼することが最も確実です。

