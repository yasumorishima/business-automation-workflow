# Step 2: PDF→Excelデータ化

Power Query と VBA を使用して、PDFファイルからテキストを抽出し、Excel形式で保存します。

## ⚠️ 重要な前提条件

**このステップは「テキスト埋め込み型PDF」のみ対応しています。**

- ✅ **対応**: デジタル生成されたPDF（Word、Excelから出力）
- ❌ **非対応**: スキャンした画像PDF、手書き文書のPDF

**確認方法**: PDFを開いてテキストをマウスで選択できるかチェック

詳細は [LIMITATIONS.md](./LIMITATIONS.md) を参照してください。

## 機能

- ✅ Power Queryによる高精度テキスト抽出
- ✅ 横向き（ランドスケープ）PDFにも対応
- ✅ 複数PDFの一括処理
- ✅ 重複ファイルの自動スキップ
- ✅ テキストデータの構造化（タブ・改行区切りを認識）

## セットアップ

### 1. Power Query の設定

1. Excelで新しいブックを作成
2. 「データ」タブ → 「データの取得」→ 「ファイルから」→ 「フォルダーから」
3. Google Driveの PDF保存フォルダを選択
4. 「データの変換」をクリック
5. Power Query エディタが開く

#### Power Query の変換ステップ
```m
let
    ソース = Folder.Files("PDFフォルダのパス"),
    フィルター済み = Table.SelectRows(ソース, each [Extension] = ".pdf"),
    追加されたカスタム = Table.AddColumn(フィルター済み, "PDFテキスト", 
        each Pdf.Tables([Content]){0}[Data])
in
    追加されたカスタム
```

6. 「閉じて読み込む」→ 「閉じて次に読み込む」
7. シート名を `CombinedPDFData` に設定

### 2. VBA マクロの設定

1. Alt + F11 でVBAエディタを開く
2. 「挿入」→ 「標準モジュール」
3. `ExportPdfDataToExcel.bas` の内容を貼り付け

#### 設定項目の編集
```vba
' Power Queryが出力したデータがあるシート名
Const TARGET_SHEET_NAME As String = "CombinedPDFData"

' Excel出力先フォルダのパス
Const OUTPUT_FOLDER_PATH As String = "C:\Your\Output\Path\"

' データ開始行（ヘッダーを除く）
Const START_ROW As Long = 2
```

### 3. 実行方法

1. Alt + F8 でマクロ一覧を表示
2. `ExportPdfDataToJson` を選択
3. 「実行」をクリック

または、ボタンを配置して実行：
1. 「開発」タブ → 「挿入」→ 「ボタン」
2. マクロ `ExportPdfDataToJson` を割り当て

## 処理フロー
```
PDF保存フォルダ（Google Drive）
    ↓
Power Query: PDFテキスト抽出
    ↓
Excelシート: CombinedPDFData
    ↓
VBA: テキストを行・列に分割
    ↓
Excel出力フォルダ: 個別の .xlsx ファイル
```

## 出力形式

各PDFから抽出されたデータは、個別のExcelファイル（.xlsx）として保存されます。

**ファイル名**: `元のファイル名.xlsx`  
**内容**: タブ・改行で分割されたテーブルデータ

## トラブルシューティング

### テキストが抽出できない

**原因1: 画像PDF**
- 対処: PDFを開いてテキスト選択を試す → できない場合はOCR処理が必要

**原因2: 暗号化PDF**
- 対処: パスワード保護を解除してから処理

**原因3: 破損したPDF**
- 対処: 元ファイルを再取得

### Power Query が動作しない

- Excel のバージョンを確認（2016以降推奨）
- Power Query アドインが有効か確認

### VBA実行エラー

- マクロのセキュリティ設定を確認
- 出力フォルダのパスが正しいか確認
- フォルダの書き込み権限があるか確認

## 次のステップ

出力されたExcelファイルは、[Step 3](../step3_data_aggregation/) でスプレッドシートに集約されます。