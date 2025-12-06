\# Step 3: データ集約（スプレッドシートへ）



複数のExcelファイルからデータを読み込み、Googleスプレッドシートに集約します。



\## 機能



\- ✅ 複数Excelファイルの一括読み込み

\- ✅ CSVとして読み込み（互換性向上）

\- ✅ 重複データの自動検知と排除

\- ✅ ヘッダーの自動認識

\- ✅ スプレッドシートへの追記（既存データを保持）

\- ✅ ファイル名の記録



\## セットアップ



\### 1. Googleスプレッドシートを作成



1\. \[Google スプレッドシート](https://sheets.google.com/) にアクセス

2\. 新しいスプレッドシートを作成

3\. シート名を `Data` に変更

4\. スプレッドシートIDをメモ

```

&nbsp;  https://docs.google.com/spreadsheets/d/【ここがスプレッドシートID】/edit

```



\### 2. Google Apps Script を設定



1\. スプレッドシートで「拡張機能」→ 「Apps Script」

2\. `transferExcelDataToSheet.gs` の内容を貼り付け

3\. プロジェクト名を設定（例：`Excel\_Data\_Aggregator`）



\### 3. 設定項目を編集

```javascript

// ExcelファイルがあるGoogle DriveフォルダのID

const CONFIG\_EXCEL\_DRIVE\_FOLDER\_ID = "YOUR\_FOLDER\_ID\_HERE";



// 転記先スプレッドシートのID

const CONFIG\_TARGET\_SPREADSHEET\_ID = "YOUR\_SPREADSHEET\_ID\_HERE";



// 転記先シート名

const CONFIG\_TARGET\_SHEET\_NAME = "Data";



// ファイル名を記録する列番号（B列=2）

const CONFIG\_FILENAME\_COLUMN\_INDEX = 2;

```



\### 4. カスタムメニューの追加



スクリプトを保存すると、スプレッドシートに「Excel転記メニュー」が追加されます。



\### 5. 初回実行



1\. スプレッドシートをリロード

2\. 「Excel転記メニュー」→ 「データ転記を実行」

3\. 権限の承認を求められるので承認

4\. 処理が完了するまで待機



\## 処理フロー

```

Excel出力フォルダ（Google Drive）

&nbsp;   ↓

GAS: ExcelファイルをCSVとして読み込み

&nbsp;   ↓

ヘッダーの自動認識

&nbsp;   ↓

データの抽出と整形

&nbsp;   ↓

重複チェック（管理番号ベース）

&nbsp;   ↓

スプレッドシート: Data シートに追記

```



\## データ形式



\### 入力（Excelファイル）



各Excelファイルには以下のような構造が想定されています：



| 列 | 内容 |

|----|------|

| A | No. |

| B | 管理番号 |

| C | 発注日 |

| D | 発注番号 |

| ... | その他の列 |



\### 出力（スプレッドシート）



| No. | ファイル名 | 発注日 | 発注番号 | 管理番号 | 得意先 | ... | 処理日時 |

|-----|-----------|--------|---------|---------|--------|-----|---------|

| 1 | file1.xlsx | 2025-01-15 | ORD-001 | MNG-001 | 顧客A | ... | 2025-01-15 10:00 |



\## カスタマイズ



\### ヘッダー名の変更



スクリプト内の `HEADERS` 配列を編集：

```javascript

const HEADERS = \[

&nbsp; 'No.', 

&nbsp; 'ファイル名', 

&nbsp; '発注日', 

&nbsp; '発注番号', 

&nbsp; '管理番号',

&nbsp; 'カスタム列1',

&nbsp; 'カスタム列2',

&nbsp; // ...

];

```



\### データ開始行の変更

```javascript

const CONFIG\_DATA\_START\_ROW = 2; // ヘッダーの次の行

```



\## トラブルシューティング



\### データが転記されない



\- ExcelフォルダIDが正しいか確認

\- スプレッドシートIDが正しいか確認

\- ファイルが `.xlsx` 形式か確認



\### ヘッダーが見つからない



\- Excelファイル内にヘッダー行があるか確認

\- ヘッダー名が `targetHeaders` 配列と一致しているか確認



\### 権限エラー



\- Google DriveとGoogle Sheetsへのアクセス権限を再承認



\## 実行頻度の推奨



\- \*\*手動実行\*\*: 週次または月次でExcelファイルがまとまったタイミング

\- \*\*トリガー実行\*\*: 毎日深夜など、定期的に実行する場合



トリガー設定方法：

1\. Apps Scriptエディタで「トリガー」アイコンをクリック

2\. 「トリガーを追加」

3\. 関数: `transferExcelDataToSheet`

4\. 時間ベース: 日次、深夜0時-1時



\## 次のステップ



集約されたスプレッドシートデータは、\[Step 4](../step4\_data\_integration/) で他のデータソースと統合されます。

