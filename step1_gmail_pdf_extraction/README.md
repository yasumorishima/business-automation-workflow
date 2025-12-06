# Step 1: Gmail PDF自動保存

Gmail の特定送信元からのメールに添付されたPDFファイルを自動的にGoogle Driveに保存します。

## 機能

- ✅ 定期的なメール検索とPDF抽出
- ✅ 重複ファイルの自動検知（既存ファイルはスキップ）
- ✅ Google Driveへの自動保存
- ✅ 実行ログの記録
- ✅ ファイル名・送信元のフィルタリング

## セットアップ

### 1. Google Apps Script プロジェクトを作成

1. [Google Apps Script](https://script.google.com/) にアクセス
2. 「新しいプロジェクト」をクリック
3. プロジェクト名を設定（例：`Gmail_PDF_Extractor`）

### 2. コードを貼り付け

`code.gs` の内容をコピーして貼り付けます。

### 3. 設定項目を編集
```javascript
// 処理対象のメール設定
const CONFIG_TARGET_FROM_EMAIL = "sender@example.com"; // 送信元メールアドレス

// Google Driveの保存先フォルダID
const CONFIG_ATTACHMENT_DRIVE_FOLDER_ID = "YOUR_DRIVE_FOLDER_ID_HERE";

// ファイル名のプレフィックス（例：「注文書_」で始まるファイルのみ処理）
const CONFIG_ATTACHMENT_PREFIX = "注文書_";
```

**フォルダIDの取得方法:**
1. Google Driveで保存先フォルダを開く
2. URLの最後の部分がフォルダID
```
   https://drive.google.com/drive/folders/【ここがフォルダID】
```

### 4. トリガーを設定

1. スクリプトエディタで「トリガー」アイコン（時計マーク）をクリック
2. 「トリガーを追加」をクリック
3. 以下を設定：
   - 実行する関数: `processXingReports`
   - イベントのソース: `時間主導型`
   - 時間ベースのトリガー: `時間ベースのタイマー`
   - 時間の間隔: `1時間おき`（推奨）

### 5. 初回実行と権限の承認

1. 関数を選択して「実行」ボタンをクリック
2. 権限の承認を求められるので、承認
3. ログを確認して正常に動作していることを確認

## 設定のカスタマイズ

### メール検索期間
```javascript
const CONFIG_SEARCH_LOOK_BACK_MONTHS = 1; // 何ヶ月前まで検索するか
```

### 既読設定
```javascript
const CONFIG_MARK_EMAIL_READ = false; // 処理後にメールを既読にするか
```

## トラブルシューティング

### PDFが保存されない

- フォルダIDが正しいか確認
- メールアドレスが正しいか確認
- ファイル名のプレフィックスが合っているか確認

### 権限エラー

- スクリプトの権限を再承認
- Google Driveへのアクセス権限があるか確認

## ログの確認方法

1. スクリプトエディタで「表示」→「ログ」
2. 実行履歴とエラーを確認

## 次のステップ

保存されたPDFファイルは、[Step 2](../step2_pdf_to_excel/) でExcelデータに変換されます。