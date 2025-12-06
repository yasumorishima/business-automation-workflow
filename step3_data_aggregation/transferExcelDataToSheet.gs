// ===============================================================================
// GASデータ転記関数（2対応・制限解除版）: SpreadsheetDataTransfer_Fixed.js
// 目的: Googleドライブフォルダ内のExcelファイル(.xlsx)をCSV形式で読み込み、
//      明細データを抽出し、スプレッドシートに転記する。
// ===============================================================================

// 修正点:
// 1. 発注番号: 「RH」始まりの制限を撤廃し、汎用的な数字に対応。
// 2. 管理番号(New): 「No.」と連結している場合、頭文字が「Z」など(BCRA以外)でも
//    拾えるように正規表現を [A-Z] に修正。
// ===============================================================================

/**
 * スプレッドシートを開いたときに実行される関数
 * メニューバーにカスタムメニューを追加します
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Excel転記メニュー') // メニューバーに表示される名前
    .addItem('データ転記を実行', 'transferExcelDataToSheet') // アイテム名と実行する関数名
    .addToUi();
}

// ★★★ 設定項目ここから ★★★
const CONFIG_EXCEL_DRIVE_FOLDER_ID = "YOUR_EXCEL_FOLDER_ID_HERE"; // フォルダID
const CONFIG_EXCEL_EXTENSION = ".xlsx";
const CONFIG_TARGET_SPREADSHEET_ID = "YOUR_SPREADSHEET_ID_HERE"; // 転記先SSID
const CONFIG_TARGET_SHEET_NAME = "Data"; // 転記先シート名
const CONFIG_DATA_START_ROW = 2; // スプレッドシートのデータ開始行
const CONFIG_FILENAME_COLUMN_INDEX = 2; // ファイル名を記録する列番号(B列=2)
// ★★★ 設定項目ここまで ★★★

// 修理商品コードの基本パターン（JR-PXXXX、JR-S50 など）
const PRODUCT_CODE_PATTERN_SOURCE = "PROD-[0-9A-Z\\-]{3,7}"; // PROD- の後に続く、数字、アルファベット、ハイフンを許容

// ★★★ 修正箇所: 管理番号の頭文字制限を撤廃（BCRA => A-Z）★★★
// No.と管理番号の連結パターン（例: 15 Z20B1266）
const NO_MANAGEMENT_PATTERN = /^(\d{1,2})\s+([A-Z]{1,2}\d{6,})$/;

/**
 * ヘッダー名、ファイル名を正規化するヘルパー関数
 */
const normalizeString = (str) => {
  return (str || '').toString().replace(/[\s\t\r\n]+/g, ' ').trim();
};

/**
 * Googleドライブの指定フォルダにあるExcelファイル(.xlsx)を読み込み、
 * その内容をスプレッドシートに転記します。
 */
function transferExcelDataToSheet() {
  Logger.log('--- Excelデータ転記スクリプト開始（2対応・汎用化版）---');
  
  // 転記先スプレッドシートの設定
  const targetSpreadsheet = SpreadsheetApp.openById(CONFIG_TARGET_SPREADSHEET_ID);
  const targetSheet = targetSpreadsheet.getSheetByName(CONFIG_TARGET_SHEET_NAME);
  if (!targetSheet) {
    Logger.log(`エラー: シート '${CONFIG_TARGET_SHEET_NAME}' が見つかりません。`);
    return;
  }
  
  // ヘッダー行の初期化（初回のみ）
  const HEADERS = ['No.', 'ファイル名', '発注日', '発注番号', '管理番号', 'お問合番号', '得意先', '修理商品', 'シリアル番号', '相談室番号', '金額', '見積回答', '処理日時'];
  if (targetSheet.getLastRow() === 0) {
    targetSheet.appendRow(HEADERS);
  }
  
  // 転記済ファイル名の確認（6列目=ファイル名を見る）
  const lastRowInTargetSheet = targetSheet.getLastRow();
  let existingFileNamesInSheet = new Set();
  if (lastRowInTargetSheet >= CONFIG_DATA_START_ROW) {
    const fileNameRange = targetSheet.getRange(CONFIG_DATA_START_ROW, CONFIG_FILENAME_COLUMN_INDEX, lastRowInTargetSheet - CONFIG_DATA_START_ROW + 1, 1);
    const fileNames = fileNameRange.getValues();
    fileNames.forEach(row => {
      if (row[0]) {
        existingFileNamesInSheet.add(normalizeString(row[0].toString()));
      }
    });
  }
  
  const excelFolder = DriveApp.getFolderById(CONFIG_EXCEL_DRIVE_FOLDER_ID);
  const files = excelFolder.getFilesByType(MimeType.MICROSOFT_EXCEL);
  
  let processedFileCount = 0;
  let errorCount = 0;
  // Mapを使用して、管理番号ごとにデータを保存し、重複を排除（1回目優先）
  const allValuesMap = new Map();
  
  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();
    
    if (!fileName.toLowerCase().endsWith(CONFIG_EXCEL_EXTENSION)) continue;
    if (existingFileNamesInSheet.has(normalizeString(fileName))) continue;
    
    Logger.log(`処理中: ${fileName}`);
    
    try {
      // 1. ExcelファイルをCSVとしてダウンロード
      const xlsxDownloadUrl = `https://docs.google.com/spreadsheets/d/${file.getId()}/export?format=csv`;
      const responseText = UrlFetchApp.fetch(xlsxDownloadUrl, {
        headers: {
          'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
        },
        muteHttpExceptions: true
      }).getContentText();
      
      // HTMLエラーページが返された場合（権限エラーなど）
      if (responseText.startsWith('<!DOCTYPE html>')) {
        throw new Error('CSVダウンロード中にHTMLエラーページが返されました。権限設定を確認してください。');
      }
      
      const rawData = Utilities.parseCsv(responseText);
      
      // 2. ヘッダー情報の初期化
      let headerIndices = {};
      let isDataExtractionStarted = false; // データ抽出開始フラグ
      
      // 3. ファイル共通ヘッダー情報（発注日、発注番号）を抽出（確定）
      let orderDate = '';
      let orderNumber = '';
      
      // 4. 明細データの抽出（CSVの先頭から最後までループ）
      for (let i = 0; i < rawData.length; i++) {
        const row = rawData[i];
        
        // A. ヘッダー検出ロジック（動的更新）
        let currentHeaderCandidates = {};
        const targetHeaders = ['管理番号', 'お問合番号', '得意先', '修理商品', 'シリアル番号', '相談室番号', '金額', '見積回答'];
        
        let foundCandidatesCount = 0;
        let isHeaderRow = false;
        
        for (let j = 0; j < row.length; j++) {
          const cellValue = normalizeString(row[j]);
          
          // 共通ヘッダー情報の抽出（データ行より前にあると仮定）
          if (!isDataExtractionStarted) {
            const rowText = row.join('|');
            const dateMatch = rowText.match(/発注日\|[^\|]*\|[^\|]*\|(\d{4})年(\d{1,2})月(\d{1,2})日/);
            if (dateMatch) {
              orderDate = `${dateMatch[1]}-${dateMatch[2].padStart(2, '0')}-${dateMatch[3].padStart(2, '0')}`;
            }
            
            // 発注番号の汎用抽出（英数字ハイフン文字列に対応）
            const orderNumMatch = rowText.match(/発注番号\|([0-9a-zA-Z\-]{3,})/);
            if (orderNumMatch) {
              orderNumber = orderNumMatch[1];
            }
          }
          
          if (targetHeaders.includes(cellValue)) {
            currentHeaderCandidates[cellValue] = j;
            foundCandidatesCount++;
            isHeaderRow = true;
          }
        }
        
        // B. ヘッダー行の確定とインデックス更新
        if (isHeaderRow && foundCandidatesCount >= 3) {
          headerIndices = currentHeaderCandidates; // 新しいヘッダー位置を更新
          isDataExtractionStarted = true;         // これ以降はデータ行と見なす
          Logger.log(`ヘッダー位置を更新しました（行 ${i + 1}）`);
          continue; // ヘッダー行自体はスキップ
        }
        
        // C. データ行の処理（ヘッダー1度でも見つかった場合のみ）
        if (!isDataExtractionStarted) {
          continue;
        }
        
        // --- データ抽出ロジックここから ---
        
        let extractedRowNo = '';
        let managementNumber = '';
        
        // A列またはB列から No. と 管理番号を抽出
        // ★★★ ここで修正したパターンを使用 ★★★
        let noAndManMatch = (row[0] && row[0].toString().match(NO_MANAGEMENT_PATTERN)) || (row[1] && row[1].toString().match(NO_MANAGEMENT_PATTERN));
        
        if (noAndManMatch) {
          extractedRowNo = noAndManMatch[1];
          managementNumber = noAndManMatch[2];
        } else if (row[0] && row[0].toString().match(/^\d{1,2}$/)) {
          extractedRowNo = row[0].toString();
          managementNumber = row[headerIndices['管理番号']] || row[headerIndices['お問合番号']];
        } else if (headerIndices['管理番号'] !== undefined && row[headerIndices['管理番号']]) {
          managementNumber = row[headerIndices['管理番号']];
        } else {
          continue;
        }
        
        // No.の妥当性チェック（1-99）
        const parsedRowNo = parseInt(extractedRowNo, 10);
        const isNoValid = !isNaN(parsedRowNo) && parsedRowNo >= 1 && parsedRowNo <= 99;
        
        if (!isNoValid) {
          continue;
        }
        
        if (!managementNumber) continue;
        
        // 重複排除: 既に Map に存在する場合はスキップ（1回目優先）
        if (allValuesMap.has(managementNumber)) {
          continue;
        }
        
        // お問合番号、シリアル番号を取得
        const inquiryNumber = row[headerIndices['お問合番号']] || '';
        const serialNumber = row[headerIndices['シリアル番号']] || '';
        
        // 得意先名と修理商品の連結を分離
        let clientProductName = row[headerIndices['得意先']] || '';
        let productName = row[headerIndices['修理商品']] || '';
        let clientName = '';
        
        // 得意先セル内でJR-****を含む場合し、見つけたらその前後で分離する
        const broadProductPattern = new RegExp(`(.*)(\${PRODUCT_CODE_PATTERN_SOURCE})`, 'i');
        
        if (clientProductName && clientProductName.match(broadProductPattern)) {
          const match = clientProductName.match(broadProductPattern);
          if (match) {
            clientName = normalizeString(match[1]);
            productName = normalizeString(match[2]);
          }
        } else {
          clientName = clientProductName;
        }
        
        clientName = normalizeString(clientName); // 分離後のクリーンアップ
        
        // 修理商品が空の場合、分離した商品名を使用
        if (!normalizeString(productName) && headerIndices['修理商品'] !== undefined && row[headerIndices['修理商品']]) {
          productName = row[headerIndices['修理商品']];
        }
        
        // --- 金額の抽出（広域探索）---
        let amount = '';
        const amountColIndex = headerIndices['金額'];
        if (amountColIndex !== undefined) {
          // 検索範囲: 規定行(i)～i+2、金額列(amountColIndex)～amountColIndex+2
          for (let r = i; r <= Math.min(i + 2, rawData.length - 1); r++) {
            for (let c = amountColIndex; c <= Math.min(amountColIndex + 2, row.length - 1); c++) {
              let rawVal = rawData[r] && rawData[r][c] ? rawData[r][c].toString() : '';
              
              // カンマ、円記号、スペースを削除
              let cleanVal = rawVal.replace(/[,¥\s]/g, '');
              
              // 1桁以上の数字（0を含む）であるか確認
              if (cleanVal.match(/^\d+$/)) {
                amount = cleanVal;
                // 金額が見つかったら、ループの残りをスキップ
                r = rawData.length;
                break;
              }
            }
          }
        }
        
        // 相談室番号、見積回答の取得
        const consultationNumber = row[headerIndices['相談室番号']] || '';
        const estimationAnswer = row[headerIndices['見積回答']] || '';
        
        // スプレッドシートへ転記する行データを構築
        const rowData = [
          extractedRowNo,
          fileName,
          orderDate,
          orderNumber,
          managementNumber,
          inquiryNumber,
          clientName,
          productName,
          serialNumber,
          consultationNumber,
          amount,
          estimationAnswer,
          new Date()
        ];
        
        // Mapに追加（1回目優先）
        allValuesMap.set(managementNumber, rowData);
      }
      
      Logger.log(`✓ 完了: ${fileName} (${allValuesMap.size}件の明細を抽出)`);
      processedFileCount++;
      
    } catch (error) {
      Logger.log(`✕ 致命的エラー: ${fileName} のデータ処理中にエラーが発生しました - ${error}`);
      errorCount++;
    }
  }
  
  // 5. データ転記（バッチ処理）
  if (allValuesMap.size > 0) {
    // Mapの値を配列に変換して転記
    const finalValuesToTransfer = Array.from(allValuesMap.values());
    const lastRow = targetSheet.getLastRow();
    const startRow = lastRow + 1;
    
    // スプレッドシートにまとめて書き込み
    targetSheet.getRange(startRow, 1, finalValuesToTransfer.length, HEADERS.length).setValues(finalValuesToTransfer);
    
    Logger.log(`処理完了: 合計 ${allValuesMap.size}件の明細を転記しました。`);
  }
}