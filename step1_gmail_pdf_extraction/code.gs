// ========================================================================
// ★★★ 設定項目ここから ★★★
// この部分の設定を変更することで、スクリプトの動作を調整できます。
// ========================================================================

// 処理対象のメール設定
const CONFIG_TARGET_FROM_EMAIL = "sender@example.com"; // 送信元メールアドレス 例: @company.com に変える
// 宛先メールアドレスは指定なし(全ての宛先を対象)

// Googleドライブの添付ファイル保存先フォルダID
// ↓共有ドライブ\SO_GAS(GAS_注文書_Xing)データ_PDF に対応します。
const CONFIG_ATTACHMENT_DRIVE_FOLDER_ID = "YOUR_DRIVE_FOLDER_ID_HERE";

// 添付ファイルの命名規則(保存対象とするファイルのパターン)
// ファイル名がこの拡張子に合致するPDFファイル(.pdf)を対象とします。
const CONFIG_ATTACHMENT_EXTENSION = ".pdf";

// 添付ファイル名の必須プレフィックス(例:「注文書_修理発注書_25.10.2-2.pdf」)
const CONFIG_ATTACHMENT_PREFIX = "注文書_修理発注書_";

// 処理済みのメールを既読にするか
// true: 処理完了したメールを自動的に既読にします。
// false: メールを未読のまま残します。
const CONFIG_MARK_EMAIL_READ = false; // 今回の要件に合わせてfalseに設定

// メール検索期間の設定
const CONFIG_SEARCH_LOOK_BACK_MONTHS = 1; // 現在から何ヶ月前までメールを検索するか(例: 6ヶ月前まで)
const CONFIG_SEARCH_END_OFFSET_DAYS = 1; // 検索終了時刻の前日の何時何分何秒の何分まるか(例: 1で実行日の当日0時0分0秒まで)

// ========================================================================
// ★★★ 設定項目ここまで ★★★
// ========================================================================

/**
 * 特定の条件に合致するメール(@example.comからのPDF添付ファイル)をGoogleドライブに自動保存します。
 * この関数は時間主導型トリガーで実行されることを想定しています。
 */
function processXingReports() {
  Logger.log('--- スクリプト開始 (XingからのPDFレポート処理とファイル保存) ---');
  
  // 現在時刻を取得
  const now = new Date();
  
  // メール検索の開始時刻を計算
  const searchStart = new Date(now);
  searchStart.setMonth(now.getMonth() - CONFIG_SEARCH_LOOK_BACK_MONTHS);
  
  // 検索終了時刻を計算
  const searchEndTime = new Date(now.getFullYear(), now.getMonth(), now.getDate() + CONFIG_SEARCH_END_OFFSET_DAYS, 0, 0, 0);
  
  Logger.log(`検索期間: ${searchStart.toLocaleDateString()} から ${searchEndTime.toLocaleDateString()} まで`);
  
  // Gmail検索クエリを構築
  // 送信元と添付ファイルがあるものなどを検索します。特定によるフィルタリングは行いません。
  let searchQuery = `from:${CONFIG_TARGET_FROM_EMAIL} has:attachment`;
  
  // Gmail検索クエリ用の {YYYY/MM/DD} 形式の文字列に変換します。
  const formatDateForGmail = (date) => {
    const year = date.getFullYear();
    const month = (date.getMonth() + 1).toString().padStart(2, '0'); // 月は0から始まるため+1
    const day = date.getDate().toString().padStart(2, '0');
    return `${year}/${month}/${day}`;
  };
  
  searchQuery += ` after:${formatDateForGmail(searchStart)} before:${formatDateForGmail(searchEndTime)}`;
  
  Logger.log(`検索クエリ: ` + searchQuery); // ログに検索クエリを出力
  
  // Googleドライブの添付ファイル保存先フォルダを取得
  const attachmentFolder = DriveApp.getFolderById(CONFIG_ATTACHMENT_DRIVE_FOLDER_ID);
  if (!attachmentFolder) {
    Logger.log(`エラー: 添付ファイル保存先Googleドライブフォルダが見つかりません。ID: ` + CONFIG_ATTACHMENT_DRIVE_FOLDER_ID);
    return; // フォルダが見つからない場合は処理を中断
  }
  
  Logger.log(`添付ファイル保存先フォルダ: ${attachmentFolder.getName()} (ID: ${CONFIG_ATTACHMENT_DRIVE_FOLDER_ID})`);
  
  // Gmailを検索
  const threads = GmailApp.search(searchQuery);
  Logger.log(`検索されたスレッド数: ${threads.length}`);
  
  // 処理済みのメールを追跡するためのセット(重複処理防止用)
  const processedMessageIds = new Set();
  
  threads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(message => {
      
      // 同じメッセージを複数回処理しないようにチェック
      if (processedMessageIds.has(message.getId())) {
        return;
      }
      
      processedMessageIds.add(message.getId());
      
      const subject = message.getSubject();
      Logger.log(`処理中の件名: "${subject}"`);
      
      // 添付ファイルを処理
      const attachments = message.getAttachments();
      if (attachments.length === 0) {
        Logger.log(`スキップ: 添付ファイルがありません。件名: "${subject}"`);
        return;
      }
      
      let pdfFileSavedForThisMessage = false; // このメッセージでPDF添付ファイルが正常に保存されたかを示すフラグ
      
      attachments.forEach(attachment => {
        const attachmentFileName = attachment.getName();
        const lowerCaseFileName = attachmentFileName.toLowerCase();
        
        // 添付ファイルがPDFファイル(.pdf)か、指定されたプレフィックスで始まるかチェック
        const isTargetPdf = lowerCaseFileName.endsWith(CONFIG_ATTACHMENT_EXTENSION) &&
                            attachmentFileName.startsWith(CONFIG_ATTACHMENT_PREFIX);
        
        if (isTargetPdf) {
          Logger.log(`対象のPDF添付ファイルが見つかりました: "${attachmentFileName}"`);
          
          // Googleドライブに同名のファイルが既に存在するかチェック(元のPDFファイル名でチェック)
          const existingFiles = attachmentFolder.getFilesByName(attachmentFileName);
          if (existingFiles.hasNext()) {
            Logger.log(`スキップ: 同名のPDFファイルがGoogleドライブに既に存在します: "${attachmentFileName}" (件名: "${subject}")`);
            pdfFileSavedForThisMessage = true; // 処理済みフラグを立てる(重複によるスキップも成功と見なす)
            return;
          }
          
          try {
            // 添付ファイル(PDFモデル)をGoogleドライブにそのまま保存
            const savedFile = attachmentFolder.createFile(attachment);
            Logger.log(`添付ファイルをPDFファイルとして保存しました: "${savedFile.getName()}" (ID: ${savedFile.getId()}) (件名: "${subject}")`);
            pdfFileSavedForThisMessage = true; // ファイル保存が成功
          } catch (e) {
            Logger.log(`エラー発生(ファイル名: "${attachmentFileName}", 件名: "${subject}"): ${e.message}`);
          }
        } else {
          Logger.log(`スキップ: 対象のPDFパターンに一致しません。ファイル名: "${attachmentFileName}" (件名: "${subject}")`);
        }
      }); // attachments.forEach
      
      // このメールから処理対象のPDF添付ファイルが一つでも正常に保存された場合
      if (pdfFileSavedForThisMessage) {
        if (CONFIG_MARK_EMAIL_READ) {
          message.markRead();
          Logger.log(`メールを既読にしました: "${subject}"`);
        } else {
          Logger.log(`このメールからは処理対象のPDF添付ファイルが見つからないか、既に保存済でした: "${subject}"`);
        }
      }
      
    }); // messages.forEach
  }); // threads.forEach
  
  Logger.log('--- すべての処理が完了しました ---');
}