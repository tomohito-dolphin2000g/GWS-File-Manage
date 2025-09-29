/**
 * @OnlyCurrentDoc
 *
 * 指定されたGoogle Driveフォルダ内のファイル一覧を取得し、
 * このスプレッドシートにファイル名、URL、最終更新日時、オーナーのメールアドレスを書き出します。
 *
 * セットアップ方法:
 * 1. スクリプトエディタの左メニュー「プロジェクトの設定」（歯車アイコン）を開きます。
 * 2. 「スクリプト プロパティ」のセクションで、「スクリプト プロパティを追加」をクリックします。
 * 3. 以下の2つのプロパティをそれぞれ追加します。
 * - 名前: FOLDER_ID,    値: [ファイル一覧を取得したいフォルダのID]
 * - 名前: SHEET_NAME,   値: [書き出し先のスプレッドシート名（例: "ファイル一覧"）]
 */
function createFileListInSheet() {
  try {
    // --- 1. IDや設定値をスクリプトプロパティから安全に取得 ---
    const scriptProperties = PropertiesService.getScriptProperties();
    const folderId = scriptProperties.getProperty('FOLDER_ID');
    const sheetName = scriptProperties.getProperty('SHEET_NAME');

    if (!folderId || !sheetName) {
      const errorMessage = 'スクリプトプロパティ「FOLDER_ID」または「SHEET_NAME」が設定されていません。セットアップ方法を確認してください。';
      console.error(errorMessage);
      return;
    }

    const folder = DriveApp.getFolderById(folderId);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    
    // --- 2. ファイル情報を一度配列にまとめてから、一括で書き込む（高速化） ---
    const files = folder.getFiles();
    
    // ▼▼▼ 変更点1: ヘッダー行に「最終更新日時」と「オーナー」を追加 ▼▼▼
    const fileDataList = [['ファイル名', 'URL', '最終更新日時', 'オーナーのメールアドレス']]; 

    while (files.hasNext()) {
      const file = files.next();
      
      // ▼▼▼ 変更点2: 最終更新日時とオーナーのメールアドレスを取得して配列に追加 ▼▼▼
      fileDataList.push([
        file.getName(),
        file.getUrl(),
        file.getLastUpdated(), // ファイルの最終更新日時を取得
        file.getOwner().getEmail() // ファイルのオーナーのメールアドレスを取得
      ]);
    }

    sheet.clear();
    
    if (fileDataList.length > 0) {
      sheet.getRange(1, 1, fileDataList.length, fileDataList[0].length).setValues(fileDataList);
    }
    
    console.info('ファイルリストの書き出しが正常に完了しました。');

  } catch (e) {
    // --- 3. エラーハンドリング ---
    const errorMessage = `エラーが発生しました: ${e.message}\nスタックトレース: ${e.stack}`;
    console.error(errorMessage);
  }
}
