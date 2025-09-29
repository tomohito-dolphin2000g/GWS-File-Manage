function listFilesInFolder() {
  // フォルダIDとスプレッドシートIDを直接指定
  var folderId = 'YOUR_FOLDER_ID'; // ここに準備したフォルダIDを設定
  var sheetId = 'YOUR_SPREADSHEET_ID'; // ここに準備したスプレッドシートIDを設定

  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();

  var sheet = SpreadsheetApp.openById(sheetId).getActiveSheet();
  sheet.clear(); // シートをクリア
  sheet.appendRow(['ファイル名', 'URL']); // ヘッダー行

  while (files.hasNext()) {
    var file = files.next();
    var fileInfo = [file.getName(), file.getUrl()];
    sheet.appendRow(fileInfo);
  }
}
