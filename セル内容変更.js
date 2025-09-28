/***** 設定ここだけ *****/
const FOLDER_ID_2 = '1t30T_cA78N0gVjrNXudsHYcoEvNwtYpy'; // 例: '13OT_ca78N0gVjnvXudsHYC0eNVwtYpy'
const TARGET_SHEET_NAME = '社員';  // 旧：正社員
/*************************/

function mainUpdateJ3() {
  const summary = {checked:0, spreadsheets:0, updated:0, errors:0};
  const root = DriveApp.getFolderById(FOLDER_ID_2);
  processFolderForJ3_(root, summary);
  Logger.log('--- 完了サマリ ---');
  Logger.log(JSON.stringify(summary, null, 2));
}

function processFolderForJ3_(folder, summary) {
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    summary.checked++;
    if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
      summary.spreadsheets++;
      try {
        updateJ3Cell_(file, summary);
      } catch (e) {
        summary.errors++;
        Logger.log(`ERROR: ${file.getName()} (${file.getId()}) -> ${e.message}`);
      }
    }
  }
  const subs = folder.getFolders();
  while (subs.hasNext()) {
    processFolderForJ3_(subs.next(), summary);
  }
}

function updateJ3Cell_(file, summary) {
  const ss = SpreadsheetApp.openById(file.getId());
  const sheet = ss.getSheetByName(TARGET_SHEET_NAME);
  if (!sheet) return; // シートが存在しない場合はスキップ

  // セルの内容を上書き
  sheet.getRange('A1').setValue('社員一覧');

  summary.updated++;
  Logger.log(`UPDATED: ${file.getName()} (${file.getId()})`);
}
