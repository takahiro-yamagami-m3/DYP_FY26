/***** 設定ここだけ *****/
const FOLDER_ID = '1t30T_cA78N0gVjrNXudsHYcoEvNwtYpy'; // 例: '13OT_ca78N0gVjnvXudsHYC0eNVwtYpy'
const OLD_NAME  = '正社員';
const NEW_NAME  = '社員';
const DRY_RUN   = false;   // 最初は true で対象確認。問題なければ false にして本実行。
/*************************/

function main() {
  const summary = {checked:0, spreadsheets:0, renamed:0, skippedNoSheet:0, skippedConflict:0, errors:0};
  const root = DriveApp.getFolderById(FOLDER_ID);
  processFolder_(root, summary);
  Logger.log('--- 完了サマリ ---');
  Logger.log(JSON.stringify(summary, null, 2));
}

function processFolder_(folder, summary) {
  // ファイル処理
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    summary.checked++;
    if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
      summary.spreadsheets++;
      try {
        processSpreadsheet_(file, summary);
      } catch (e) {
        summary.errors++;
        Logger.log(`ERROR: ${file.getName()} (${file.getId()}) -> ${e.message}`);
      }
    }
  }
  // サブフォルダ再帰
  const subs = folder.getFolders();
  while (subs.hasNext()) {
    processFolder_(subs.next(), summary);
  }
}

function processSpreadsheet_(file, summary) {
  const ss = SpreadsheetApp.openById(file.getId());
  const target = ss.getSheetByName(OLD_NAME);
  if (!target) {
    summary.skippedNoSheet++;
    return;
  }
  // すでに「社員」シートがあるケースはスキップ（衝突回避）
  if (ss.getSheetByName(NEW_NAME)) {
    summary.skippedConflict++;
    Logger.log(`SKIP（同名あり）: ${file.getName()} (${file.getId()})`);
    return;
  }
  if (DRY_RUN) {
    Logger.log(`[DRY-RUN] ${file.getName()} の「${OLD_NAME}」→「${NEW_NAME}」`);
    return;
  }
  target.setName(NEW_NAME);
  summary.renamed++;
  Logger.log(`RENAMED: ${file.getName()} (${file.getId()})`);
}
