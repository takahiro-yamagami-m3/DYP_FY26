const TEMPLATE_ID = "1RFNqrJwwUJVtGWSOjEni38SqMvhkXJkBFZy3NJOVqTk";
const TARGET_FOLDER_ID = "1uvfjiG5KPIIkkrZyLOw1d_8cyi4q1ZjC";
const SPREADSHEET_NAME_SUFFIX = "_FY26予算フォーマット";

const DEPT_CELLS = {
  sheetName: "設定",
  cells: ["C2", "C3"]
};

/**
 * テンプレートをコピーし、各部署と共有
 */
function createAndShareSpreadsheet() {
  Logger.log('スプレッドシートの作成と共有を開始します。');

  const spreadsheet = SpreadsheetApp.openById(TEMPLATE_ID);
  const settingsSheet = spreadsheet.getSheetByName('部署設定');
  if (!settingsSheet) {
    Logger.log('エラー： 「部署設定」という名前のシートが見つかりません。');
    return;
  }

  const departmentPermissions = getPermissionsFromSettingsSheet(settingsSheet);
  if (departmentPermissions.length === 0) {
    Logger.log('エラー： 「部署設定」シートに部署情報が見つかりません。');
    return;
  }

  departmentPermissions.forEach((permission, index) => {
    Logger.log(`Processing department combination: ${permission.departmentValues.join(" / ")}`);

    let newFile = null;
    let newSpreadsheetName = '';

    try {
      const sourceSheet = spreadsheet.getSheetByName(DEPT_CELLS.sheetName);
      if (!sourceSheet) {
        Logger.log(`エラー： 「${DEPT_CELLS.sheetName}」という名前のシートが見つかりません。スキップします。`);
        return; // このreturnはループ全体を終了させます
      }

      // 部署名（C2, C3）の値を一括で変更
      const valuesToSet = [[permission.departmentValues[0]], [permission.departmentValues[1]]];
      sourceSheet.getRange('C2:C3').setValues(valuesToSet);

      SpreadsheetApp.flush();

      // 新しい命名規則に基づいてファイル名を生成
      let [valueC2, valueC3] = permission.departmentValues;

      // アスタリスクを削除
      valueC2 = valueC2.replace(/\*/g, '');
      valueC3 = valueC3.replace(/\*/g, '');

      if (valueC2 === valueC3) {
        newSpreadsheetName = `${valueC2}${SPREADSHEET_NAME_SUFFIX}`;
      } else {
        newSpreadsheetName = `${valueC2}_${valueC3}${SPREADSHEET_NAME_SUFFIX}`;
      }

      // 連続するアンダースコアを一つにまとめ、先頭・末尾のアンダースコアを削除
      newSpreadsheetName = newSpreadsheetName.replace(/_+/g, '_').replace(/^_|_$/g, '');

      const sourceFile = DriveApp.getFileById(TEMPLATE_ID);
      const targetFolder = DriveApp.getFolderById(TARGET_FOLDER_ID);
      newFile = sourceFile.makeCopy(newSpreadsheetName, targetFolder);

      // 新しく作成したスプレッドシートのURLを「部署設定」シートのD列に書き込む
      settingsSheet.getRange(index + 2, 4).setValue(newFile.getUrl());

      Logger.log(`File copied: ${newFile.getUrl()}`);

      // 非表示にするシートを変更
      const sheetsToHide = [
        '科目MAP',
        '設定',
        '部署設定'
      ];

      const newSpreadsheet = SpreadsheetApp.openById(newFile.getId());
      sheetsToHide.forEach(sheetName => {
        const sheet = newSpreadsheet.getSheetByName(sheetName);
        if (sheet) {
          sheet.hideSheet();
        } else {
          Logger.log(`警告： 「${sheetName}」という名前のシートが見つかりませんでした。`);
        }
      });
      
      // ここに共有設定を追加する場合は、permissionオブジェクトにメールアドレスも格納する必要があります
      // 例: newFile.addEditor(permission.email);

      Logger.log(`成功： ${newSpreadsheetName} を共有ドライブに作成しました。`);

    } catch (error) {
      Logger.log(`Error processing for combination ${permission.departmentValues.join(" / ")}: ${error.message}`);
      if (newFile) {
        newFile.setTrashed(true);
        Logger.log(`エラーのため、作成中のファイルを削除しました。`);
      }
    }
  });

  Logger.log('完了：すべてのスプレッドシートの作成と共有が完了しました。');
}

/**
 * 「部署設定」シートから部署情報を取得
 */
function getPermissionsFromSettingsSheet(sheet) {
  const startRow = 2;
  const numRows = sheet.getLastRow() - startRow + 1;
  const numCols = sheet.getLastColumn();

  if (numRows <= 0 || numCols <= 0) {
    return [];
  }

  const values = sheet.getRange(startRow, 1, numRows, numCols).getValues();

  const permissions = [];
  values.forEach(row => {
    // A列とB列の部署情報を取得
    const departmentValues = [row[0], row[1]];
    permissions.push({ departmentValues: departmentValues });
  });
  
  return permissions;
}