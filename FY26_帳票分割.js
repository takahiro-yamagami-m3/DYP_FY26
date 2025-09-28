const TEMPLATE_SS_ID = "1qzQJJJXAGcDoyao0ChSx0I8Pb8rGgujS6IY30EzCHa8";
const TEMPLATE_FILE_ID = "1qzQJJJXAGcDoyao0ChSx0I8Pb8rGgujS6IY30EzCHa8";
const TARGET_FOLDER_ID = "15ewFOcCSo-FUGi8exM-kyf0yf-Yz1WC2";
const SPREADSHEET_NAME_SUFFIX = "_DYP申請フォーマット";

const DEPT_CELLS = {
  sheetName: "申請サマリ",
  cells: ["B3", "B4", "B5"]
};
  
/**
 * テンプレートをコピーし、各部署と共有
 */
function createAndShareSpreadsheet() {
  Logger.log('スプレッドシートの作成と共有を開始します。');

  const spreadsheet = SpreadsheetApp.openById(TEMPLATE_SS_ID);
  const settingsSheet = spreadsheet.getSheetByName('設定');
  if (!settingsSheet) {
    Logger.log('エラー： 「設定」という名前のシートが見つかりません。');
    return;
  }

  const permissions = getPermissionsFromSettingsSheet(settingsSheet);
  if (permissions.length === 0) {
    Logger.log('エラー： 「設定」シートのA2セル以降に部署情報が見つかりません。');
    return;
  }

  // 権限情報を取得した後に、新しいスプレッドシートを作成し、URLを「設定」シートに書き込む
  const settingRows = settingsSheet.getRange(2, 1, permissions.length, settingsSheet.getLastColumn());
  const settingValues = settingRows.getValues();

  permissions.forEach((departmentPermission, index) => {
    Logger.log(`Processing department combination: ${departmentPermission.departmentValues.join(" / ")}`);

    let newFile = null;

    try {
      const sourceSheet = spreadsheet.getSheetByName(DEPT_CELLS.sheetName);
      if (!sourceSheet) {
        Logger.log(`エラー： 「${DEPT_CELLS.sheetName}」という名前のシートが見つかりません。スキップします。`);
        return;
      }

      // スプレッドシートのコピー前にB3, B4, B5セルの値を一括で変更
      const valuesToSet = [[departmentPermission.departmentValues[0]], [departmentPermission.departmentValues[1]], [departmentPermission.departmentValues[2]]];
      sourceSheet.getRange('B3:B5').setValues(valuesToSet);

      SpreadsheetApp.flush();
      
      // 新しい命名規則に基づいてファイル名を生成
      let [valueB3, valueB4, valueB5] = departmentPermission.departmentValues;
      
      valueB3 = valueB3.replace(/\*/g, '');
      valueB4 = valueB4.replace(/\*/g, '');
      valueB5 = valueB5.replace(/\*/g, '');
      
      let newSpreadsheetName;
      if (valueB3 === valueB4) {
        newSpreadsheetName = `${valueB3}_${valueB5}${SPREADSHEET_NAME_SUFFIX}`;
      } else {
        newSpreadsheetName = `${valueB3}_${valueB4}_${valueB5}${SPREADSHEET_NAME_SUFFIX}`;
      }
      
      // 連続するアンダースコアを一つにまとめる
      newSpreadsheetName = newSpreadsheetName.replace(/_+/g, '_');
      newSpreadsheetName = newSpreadsheetName.replace(/^_|_$/g, '');

      const sourceFile = DriveApp.getFileById(TEMPLATE_FILE_ID);
      const targetFolder = DriveApp.getFolderById(TARGET_FOLDER_ID);
      newFile = sourceFile.makeCopy(newSpreadsheetName, targetFolder);
      
      // 新しく作成したスプレッドシートのURLを「設定」シートのD列に書き込む
      settingsSheet.getRange(index + 2, 4).setValue(newFile.getUrl());
      
      Logger.log(`File copied: ${newFile.getUrl()}`);

      // 非表示にするシートを変更
      const sheetsToHide = [
        '設定',
        '部署マスタ',
        '部署一覧',
        'SPBU_Gマスタ'
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

      // SpreadsheetApp を使って権限を付与
      departmentPermission.users.forEach(user => {
        try {
          if (user.role === '編集者') {
            // 不要な引数を削除
            newSpreadsheet.addEditor(user.email);
          } else if (user.role === '閲覧者') {
            // 不要な引数を削除
            newSpreadsheet.addViewer(user.email);
          }
        } catch (e) {
          Logger.log(`Error adding permission for ${user.email}: ${e.message}`);
        }
      });
    
      Logger.log(`成功： ${newSpreadsheetName} を共有ドライブに作成し、権限を付与しました。`);

    } catch (error) {
      Logger.log(`Error processing for combination ${departmentPermission.departmentValues.join(" / ")}: ${error.message}`);
      if (newFile) {
        newFile.setTrashed(true);
        Logger.log(`Trashed partially created file due to error.`);
      }
    }
  });

  Logger.log('完了：すべてのスプレッドシートの作成と共有が完了しました。');
}

/**
 * 「設定」シートからユーザー権限を取得
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
    const departmentValues = [row[0], row[1], row[2]];
    
    // 部署情報が空でないことを確認
    if (departmentValues.every(val => val && val.toString().trim() !== '')) {
      const users = [];
      // メールアドレスはE列(インデックス4)から開始
      for (let i = 4; i < 9; i++) { // E列からI列までがメールアドレスの想定
        const email = row[i];
        if (email && email.toString().trim() !== '') {
          // 権限種別はメールアドレスの数だけI列から取得
          const role = row[i + 5]; 
          if (role && role.toString().trim() !== '') {
            users.push({
              email: email.toString().trim(),
              role: role.toString().trim()
            });
          }
        }
      }
      permissions.push({
        departmentValues: departmentValues,
        users: users
      });
    }
  });
  return permissions;
}
