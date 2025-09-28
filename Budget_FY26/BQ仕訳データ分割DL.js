/**
 * 部署マスタに基づき、BigQueryからデータを取得し、
 * 各部署のデータを指定されたスプレッドシートに6ヶ月ごとに書き出すスクリプト。
 */ 
function exportDataByDepartment() {
  // --------------------------------------------------
  // ユーザー設定
  // --------------------------------------------------
  const projectId = 'm3-coe-aiagent';
  const masterSpreadsheetId = '1RFNqrJwwUJVtGWSOjEni38SqMvhkXJkBFZy3NJOVqTk';
  const masterSheetName = '部署設定';
  const tableId = 'm3-coe-aiagent.fpa.Loglass_Actual';
  const startDate = new Date('2024-04-01'); // 開始年月
  const endDate = new Date('2025-08-31');   // 終了年月
  // --------------------------------------------------

  try {
    const masterSpreadsheet = SpreadsheetApp.openById(masterSpreadsheetId);
    const masterSheet = masterSpreadsheet.getSheetByName(masterSheetName);

    if (!masterSheet) {
      throw new Error('指定されたシート名「' + masterSheetName + '」が見つかりません。');
    }

    // A列（部署第2階層）、B列（部署第3階層）、D列（スプレッドシートID）のデータを取得
    // データは2行目から開始（ヘッダー行を除く）
    // rangeの列を4から5に変更してD列（4番目の要素）を確実に取得する
    const masterData = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 5).getValues();

    // 期間内の6ヶ月ごとの日付リストを生成
    const quarterlyDates = getQuarterlyDates(startDate, endDate, 6);
    if (quarterlyDates.length === 0) {
      console.warn('指定された期間内にデータがありません。');
      return;
    }

    // 各部署マスタ行に対してループ
    masterData.forEach(row => {
      const bu_sho_2 = row[0]; // A列: 部署第2階層
      const bu_sho_3 = row[1]; // B列: 部署第3階層
      const targetSpreadsheetUrl = row[3]; // D列: 吐き出し先スプレッドシートのURLを取得

      if (!targetSpreadsheetUrl) {
        console.warn(`部署「${bu_sho_2}」のスプレッドシートURLが設定されていないため、スキップします。`);
        return;
      }
      
      // URLからID部分を抽出する正規表現
      const match = targetSpreadsheetUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
      if (!match || !match[1]) {
        console.warn(`部署「${bu_sho_2}」のURLからスプレッドシートIDを抽出できませんでした。スキップします。`);
        return;
      }
      const targetSpreadsheetId = match[1];

      const targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
      const targetSheetName = '仕訳データ'; // ここを固定のシート名に変更
      let sheet = targetSpreadsheet.getSheetByName(targetSheetName);

      if (!sheet) {
        sheet = targetSpreadsheet.insertSheet(targetSheetName);
      }
      
      // 各部署の処理を開始する前に、シートを一度だけクリア
      sheet.clearContents();
      let isFirstQuery = true; // ヘッダー書き込み用のフラグ

      // 各期間に対してデータを取得し、書き出す
      quarterlyDates.forEach(period => {
        const start = Utilities.formatDate(period.start, 'JST', 'yyyy-MM-dd');
        const end = Utilities.formatDate(period.end, 'JST', 'yyyy-MM-dd');

        console.log(`部署「${bu_sho_2}」の${start}から${end}までのデータをスプレッドシートID「${targetSpreadsheetId}」にエクスポートします。`);

        // SQLクエリを動的に生成
        let sqlQuery;
        if (bu_sho_3 === '*') {
          sqlQuery = `
            SELECT
              *
            FROM
              \`${tableId}\`
            WHERE
              \`部署第2階層\` = '${bu_sho_2}' AND \`年月\` BETWEEN DATE('${start}') AND DATE('${end}')
          `;
        } else {
          sqlQuery = `
            SELECT
              *
            FROM
              \`${tableId}\`
            WHERE
              \`部署第2階層\` = '${bu_sho_2}' AND \`部署第3階層\` = '${bu_sho_3}' AND \`年月\` BETWEEN DATE('${start}') AND DATE('${end}')
          `;
        }

        // データのエクスポート処理を実行
        exportBigQueryData(projectId, sqlQuery, targetSpreadsheetId, targetSheetName, isFirstQuery);
        if (isFirstQuery) {
          isFirstQuery = false; // ヘッダーを1度だけ書き込むためのフラグ
        }
      });
    });

    console.log('すべての部署と期間のデータエクスポートが完了しました。');

  } catch (e) {
    console.error('エラー:', e.message);
  }
}

/**
 * 開始日から終了日までの指定された間隔（月）ごとの日付リストを生成するヘルパー関数
 * @param {Date} startDate 開始日
 * @param {Date} endDate 終了日
 * @param {number} intervalMonths 間隔（月）
 * @return {object[]} 開始日と終了日のペアの配列
 */
function getQuarterlyDates(startDate, endDate, intervalMonths) {
  const periods = [];
  let currentStart = new Date(startDate.getFullYear(), startDate.getMonth(), 1);

  while (currentStart <= endDate) {
    const currentEnd = new Date(currentStart);
    currentEnd.setMonth(currentEnd.getMonth() + intervalMonths);
    currentEnd.setDate(currentEnd.getDate() - 1); 

    if (currentEnd > endDate) {
      periods.push({
        start: new Date(currentStart),
        end: new Date(endDate)
      });
    } else {
      periods.push({
        start: new Date(currentStart),
        end: new Date(currentEnd)
      });
    }
    currentStart.setMonth(currentStart.getMonth() + intervalMonths);
  }
  return periods;
}

/**
 * BigQueryからデータを取得し、指定されたスプレッドシートに書き出すヘルパー関数
 * @param {string} projectId BigQueryプロジェクトID
 * @param {string} sqlQuery 実行するSQLクエリ
 * @param {string} targetSpreadsheetId 書き込み先のスプレッドシートID
 * @param {string} targetSheetName 書き込み先のシート名
 * @param {boolean} writeHeader ヘッダーを書き込むかどうか
 */
function exportBigQueryData(projectId, sqlQuery, targetSpreadsheetId, targetSheetName, writeHeader) {
  const spreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
  const sheet = spreadsheet.getSheetByName(targetSheetName);

  if (!sheet) {
    throw new Error('指定されたシート名が見つかりません: ' + targetSheetName);
  }

  const queryConfig = {
    query: sqlQuery,
    useLegacySql: false
  };

  try {
    let queryResults = BigQuery.Jobs.query(queryConfig, projectId);
    const jobId = queryResults.jobReference.jobId;
    let pageToken = queryResults.pageToken;

    let allRows = [];
    let headerRow = [];

    // ヘッダー情報の取得（初回のみ）
    if (writeHeader) {
      if (!queryResults.schema || !queryResults.schema.fields) {
        sheet.getRange(1, 1).setValue('クエリ結果にスキーマ情報がありませんでした。');
        return;
      }
      headerRow = queryResults.schema.fields.map(field => field.name);
      sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
    }
    
    // データ行の取得
    if (queryResults.rows) {
      allRows = allRows.concat(queryResults.rows.map(row => row.f.map(field => field.v)));
    }

    // ページング処理
    while (pageToken) {
      const options = { pageToken: pageToken };
      queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId, options);
      if (queryResults.rows) {
        allRows = allRows.concat(queryResults.rows.map(row => row.f.map(field => field.v)));
      }
      pageToken = queryResults.pageToken;
    }

    if (allRows.length === 0) {
      // データがない場合でも、ヘッダー行の次の行にメッセージを追記
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1).setValue('この期間にはデータがありませんでした。');
      return;
    }
    
    // データ書き込み
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, allRows.length, allRows[0].length).setValues(allRows);
    
  } catch (e) {
    console.error(`データエクスポートエラー: ${targetSpreadsheetId} - ${targetSheetName}:`, e.message);
    sheet.clear();
    sheet.getRange(1, 1).setValue('エラーが発生しました: ' + e.message);
  }
}