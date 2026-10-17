const SHEET_NAME = '申込情報';

/**
 * LIFF画面を表示
 */
function doGet() {
  console.log("スタート！")
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('申込情報');
}


/**
 * POST処理
 */
function doPost(e) {
  console.log("スタート２！")

  try {

    const data = JSON.parse(e.postData.contents);

    switch (data.action) {

      case 'create':
        return createRecord(data);

      case 'read':
        return readRecords();

      case 'update':
        return updateRecord(data);

      case 'delete':
        return deleteRecord(data);

      default:
        return response({
          success: false,
          message: '不正なactionです'
        });
    }

  } catch (error) {

    return response({
      success: false,
      message: error.message
    });

  }
}


/**
 * CREATE
 * 新規登録
 */
function createRecord(data) {

  const sheet = getSheet();

  // 申込番号の重複チェック
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {

    if (String(values[i][0]) === String(data.applicationNo)) {

      return response({
        success: false,
        message: 'この申込番号は既に登録されています'
      });

    }
  }

  sheet.appendRow([
    data.applicationNo,
    data.lineId,
    data.displayName
  ]);

  return response({
    success: true,
    message: '登録しました'
  });
}


/**
 * READ
 * 全件取得
 */
function readRecords() {

  const sheet = getSheet();

  const values = sheet.getDataRange().getValues();

  const records = [];

  for (let i = 1; i < values.length; i++) {

    if (values[i][0] === '') {
      continue;
    }

    records.push({
      applicationNo: String(values[i][0]),
      lineId: String(values[i][1]),
      displayName: String(values[i][2])
    });
  }

  return response({
    success: true,
    records: records
  });
}


/**
 * UPDATE
 * 申込番号をキーに更新
 */
function updateRecord(data) {

  const sheet = getSheet();

  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {

    if (String(values[i][0]) === String(data.applicationNo)) {

      sheet.getRange(i + 1, 2).setValue(data.lineId);
      sheet.getRange(i + 1, 3).setValue(data.displayName);

      return response({
        success: true,
        message: '更新しました'
      });
    }
  }

  return response({
    success: false,
    message: '対象の申込番号がありません'
  });
}


/**
 * DELETE
 * 申込番号をキーに削除
 */
function deleteRecord(data) {

  const sheet = getSheet();

  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {

    if (String(values[i][0]) === String(data.applicationNo)) {

      sheet.deleteRow(i + 1);

      return response({
        success: true,
        message: '削除しました'
      });
    }
  }

  return response({
    success: false,
    message: '対象の申込番号がありません'
  });
}


/**
 * スプレッドシート取得
 */
function getSheet() {

  return SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(SHEET_NAME);
}


/**
 * JSONレスポンス
 */
function response(data) {

  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}