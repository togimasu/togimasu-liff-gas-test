const SHEET_NAME = '申込情報';
const HEADERS = ['申込番号', '申込日時', 'LINE表示名', 'LINE ID', '受取り希望時間', '備考', '作業状況', '本数'];
const LEGACY_HEADERS = HEADERS.slice(0, -1);
const TIME_ZONE = 'Asia/Tokyo';

/** 申込者向けフォームをGASから直接表示する場合の入口。 */
function doGet() {
  console.log("★doGet Start★");
  return HtmlService.createHtmlOutputFromFile('index').setTitle('包丁研ぎのお申込み');
}

/** Googleログインボタンの初期化に使う公開OAuthクライアントIDを返す。 */
function getAdminAuthConfig() {
  const properties = PropertiesService.getScriptProperties();
  const clientId = properties.getProperty('ADMIN_GOOGLE_CLIENT_ID');
  const allowedEmails = properties.getProperty('ADMIN_EMAILS');
  if (!clientId || !allowedEmails) throw new Error('管理用Google認証の設定が未完了です。');
  return { clientId: clientId };
}

/** Google IDトークンと許可メール一覧を確認して申込一覧を返す。 */
function getAdminApplications(idToken) {
  const properties = PropertiesService.getScriptProperties();
  const clientId = properties.getProperty('ADMIN_GOOGLE_CLIENT_ID');
  const allowedEmails = (properties.getProperty('ADMIN_EMAILS') || '')
    .split(',').map(email => email.trim().toLowerCase()).filter(Boolean);
  if (!clientId || allowedEmails.length === 0) {
    throw new Error('管理用Google認証が未設定です。');
  }
  if (!idToken) throw new Error('Googleアカウントでログインしてください。');

  const verification = UrlFetchApp.fetch(
    'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken),
    { muteHttpExceptions: true }
  );
  if (verification.getResponseCode() !== 200) throw new Error('Googleログインを確認できませんでした。再度ログインしてください。');
  const claims = JSON.parse(verification.getContentText());
  const email = String(claims.email || '').toLowerCase();
  const isEmailVerified = claims.email_verified === true || claims.email_verified === 'true';
  const validIssuer = claims.iss === 'accounts.google.com' || claims.iss === 'https://accounts.google.com';
  if (claims.aud !== clientId || !validIssuer || !isEmailVerified || !allowedEmails.includes(email)) {
    throw new Error('このGoogleアカウントには管理画面の利用権限がありません。');
  }
  const rows = getSheet().getDataRange().getValues();
  return rows.slice(1).map(row => {
    const record = rowToRecord(row);
    delete record.lineId;
    return record;
  }).reverse();
}

/** 初回セットアップ時に、エディタから外部通信権限を承認するための関数。 */
function authorizeExternalRequest() {
  UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true });
}

/** 管理APIと申込者向けLIFF APIを振り分けて処理する。 */
function doPost(e) {
  try {
    console.log("★doPost Start★");
    const data = JSON.parse(e.postData.contents || '{}');
    if (data.action === 'adminAuthConfig') {
      return response({ success: true, config: getAdminAuthConfig() });
    }
    if (data.action === 'adminList') {
      return response({ success: true, records: getAdminApplications(data.idToken) });
    }
    const identity = verifyLiffToken(data.idToken);
    let result;
    console.log("★doPost ★★" + data);
    switch (data.action) {
      case 'findToday':
        result = findToday(identity.userId);
        break;
      case 'create':
        data.lineId = identity.userId;
        data.displayName = identity.displayName || data.displayName;
        result = createApplication(data);
        break;
      default:
        result = { success: false, message: '不正な操作です。' };
    }
    return response(result);
  } catch (error) {
    console.error(error);
    return response({ success: false, message: error.message || 'サーバー処理に失敗しました。' });
  }
}

/** LIFF IDトークンをLINEで検証し、ユーザーIDをサーバー側で確定する。 */
function verifyLiffToken(idToken) {
  if (!idToken) throw new Error('LINEの認証情報がありません。LIFFを開き直してください。');
  const channelId = '2011920848';
  const result = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: { id_token: idToken, client_id: channelId },
    muteHttpExceptions: true
  });
  if (result.getResponseCode() !== 200) {
    let detail = result.getContentText();
    try {
      const errorBody = JSON.parse(detail);
      detail = errorBody.error_description || errorBody.error || detail;
    } catch (ignored) {}
    console.error('LINE ID token verification failed: HTTP ' + result.getResponseCode() + ', ' + detail);
    throw new Error('LINE認証の検証に失敗しました（HTTP ' + result.getResponseCode() + '）。' + detail);
  }
  const profile = JSON.parse(result.getContentText());
  if (!profile.sub) throw new Error('LINEユーザーを確認できませんでした。');
  return { userId: profile.sub, displayName: profile.name || '' };
}

/** 本日、指定LINE IDで申込済みかを確認する。 */
function findToday(lineId) {
  console.log("★findToday★");
  
  if (!lineId) return { success: false, message: 'LINE IDを取得できませんでした。' };
  const today = todayKey();
  const rows = getSheet().getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][3]) === String(lineId) && dateKey(rows[i][1]) === today) {
      return { success: true, record: rowToRecord(rows[i]) };
    }
  }
  return { success: true, record: null };
}

/** 申込番号を採番して、本日の申込として登録する。 */
function createApplication(data) {
  if (!data.lineId || !data.displayName) throw new Error('LINEのプロフィールを取得できませんでした。');
  const pickupTime = String(data.pickupTime || '');
  if (!isValidPickupTime(pickupTime)) throw new Error('受取り希望時間を選択し直してください。');
  const knifeCount = Number(data.knifeCount);
  if (!Number.isInteger(knifeCount) || knifeCount < 1) throw new Error('本数は1以上の整数で入力してください。');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet();
    const today = todayKey();
    const rows = sheet.getDataRange().getValues();
    let sequence = 0;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][3]) === String(data.lineId) && dateKey(rows[i][1]) === today) {
        return { success: true, record: rowToRecord(rows[i]) };
      }
      const existingNo = String(rows[i][0] || '');
      const match = existingNo.match(new RegExp('^' + today + '-(\\d+)$'));
      if (match) sequence = Math.max(sequence, Number(match[1]));
    }

    const record = {
      applicationNo: today + '-' + String(sequence + 1).padStart(3, '0'),
      applicationDate: today,
      displayName: String(data.displayName).slice(0, 100),
      lineId: String(data.lineId),
      pickupTime: pickupTime,
      knifeCount: knifeCount,
      notes: String(data.notes || '').slice(0, 1000),
      workStatus: '受付済み'
    };
    sheet.appendRow([
      record.applicationNo, record.applicationDate, record.displayName,
      record.lineId, record.pickupTime, record.notes, record.workStatus, record.knifeCount
    ]);
    return { success: true, record: record };
  } finally {
    lock.releaseLock();
  }
}

function getSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
  else {
    const current = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), LEGACY_HEADERS.length)).getValues()[0];
    const sampleHeader = ['申込番号', 'LINE ID', 'LINE表示名'];
    if (sampleHeader.every((value, index) => String(current[index] || '') === value)) {
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    } else if (LEGACY_HEADERS.every((value, index) => String(current[index] || '') === value)) {
      sheet.getRange(1, HEADERS.length).setValue(HEADERS[HEADERS.length - 1]);
    } else if (!HEADERS.every((value, index) => String(current[index] || '') === value)) {
      throw new Error('「申込情報」シートの1行目を申込項目の見出しにしてください。');
    }
  }
  return sheet;
}

function todayKey() {
  return Utilities.formatDate(new Date(), TIME_ZONE, 'yyyyMMdd');
}

function dateKey(value) {
  if (value instanceof Date) return Utilities.formatDate(value, TIME_ZONE, 'yyyyMMdd');
  return String(value || '').replace(/[^0-9]/g, '').slice(0, 8);
}

function isValidPickupTime(value) {
  if (value === '希望なし') return true;
  const match = value.match(/^(\d{2}):(00|30)$/);
  if (!match) return false;
  const hour = Number(match[1]);
  const minute = Number(match[2]);
  return (hour > 12 || (hour === 12 && minute === 0)) && (hour < 20 || (hour === 20 && minute === 0));
}

function rowToRecord(row) {
  return {
    applicationNo: String(row[0]),
    applicationDate: dateKey(row[1]),
    displayName: String(row[2] || ''),
    lineId: String(row[3] || ''),
    pickupTime: String(row[4] || '希望なし'),
    notes: String(row[5] || ''),
    workStatus: String(row[6] || '受付済み'),
    knifeCount: Number(row[7]) || 1
  };
}

function response(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
