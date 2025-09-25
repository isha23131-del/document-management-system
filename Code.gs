// スプレッドシートのIDと名前を動的に取得
const SCRIPT_PROPS = PropertiesService.getScriptProperties();
const SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const SHEET_NAME = "Submissions"; // 固定のシート名

function getSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, 9).setValues([[
      'ID', 'Timestamp', 'UserEmail', '氏名', '書類名', '提出予定日', '提出日(実績)', 'ステータス', '備考'
    ]]);
  }
  return sheet;
}

function getCurrentUser() {
  const email = Session.getActiveUser().getEmail(); // 実際のユーザーメールを使用
  // デバッグ用: emailが取得できない場合
  // const email = Session.getActiveUser().getEmail() || "test@example.com";
  const name = email.includes("@") ? email.split("@")[0].trim() : email.trim();
  return { email, name };
}

function submitForm(payload) {
  const sheet = getSheet_();
  const timestamp = new Date();
  const user = getCurrentUser();

  const row = [
    new Date().getTime(), // ユニークID
    timestamp,
    user.email,
    user.name,
    payload.docName,
    payload.dueDate,
    "", // 提出日(実績)
    "未提出",
    "" // 備考
  ];
  sheet.appendRow(row);

  return { success: true };
}

function getSubmissions() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return []; // ヘッダーのみの場合は空の配列を返す
  
  const headers = values[0]; // ヘッダー行
  values.shift(); // ヘッダー行を削除

  const user = getCurrentUser();
  const today = new Date();
  today.setHours(0, 0, 0, 0); // 日付のみで比較するため時間をリセット

  return values
    .filter(row => row[3] && row[3].toString().trim() === user.name)
    .sort((a, b) => b[1].getTime() - a[1].getTime()) // Timestampで降順にソート
    .map((row, index) => ({
      rowIndex: index + 2, // スプレッドシートの行番号 (ヘッダーと0-indexedを考慮)
      ID: row[0],
      氏名: row[3],
      書類名: row[4],
      提出予定日: Utilities.formatDate(new Date(row[5]), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      提出日実績: row[6] ? Utilities.formatDate(new Date(row[6]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
      ステータス: row[7],
      isOverdue: (new Date(row[5]) < today && row[7] === "未提出") // 提出予定日超過かつ未提出
    }));
}

function updateSubmission(payload) {
  const sheet = getSheet_();
  // スプレッドシートの行番号は1から始まるため、payload.rowIndexはそのまま使用
  // ID, Timestamp, UserEmail, 氏名, 書類名, 提出予定日, 提出日(実績), ステータス, 備考
  const range = sheet.getRange(payload.rowIndex, 5, 1, 4); // 書類名からステータスまでの4列
  const newValues = [payload.docName, payload.dueDate, payload.actualDate, payload.status];
  range.setValues([newValues]);
  return { success: true };
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index").setTitle("書類提出システム");
}
