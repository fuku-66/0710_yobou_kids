// グローバル定数
const BIRTHDAY_CELL = 'B1';
const DATA_START_ROW = 4;  // データの開始行（ヘッダーを除く）
const STATUS_COLUMN = 6;   // F列：ステータス
const APPOINTMENT_COLUMN = 7;  // G列：予約日

// スプレッドシートが編集されたときに実行
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  
  // 生年月日が入力された場合
  if (range.getA1Notation() === BIRTHDAY_CELL) {
    calculateVaccinationDates(sheet);
  }
  
  // 予約日が入力された場合
  if (range.getColumn() === APPOINTMENT_COLUMN && range.getRow() >= DATA_START_ROW) {
    addToCalendar(sheet, range.getRow());
  }
}

// 生年月日から予防接種日を計算
function calculateVaccinationDates(sheet) {
  const birthday = sheet.getRange(BIRTHDAY_CELL).getValue();
  if (!birthday || !(birthday instanceof Date)) return;
  
  const lastRow = sheet.getLastRow();
  
  for (let row = DATA_START_ROW; row <= lastRow; row++) {
    const startAge = sheet.getRange(row, 3).getValue(); // C列：標準的な接種開始時期
    if (!startAge) continue;
    
    // 開始日と終了日を計算
    const startDate = calculateDate(birthday, startAge);
    const endDate = calculateEndDate(birthday, startAge);
    
    // D列とE列に日付を設定
    sheet.getRange(row, 4).setValue(startDate);
    sheet.getRange(row, 5).setValue(endDate);
  }
}

// 予約をカレンダーに追加
function addToCalendar(sheet, row) {
  const calendar = CalendarApp.getDefaultCalendar();
  const vaccineName = sheet.getRange(row, 1).getValue();
  const appointmentDate = sheet.getRange(row, APPOINTMENT_COLUMN).getValue();
  
  if (!appointmentDate || !(appointmentDate instanceof Date)) return;
  
  // 設定から予定の長さを取得
  const settings = getSettings();
  const durationMinutes = settings.appointmentDuration || 60;
  
  // 予定の終了時刻を計算
  const endTime = new Date(appointmentDate.getTime() + (durationMinutes * 60 * 1000));
  
  calendar.createEvent(
    `予防接種：${vaccineName}`,
    appointmentDate,
    endTime,
    {description: '予防接種の予約'}
  );
}

// リマインダーチェック（毎日実行）
function checkReminders() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  const birthday = sheet.getRange(BIRTHDAY_CELL).getValue();
  
  if (!birthday || !(birthday instanceof Date)) return;
  
  // 設定を取得
  const settings = getSettings();
  if (settings.mailNotification !== 'はい') return;
  
  const today = new Date();
  const ageInMonths = monthsDiff(birthday, today);
  
  for (let row = DATA_START_ROW; row <= lastRow; row++) {
    const status = sheet.getRange(row, STATUS_COLUMN).getValue();
    if (status === '済み') continue;
    
    const endDate = sheet.getRange(row, 5).getValue();
    if (!endDate || !(endDate instanceof Date)) continue;
    
    const daysUntil = daysDiff(today, endDate);
    
    // リマインダー条件チェック
    if (shouldSendReminder(daysUntil, ageInMonths, settings)) {
      sendReminderEmail(sheet, row, daysUntil, settings);
    }
  }
}

// リマインダー送信条件をチェック
function shouldSendReminder(daysUntil, ageInMonths, settings) {
  const reminderDays = {
    '3日前': 3,
    '1週間前': 7,
    '2週間前': 14,
    '1ヶ月前': 30,
    '2ヶ月前': 60,
    '3ヶ月前': 90
  };
  
  // 全年齢のリマインド
  if (daysUntil === reminderDays[settings.reminderAll]) return true;
  
  // 1歳未満のリマインド
  if (ageInMonths < 12 && daysUntil === reminderDays[settings.reminderUnderOne]) return true;
  
  // 2歳以降のリマインド
  if (ageInMonths >= 24 && daysUntil === reminderDays[settings.reminderOverTwo]) return true;
  
  return false;
}

// リマインダーメールを送信
function sendReminderEmail(sheet, row, daysUntil, settings) {
  const vaccineName = sheet.getRange(row, 1).getValue();
  const userEmail = settings.notificationEmail;
  
  const subject = `【予防接種リマインド】${vaccineName}の接種期間が近づいています`;
  const body = `
${vaccineName}の推奨接種期間まであと${daysUntil}日です。

医療機関に予約することをお勧めします。
予約が完了しましたら、スプレッドシートの予約日欄に日時を入力してください。
`;
  
  MailApp.sendEmail(userEmail, subject, body);
}

// 設定を取得
function getSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('設定');
  
  if (!settingsSheet) {
    // 設定シートがない場合はデフォルト値を返す
    return {
      mailNotification: 'はい',
      notificationEmail: Session.getActiveUser().getEmail(),
      reminderAll: '1ヶ月前',
      reminderUnderOne: '1週間前',
      reminderOverTwo: '3ヶ月前',
      appointmentDuration: 60
    };
  }
  
  // 設定値を取得
  return {
    mailNotification: settingsSheet.getRange('B3').getValue(),
    notificationEmail: settingsSheet.getRange('B4').getValue(),
    reminderAll: settingsSheet.getRange('B5').getValue(),
    reminderUnderOne: settingsSheet.getRange('B6').getValue(),
    reminderOverTwo: settingsSheet.getRange('B7').getValue(),
    appointmentDuration: parseInt(settingsSheet.getRange('B8').getValue())
  };
}

// ユーティリティ関数
function monthsDiff(date1, date2) {
  const yearDiff = date2.getFullYear() - date1.getFullYear();
  const monthDiff = date2.getMonth() - date1.getMonth();
  return yearDiff * 12 + monthDiff;
}

function daysDiff(date1, date2) {
  const diffTime = Math.abs(date2 - date1);
  return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
}

function calculateDate(birthday, ageString) {
  // "生後2ヶ月"のような文字列から月数を抽出
  const months = parseInt(ageString.match(/\d+/)[0]);
  const date = new Date(birthday);
  date.setMonth(date.getMonth() + months);
  return date;
}

function calculateEndDate(birthday, ageString) {
  // 開始日から6ヶ月後を終了日とする（カスタマイズ可能）
  const startDate = calculateDate(birthday, ageString);
  const endDate = new Date(startDate);
  endDate.setMonth(endDate.getMonth() + 6);
  return endDate;
} 