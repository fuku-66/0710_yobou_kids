/**
 * カレンダー管理クラス
 */
class CalendarManager {
  constructor() {
    this.sheet = SpreadsheetApp.getActiveSpreadsheet();
    this.settings = this.loadSettings();
  }

  /**
   * 設定を読み込む
   */
  loadSettings() {
    const settingsSheet = this.sheet.getSheetByName('設定');
    if (!settingsSheet) return {};

    const settings = {};
    const data = settingsSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === 'カレンダーID') {
        settings.calendarId = data[i][1];
      }
      if (data[i][0] === 'カレンダー予定の長さ') {
        settings.eventDuration = data[i][1] || 60; // デフォルト60分
      }
    }
    return settings;
  }

  /**
   * カレンダーIDを取得または作成
   */
  getOrCreateCalendar() {
    if (this.settings.calendarId) {
      try {
        const calendar = CalendarApp.getCalendarById(this.settings.calendarId);
        if (calendar) return calendar;
      } catch (e) {
        console.error('保存されていたカレンダーIDが無効です:', e);
      }
    }

    // 新しいカレンダーを作成
    const calendar = CalendarApp.createCalendar('予防接種スケジュール', {
      summary: '子供の予防接種スケジュール管理用カレンダー',
      timeZone: 'Asia/Tokyo'
    });

    // カレンダーIDを設定に保存
    this.saveCalendarId(calendar.getId());
    return calendar;
  }

  /**
   * カレンダーIDを設定に保存
   */
  saveCalendarId(calendarId) {
    const settingsSheet = this.sheet.getSheetByName('設定');
    if (!settingsSheet) return;

    const data = settingsSheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === 'カレンダーID') {
        settingsSheet.getRange(i + 1, 2).setValue(calendarId);
        found = true;
        break;
      }
    }

    if (!found) {
      const lastRow = settingsSheet.getLastRow();
      settingsSheet.getRange(lastRow + 1, 1, 1, 2).setValues([['カレンダーID', calendarId]]);
    }

    this.settings.calendarId = calendarId;
  }

  /**
   * 予防接種予定をカレンダーに追加
   */
  addVaccinationEvent(childName, vaccineName, date, memo = '') {
    if (!date) return null;

    const calendar = this.getOrCreateCalendar();
    const duration = this.settings.eventDuration || 60;
    const endTime = new Date(date.getTime() + duration * 60000);

    const title = `${childName}の予防接種: ${vaccineName}`;
    const description = memo ? `メモ: ${memo}` : '';

    try {
      const event = calendar.createEvent(title, date, endTime, {
        description: description,
        location: '',
        sendInvites: false
      });
      return event.getId();
    } catch (e) {
      console.error('カレンダーイベントの作成に失敗しました:', e);
      return null;
    }
  }

  /**
   * カレンダーの予定を更新
   */
  updateVaccinationEvent(eventId, childName, vaccineName, date, memo = '') {
    if (!eventId || !date) return false;

    const calendar = this.getOrCreateCalendar();
    try {
      const event = calendar.getEventById(eventId);
      if (!event) return false;

      const duration = this.settings.eventDuration || 60;
      const endTime = new Date(date.getTime() + duration * 60000);

      const title = `${childName}の予防接種: ${vaccineName}`;
      const description = memo ? `メモ: ${memo}` : '';

      event.setTitle(title);
      event.setTime(date, endTime);
      event.setDescription(description);
      return true;
    } catch (e) {
      console.error('カレンダーイベントの更新に失敗しました:', e);
      return false;
    }
  }

  /**
   * カレンダーの予定を削除
   */
  deleteVaccinationEvent(eventId) {
    if (!eventId) return false;

    const calendar = this.getOrCreateCalendar();
    try {
      const event = calendar.getEventById(eventId);
      if (!event) return false;

      event.deleteEvent();
      return true;
    } catch (e) {
      console.error('カレンダーイベントの削除に失敗しました:', e);
      return false;
    }
  }
}

/**
 * スケジュールシートの予定をカレンダーに同期
 */
function syncScheduleToCalendar() {
  const calendarManager = new CalendarManager();
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const scheduleSheets = sheet.getSheets().filter(s => 
    s.getName().match(/スケジュール/)
  );

  scheduleSheets.forEach(scheduleSheet => {
    const childName = scheduleSheet.getName().replace('スケジュール', '').trim();
    const data = scheduleSheet.getDataRange().getValues();
    const headerRow = data[0];
    
    // 列のインデックスを取得
    const vaccineNameCol = headerRow.indexOf('ワクチン名');
    const dateCol = headerRow.indexOf('予約日');
    const memoCol = headerRow.indexOf('メモ');
    const eventIdCol = headerRow.indexOf('カレンダーEventID');
    
    // 必要な列が見つからない場合はスキップ
    if (vaccineNameCol === -1 || dateCol === -1) return;

    // 各行を処理
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const vaccineName = row[vaccineNameCol];
      const date = row[dateCol];
      const memo = memoCol !== -1 ? row[memoCol] : '';
      const eventId = eventIdCol !== -1 ? row[eventIdCol] : '';

      if (date && date instanceof Date) {
        if (eventId) {
          // 既存の予定を更新
          const success = calendarManager.updateVaccinationEvent(
            eventId,
            childName,
            vaccineName,
            date,
            memo
          );
          if (!success) {
            // 更新に失敗した場合は新規作成
            const newEventId = calendarManager.addVaccinationEvent(
              childName,
              vaccineName,
              date,
              memo
            );
            if (newEventId && eventIdCol !== -1) {
              scheduleSheet.getRange(i + 1, eventIdCol + 1).setValue(newEventId);
            }
          }
        } else {
          // 新規予定を作成
          const newEventId = calendarManager.addVaccinationEvent(
            childName,
            vaccineName,
            date,
            memo
          );
          if (newEventId && eventIdCol !== -1) {
            scheduleSheet.getRange(i + 1, eventIdCol + 1).setValue(newEventId);
          }
        }
      } else if (eventId) {
        // 日付が削除された場合は予定も削除
        calendarManager.deleteVaccinationEvent(eventId);
        if (eventIdCol !== -1) {
          scheduleSheet.getRange(i + 1, eventIdCol + 1).setValue('');
        }
      }
    }
  });
}

/**
 * カレンダーのイベントが更新されたときのトリガー
 */
function onCalendarEventUpdated(event) {
  const calendarManager = new CalendarManager();
  const calendar = calendarManager.getOrCreateCalendar();
  
  // イベントがこのカレンダーのものでない場合はスキップ
  if (event.calendarId !== calendar.getId()) return;
  
  // イベントタイトルから子供の名前とワクチン名を抽出
  const match = event.title.match(/^(.+)の予防接種: (.+)$/);
  if (!match) return;
  
  const childName = match[1];
  const vaccineName = match[2];
  
  // スプレッドシートを更新
  updateScheduleSheet(childName, vaccineName, event);
}

/**
 * スプレッドシートの更新
 */
function updateScheduleSheet(childName, vaccineName, event) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const scheduleSheet = sheet.getSheetByName(`${childName}のスケジュール`);
  if (!scheduleSheet) return;

  const data = scheduleSheet.getDataRange().getValues();
  const headerRow = data[0];
  
  // 列のインデックスを取得
  const vaccineNameCol = headerRow.indexOf('ワクチン名');
  const dateCol = headerRow.indexOf('予約日');
  const statusCol = headerRow.indexOf('ステータス');
  const eventIdCol = headerRow.indexOf('カレンダーEventID');
  
  if (vaccineNameCol === -1 || dateCol === -1 || statusCol === -1 || eventIdCol === -1) return;

  // イベントの日付が過去の場合は「済み」、未来の場合は「予約済」
  const now = new Date();
  const eventDate = new Date(event.start.dateTime);
  const newStatus = eventDate < now ? '済み' : '予約済';

  // 該当する行を探して更新
  for (let i = 1; i < data.length; i++) {
    if (data[i][vaccineNameCol] === vaccineName && data[i][eventIdCol] === event.id) {
      scheduleSheet.getRange(i + 1, dateCol + 1).setValue(eventDate);
      scheduleSheet.getRange(i + 1, statusCol + 1).setValue(newStatus);
      
      // 接種済みの場合はメール通知
      if (newStatus === '済み') {
        sendVaccinationCompletionEmail(childName, vaccineName, eventDate);
      }
      break;
    }
  }
}

/**
 * 接種完了メールの送信
 */
function sendVaccinationCompletionEmail(childName, vaccineName, date) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = sheet.getSheetByName('設定');
  if (!settingsSheet) return;

  // メールアドレスの取得
  const data = settingsSheet.getDataRange().getValues();
  let emailAddresses = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'メールアドレス' && data[i][1]) {
      emailAddresses.push(data[i][1]);
    }
  }

  if (emailAddresses.length === 0) return;

  // メール本文の作成
  const formattedDate = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy年M月d日');
  const subject = `【予防接種完了】${childName}の${vaccineName}`;
  const body = `
${childName}の${vaccineName}の接種が完了しました。

接種日: ${formattedDate}

スプレッドシートの状態を「済み」に更新しました。
次回の予防接種の予定も確認してください。

このメールは自動送信されています。
`;

  // メール送信
  emailAddresses.forEach(email => {
    try {
      GmailApp.sendEmail(email, subject, body);
    } catch (e) {
      console.error('メール送信に失敗しました:', e);
    }
  });
}

/**
 * カレンダートリガーの設定
 */
function setupCalendarTrigger() {
  const calendarManager = new CalendarManager();
  const calendar = calendarManager.getOrCreateCalendar();
  
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onCalendarEventUpdated') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 新しいトリガーを設定
  ScriptApp.newTrigger('onCalendarEventUpdated')
    .forCalendar(calendar.getId())
    .onEventUpdated()
    .create();
}

/**
 * カレンダーのURLを取得
 */
function getCalendarUrl() {
  const calendarManager = new CalendarManager();
  const calendar = calendarManager.getOrCreateCalendar();
  
  try {
    // カレンダーのURLを取得
    const calendarId = calendar.getId();
    const encodedCalendarId = encodeURIComponent(calendarId);
    const url = `https://calendar.google.com/calendar/embed?src=${encodedCalendarId}`;
    
    // UIに表示
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      'カレンダーURL',
      `以下のURLからカレンダーにアクセスできます：\n\n${url}\n\n` +
      'このURLを共有すると、他の人もカレンダーを閲覧できます。\n' +
      '※ カレンダーの編集権限は共有されません。',
      ui.ButtonSet.OK
    );
  } catch (e) {
    console.error('カレンダーURLの取得に失敗しました:', e);
    SpreadsheetApp.getUi().alert(
      'エラー',
      'カレンダーURLの取得に失敗しました。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
} 