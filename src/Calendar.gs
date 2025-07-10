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