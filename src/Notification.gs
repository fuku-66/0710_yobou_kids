// メール通知を送信
function sendNotifications() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('設定');
  
  // 設定を取得
  const settings = getSettings(settingsSheet);
  if (!settings.enableNotification) return;
  
  // メールアドレスのリストを取得
  const emailAddresses = settings.notificationEmails.split(',').map(email => email.trim());
  if (emailAddresses.length === 0) return;
  
  // 子供の情報シートを取得
  const childrenSheet = ss.getSheetByName('子供の情報');
  if (!childrenSheet) return;
  
  // 各子供のスケジュールをチェック
  const children = childrenSheet.getRange(2, 1, childrenSheet.getLastRow() - 1, 2).getValues();
  children.forEach(([name, birthday]) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    
    // 通知が必要なワクチンを取得
    const notifications = checkVaccineSchedule(sheet, birthday, settings);
    if (notifications.length === 0) return;
    
    // 各メールアドレスに通知を送信
    emailAddresses.forEach(email => {
      sendNotificationEmail(email, name, notifications);
    });
  });
}

// 設定を取得
function getSettings(settingsSheet) {
  const settings = {
    enableNotification: false,
    notificationEmails: '',
    reminderPeriods: {
      underOne: '1週間前',
      oneToTwo: '1ヶ月前',
      overTwo: '3ヶ月前'
    },
    appointmentDuration: 60,
    showOptional: true,
    showSimultaneous: true
  };
  
  // メール通知設定
  settings.enableNotification = settingsSheet.getRange('B4').getValue() === 'はい';
  settings.notificationEmails = settingsSheet.getRange('B5').getValue();
  
  // リマインド設定
  settings.reminderPeriods.underOne = settingsSheet.getRange('B9').getValue();
  settings.reminderPeriods.oneToTwo = settingsSheet.getRange('B10').getValue();
  settings.reminderPeriods.overTwo = settingsSheet.getRange('B11').getValue();
  
  // 予約設定
  const duration = settingsSheet.getRange('B14').getValue();
  settings.appointmentDuration = parseInt(duration);
  
  // 表示設定
  settings.showOptional = settingsSheet.getRange('B17').getValue() === 'はい';
  settings.showSimultaneous = settingsSheet.getRange('B18').getValue() === 'はい';
  
  return settings;
}

// ワクチンスケジュールをチェック
function checkVaccineSchedule(sheet, birthday, settings) {
  const notifications = [];
  const today = new Date();
  
  // データ範囲を取得
  const dataRange = sheet.getRange(4, 1, sheet.getLastRow() - 3, 7);
  const values = dataRange.getValues();
  
  values.forEach((row, index) => {
    const [type, name, period, appointment, status, note, group] = row;
    
    // 既に接種済みまたは予約済みの場合はスキップ
    if (status === '済み' || status === '予約済') return;
    
    // 任意接種が非表示の場合、任意接種をスキップ
    if (!settings.showOptional && type === '任意') return;
    
    // 推奨接種期間を解析
    const [startDate, endDate] = period.split(' 〜 ').map(d => new Date(d));
    
    // 通知期間を決定
    let reminderPeriod;
    const ageInMonths = (today - birthday) / (1000 * 60 * 60 * 24 * 30.44);
    
    if (ageInMonths < 12) {
      reminderPeriod = parsePeriod(settings.reminderPeriods.underOne);
    } else if (ageInMonths >= 24) {
      reminderPeriod = parsePeriod(settings.reminderPeriods.overTwo);
    } else {
      reminderPeriod = parsePeriod(settings.reminderPeriods.oneToTwo);
    }
    
    // 通知が必要か確認
    const notificationDate = new Date(startDate);
    notificationDate.setDate(notificationDate.getDate() - reminderPeriod);
    
    if (today >= notificationDate && today <= endDate) {
      const notification = {
        name,
        type,
        period,
        group
      };
      
      // 同時接種の推奨情報を追加
      if (settings.showSimultaneous && group) {
        const simultaneousVaccines = values
          .filter((r, i) => i !== index && r[6] === group && r[4] === '未接種')
          .map(r => r[1]);
        
        if (simultaneousVaccines.length > 0) {
          notification.simultaneous = simultaneousVaccines;
        }
      }
      
      notifications.push(notification);
    }
  });
  
  return notifications;
}

// 通知メールを送信
function sendNotificationEmail(email, childName, notifications) {
  const subject = `【予防接種リマインダー】${childName}の予防接種予定`;
  
  let body = `${childName}の予防接種リマインダーをお知らせします。\n\n`;
  
  notifications.forEach(notification => {
    body += `■ ${notification.name}（${notification.type}）\n`;
    body += `推奨接種期間：${notification.period}\n`;
    
    if (notification.simultaneous && notification.simultaneous.length > 0) {
      body += `同時接種可能なワクチン：${notification.simultaneous.join('、')}\n`;
    }
    
    body += '\n';
  });
  
  body += '※ このメールは自動送信されています。\n';
  body += '※ 実際の接種スケジュールについては、医師にご相談ください。';
  
  GmailApp.sendEmail(email, subject, body);
}

// 期間をパース（例：'1週間前' → 7）
function parsePeriod(periodString) {
  const matches = periodString.match(/(\d+)(.+)前/);
  if (!matches) return 30;  // デフォルトは30日
  
  const value = parseInt(matches[1]);
  const unit = matches[2];
  
  switch (unit) {
    case '日':
      return value;
    case '週間':
      return value * 7;
    case 'ヶ月':
      return value * 30;
    default:
      return 30;
  }
} 