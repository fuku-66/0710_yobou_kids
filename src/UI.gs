// スプレッドシートが開かれたときに実行
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // カスタムメニューを作成
  ui.createMenu('予防接種管理')
    .addItem('新規スケジュール作成', 'createNewVaccinationSheet')
    .addSeparator()
    .addSubMenu(ui.createMenu('子供の管理')
      .addItem('新しい子供を追加', 'addNewChild')
      .addItem('子供の情報を編集', 'editChild')
      .addItem('子供の情報を削除', 'deleteChild'))
    .addSeparator()
    .addItem('設定', 'showSettings')
    .addSeparator()
    .addItem('使い方', 'showHelp')
    .addToUi();
}

// 子供の管理関連の関数
function addNewChild() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let childrenSheet = ss.getSheetByName('子供の情報');
  
  if (!childrenSheet) {
    childrenSheet = createChildrenSheet();
  }
  
  const result = ui.prompt(
    '新しい子供を追加',
    '子供の名前を入力してください：',
    ui.ButtonSet.OK_CANCEL);
  
  const button = result.getSelectedButton();
  if (button === ui.Button.OK) {
    const name = result.getResponseText();
    const birthdayResult = ui.prompt(
      '生年月日の入力',
      '生年月日を入力してください（yyyy/mm/dd形式）：',
      ui.ButtonSet.OK_CANCEL);
    
    if (birthdayResult.getSelectedButton() === ui.Button.OK) {
      const birthday = new Date(birthdayResult.getResponseText());
      if (birthday.toString() === 'Invalid Date') {
        ui.alert('エラー', '正しい日付形式で入力してください。', ui.ButtonSet.OK);
        return;
      }
      
      // 子供の情報を追加
      const lastRow = childrenSheet.getLastRow();
      childrenSheet.getRange(lastRow + 1, 1, 1, 2).setValues([[name, birthday]]);
      
      // スケジュールシートを作成
      createScheduleSheet(name, birthday);
      
      ui.alert('完了', '新しい子供の情報が追加されました。', ui.ButtonSet.OK);
    }
  }
}

function editChild() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const childrenSheet = ss.getSheetByName('子供の情報');
  
  if (!childrenSheet || childrenSheet.getLastRow() <= 1) {
    ui.alert('エラー', '編集可能な子供の情報がありません。', ui.ButtonSet.OK);
    return;
  }
  
  const children = childrenSheet.getRange(2, 1, childrenSheet.getLastRow() - 1, 2).getValues();
  let childList = '現在登録されている子供：\n';
  children.forEach((child, index) => {
    childList += `${index + 1}. ${child[0]} (${formatDate(child[1])})\n`;
  });
  
  const result = ui.prompt(
    '子供の情報を編集',
    childList + '\n編集したい子供の番号を入力してください：',
    ui.ButtonSet.OK_CANCEL);
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const index = parseInt(result.getResponseText()) - 1;
    if (isNaN(index) || index < 0 || index >= children.length) {
      ui.alert('エラー', '正しい番号を入力してください。', ui.ButtonSet.OK);
      return;
    }
    
    const nameResult = ui.prompt(
      '名前の編集',
      `現在の名前: ${children[index][0]}\n新しい名前を入力してください（変更しない場合は空欄）：`,
      ui.ButtonSet.OK_CANCEL);
    
    if (nameResult.getSelectedButton() === ui.Button.OK) {
      const newName = nameResult.getResponseText() || children[index][0];
      
      const birthdayResult = ui.prompt(
        '生年月日の編集',
        `現在の生年月日: ${formatDate(children[index][1])}\n新しい生年月日を入力してください（yyyy/mm/dd形式、変更しない場合は空欄）：`,
        ui.ButtonSet.OK_CANCEL);
      
      if (birthdayResult.getSelectedButton() === ui.Button.OK) {
        const newBirthday = birthdayResult.getResponseText() ? 
          new Date(birthdayResult.getResponseText()) : children[index][1];
        
        if (newBirthday.toString() === 'Invalid Date') {
          ui.alert('エラー', '正しい日付形式で入力してください。', ui.ButtonSet.OK);
          return;
        }
        
        // 情報を更新
        childrenSheet.getRange(index + 2, 1, 1, 2).setValues([[newName, newBirthday]]);
        
        // シート名を更新
        const oldSheet = ss.getSheetByName(children[index][0]);
        if (oldSheet) {
          oldSheet.setName(newName);
          // 生年月日を更新して再計算
          updateScheduleSheet(oldSheet, newName, newBirthday);
        }
        
        ui.alert('完了', '子供の情報が更新されました。', ui.ButtonSet.OK);
      }
    }
  }
}

function deleteChild() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const childrenSheet = ss.getSheetByName('子供の情報');
  
  if (!childrenSheet || childrenSheet.getLastRow() <= 1) {
    ui.alert('エラー', '削除可能な子供の情報がありません。', ui.ButtonSet.OK);
    return;
  }
  
  const children = childrenSheet.getRange(2, 1, childrenSheet.getLastRow() - 1, 2).getValues();
  let childList = '現在登録されている子供：\n';
  children.forEach((child, index) => {
    childList += `${index + 1}. ${child[0]} (${formatDate(child[1])})\n`;
  });
  
  const result = ui.prompt(
    '子供の情報を削除',
    childList + '\n削除したい子供の番号を入力してください：',
    ui.ButtonSet.OK_CANCEL);
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const index = parseInt(result.getResponseText()) - 1;
    if (isNaN(index) || index < 0 || index >= children.length) {
      ui.alert('エラー', '正しい番号を入力してください。', ui.ButtonSet.OK);
      return;
    }
    
    const confirmResult = ui.alert(
      '確認',
      `${children[index][0]}の情報を削除します。この操作は取り消せません。\n続行しますか？`,
      ui.ButtonSet.YES_NO);
    
    if (confirmResult === ui.Button.YES) {
      // シートを削除
      const sheet = ss.getSheetByName(children[index][0]);
      if (sheet) {
        ss.deleteSheet(sheet);
      }
      
      // 子供の情報を削除
      childrenSheet.deleteRow(index + 2);
      
      ui.alert('完了', '子供の情報が削除されました。', ui.ButtonSet.OK);
    }
  }
}

// 子供の情報シートを作成
function createChildrenSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.insertSheet('子供の情報');
  
  // ヘッダーの設定
  sheet.getRange('A1:B1').setValues([['名前', '生年月日']]);
  sheet.getRange('A1:B1').setBackground('#f3f3f3').setFontWeight('bold');
  
  // 書式設定
  sheet.getRange('B:B').setNumberFormat('yyyy/mm/dd');
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 150);
  
  return sheet;
}

// 設定画面を表示
function showSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let settingsSheet = ss.getSheetByName('設定');
  
  // 設定シートが存在しない場合は作成
  if (!settingsSheet) {
    settingsSheet = createSettingsSheet();
  }
  
  // 設定シートをアクティブにする
  settingsSheet.activate();
}

// 設定シートを作成
function createSettingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.insertSheet('設定');
  
  // ヘッダーの設定
  settingsSheet.getRange('A1:B1').merge()
    .setValue('基本設定')
    .setBackground('#4285f4')
    .setFontColor('white')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  // セクション1: メール通知設定
  settingsSheet.getRange('A3').setValue('📧 メール通知設定')
    .setFontWeight('bold')
    .setBackground('#e8f0fe');
  settingsSheet.getRange('A4:B4').setValues([['メール通知を有効にする', 'はい']]);
  settingsSheet.getRange('A5:B5').setValues([['通知先メールアドレス', '']]);
  settingsSheet.getRange('A6:B6').setValues([['（複数の場合は「,」で区切って入力）', '例：mama@example.com, papa@example.com']]);
  
  // セクション2: リマインド設定
  settingsSheet.getRange('A8').setValue('⏰ リマインド設定')
    .setFontWeight('bold')
    .setBackground('#e8f0fe');
  settingsSheet.getRange('A9:B9').setValues([['リマインド時期（1歳未満）', '1週間前']]);
  settingsSheet.getRange('A10:B10').setValues([['リマインド時期（1-2歳）', '1ヶ月前']]);
  settingsSheet.getRange('A11:B11').setValues([['リマインド時期（2歳以降）', '3ヶ月前']]);
  
  // セクション3: 予約設定
  settingsSheet.getRange('A13').setValue('📅 予約設定')
    .setFontWeight('bold')
    .setBackground('#e8f0fe');
  settingsSheet.getRange('A14:B14').setValues([['予約時間の長さ', '60分']]);
  
  // セクション4: 表示設定
  settingsSheet.getRange('A16').setValue('👀 表示設定')
    .setFontWeight('bold')
    .setBackground('#e8f0fe');
  settingsSheet.getRange('A17:B17').setValues([['任意接種を表示する', 'はい']]);
  settingsSheet.getRange('A18:B18').setValues([['同時接種の推奨を表示する', 'はい']]);
  
  // 選択肢の設定
  // はい/いいえの選択肢
  const yesNoRanges = [
    settingsSheet.getRange('B4'),  // メール通知
    settingsSheet.getRange('B17'), // 任意接種
    settingsSheet.getRange('B18')  // 同時接種
  ];
  
  const yesNoRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['はい', 'いいえ'])
    .build();
  
  yesNoRanges.forEach(range => range.setDataValidation(yesNoRule));
  
  // リマインド時期の選択肢
  const reminderRanges = [
    settingsSheet.getRange('B9'),  // 1歳未満
    settingsSheet.getRange('B10'), // 1-2歳
    settingsSheet.getRange('B11')  // 2歳以降
  ];
  
  const reminderOptions = [
    '3日前',
    '1週間前',
    '2週間前',
    '1ヶ月前',
    '2ヶ月前',
    '3ヶ月前'
  ];
  
  const reminderRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(reminderOptions)
    .build();
  
  reminderRanges.forEach(range => range.setDataValidation(reminderRule));
  
  // 予約時間の選択肢
  const durationRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['30分', '45分', '60分', '90分'])
    .build();
  settingsSheet.getRange('B14').setDataValidation(durationRule);
  
  // 書式設定
  settingsSheet.getRange('A1:B18').setBorder(true, true, true, true, true, true);
  settingsSheet.setColumnWidth(1, 300);
  settingsSheet.setColumnWidth(2, 300);
  
  // 説明文を追加
  const explanations = [
    ['💡 設定の使い方'],
    ['1. メール通知設定'],
    ['   ・「メール通知を有効にする」を「はい」にすると、予防接種の時期が近づいたときにメールでお知らせします'],
    ['   ・通知先のメールアドレスを入力してください（ご両親それぞれのアドレスを入力できます）'],
    ['2. リマインド設定'],
    ['   ・お子様の年齢に応じて、いつ前にお知らせするかを選べます'],
    ['   ・1歳未満は接種が多いので、1週間前がおすすめです'],
    ['3. 予約設定'],
    ['   ・病院での予約時間の長さを選択します（カレンダーに予定を追加するときに使用）'],
    ['4. 表示設定'],
    ['   ・任意接種（おたふくかぜなど）を表示するかどうかを選べます'],
    ['   ・同時接種の推奨を表示すると、一緒に受けられる予防接種がわかります'],
    [''],
    ['※ 設定を変更したら、すぐに反映されます'],
    ['※ 実際の予防接種スケジュールは、必ず医師にご相談ください']
  ];
  
  const startRow = 20;
  explanations.forEach((text, index) => {
    const row = startRow + index;
    settingsSheet.getRange(row, 1, 1, 2).merge()
      .setValue(text[0])
      .setFontColor(index === 0 ? '#4285f4' : '#666666')
      .setFontWeight(index === 0 ? 'bold' : 'normal');
  });
  
  // 入力規則の説明を追加
  settingsSheet.getRange('B5').setNote('メールアドレスを入力してください\n例：mama@example.com');
  settingsSheet.getRange('B6').setFontColor('#666666').setFontStyle('italic');
  
  return settingsSheet;
}

// 日付のフォーマット
function formatDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd');
}

// 使い方を表示
function showHelp() {
  const ui = SpreadsheetApp.getUi();
  const helpText = 
    '【予防接種管理ツールの使い方】\n\n' +
    '1. 基本的な使い方\n' +
    '   - メニューから「子供の管理」→「新しい子供を追加」で子供の情報を登録\n' +
    '   - 各子供専用のシートが作成され、予防接種スケジュールが自動計算されます\n' +
    '   - 予約が確定したら、予約日列に日時を入力してください\n' +
    '   - 接種が完了したら、ステータス列を「済み」に更新してください\n\n' +
    '2. 視覚的な進捗管理\n' +
    '   - 接種期間が1ヶ月以内：黄色で表示\n' +
    '   - 接種期限切れ：赤色で表示\n' +
    '   - 接種済み：グレー色で表示\n' +
    '   - 定期接種：白背景\n' +
    '   - 任意接種：グレー背景\n\n' +
    '3. 同時接種の推奨\n' +
    '   - 同じ時期に接種可能なワクチンがグループ化されて表示されます\n' +
    '   - グループごとに推奨される同時接種の組み合わせが表示されます\n\n' +
    '4. リマインダー機能\n' +
    '   - 設定した期間に応じて、自動的にメール通知が送信されます\n' +
    '   - 通知の設定は「設定」メニューから変更できます\n' +
    '   - 複数のメールアドレスに通知を送信できます\n\n' +
    '5. カレンダー連携\n' +
    '   - 予約日を入力すると、自動的にGoogleカレンダーに予定が追加されます\n\n' +
    '6. 注意事項\n' +
    '   - このツールは参考情報として使用してください\n' +
    '   - 実際の予防接種スケジュールは医師に相談の上で決定してください\n' +
    '   - 予防接種のスケジュールは地域や状況により異なる場合があります';
  
  ui.alert('使い方', helpText, ui.ButtonSet.OK);
} 