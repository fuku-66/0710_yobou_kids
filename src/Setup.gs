// スプレッドシートの初期設定
function setupSpreadsheet() {
  const sheet = SpreadsheetApp.getActiveSheet();
  
  // ヘッダーの設定
  const headers = [
    ['予防接種名', '接種回数', '標準的な接種開始時期', '接種推奨期間（開始）', '接種推奨期間（終了）', 'ステータス', '予約日', 'メモ', '種別']
  ];
  sheet.getRange(3, 1, 1, headers[0].length).setValues(headers);
  
  // 生年月日入力欄の設定
  sheet.getRange('A1').setValue('生年月日：');
  sheet.getRange(BIRTHDAY_CELL).setNumberFormat('yyyy/mm/dd');
  
  // 予防接種データの設定
  const vaccineData = [
    // 乳児期（生後2か月〜1歳）
    ['B型肝炎', '1回目', '生後2ヶ月', '', '', '未接種', '', '', '定期'],
    ['B型肝炎', '2回目', '生後3ヶ月', '', '', '未接種', '', '', '定期'],
    ['B型肝炎', '3回目', '生後7ヶ月', '', '', '未接種', '', '', '定期'],
    ['ロタウイルス', '1回目', '生後2ヶ月', '', '', '未接種', '', '', '定期'],
    ['ロタウイルス', '2回目', '生後3ヶ月', '', '', '未接種', '', '', '定期'],
    ['ロタウイルス', '3回目', '生後4ヶ月', '', '', '未接種', '', '', '定期'],
    ['小児用肺炎球菌', '1回目', '生後2ヶ月', '', '', '未接種', '', '', '定期'],
    ['小児用肺炎球菌', '2回目', '生後3ヶ月', '', '', '未接種', '', '', '定期'],
    ['小児用肺炎球菌', '3回目', '生後4ヶ月', '', '', '未接種', '', '', '定期'],
    ['小児用肺炎球菌', '4回目', '生後12ヶ月', '', '', '未接種', '', '', '定期'],
    ['五種混合', '1回目', '生後2ヶ月', '', '', '未接種', '', '', '定期'],
    ['五種混合', '2回目', '生後3ヶ月', '', '', '未接種', '', '', '定期'],
    ['五種混合', '3回目', '生後4ヶ月', '', '', '未接種', '', '', '定期'],
    ['五種混合', '4回目', '生後12ヶ月', '', '', '未接種', '', '', '定期'],
    ['BCG', '1回目', '生後5ヶ月', '', '', '未接種', '', '', '定期'],
    
    // 1歳の誕生日すぐ
    ['MR（麻しん・風しん）', '1期', '生後12ヶ月', '', '', '未接種', '', '', '定期'],
    ['水痘', '1回目', '生後12ヶ月', '', '', '未接種', '', '', '定期'],
    ['水痘', '2回目', '生後15ヶ月', '', '', '未接種', '', '', '定期'],
    ['おたふくかぜ', '1回目', '生後12ヶ月', '', '', '未接種', '', '', '任意'],
    ['おたふくかぜ', '2回目', '生後24ヶ月', '', '', '未接種', '', '', '任意'],
    
    // 1歳6か月頃
    ['ヒブ', '追加接種', '生後18ヶ月', '', '', '未接種', '', '', '定期'],
    ['小児用肺炎球菌', '追加接種', '生後18ヶ月', '', '', '未接種', '', '', '定期'],
    ['五種混合', '追加接種', '生後18ヶ月', '', '', '未接種', '', '', '定期'],
    
    // 3歳
    ['日本脳炎', '1期初回1回目', '生後36ヶ月', '', '', '未接種', '', '', '定期'],
    ['日本脳炎', '1期初回2回目', '生後37ヶ月', '', '', '未接種', '', '', '定期'],
    ['日本脳炎', '1期追加', '生後48ヶ月', '', '', '未接種', '', '', '定期'],
    
    // 5〜6歳
    ['MR（麻しん・風しん）', '2期', '生後60ヶ月', '', '', '未接種', '', '', '定期'],
    
    // 9〜10歳
    ['日本脳炎', '2期', '生後108ヶ月', '', '', '未接種', '', '', '定期'],
    
    // 11〜12歳
    ['DT（ジフテリア・破傷風）', '2期', '生後132ヶ月', '', '', '未接種', '', '', '定期'],
    ['HPV', '1回目', '生後132ヶ月', '', '', '未接種', '', '', '定期'],
    ['HPV', '2回目', '生後138ヶ月', '', '', '未接種', '', '', '定期'],
    
    // 毎年の接種
    ['インフルエンザ', '毎年1回目', '生後6ヶ月', '', '', '未接種', '', '', '任意'],
    ['インフルエンザ', '毎年2回目', '生後6ヶ月', '', '', '未接種', '', '', '任意']
  ];
  
  // データの入力
  sheet.getRange(DATA_START_ROW, 1, vaccineData.length, vaccineData[0].length).setValues(vaccineData);
  
  // ステータス列にドロップダウンリストを設定
  const statusRange = sheet.getRange(DATA_START_ROW, STATUS_COLUMN, vaccineData.length, 1);
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['未接種', '予約済', '済み'])
    .build();
  statusRange.setDataValidation(statusRule);
  
  // 書式設定
  const dataRange = sheet.getRange(DATA_START_ROW, 1, vaccineData.length, headers[0].length);
  dataRange.setBorder(true, true, true, true, true, true);
  
  // 日付列の書式設定
  const dateRanges = [
    sheet.getRange(DATA_START_ROW, 4, vaccineData.length, 2),  // D列とE列
    sheet.getRange(DATA_START_ROW, APPOINTMENT_COLUMN, vaccineData.length, 1)  // G列
  ];
  dateRanges.forEach(range => range.setNumberFormat('yyyy/mm/dd'));
  
  // ヘッダーの書式設定
  const headerRange = sheet.getRange(3, 1, 1, headers[0].length);
  headerRange.setBackground('#f3f3f3');
  headerRange.setFontWeight('bold');
  headerRange.setBorder(true, true, true, true, true, true);
  
  // 種別による色分け
  const typeRange = sheet.getRange(DATA_START_ROW, 9, vaccineData.length, 1);
  const types = typeRange.getValues();
  const colors = types.map(type => {
    return [type[0] === '定期' ? '#ffffff' : '#f3f3f3'];
  });
  typeRange.setBackgrounds(colors);
  
  // 列幅の自動調整
  sheet.autoResizeColumns(1, headers[0].length);
  
  // 説明を追加
  sheet.getRange('A2').setValue('※ 定期接種は白背景、任意接種はグレー背景で表示されています。');
  sheet.getRange('A2').setFontColor('#666666');
}

// トリガーの設定
function setupTriggers() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  // 毎日午前9時にリマインダーチェックを実行
  ScriptApp.newTrigger('checkReminders')
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .create();
}

/**
 * 設定シートを作成または更新
 */
function setupSettingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('設定');
  
  if (!sheet) {
    sheet = ss.insertSheet('設定');
  }
  
  // ヘッダー行の設定
  const headers = [['項目', '値', '説明']];
  const headerRange = sheet.getRange(1, 1, 1, 3);
  headerRange.setValues(headers);
  headerRange.setBackground('#f3f3f3');
  headerRange.setFontWeight('bold');
  
  // 設定項目の定義
  const settings = [
    // メール通知設定
    ['📧 メール通知設定', '', ''],
    ['メール通知', 'ON', 'メール通知のON/OFF'],
    ['メールアドレス', '', '通知先のメールアドレス（複数の場合は改行で区切る）'],
    
    // リマインド設定
    ['⏰ リマインド設定', '', ''],
    ['全年齢のリマインド', '1ヶ月前', 'すべての予防接種の通知タイミング'],
    ['1歳未満のリマインド', '1週間前', '1歳未満の予防接種の通知タイミング'],
    ['2歳以降のリマインド', '3ヶ月前', '2歳以降の予防接種の通知タイミング'],
    
    // カレンダー設定
    ['📅 カレンダー設定', '', ''],
    ['カレンダーID', '', 'Googleカレンダーの連携用ID'],
    ['カレンダー予定の長さ', '60', '予定の長さ（分）'],
    
    // 表示設定
    ['👀 表示設定', '', ''],
    ['任意接種の表示', 'ON', '任意接種の表示/非表示'],
    ['同時接種の推奨表示', 'ON', '同時接種可能な組み合わせの表示']
  ];
  
  // 設定値の入力
  const settingsRange = sheet.getRange(2, 1, settings.length, 3);
  settingsRange.setValues(settings);
  
  // 列幅の設定
  sheet.setColumnWidth(1, 200);  // 項目
  sheet.setColumnWidth(2, 150);  // 値
  sheet.setColumnWidth(3, 300);  // 説明
  
  // データ入力規則の設定
  const onOffRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['ON', 'OFF'], true)
    .build();
  
  const reminderRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([
      '3日前',
      '1週間前',
      '2週間前',
      '1ヶ月前',
      '2ヶ月前',
      '3ヶ月前'
    ], true)
    .build();
  
  const durationRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['30', '60', '90', '120'], true)
    .build();
  
  // ON/OFF選択肢の設定
  sheet.getRange('B2').setDataValidation(onOffRule);  // メール通知
  sheet.getRange('B12').setDataValidation(onOffRule); // 任意接種の表示
  sheet.getRange('B13').setDataValidation(onOffRule); // 同時接種の推奨表示
  
  // リマインド時期の選択肢
  sheet.getRange('B5').setDataValidation(reminderRule); // 全年齢
  sheet.getRange('B6').setDataValidation(reminderRule); // 1歳未満
  sheet.getRange('B7').setDataValidation(reminderRule); // 2歳以降
  
  // カレンダー予定の長さの選択肢
  sheet.getRange('B10').setDataValidation(durationRule);
  
  // セクション見出しの書式設定
  const sectionRows = [2, 5, 9, 12]; // セクション見出しの行番号
  sectionRows.forEach(row => {
    sheet.getRange(row, 1).setFontWeight('bold');
    sheet.getRange(row, 1, 1, 3).setBackground('#e8eaf6');
  });
  
  // 説明セルの書式設定
  const lastRow = settings.length + 1;
  sheet.getRange(2, 3, lastRow - 1, 1).setWrap(true);
  
  // シートの保護
  const protection = sheet.protect();
  protection.setDescription('設定シートの保護');
  protection.setWarningOnly(true);
} 