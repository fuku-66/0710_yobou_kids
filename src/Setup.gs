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