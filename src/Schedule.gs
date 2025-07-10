// スケジュールシートを作成
function createScheduleSheet(childName, birthday) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.insertSheet(childName);
  
  // ヘッダーの設定
  const headers = [
    ['対象者', childName, '', '生年月日', formatDate(birthday)],
    ['', '', '', '', ''],
    ['種類', 'ワクチン名', '推奨接種期間', '予約日', 'ステータス', '備考', 'グループ']
  ];
  sheet.getRange('A1:E2').setValues(headers.slice(0, 2));
  sheet.getRange('A3:G3').setValues([headers[2]]);
  
  // ヘッダーの書式設定
  sheet.getRange('A1:E2').setBackground('#f3f3f3');
  sheet.getRange('A3:G3').setBackground('#f3f3f3').setFontWeight('bold');
  
  // 列幅の設定
  sheet.setColumnWidth(1, 100);  // 種類
  sheet.setColumnWidth(2, 200);  // ワクチン名
  sheet.setColumnWidth(3, 200);  // 推奨接種期間
  sheet.setColumnWidth(4, 150);  // 予約日
  sheet.setColumnWidth(5, 100);  // ステータス
  sheet.setColumnWidth(6, 200);  // 備考
  sheet.setColumnWidth(7, 100);  // グループ
  
  // ワクチン情報を追加
  const vaccines = getVaccineSchedule();
  const values = vaccines.map(v => [
    v.type,
    v.name,
    calculateRecommendedPeriod(birthday, v.startAge, v.endAge),
    '',
    '未接種',
    v.note || '',
    v.group || ''
  ]);
  
  const dataRange = sheet.getRange(4, 1, values.length, 7);
  dataRange.setValues(values);
  
  // ステータスの選択肢を設定
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['未接種', '予約済', '済み'])
    .build();
  sheet.getRange(4, 5, values.length, 1).setDataValidation(statusRule);
  
  // 予約日の書式設定
  sheet.getRange(4, 4, values.length, 1).setNumberFormat('yyyy/mm/dd');
  
  // 条件付き書式の設定
  setConditionalFormatting(sheet, 4, values.length);
  
  // 同時接種のグループ化
  groupSimultaneousVaccinations(sheet, 4, values.length);
  
  return sheet;
}

// ワクチンスケジュールの定義
function getVaccineSchedule() {
  return [
    // 定期接種（A類疾病）
    { type: '定期', name: 'ロタウイルス（1回目）', startAge: '生後2ヶ月', endAge: '生後14週6日', group: 'A' },
    { type: '定期', name: 'ロタウイルス（2回目）', startAge: '1回目から4週間後', endAge: '生後24週0日', group: 'A' },
    { type: '定期', name: 'B型肝炎（1回目）', startAge: '生後2ヶ月', endAge: '生後9ヶ月', group: 'A' },
    { type: '定期', name: 'B型肝炎（2回目）', startAge: '1回目から27日以上', endAge: '生後9ヶ月', group: 'B' },
    { type: '定期', name: 'B型肝炎（3回目）', startAge: '1回目から139日以上', endAge: '生後12ヶ月', group: 'C' },
    { type: '定期', name: 'ヒブ（1回目）', startAge: '生後2ヶ月', endAge: '生後7ヶ月', group: 'A' },
    { type: '定期', name: 'ヒブ（2回目）', startAge: '1回目から27日以上', endAge: '生後7ヶ月', group: 'B' },
    { type: '定期', name: 'ヒブ（3回目）', startAge: '1回目から139日以上', endAge: '生後12ヶ月', group: 'C' },
    { type: '定期', name: 'ヒブ（4回目）', startAge: '生後12ヶ月', endAge: '生後15ヶ月', group: 'D' },
    { type: '定期', name: '小児用肺炎球菌（1回目）', startAge: '生後2ヶ月', endAge: '生後7ヶ月', group: 'A' },
    { type: '定期', name: '小児用肺炎球菌（2回目）', startAge: '1回目から27日以上', endAge: '生後7ヶ月', group: 'B' },
    { type: '定期', name: '小児用肺炎球菌（3回目）', startAge: '1回目から139日以上', endAge: '生後12ヶ月', group: 'C' },
    { type: '定期', name: '小児用肺炎球菌（4回目）', startAge: '生後12ヶ月', endAge: '生後15ヶ月', group: 'D' },
    { type: '定期', name: '4種混合（1回目）', startAge: '生後3ヶ月', endAge: '生後12ヶ月', group: 'B' },
    { type: '定期', name: '4種混合（2回目）', startAge: '1回目から20日以上', endAge: '生後12ヶ月', group: 'C' },
    { type: '定期', name: '4種混合（3回目）', startAge: '2回目から20日以上', endAge: '生後12ヶ月', group: 'D' },
    { type: '定期', name: '4種混合（追加）', startAge: '3回目から6ヶ月以上', endAge: '1歳6ヶ月', group: 'E' },
    { type: '定期', name: 'BCG', startAge: '生後5ヶ月', endAge: '生後8ヶ月', group: 'C' },
    { type: '定期', name: 'MR（1期）', startAge: '1歳', endAge: '2歳', group: 'D' },
    { type: '定期', name: 'MR（2期）', startAge: '5歳', endAge: '7歳', group: 'F' },
    { type: '定期', name: '水痘（1回目）', startAge: '1歳', endAge: '1歳3ヶ月', group: 'D' },
    { type: '定期', name: '水痘（2回目）', startAge: '1回目から6ヶ月以上', endAge: '3歳', group: 'E' },
    { type: '定期', name: '日本脳炎（1期初回1回目）', startAge: '3歳', endAge: '4歳', group: 'E' },
    { type: '定期', name: '日本脳炎（1期初回2回目）', startAge: '1回目から6日以上', endAge: '4歳', group: 'E' },
    { type: '定期', name: '日本脳炎（1期追加）', startAge: '2回目から6ヶ月以上', endAge: '5歳', group: 'F' },
    { type: '定期', name: '日本脳炎（2期）', startAge: '9歳', endAge: '13歳', group: 'G' },
    
    // 任意接種
    { type: '任意', name: 'おたふくかぜ（1回目）', startAge: '1歳', endAge: '2歳', group: 'D', note: '任意接種' },
    { type: '任意', name: 'おたふくかぜ（2回目）', startAge: '1回目から1ヶ月以上', endAge: '7歳6ヶ月', group: 'E', note: '任意接種' },
    { type: '任意', name: 'インフルエンザ', startAge: '生後6ヶ月', endAge: '13歳', group: 'H', note: '毎年10月〜12月に接種' },
    { type: '任意', name: 'A型肝炎（1回目）', startAge: '1歳', endAge: '13歳', group: 'E', note: '任意接種' },
    { type: '任意', name: 'A型肝炎（2回目）', startAge: '1回目から2週間以上', endAge: '13歳', group: 'F', note: '任意接種' },
    { type: '任意', name: 'A型肝炎（3回目）', startAge: '1回目から20週間以上', endAge: '13歳', group: 'G', note: '任意接種' }
  ];
}

// 推奨接種期間を計算
function calculateRecommendedPeriod(birthday, startAge, endAge) {
  const startDate = calculateDate(birthday, startAge);
  const endDate = calculateDate(birthday, endAge);
  return `${formatDate(startDate)} 〜 ${formatDate(endDate)}`;
}

// 日付を計算
function calculateDate(baseDate, ageString) {
  const date = new Date(baseDate);
  
  if (ageString.includes('から')) {
    // 相対的な期間（例：「1回目から27日以上」）の場合は、基準日をそのまま返す
    return date;
  }
  
  // 生後X週の場合
  const weekMatch = ageString.match(/生後(\d+)週(\d+)?日?/);
  if (weekMatch) {
    const weeks = parseInt(weekMatch[1]);
    const days = weekMatch[2] ? parseInt(weekMatch[2]) : 0;
    date.setDate(date.getDate() + (weeks * 7) + days);
    return date;
  }
  
  // 生後Xヶ月の場合
  const monthMatch = ageString.match(/生後(\d+)ヶ月/);
  if (monthMatch) {
    const months = parseInt(monthMatch[1]);
    date.setMonth(date.getMonth() + months);
    return date;
  }
  
  // X歳Y(ヶ月)の場合
  const yearMonthMatch = ageString.match(/(\d+)歳(?:(\d+)ヶ月)?/);
  if (yearMonthMatch) {
    const years = parseInt(yearMonthMatch[1]);
    const months = yearMonthMatch[2] ? parseInt(yearMonthMatch[2]) : 0;
    date.setFullYear(date.getFullYear() + years);
    date.setMonth(date.getMonth() + months);
    return date;
  }
  
  return date;
}

// 条件付き書式を設定
function setConditionalFormatting(sheet, startRow, rowCount) {
  const range = sheet.getRange(startRow, 1, rowCount, 7);
  
  // 任意接種の背景色を設定
  const typeRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormula(`=$A${startRow}="任意"`)
    .setBackground('#f3f3f3')
    .setRanges([range])
    .build();
  
  // ステータスに応じた行の色を設定
  const statusRules = [
    // 済み
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormula(`=$E${startRow}="済み"`)
      .setBackground('#e0e0e0')
      .setRanges([range])
      .build(),
    
    // 予約済
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormula(`=$E${startRow}="予約済"`)
      .setBackground('#fff3e0')
      .setRanges([range])
      .build(),
    
    // 期限切れ
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormula(`=AND($E${startRow}="未接種", RIGHT($C${startRow}, 10) < TEXT(TODAY(), "yyyy/mm/dd"))`)
      .setBackground('#ffebee')
      .setRanges([range])
      .build(),
    
    // 1ヶ月以内
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormula(`=AND($E${startRow}="未接種", LEFT($C${startRow}, 10) <= TEXT(TODAY()+30, "yyyy/mm/dd"), RIGHT($C${startRow}, 10) >= TEXT(TODAY(), "yyyy/mm/dd"))`)
      .setBackground('#fff9c4')
      .setRanges([range])
      .build()
  ];
  
  // 条件付き書式を適用
  const rules = sheet.getConditionalFormatRules();
  sheet.setConditionalFormatRules([...rules, typeRule, ...statusRules]);
}

// 同時接種のグループ化
function groupSimultaneousVaccinations(sheet, startRow, rowCount) {
  const range = sheet.getRange(startRow, 1, rowCount, 7);
  const values = range.getValues();
  
  // グループごとに推奨される同時接種の組み合わせを追加
  const groups = {};
  values.forEach((row, index) => {
    const group = row[6];  // グループ列
    if (group && row[0] !== '済み') {  // 未接種のワクチンのみ
      if (!groups[group]) {
        groups[group] = [];
      }
      groups[group].push({
        name: row[1],
        index: index + startRow
      });
    }
  });
  
  // グループ情報を備考欄に追加
  Object.values(groups).forEach(vaccines => {
    if (vaccines.length > 1) {
      const names = vaccines.map(v => v.name).join('、');
      const note = `推奨同時接種グループ：${names}`;
      vaccines.forEach(v => {
        const currentNote = sheet.getRange(v.index, 6).getValue();
        sheet.getRange(v.index, 6).setValue(currentNote ? `${currentNote}\n${note}` : note);
      });
    }
  });
}

// スケジュールシートを更新
function updateScheduleSheet(sheet, childName, birthday) {
  // ヘッダー情報を更新
  sheet.getRange('B1').setValue(childName);
  sheet.getRange('E1').setValue(formatDate(birthday));
  
  // 既存のワクチン情報を取得
  const dataRange = sheet.getRange(4, 1, sheet.getLastRow() - 3, 7);
  const values = dataRange.getValues();
  
  // 推奨接種期間を再計算
  const vaccines = getVaccineSchedule();
  values.forEach((row, index) => {
    const vaccine = vaccines.find(v => v.name === row[1]);
    if (vaccine) {
      row[2] = calculateRecommendedPeriod(birthday, vaccine.startAge, vaccine.endAge);
    }
  });
  
  // 更新した値を設定
  dataRange.setValues(values);
  
  // 条件付き書式を再設定
  setConditionalFormatting(sheet, 4, values.length);
  
  // 同時接種のグループを再設定
  groupSimultaneousVaccinations(sheet, 4, values.length);
}

/**
 * 月齢を年齢表示に変換
 * @param {number} months 月齢
 * @return {string} 変換後の表示
 */
function formatAgeDisplay(months) {
  if (months <= 12) {
    return `生後${months}ヶ月`;
  } else {
    const years = Math.floor(months / 12);
    const remainingMonths = months % 12;
    if (remainingMonths === 0) {
      return `${years}歳`;
    } else {
      return `${years}歳${remainingMonths}ヶ月`;
    }
  }
}

/**
 * デフォルトの予防接種データを取得
 */
function getDefaultVaccineData() {
  return [
    ['ヒブ', '3回目', '生後4ヶ月', '', '', '未接種', '', '', '定期'],
    ['ヒブ', '4回目', formatAgeDisplay(12), '', '', '未接種', '', '', '定期'],
    ['小児用肺炎球菌', '1回目', '生後5ヶ月', '', '', '未接種', '', '', '定期'],
    ['小児用肺炎球菌', '1期', formatAgeDisplay(12), '', '', '未接種', '', '', '定期'],
    ['B型肝炎', '1回目', formatAgeDisplay(12), '', '', '未接種', '', '', '定期'],
    ['B型肝炎', '2回目', formatAgeDisplay(15), '', '', '未接種', '', '', '定期'],
    ['ロタウイルス', '1回目', formatAgeDisplay(12), '', '', '未接種', '', '', '任意'],
    ['ロタウイルス', '2回目', formatAgeDisplay(24), '', '', '未接種', '', '', '任意'],
    ['四種混合', '追加接種', formatAgeDisplay(18), '', '', '未接種', '', '', '定期'],
    ['麻しん・風しん混合', '追加接種', formatAgeDisplay(18), '', '', '未接種', '', '', '定期'],
    ['水痘', '追加接種', formatAgeDisplay(18), '', '', '未接種', '', '', '定期'],
    ['日本脳炎', '1期初回', formatAgeDisplay(36), '', '', '未接種', '', '', '定期'],
    ['日本脳炎', '1期初回', formatAgeDisplay(37), '', '', '未接種', '', '', '定期'],
    ['日本脳炎', '1期追加', formatAgeDisplay(48), '', '', '未接種', '', '', '定期'],
    ['日本脳炎', '2期', formatAgeDisplay(60), '', '', '未接種', '', '', '定期'],
    ['二種混合', '2期', formatAgeDisplay(108), '', '', '未接種', '', '', '定期'],
    ['日本脳炎', '2期', formatAgeDisplay(132), '', '', '未接種', '', '', '定期'],
    ['HPV', '1回目', formatAgeDisplay(132), '', '', '未接種', '', '', '定期'],
    ['HPV', '2回目', formatAgeDisplay(138), '', '', '未接種', '', '', '定期'],
    ['インフルエンザ', '毎年1回', '生後6ヶ月', '', '', '未接種', '', '', '任意']
  ];
} 