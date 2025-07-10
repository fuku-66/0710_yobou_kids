// 新規スプレッドシートを作成
function createNewVaccinationSheet() {
  // 新しいスプレッドシートを作成
  const newSpreadsheet = SpreadsheetApp.create('予防接種管理表');
  const sheet = newSpreadsheet.getActiveSheet();
  
  // シート名を設定
  sheet.setName('予防接種スケジュール');
  
  // 説明文を追加
  sheet.getRange('A2').setValue('※このスプレッドシートは予防接種のスケジュール管理用です。');
  sheet.getRange('A2').setFontColor('#666666');
  
  // スプレッドシートのURLを取得
  const url = newSpreadsheet.getUrl();
  
  // スプレッドシートの初期設定を実行
  setupSpreadsheet();
  
  // トリガーの設定
  setupTriggers();
  
  // 作成したスプレッドシートのURLをログに出力
  Logger.log('新しいスプレッドシートを作成しました: ' + url);
  
  return url;
}

function createScheduleSheet(childName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = `${childName}のスケジュール`;
  let sheet = ss.getSheetByName(sheetName);
  
  if (sheet) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '確認',
      'すでに同じ名前のシートが存在します。新しく作り直しますか？',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      ss.deleteSheet(sheet);
    } else {
      return null;
    }
  }
  
  sheet = ss.insertSheet(sheetName);
  
  // ヘッダー行の設定
  const headers = [
    'ワクチン名',
    '接種回数',
    '推奨時期',
    '接種間隔',
    '予約日',
    'ステータス',
    'メモ',
    'カレンダーEventID'  // カレンダー連携用の列を追加
  ];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#f3f3f3');
  headerRange.setFontWeight('bold');
  
  // 列幅の設定
  sheet.setColumnWidth(1, 150);  // ワクチン名
  sheet.setColumnWidth(2, 80);   // 接種回数
  sheet.setColumnWidth(3, 150);  // 推奨時期
  sheet.setColumnWidth(4, 150);  // 接種間隔
  sheet.setColumnWidth(5, 100);  // 予約日
  sheet.setColumnWidth(6, 80);   // ステータス
  sheet.setColumnWidth(7, 200);  // メモ
  sheet.setColumnWidth(8, 0);    // カレンダーEventID（非表示）
  
  // データ入力規則の設定
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['未接種', '予約済', '済み'], true)
    .build();
  const statusRange = sheet.getRange('F2:F100');
  statusRange.setDataValidation(statusRule);
  
  // 条件付き書式の設定
  const lastColumn = sheet.getLastColumn();
  const dataRange = sheet.getRange(2, 1, 99, lastColumn);
  
  // 期限切れ（赤色）
  const redRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormula(isOverdue)
    .setBackground('#ffcdd2')
    .setRanges([dataRange])
    .build();
  
  // 1ヶ月以内（黄色）
  const yellowRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormula(isWithinOneMonth)
    .setBackground('#fff9c4')
    .setRanges([dataRange])
    .build();
  
  // 接種済み（グレー）
  const greyRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormula(isCompleted)
    .setBackground('#f5f5f5')
    .setRanges([dataRange])
    .build();
  
  sheet.setConditionalFormatRules([redRule, yellowRule, greyRule]);
  
  // デフォルトの予防接種データを入力
  const vaccineData = getDefaultVaccineData();
  if (vaccineData.length > 0) {
    const dataRange = sheet.getRange(2, 1, vaccineData.length, vaccineData[0].length);
    dataRange.setValues(vaccineData);
  }
  
  return sheet;
} 