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