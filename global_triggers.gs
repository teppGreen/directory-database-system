function setTriggers() { //このシートのトリガーを設定
  deleteTriggers();
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('onFormSubmitFunctions').forSpreadsheet(sheet).onFormSubmit().create();
  ScriptApp.newTrigger('checkExpiration').timeBased().everyDays(1).atHour(0).create();
}

function deleteTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) ScriptApp.deleteTrigger(trigger);
}

function onFormSubmitFunctions(e) {
  try {
    console.log(`▼${arguments.callee.name}`);

    addPersons(e.range.getRow());
  } catch(error) {
    sendNotificationToSlack(error);
    throw new Error(error.stack);
  }
}

function notifyError(error) {
  // return;
  
  const now = new Date();
  const datetime = Utilities.formatDate(now, 'JST', 'MM/dd HH:mm');
  const sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const scriptUrl = 'https://script.google.com/home/projects/' + ScriptApp.getScriptId() + '/executions';

  const to = PropertiesService.getScriptProperties().getProperty('systemManagerEmails');
  const subject = `【${datetime}】【障害】名簿データベースでエラー発生`;
  const body = 
    '名簿データベースでエラーが発生しました。対応が必要な可能性がありますので、以下の内容を確認してください。' + 
    '\n\nエラー発生日時: ' + Utilities.formatDate(now, 'JST', 'yyyy/MM/dd(E) HH:mm') + 
    '\n' + error.stack + 
    '\n\nリソース管理シート: ' + sheetUrl + 
    '\nAppsScript: ' + scriptUrl

  const options = { name: 'リソース管理シート GASトリガー' };
  GmailApp.sendEmail(to,subject,body,options);
  
}