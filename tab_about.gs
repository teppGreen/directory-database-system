function copyCode() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('about');
  const code = sheet.getRange('F3').getValue().replace(/\n/g, '\\n');
  Browser.msgBox("下記の関数をコピーしてください",code, Browser.Buttons.OK);
}