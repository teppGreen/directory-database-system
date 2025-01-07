function test1() {
  addPersons(2);
}

function test2() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('persons');
  let personSheetRow = sheet.getRange(1,2).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow() + 1;
  console.log(personSheetRow)
  console.log(sheet.getMaxRows())
}