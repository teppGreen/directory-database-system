function addPersons(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName('form');
  const personSheet = ss.getSheetByName('persons');

  let items = {
    'Nickname': 'Slack 表示名',
    'Family Name': '氏名（姓）',
    'Given Name': '氏名（名）',
    'Family Name Yomi': '氏名よみがな（姓）',
    'Given Name Yomi': '氏名よみがな（名）',
    'E-mail 1 - Value': 'メールアドレス',
    'Phone 1 - Value': '電話番号',
    'Location': '居住都道府県',
    'Birthday': '生年月日',
    'Organization 1 - Department': '学び方',
    'Organization 1 - Job Description': 'コース',
    'Organization 1 - Location': 'キャンパス',
    'Organization 1 - Title': '学年',
    'Organization 2 - Name': '部署',
    'Organization 2 - Title': '役職',
    'Website 1 - Value': '関連リンク',
    'Slack ID': 'Slack メンバーID',
    'Slack Image': 'Slack プロフィール画像',
    'Notes': '備考',
    'Last Submission': 'タイムスタンプ',
    'Expiration': '有効期限'
  }

  let items_column = {};
  let items_value = {};

  for (const key in items) {
    const header_form = formSheet.getRange(1,1,1,formSheet.getLastColumn()).getValues().flat();
    const header_person = personSheet.getRange(1,1,1,personSheet.getLastColumn()).getValues().flat();
    const column_form = header_form.indexOf(items[key]) + 1;

    items_value[key] = formSheet.getRange(row,column_form).getValue();
    items_column[key] = header_person.indexOf(key) + 1;
  }

  let personSheetRow = personSheet.getRange(1,items_column['E-mail 1 - Value']).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow() + 1;
  if (personSheetRow === personSheet.getMaxRows()+1) personSheetRow = 2;

  //メアド検索して上書きか新規か判定
  const currentEmails = personSheet.getRange(1,items_column['E-mail 1 - Value'],personSheetRow,1).getValues().flat();
  const currentEmails_row = currentEmails.indexOf(items_value['E-mail 1 - Value']) + 1;

  if (currentEmails_row > 0) {
    personSheetRow = currentEmails_row;
  }

  for (const key in items) {
    personSheet.getRange(personSheetRow,items_column[key]).setValue(items_value[key]);
  }

  checkExpiration();
}

function checkExpiration() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const personSheet = ss.getSheetByName('persons');
  const header_person = personSheet.getRange(1,1,1,personSheet.getLastColumn()).getValues().flat();
  const column_person_email = header_person.indexOf('E-mail 1 - Value') + 1;
  
  let personSheetLastRow = personSheet.getRange(1,column_person_email).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  if (personSheetLastRow === personSheet.getMaxRows()) personSheetLastRow = 2;

  const header = personSheet.getRange(1,1,1,personSheet.getLastColumn()).getValues().flat();
  const expirationColumn = header.indexOf('Expiration') + 1;
  const statusColumn = header.indexOf('Status') + 1;
  const nowDate = new Date();

  for (let i = 2; i <= personSheetLastRow; i++) {
    const expirationValue = personSheet.getRange(i,expirationColumn).getValue();
    if (expirationValue !== '' && expirationValue < nowDate) {
      personSheet.getRange(i,statusColumn).setValue('無効');
    } else {
      personSheet.getRange(i,statusColumn).setValue('有効');
    }
  }
}