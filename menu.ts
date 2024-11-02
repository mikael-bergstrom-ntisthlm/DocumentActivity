// TODO: GetUserIdFromEmail
//  Google limits number of calls to Person API, so we need to limit ourselves
//  No cache except what exists in sheet = save user ID and email in sheet

function onOpen() {
  let ui = SpreadsheetApp.getUi(); 

  ui.createMenu('DocumentActivity')
    .addItem('Get activity (weeks)', 'GetWeeks')
    .addItem('Get resource names for users', 'GetResourceNames')
    .addToUi();
}

function GetResourceNames()
{
  let sheet = SpreadsheetApp.getActiveSheet();
  let range = sheet.getActiveRange();
  if (!range) return;

  const values = range.getValues();

  const colStart = range.getColumn();
  const rowStart = range.getRow();

  if (values[0].length != 1) { // !
    SpreadsheetApp.getUi().alert("Expecting exactly one column, containing user queries"); // !
    return;
  }

  for (let rNum = 0; rNum < values.length; rNum++) {
    const row = values[rNum];
    const userQuery: string = String(row[0]); // !

    const resourceName = GetUserResourceName(userQuery); // !

    let targetCell = sheet.getRange(rowStart + rNum, colStart + 1, 1, 1); // !

    targetCell.setValue(resourceName); // !
  }
}


function GetWeeks() {
  let sheet = SpreadsheetApp.getActiveSheet();
  let range = sheet.getActiveRange();
  if (!range) return;

  const values = range.getValues();

  const colStart = range.getColumn();
  const rowStart = range.getRow();

  if (values[0].length != 2) {
    SpreadsheetApp.getUi().alert("Expecting exactly two columns; first with gdocs links, second with email addresses");
    return;
  }

  for (let rNum = 0; rNum < values.length; rNum++) {
    const row = values[rNum];
    const docUrl: string = String(row[0]);
    const userResourceName: string = String(row[1]);

    const dates = GetHistory(docUrl, userResourceName)

    if (dates.length == 0) continue;
    
    const weeks: Set<string> = new Set(
      dates.map(date => {
        return Utilities.formatDate(date, Session.getScriptTimeZone(), "w");
      })
    );

    let targetCell = sheet.getRange(rowStart + rNum, colStart + 2, 1, weeks.size);

    targetCell.setValues([Array.from(weeks).sort()]);
  }

  // SpreadsheetApp.getUi().alert(r.offset(0,r.getWidth()).getValue());
  // SpreadsheetApp.getUi().alert(r?.getA1Notation());
}