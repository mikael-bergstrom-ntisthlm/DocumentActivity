// TODO: Get class roster (Class, Name, Surname, ResourceName) from Classroom

type RowProcessor = (row: any[]) => string[];

function onOpen() {
  let ui = SpreadsheetApp.getUi();

  ui.createMenu('DocumentActivity')
    .addItem('Get activity (weeks)', 'GetWeeks')
    .addItem('Get resource names for users', 'GetResourceNames')
    .addToUi();
}

function GetResourceNames() {
  ProcessCurrentRange(
    1, "Containing user queries (names/email)",
    row => {
    const userQuery: string = String(row[0]);
    const resourceName = GetUserResourceName(userQuery)
    return [resourceName];
  });
}


function GetWeeks() {
  ProcessCurrentRange(
    2, "First with gdocs links, second with email addresses",
    row => {

    const docUrl: string = String(row[0]);
    const userResourceName: string = String(row[1]);

    const dates = GetHistory(docUrl, userResourceName);

    if (dates.length == 0) return [];

    const weeks: Set<string> = new Set(
      dates.map(date => {
        return Utilities.formatDate(date, Session.getScriptTimeZone(), "w");
      })
    );

    return Array.from(weeks).sort();
  })
}

function ProcessCurrentRange(expectedColumns: number, expectedColumnDesc: string, processor: RowProcessor) {
  let sheet = SpreadsheetApp.getActiveSheet();
  let range = sheet.getActiveRange();
  if (!range) return;

  const values = range.getValues();

  if (values[0].length != expectedColumns) {
    SpreadsheetApp.getUi().alert("Expected exactly " + expectedColumns + " columns. " + expectedColumnDesc);
    return;
  }

  const colStart = range.getColumn();
  const rowStart = range.getRow();

  for (let rNum = 0; rNum < values.length; rNum++) {
    const row = values[rNum];
    const result: string[] = processor(row);

    let targetCells = sheet.getRange(rowStart + rNum, colStart + row.length, 1, result.length);
    
    targetCells.setValues([result])
  }
}