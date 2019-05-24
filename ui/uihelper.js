//shows the full export dialog
function showExportDialog() {
  var html = HtmlService.createTemplateFromFile("ui/export.html")
  html.nextDate = nextEventDate();
  html = html.evaluate()
    .setWidth(601)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, "Export");
}

//evaluates the templated html content in fName and returns it as a string
function include(fName) {
  return HtmlService.createTemplateFromFile(fName).evaluate().getContent();
}

function exportPrintSheet(sourceSheet, triviaNightDate) {
  SpreadsheetApp.setActiveSheet(writePrintSheet(getAllCategories(sourceSheet)));
};

//gets the sheet with the specified id (string or int) from the active spreadsheet
function getSheetById(id) {
  //parse a string to a number
  id = parseInt(id);
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  //map an array of sheets to an array of sheet IDs
  var sheetIds = sheets.map(function(s) { return s.getSheetId(); });
  //return the sheet whose index matches
  return sheets[sheetIds.indexOf(id)];
}

//returns to the client the number of categories in the sheet with the provided id
function getCategoryCount(sheetId) {
  try {
    //return the length of the array of parsed categories
    return getAllCategories(getSheetById(sheetId)).length;
  } catch(err) {
    Logger.log(err);
    //if anything goes wrong, return -1
    return -1;
  }
}

//the day of the week for the default trivia night date: 0 is Sunday, 5 (default) is Friday
//TODO make this an option
var EVENT_WEEKDAY = 5;

//returns the next date a trivia night is likely to be held -- next Friday, by default
function nextEventDate() {
  var next = new Date();
  //moves the current date, the mod is necessary to avoid negative moves
  next.setDate(next.getDate() + (7 + EVENT_WEEKDAY - next.getDay()) % 7);
  
  return next;
}

//returns a string representing the next day a trivia night is likely to be held (see nextEventDate)
function nextEventDateString() {
  var d = nextEventDate
}