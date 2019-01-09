var PRINT_SHEET_NAME = "Questions and Answers (To Print)";

var COLUMN_WIDTHS = {"NUMBER":31 ,"QUESTION":437, "ANSWER":156};
var FONT_SIZES = {"HEADER":12, "CATEGORY_NAME":20, "QUESTION":10, "ANSWER":10, "NUMBER":10};
var BOLD = {"HEADER":"normal", "CATEGORY_NAME":"bold", "QUESTION":"normal", "ANSWER":"bold", "NUMBER":"bold"};

var HEADER_HEIGHT = 2;

function foo() {
  writePrintSheet(getAllCategories(SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]));
}

function writePrintSheet(categories) {
  //many things might break if no categories are presented
  if(!categories || categories.length == 0) {
    showError("No Categories!", "You need to provide at least one category. Do you have the wrong sheet selected? For help go to: https://urlred.page.link/help");
    return;
  }
  
  //get an existing or create a new print sheet
  var printSheet = getPrintSheet();
  
  //resize the columns being written to
  //TODO remove
  //resizeColumns(printSheet);
  
  //format each column
  //TODO - remove
  //formatContents(printSheet); (renamed)
  
  //get the range for the header and write the header
  writeHeader(printSheet.getRange(1, 1, 1, 3));
  
  //the first number of each category is stored here
  var firstQuestionNumber = 1;
  //represents rows and columns
  var questionData = [];
  //where the category headers are written
  var categoryHeaderRows = [];
  
  //write each category to the array
  for(var categoryI in categories) {
    
    if(categories[categoryI].questions.length < 1) {
      Logger.log("Empty Category, skipped!");
      continue;
    }
    
    //push the row of the next category header - the current number
    //  of rows in questionData plus an offset for the header plus an offset for off by one
    categoryHeaderRows.push(questionData.length + HEADER_HEIGHT + 1);
    
    firstQuestionNumber = writeCategory(categories[categoryI], questionData, firstQuestionNumber);
  }
  
  //write the array to the sheet
  printSheet.getRange(HEADER_HEIGHT + 1, 1, questionData.length, 3).setValues(questionData);
  
  //the range containing the category header formatting (immediately after the header)
  catHeaderFormatR = printSheet.getRange(HEADER_HEIGHT + 1, 1, 1, 3);
  
  //the range containing the question formatting (immediately after the cat header format)
  questionFormatR = printSheet.getRange(HEADER_HEIGHT + 2, 1, 1, 3);
  
  //ID identifying the printSheet (for copying formatting)
  printSheetID = printSheet.getSheetId();
  
  //format the category headers
  for(var i in categoryHeaderRows) {
    //get the header range range
    var catHeaderRange = printSheet.getRange(categoryHeaderRows[i], 1, 1, 3);
    
    //the first row of the category
    firstRow = catHeaderRange.getRow();
    
    //copy the header format
    catHeaderFormatR.copyFormatToRange(printSheetID, 1, 3, firstRow, firstRow);
    
    //determine the length of the category
    catLength = categories[i].questions.length;
    
    //copy the question format
    questionFormatR.copyFormatToRange(printSheetID, 1, 3, firstRow + 1, firstRow + catLength);
    
    //merge the category header
    catHeaderRange.merge();
  }
}

//creates an empty print sheet, with default formatting
function createNewPrintSheet() {
  //create the sheet
  writeSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(PRINT_SHEET_NAME);
  
  //format the sheet
  resizeColumns(writeSheet);
  formatContents(writeSheet);
  formatHeader(writeSheet.getRange(1, 1, 1, 3));
  
  //return the sheet
  return writeSheet;
}

function getPrintSheet() {
  //try and get an existing print sheet
  var writeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PRINT_SHEET_NAME);
  
  //if it doesn't exist, create it
  if(!writeSheet) {
    writeSheet = createNewPrintSheet();
  }
  //otherwise clear it
  else {
    
    //clear only the content in the template range, except the header (first cat header, first question)
    writeSheet.getRange(2, 1, HEADER_HEIGHT + 1, 3).clearContent();
    //clear the rest of the sheet entirely
    writeSheet.getRange(HEADER_HEIGHT + 3, 1, writeSheet.getDataRange().getNumRows(), 3);
    
    //unmerge the entire sheet
    writeSheet.getRange("A1:C").breakApart();
  }
  
  return writeSheet;
}

function resizeColumns(sheet) {
  //resize the three columns
  sheet.setColumnWidth(1, COLUMN_WIDTHS.NUMBER);
  sheet.setColumnWidth(2, COLUMN_WIDTHS.QUESTION);
  sheet.setColumnWidth(3, COLUMN_WIDTHS.ANSWER);
}

//formats the two rows directly below the header
//first for category name format, second for question/ans format
//these are then copied to the rest of the sheet
function formatContents(sheet) {
  
  //format the category name
  var catHeaderRange = sheet.getRange(HEADER_HEIGHT + 1, 1)
    //set the font weight (bold/not bold)
    .setFontWeight(BOLD.CATEGORY_NAME)
    //set the font size
    .setFontSize(FONT_SIZES.CATEGORY_NAME);
  
  //format the question numbers
  sheet.getRange(HEADER_HEIGHT + 2, 1)
    //set the font weight (bold/not bold)
    .setFontWeight(BOLD.NUMBER)
    //set the font size
    .setFontSize(FONT_SIZES.NUMBER);
  
  //format the questions
  sheet.getRange(HEADER_HEIGHT + 2, 2)
    //set the font weight (bold/not bold)
    .setFontWeight(BOLD.QUESTION)
    //set the font size
    .setFontSize(FONT_SIZES.QUESTION);
  
  //format the answers
  sheet.getRange(HEADER_HEIGHT + 2, 3)
    //set the font weight (bold/not bold)
    .setFontWeight(BOLD.ANSWER)
    //set the font size
    .setFontSize(FONT_SIZES.ANSWER);
  
  //number, question and answer
  //location: 2nd row after HEADER_HEIGHT, first three columns
  sheet.getRange(HEADER_HEIGHT + 2, 1, 1, 3)
    //add a top and bottom border for visual seperation
    .setBorder(true, null, true, null, null, null, null, null);
  
  //set the wrap strategy for all cells
  sheet.getRange("A1:C").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
}

//writes the header to the specified range, merging it
function writeHeader(range) {
  //set the header text
  range.getCell(1, 1).setValue(getHeaderText());
  //merge the header cells
  range.merge();
}

function formatHeader(range) {
  
  //format the header weight
  range.setFontWeight(BOLD.HEADER);
  //set the font size
  range.setFontSize(FONT_SIZES.HEADER);
  //add the cell border
  range.setBorder(true, true, true, true, null, null, null, null);
  
  //also add a cell border on the cell immediately to the right
  range.offset(0, range.getNumColumns()).getCell(1,1).setBorder(null, true, null, null, null, null, null, null);
  
  //center the header
  range.setHorizontalAlignment("center");
  
  //create a range to make the spacer row
  range = range.offset(range.getHeight(), 0, 1);
  
  //set the spacer row height
  range.getSheet().setRowHeight(range.getRow(), 10);
}

function getHeaderText() {
  return "Trivia Night\n" + Utilities.formatDate(new Date(), "GMT+1", "MMMM d, YYYY") + "\nANSWER DOCUMENT";
}

function writeCategory(cat, writeArray, firstQNumber) {
  
  //write the name of the category
  writeArray.push([cat.name, "", ""]);
  
  //for each question
  for(var questionI in cat.questions) {
    
    //write a row with question number, question and answer, then increment question number
    writeArray.push([firstQNumber, cat.questions[questionI].question, formatAnswer(cat.questions[questionI].answer, false)]);
    firstQNumber++;
    
  }
  
  //return the first question number for the next category
  return firstQNumber;
}