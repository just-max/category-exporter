var CAT_LABEL_REG = /^category:/i;
var CAT_NAME_REG = /[^:]*:\s*(.*)/;

//checks whether the specified cell starts a category
function isCategoryStart(content) {
  //if the cell starts with 'Category:'
  return content.toString().match(CAT_LABEL_REG)
}

//checks whether the specified cell ends a category
function isCategoryEnd(content) {
  //determined by a blank cell
  return content == "";
}

function formatAnswer(rawAnswer, keepCase) {
  answer = rawAnswer.replace(/([^\\])\s*>\s*/g, "$1 ACCEPT ") //format non escaped > to ACCEPT
                    .replace(/([^\\])\s*#\s*/g, "$1 OR ") //format non escaped # to OR
                    .replace(/\\(>|#)/g, "$1"); //remove \ from escaped characters
  
  if (keepCase) { return answer; }
  else          { return answer.toUpperCase(); }
}

//gets all regions that represent categories
function getAllCategories(sheet) {
  
  //get all data in the sheet
  ///var sheetDataRange = sheet.getDataRange();
  var sheetData = sheet.getDataRange().getValues();
  
  //array for all the found categories
  var categories = [];
  
  //read the cells one by one to find starts
  for(var row = 0; row < sheetData.length; row++) {
    for(var col = 0; col < sheetData[0].length; col++) {
      
      //get the cell at this position
      var currentCell = sheetData[row][col];
      
      //ignore if blank, then check if start of a category
      if(currentCell != "" && isCategoryStart(currentCell)) {
        
        //we are at the start of a category:
        
        //start by assuming only the first cell ends the category
        var assumedEndRow = row;
        
        //search for the end of the category, starting right below the start
        //keep assuming the current cell is the end, then move to the next, if it isn't
        for(var nextEndRow = row + 1; nextEndRow < sheetData.length; nextEndRow++) {
          
          //the next cell to check
          var nextCell = sheetData[nextEndRow][col];
          
          //if the cell here starts a new category or ends a category, stop searching
          if(isCategoryEnd(nextCell) || isCategoryStart(nextCell)) { break; }
          
          //otherwise move the end to the next cell
          assumedEndRow = nextEndRow;
        }
        
        //note: the assumed end row is now the actual end row
        
        //check that the category has a length of at least two lines
        if(assumedEndRow > row) {
          
          //to store q/a pair
          catQuestions = [];
          
          //for each row in the category's range, not including the category name
          for(var currRow = row + 1; currRow <= assumedEndRow; currRow++) {
            
            //push the question here to the array
            catQuestions.push({
                //the question is at the current position, the answer one over (to the right)
                question: sheetData[currRow][col].toString(),
                //the answer
                answer: sheetData[currRow][col + 1].toString()
              });
          }
          
          //extract the category name (anything after the first colon in the top left cell)
          catName = CAT_NAME_REG.exec(sheetData[row][col].toString())[1];
          
          //push the category object
          categories.push({
            name: catName,
            questions: catQuestions
          })
        }
      }
      
    }
  }
  
  return categories;
}
