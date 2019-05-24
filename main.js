function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  var addOnMenu = SpreadsheetApp.getUi().createAddonMenu();
  
  addOnMenu
    //add the menu item to export the print sheet
    .addItem("Create question/answer sheet", "createPrintSheet")
    //and to export the slideshow
    .addItem("Create question/answer slideshow", "pickPresentation")
    //to export the slideshow using a url
    .addItem("Create slideshow (from URL)", "createPresentationURL")
    //to display the export gui
    .addItem("Export (beta)", "exportWithGui")
    //to display version
    .addItem("Version information", "showInfo")
  //add the menu
  .addToUi();
}

function showError(title, prompt) {
  //get the sheet's UI
  ui = SpreadsheetApp.getUi();
  
  //if a detailed message is desired
  if(prompt) {
    ui.alert(title, prompt, ui.ButtonSet.OK);
  }
  //title only
  else {
    ui.alert(title)
  }
}

//creates the print sheet, with questions and answers, using the currently active sheet
function createPrintSheet() {
  SpreadsheetApp.setActiveSheet(writePrintSheet(getAllCategories(SpreadsheetApp.getActiveSheet())));
}

//picks the target presentation using google picker, then returns the id to writePresentationCallback
function pickPresentation() {
  showPicker();
}

//creates a presentation by asking for its url
function createPresentationURL() {
  ui = SpreadsheetApp.getUi();
  
  //query the URL of the presentation
  response = ui.prompt("Please enter presentation URL");
  
  var respText = response.getResponseText();
  
  if(response.getSelectedButton() == ui.Button.OK) {
    
    try {
      //try to open the specified url
      target = SlidesApp.openByUrl(respText);
    } catch (err) {
      //report an invalid URL
      ui.alert("[Error] An error occurred opening the slideshow at that URL!");
      return;
    }
    
    try {
      //write the presentation
      writePresentation(target, getAllCategories(SpreadsheetApp.getActiveSheet()));
      ui.alert("[Success] Slides created!");
    } catch (err) {
      ui.alert("[Error] An error occurred writing the slideshow. Do you have edit access?");
    }
  }
}

//writes the presentation to the provided id, from the currently active sheet
function writePresentationCallback(id) {
  target = SlidesApp.openById(id);
  writePresentation(target, getAllCategories(SpreadsheetApp.getActiveSheet()));
  SpreadsheetApp.getUi().alert("[Success] Slides created!");
}

function exportWithGui() {
  showExportDialog();
}

function showInfo() {
  SpreadsheetApp.getUi().alert("Version 13, 5 March 2019");
}