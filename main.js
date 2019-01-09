function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  var addOnMenu = SpreadsheetApp.getUi().createAddonMenu();
  
  addOnMenu
    //add the menu item to export the print sheet
    .addItem("Create Question/Answer Sheet", "createPrintSheet")
    //and to export the slideshow
    .addItem("Create Question/Answer Slideshow", "pickPresentation")
    //to export the slideshow using a url
    .addItem("Create Slideshow (from URL)", "createPresentationURL")
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
  writePrintSheet(getAllCategories(SpreadsheetApp.getActiveSheet()));
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
  
  if(response.getSelectedButton() == ui.Button.OK) {
    
    try {
      //try to open the specified url
      target = SlidesApp.openByUrl(response.getResponseText());
    } catch (err) {
      //report an invalid URL
      ui.alert("Invalid URL!");
      return;
    }
    
    try {
      //try and get the list of editors, (((this throws an exception if the current user can't edit)))
      //this SHOULD fail if the current user can't edit, but it doesn't(?)
      editors = target.getEditors();
      
      //get the effective current user (the user whose permissions control this script)
      effectiveUser = Session.getEffectiveUser();
      
      editAccess = false;
      for(var editorI in editors) {
        if(editors[editorI].getEmail() == effectiveUser.getEmail()) { editAccess = true; }
      }
      
      if(!editAccess) {
        throw "No edit access!";
      }
      
      /*actUser = Session.getActiveUser();
      
      Logger.log("Effective: " + effUser);
      Logger.log("Active: " + actUser);
      for(var editorI in editors) {
        editor = editors[editorI];
        Logger.log("For editor " + editor);
        Logger.log(Object.keys(editor));
        Logger.log("Equal eff: " + (editor == effUser));
        Logger.log("Equal act: " + (editor == actUser));
      }
      //Logger.log(editors);
      //Logger.log(editors[0] == Session.getEffectiveUser());
      
      //so instead check if the current user is an editor
      if (editors.indexOf(Session.getEffectiveUser()) < 0) { throw "Not an editor!"; }*/
    } catch (err) {
      ui.alert("[Error] Cannot edit slideshow at that URL!");
      return;
    }
    
    //once all checks have passed, write the presentation
    writePresentation(target, getAllCategories(SpreadsheetApp.getActiveSheet()));
    ui.alert("[Success] Slides created!");
  }
}

//writes the presentation to the provided id, from the currently active sheet
function writePresentationCallback(id) {
  target = SlidesApp.openById(id);
  writePresentation(target, getAllCategories(SpreadsheetApp.getActiveSheet()));
  SpreadsheetApp.getUi().alert("[Success] Slides created!");
}