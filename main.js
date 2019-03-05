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
      ui.alert("Invalid URL!");
      return;
    }
    
    /*
    try {
      //try and get the list of editors, (((this throws an exception if the current user can't edit)))
      //this SHOULD fail if the current user can't edit, but it doesn't(?)
      editors = target.getEditors();
      
      Logger.log(editors);
      
      //get the effective current user (the user whose permissions control this script)
      effectiveUser = Session.getEffectiveUser();
      
      editAccess = false;
      for(var editorI in editors) {
        if(editors[editorI].getEmail() == effectiveUser.getEmail()) { editAccess = true; }
      }
      
      if(!editAccess) {
        throw "No edit access!";
      }
      
      /block comment start
      actUser = Session.getActiveUser();
      
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
      if (editors.indexOf(Session.getEffectiveUser()) < 0) { throw "Not an editor!"; }
      /block comment end
    } catch (err) {
      ui.alert("[Error] Cannot edit slideshow at that URL!");
      Logger.log(err);
      return;
    }*/
    
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

function showInfo() {
  SpreadsheetApp.getUi().alert("Version 12, 21 February 2019");
}