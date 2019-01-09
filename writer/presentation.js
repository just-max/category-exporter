var TITLE_L = SlidesApp.PredefinedLayout.TITLE;
var CATEGORY_L = SlidesApp.PredefinedLayout.MAIN_POINT;
var QUESTION_L = SlidesApp.PredefinedLayout.SECTION_HEADER;
var ANSWER_L = SlidesApp.PredefinedLayout.TITLE_AND_TWO_COLUMNS;

var weekdays = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
var months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

var mainTitleText = "Trivia Night";
var answerPrefix = "Answers for: ";

function writePresentation(target, categories) {
  
  //the first slide, the title slide
  titleSlideElements = target.appendSlide(TITLE_L).getPageElements();
  
  //set the title and today's date
  titleSlideElements[0].asShape().getText().appendText(mainTitleText);
  today = new Date();
  dateString = weekdays[today.getDay()] + ", " + months[today.getMonth()] + " " + today.getDate();
  titleSlideElements[1].asShape().getText().appendText(dateString);
  
  //the current question number
  var questionNumber = 1;
  
  //for each category
  for(var categoryI in categories) {
    
    //the first question number in the category, used for creating the answers later
    var answerQNumber = questionNumber;
    
    //append a slide with the category title layout
    catTitleElements = target.appendSlide(CATEGORY_L).getPageElements();
    //set the slide to display the name of the category
    catTitleElements[0].asShape().getText().appendText(categories[categoryI].name);
    
    //for each question in this category
    for(var questionI in categories[categoryI].questions) {
      //append a slide with the question slide layout
      questionElements = target.appendSlide(QUESTION_L).getPageElements();
      //set the slide to display the question
      questionElements[0].asShape().getText().appendText(questionNumber + ". " + categories[categoryI].questions[questionI].question);
      
      //increment the question number for the next question
      questionNumber++;
    }
    
    //append the answer slide
    ansElements = target.appendSlide(ANSWER_L).getPageElements();
    //set the slide title ("Answers for: " followed by the category name)
    ansElements[0].asShape().getText().appendText(answerPrefix + categories[categoryI].name);
    
    //split the answers in two
    splitAt = Math.ceil(categories[categoryI].questions.length / 2);
    
    //the current column for answers (incremented once the question index reaches splitAt
    ansColumn = 1;
    
    for(questionI in categories[categoryI].questions) {
      //whether this answer goes in element 1 (left column) or 2 (right column)
      if(questionI == splitAt) { ansColumn = 2; }
      
      //append text with the answer, followed by a new line
      ansElements[ansColumn].asShape().getText().appendText(answerQNumber + ". " + formatAnswer(categories[categoryI].questions[questionI].answer, true) + "\n");
      
      //increment the question number for the next answer
      answerQNumber++;
    }
  }
  
  target.saveAndClose();
}
