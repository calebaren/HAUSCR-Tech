var SCRIPT_PROP = PropertiesService.getScriptProperties();
var PRESO = SlidesApp.getActivePresentation();

function setup() {
  var pres = SlidesApp.getActivePresentation();
  SCRIPT_PROP.setProperty("key", preso.getId())
}

var UI = SlidesApp.getUi();

/* creates menu upon opening */
function onOpen(e) {
  UI.createAddonMenu()
  .addItem('Add/Update table of contents slide', 'updateRoadmap')
  .addItem('Add name and position', 'manualNamePosition')
    .addToUi();
}

/* removes all roadmaps */
function removeRoadmap() {
  var slidesList = PRESO.getSlides();
  for (var i = 0; i < slidesList.length; ++i) {
    if (slidesList[i].getLayout().getLayoutName() == 'SECTION_HEADER_1_1') {
      slidesList[i].remove();
    }
  }
}

/* updates all roadmaps */
function updateRoadmap() {
  var subtitles = getSubtitles();
  if (subtitles.length !== 0) {
    removeRoadmap();
    createRoadmap();
    return;
  }
  UI.alert('Add a subtitle slide to create a table of contents!');
  return;
}

/* adds name + position */
function manualNamePosition() {
  var slidesList = PRESO.getSlides();
  
  // prompts user for name + position
  var name = UI.prompt('Please enter your full name:',
                         '',
                        UI.ButtonSet.OK_CANCEL);
  if (name.getSelectedButton() == UI.Button.OK) {
    var position = UI.prompt('Please enter your HAUSCR position:',
                         '',
                         UI.ButtonSet.OK_CANCEL);
    if (position.getSelectedButton() == UI.Button.OK) {
    }
    else {
      return;
    }
  }
  else {
    return;
  }
  
  // fills placeholders
  for (var i = 0; i < slidesList.length; ++i) {
    if (slidesList[i].getLayout().getLayoutName() == 'TITLE') {
      slidesList[i].getPageElements()[2].asShape().getText().setText(name.getResponseText().toUpperCase());
      slidesList[i].getPageElements()[3].asShape().getText().setText(position.getResponseText().toUpperCase());
      var oldDate = slidesList[i].getPageElements()[4].asShape().getText();
      var date = UI.alert('Do you want to update the date?',
               UI.ButtonSet.YES_NO);
      if (date == UI.Button.YES) {
        slidesList[i].getPageElements()[4].asShape().getText().setText(hauscrDate(0)); //hauscrDate(1) for date created, hauscrDate(0) for current date
      }
      else {
        slidesList[i].getPageElements()[4].asShape().getText().setText(oldDate); //hauscrDate(1) for date created, hauscrDate(0) for current date
      }
    }
  }
}

// helper functions

/* getSubtitles function */
function getSubtitles() {
  var slidesList = PRESO.getSlides();
  var subtitles = [];
  var indexer = 0;
  
  for (var i = 0; i < slidesList.length; ++i) {
    var shapes;
    var groups;
    if (slidesList[i].getLayout().getLayoutName() == 'SECTION_HEADER') {
      subtitles[indexer] = slidesList[i]
      .getPlaceholders()[0]
      .asShape()
      .getText()
      .asString()
      .replace(/(\r\n\t|\n|\r\t)/gm,"");
      ++indexer;
    }
  }
  return subtitles;
}

/* creates the roadmap */
function createRoadmap() {
  var LEFT = 34;
  var TOP = 100;
  var WIDTH = 500;
  var HEIGHT = 44;
  var FONT = 'Nunito';
    var FONT_SIZE = 20;
    var FONT_WEIGHT = 600;
    var FONT_COLOR = '#FFFFFF';
  
  var template = PRESO.insertSlide(1,SlidesApp.openById("1pPP45wrU2Yh0QkM3JwPrpPQjmo57cx-TjBUe-9Vg7yw").getSlides()[1]);
  var roadmap = PRESO.getSlides()[1];
  var templateGroup = template.getGroups()[0];
  var subtitles = getSubtitles();
  var groups = [];
  
  for (var i = 0; i < subtitles.length; ++i) {
    // inserts subtitle textboxes
    roadmap
    .insertShape(SlidesApp.ShapeType.TEXT_BOX, LEFT + 96, TOP - 7 + HEIGHT * i, WIDTH, HEIGHT)
    .getText()
    .setText(subtitles[i] + ' ')
    .getTextStyle()
      .setFontFamilyAndWeight(FONT, FONT_WEIGHT)
      .setFontSize(FONT_SIZE)
      .setForegroundColor(FONT_COLOR);
    
    // inserts shape groups
    groups[groups.length] = roadmap
                    .insertGroup(templateGroup)
                    .setTop(TOP + HEIGHT * i)
                    .setLeft(LEFT);
  }
  
  // adds numbers to white boxes
  for (var i = 0; i < groups.length; ++i) {
    groups[i].getChildren()[0]
    .asShape()
    .getText()
    .setText(i+1);
  }
  roadmap.getGroups()[0].remove();
}

/* creates date of form "MONTH DD, YYYY" */
function hauscrDate(created) {
  if (created) {
    var date = DriveApp.getFileById(PRESO.getId()).getDateCreated();
  }
  else {
    var date = new Date();
  }
  var months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  return months[date.getMonth()] + ' ' + date.getDate() + ', ' + date.getFullYear();
}

// note: 'HAUSCR TEMPLATE' ID = 1pPP45wrU2Yh0QkM3JwPrpPQjmo57cx-TjBUe-9Vg7yw;

function initialCopy() {

//  copy title slide
  var titleSlide = SlidesApp.openById("1UVBJHBRt_s2eokSWMXw6I_M3T1zLEMT_gv3I4KCnWi8").getSlides()[0];
  var subtitleSlide = SlidesApp.openById("1UVBJHBRt_s2eokSWMXw6I_M3T1zLEMT_gv3I4KCnWi8").getSlides()[1];
  SlidesApp.getActivePresentation().appendSlide(titleSlide);
  SlidesApp.getActivePresentation().appendSlide(subtitleSlide);  
  
// delete first blank slide
  SlidesApp.getActivePresentation().getSlides()[0].remove();
  
//  autofill title placeholder with title of presentation
  var presentationTitle = SlidesApp.getActivePresentation().getName();
  SlidesApp.getActivePresentation().getSlides()[0].getPlaceholders()[0].asShape().getText().insertText(0, presentationTitle);
  
//  Logger.log("insert image here");
  
}

// here's some code that may be helpful

function initialCopyP() {
//copy title/subtitle slides
  var titleSlide = SlidesApp.openById("1UVBJHBRt_s2eokSWMXw6I_M3T1zLEMT_gv3I4KCnWi8").getSlides()[0];
  var subtitleSlide = SlidesApp.openById("1UVBJHBRt_s2eokSWMXw6I_M3T1zLEMT_gv3I4KCnWi8").getSlides()[1];
  var preso = SlidesApp.getActivePresentation();
  preso.insertSlide(0, titleSlide);
  preso.insertSlide(1, subtitleSlide);

/*
autofill but searches the entire slidedeck for the title & searches slide for 'TITLE' placeholder
*/
  var title = preso.getName();
  for (var i = 0; i < preso.getSlides().length; ++i) {
    if (preso.getSlides()[i].getLayout().getLayoutName() == 'TITLE') {
      preso.getSlides()[i]
      .getPlaceholder(SlidesApp.PlaceholderType.TITLE)
      .asShape()
      .getText()
      .insertText(0, title);
    }
  }
}