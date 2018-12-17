/*
****
<NAME OF CLASS/EVENT>
<HH:00>-<HH:00>#BOLD <DAYOFWEEK#BOLD>, <MONTH>, <DATE>
Location: <LOCATION>
Attendance (optional for class events): <"EVERYONE" or else>
Additional Information: <CLASS DESCRIPTION BLURB>
****

documentation: https://developers.google.com/apps-script/reference/sites/
test site link: https://sites.google.com/site/hauscrtest2/
final site link: https://sites.google.com/site/xweek2018/
important methods:

*/

var classArray = [];
var studentArray = [];
var site = SitesApp.getSiteByUrl('https://sites.google.com/site/xweek2018/');
var pages = site.getChildren();

// creates all pages at once
function createAllPages() {
  openSpreadsheet();
  for each (student in studentArray) {
    createPage(student);
  }
}

// deletes all pages at once
function deleteAllPages() {
  for each (page in pages) {
    page.deletePage();
  }
}

// debugging reasons
function testOnePageUI() {
  openSpreadsheet();
  createPage(studentArray[0]);
}

// initializes
function openSpreadsheet() {
  var spreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1EebkWd_BpFiXCm9Znnd9Z9HT99HctrU3DFkTYR24a5Y/edit#gid=0");
  var classes = spreadsheet.getRange("'Class List'!A2:AA88");
                                     
  var rows = classes.getNumRows();
  var cols = classes.getNumColumns();

  for (var i = 1; i <= rows; i++) {
    var class = {
      name:classes.getCell(i, 6).getDisplayValue(),
      code:classes.getCell(i, 5).getDisplayValue(),
      dscrp:classes.getCell(i, 19).getDisplayValue(),
      date:classes.getCell(i, 9).getDisplayValue(),
      time:classes.getCell(i, 10).getDisplayValue(),
      location:classes.getCell(i, 11).getDisplayValue(),
      attendance:"Just you!",
      addnInfo:"",
      h:""
    }
    if (class.date == 'MW') {
      class.h = classes.getCell(i, 8).getDisplayValue();
    }
    if (class.date == 'TTh') {
      class.h = classes.getCell(i, 7).getDisplayValue();
    }
    classArray.push(class);
  }
  
  // classArray now stores an array of classes
  
  var students = spreadsheet.getRange("'XWeek'!A2:T79");
                                     
  rows = students.getNumRows();
  var numstudents = rows;
  cols = students.getNumColumns();
  
  for (var i = 2; i <= rows; i++) {
    var student = {
      id:students.getCell(i, 4).getDisplayValue(),
      group:students.getCell(i, 7).getDisplayValue(),
      name_eng:students.getCell(i, 6).getDisplayValue(),
      name_chn:students.getCell(i, 5).getDisplayValue(),
      wed:[],
      ths:[],
      hash:students.getCell(i,20).getDisplayValue()
    }
    
    Logger.log(student.name_eng);
    
    // wednesday columns
    k = students.getCell(i, 11).getDisplayValue();
    l = students.getCell(i, 12).getDisplayValue();
    
    // thursday columns
    m = students.getCell(i, 13).getDisplayValue();
    n = students.getCell(i, 14).getDisplayValue();
    o = students.getCell(i, 15).getDisplayValue();
    p = students.getCell(i, 16).getDisplayValue();
    
    // grabs wednesday classes
    var wedCols = [11, 12];
    for each (col in wedCols) {
      var time = students.getCell(1, col).getDisplayValue();
      addClassToSched(student['wed'], students.getCell(i, col).getDisplayValue(), time);
    }
    
    // grabs thursday classes
    var thursCols = [13, 14, 15, 16];
    for each (col in thursCols) {
      var time = students.getCell(1, col).getDisplayValue();
      addClassToSched(student['ths'], students.getCell(i, col).getDisplayValue(), time);
    }

    Logger.log(student.hash);
    studentArray.push(student);
  }
  
  // studentArray now stores an array of all students w/ their assigned classes
  Logger.log('success');
}

// pushes classes
function addClassToSched(studentDayArray, event, time) {
  if (event == "Professor Seminar") {
    studentDayArray.push({
      name: "Professor Seminar",
      dscrp: 'Featuring discussions with scholars from the Harvard community, seminars will provide students with in-depth discussions about how the themes of liberal arts and perspective influence education and work.',
      time: time,
      location: '',
    });
  } else if (event == "Office Hours") {
    studentDayArray.push({
      name: 'Office Hours',
      dscrp: 'Office hours are a chance to work on your final papers and projects. You may choose any study space to work at; common study spaces include the Smith Campus Center, Science Center, Memorial Hall Basement, and Ticknor Lounge.',
      time: time,
      location: 'Your Choice',
    });
  } else if (event == "AO Info") {
    studentDayArray.push({
      name: "Admission Officer Information Sessions",
      dscrp: "Learn more about the admissions process at Harvard by interacting with faculty from the Admissions Office.",
      time: time,
      location: '',
    });
  } else if (event == "MIT Hackathon") {
    studentDayArray.push({
      name: 'MIT Hackathon Office Hours',
      dscrp: 'This is a chance to collaborate with other XWeek students to complete your group projects for your hackathon. During this time, you can also ask questions or discuss ideas with Harvard undergraduates.',
      time: time,
      location: 'Sever Hall',
    });
  } else if (event == "Writing Workshop") {
    studentDayArray.push({
      name: 'Writing Workshop',
      dscrp: 'Writing and being able to express oneself through words is a critical part of being a successful Harvard student. In this Writing Workshop, students will learn how to engage in research, analysis, and critical thinking to write an academic paper that synthesizes information from a variety of sources.',
      time: time,
      location: '',
    });
  } else if (event == "") {
    
  } else {
    var obj = classArray.filter( function(class) {return class.h == event;} );
    studentDayArray.push(obj[0]);
  }
            
        
}

// creates pages for each student
function createPage(student) {
  var html = "<h1 style='margin-top:0px'>Wednesday Classes</h1>" + createClassTable(student.wed);
  html += "<h1>Thursday Classes</h1>" + createClassTable(student.ths);
  
  html += "<a href=\"https://sites.google.com/site/xweek2018/\">Return to Home";
  
  if(site.getChildByName(student.hash) != null) {
    // If the page already exists, just update the HTML
    site.getChildByName(student.hash).setHtmlContent(html);
    Logger.log(site.getChildByName(student.hash));
  } else {
    // If the page does not exist, create it and set HTML
    site.createWebPage(student.name_eng + "'s schedule", student.hash, html);
  }
}

// deletes pages with hashes
function deleteIDPages(student) {
    if (site.getChildByName(student.id)) {
      site.getChildByName(student.id).deletePage();
    }
}

// creates a table schedule
function createClassTable(classes) {
  var tableheader = HtmlService.createHtmlOutputFromFile('tableheader.html').getBlob().getDataAsString();
  var table = tableheader;
  for each (class in classes) {
    table += createClassRow(class);
  }
  table += "</table>";
  return table;
}

// create the table row for an individual class
function createClassRow(class) {
  var rowHtml = "<tr>";
  var row = {
    nameCell:createClassCell(class.name),
    dscrpCell:createClassCell(class.dscrp),
    timeCell:createClassCell(class.time),
    locCell:createClassCell(class.location),
  }
  
  for each (cell in row) {
    rowHtml += cell;
  }
  rowHtml += "</tr>";
  return rowHtml;
}

// Creates the cell string for a content
function createClassCell(contents) {
  return "<td style='border-bottom:2px solid #ffffff;padding:5px;'>" + contents + "</td>";
}