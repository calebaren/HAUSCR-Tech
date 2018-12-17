/* Streamlines deliberation for HAUSCR associates by parsing spreadsheets into formatted documents. */
function onOpen(e) {
  SpreadsheetApp.getUi()
  .createAddonMenu()
  .addItem('/response-doc','documentize')
  .addToUi();
}


function documentize() {
  var sheet = SpreadsheetApp.getActiveSheet();
  values = sheet.getDataRange().getValues();
  var doc = DocumentApp.openById('1bV9jiTfnUHpOarzi1hzzBX-1Qj4u8i2j8qfApBBrewc');
  doc.getBody().clear();
  
  /* Defining styles */
  var nameStyle = {};
  nameStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.CENTER;
  nameStyle[DocumentApp.Attribute.FONT_FAMILY] = 'NUNITO';
  nameStyle[DocumentApp.Attribute.FONT_SIZE] = 17;
  nameStyle[DocumentApp.Attribute.BOLD] = false;
  
  var emailStyle = {};
  emailStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.CENTER;
  emailStyle[DocumentApp.Attribute.FONT_FAMILY] = 'NUNITO';
  emailStyle[DocumentApp.Attribute.FONT_SIZE] = 12;
  
  var boldStyle = {};
  boldStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
  boldStyle[DocumentApp.Attribute.FONT_SIZE] = 20;
  boldStyle[DocumentApp.Attribute.BOLD] = true;
  
  var normalStyle = {};
  normalStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
  normalStyle[DocumentApp.Attribute.FONT_SIZE] = 14;
  normalStyle[DocumentApp.Attribute.BOLD] = false;
  normalStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = 
    DocumentApp.HorizontalAlignment.LEFT;
  normalStyle[DocumentApp.Attribute.ITALIC] = false;
  
  var smallText = {};
  smallText[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
  smallText[DocumentApp.Attribute.FONT_SIZE] = 9;
  smallText[DocumentApp.Attribute.BOLD] = false;
  smallText[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = 
    DocumentApp.HorizontalAlignment.LEFT;
  
  var ital = {};
  ital[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
  ital[DocumentApp.Attribute.FONT_SIZE] = 9;
  ital[DocumentApp.Attribute.BOLD] = false;
  ital[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = 
    DocumentApp.HorizontalAlignment.LEFT;
  ital[DocumentApp.Attribute.ITALIC] = true;
  
  /* Appends paragraphs */
  doc.getBody().appendParagraph('HAUSCR Associate Interviewer Responses (/response-doc)').setAttributes(nameStyle).setFontSize(25);
  doc.getBody().appendParagraph('Applicants').setAttributes(boldStyle).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setFontSize(19);
  for (var i = 1; i < values.length; i++) {
    par = doc.getBody().appendParagraph(i + '.   ').setAttributes(normalStyle).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    par.appendText(values[i][2]);
    
  }
  doc.getBody().appendPageBreak();
  
  /* Loops through all associates */
  for (var i = 1; i < values.length; i++) {
    //Interviewer name: values[i][1] DONE
    //Applicant name: values[i][2] DONE
    //Email notes: values[i][3]    DONE
    //Initial rating: values[i][4] DONE
    //General notes: [values[5]    DONE
    //Committee interests: [6]     DONE
    //Availability on tues: [7]    DONE
    //Interest in HAUSCR: [8]      DONE
    //Personality fit: [9]         DONE
    //Commitment: [10]             DONE
    //Overall should we take: [11] DONE
    
    doc.getBody().appendParagraph(i + ".  "+values[i][2]).setAttributes(nameStyle); //name
    doc.getBody().appendParagraph(values[i][3]).setAttributes(emailStyle); //email
    doc.getBody().appendParagraph('Interviewed by ' + values[i][1]).setAttributes(ital).setAlignment(DocumentApp.HorizontalAlignment.CENTER); //interviewer
    doc.getBody().appendParagraph('');
    par = doc.getBody().appendParagraph('Initial rating: ').setAttributes(normalStyle);
    par.appendText(values[i][4]).setAttributes(boldStyle); //initial rating
    doc.getBody().appendParagraph('General notes: ').setAttributes(normalStyle);
    doc.getBody().appendParagraph(values[i][5]).setAttributes(smallText);
    doc.getBody().appendParagraph('');
    doc.getBody().appendParagraph('Committee interests: ').setAttributes(normalStyle);
    doc.getBody().appendParagraph(values[i][6]).setAttributes(smallText);
    doc.getBody().appendParagraph('');
    par = doc.getBody().appendParagraph('Interest in HAUSCR: ').setAttributes(normalStyle);
    par.appendText(values[i][8]).setAttributes(boldStyle); //interest in hauscr
    par = doc.getBody().appendParagraph('Personality fit: ').setAttributes(normalStyle);
    par.appendText(values[i][9]).setAttributes(boldStyle); //personality fit
    par = doc.getBody().appendParagraph('Commitment: ').setAttributes(normalStyle);
    par.appendText(values[i][10]).setAttributes(boldStyle); //commitment
    doc.getBody().appendParagraph('');
    par = doc.getBody().appendParagraph('Should we take them? ').setAttributes(normalStyle).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    
    if (values[i][11].toString() == '4') {
      doc.appendParagraph(values[i][11]).setAttributes(boldStyle).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setFontSize(30).setForegroundColor('#39b25d'); //Overall should we take
    } else if (values[i][11].toString() == '1') {
      doc.appendParagraph(values[i][11]).setAttributes(boldStyle).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setFontSize(30).setForegroundColor('#e0381a'); //Overall should we take
    } else {
      doc.appendParagraph(values[i][11]).setAttributes(boldStyle).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setFontSize(30).setForegroundColor('#000000');//Overall should we take
    }
    if (values[i][7].toString() == 'Yes') {
      doc.getBody().appendParagraph('They are available on Tuesday 7:30-8:30 for general meeting.').setAttributes(normalStyle).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setForegroundColor('#000000');
    }
    else {
      doc.getBody().appendParagraph('They are NOT available on Tuesday 7:30-8:30 for general meeting.').setAttributes(boldStyle).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setForegroundColor('#000000');
    }
    doc.getBody().appendPageBreak();
  }
  doc.saveAndClose();
  SpreadsheetApp.getUi().alert('go.hauscr.org/response-doc');
}
