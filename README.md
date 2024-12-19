# PM-App-Scripts
Creating Memo and Send to Calendar to anyone participating in a specific project

/**
* @OnlyCurrentDoc
* @param {Object} e The onOpen() event object.
*/

function onOpen() {
 SpreadsheetApp.getUi().createMenu("⚙️ Admin")
   .addItem("Create Memo", "createMemo")
   .addItem("Book Memo in Calendar", "bookMemo")
   .addToUi();
}

/**
* @OnlyCurrentDoc
* @param {Object} e The createMemo() event object.
*/

function createMemo() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = sheet.getActiveRange();

  var owner = row.getValues()[0][0];
  var topic = row.getValues()[0][1];
  var objective = row.getValues()[0][2];
  var method = row.getValues()[0][3];
  var condition = row.getValues()[0][4];
  var start = row.getDisplayValues()[0][5];
  var end = row.getDisplayValues()[0][6];
  var evaluation = row.getValues()[0][7];
  var docLink = row.getValues()[0][8];

  // Create a new Google Doc for the column if empty
  if (docLink === ""){
    var doc = DocumentApp.create(topic);
    docLink = doc.getUrl();
    sheet.getRange(row.getRowIndex(),9).setValue(docLink);
  } else {
    var doc = DocumentApp.openByUrl(docLink);
    doc.getBody().clear();
  }

  var body = doc.getBody();
  var imageStyle = {};
  imageStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
  var image = "";
  switch(owner){
    case 1:
      image = DriveApp.getFileById("1FYk0DZPHt-FpLfY3d6yUNUD7pB1vJOfl").getBlob();
      break;
    case 2:
      image = DriveApp.getFileById("1Yc1p1iKzOX6BbROLES5YbXRuYZreRNdE").getBlob();
      break;
    case 3:
      image = DriveApp.getFileById("1_bjlYID2M4vNg7zYbou_XBNaINRrsGAI").getBlob();
      break;
    case 4:
      image = DriveApp.getFileById("1B1ZWEF5MYCXThbIkiR0YbLfnZiFal1Mw").getBlob();
      break;
    case 5:
      image = DriveApp.getFileById("1-NSzCf4PRn34-esP7XZ1wPgcWaYAqIvi").getBlob();
      break;
    case 6:
      image = DriveApp.getFileById("1udQG-hiG2QaGrPADXQaENbNwI00iaWZx").getBlob();
      break;
    default:
      image = DriveApp.getFileById("1FYk0DZPHt-FpLfY3d6yUNUD7pB1vJOfl").getBlob();
      break;
  }
  var imgPara = body.appendImage(image);

  var ratio = imgPara.getHeight()/imgPara.getWidth();
  imgPara.setHeight(106).setWidth(106/ratio).getParent().setAttributes(imageStyle);

  body.appendParagraph("Project Detail " + topic).setHeading(DocumentApp.ParagraphHeading.HEADING1).setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  var table = body.appendTable().setBorderColor("#ffffff");
  var head = ["Objective: ", "Method: ", "Condition: ", "Start-date: ", "End-date: ", "Evaluation: "];
  var content = [objective, method, condition, start, end, evaluation];
  for (var k = 0; k < 6; k++){
    var row = table.appendTableRow();
    row.appendTableCell(head[k]).setWidth(75);
    row.appendTableCell(content[k]);
  }
  body.appendHorizontalRule();
  body.appendPageBreak();

  body.appendParagraph("Media Example").setHeading(DocumentApp.ParagraphHeading.HEADING1).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  var imgPR = DriveApp.getFilesByName(owner+"_"+topic+"(Any Suffix)");
  if (imgPR.hasNext() == 1){
    imgPara = body.appendImage(imgPR.next());
    ratio = imgPara.getHeight()/imgPara.getWidth();
    imgPara.setHeight(650).setWidth(650/ratio).getParent().setAttributes(imageStyle);
  } else {body.appendParagraph("Not Yet Image");}
  body.appendPageBreak();

  body.appendParagraph("รูปภาพแสดงการกดปุ่มใน POS").setHeading(DocumentApp.ParagraphHeading.HEADING1).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  imgPR = DriveApp.getFilesByName(owner+"_"+topic+"(Any Suffix");
  if (imgPR.hasNext() == 1){
    imgPara = body.appendImage(imgPR.next());
    ratio = imgPara.getHeight()/imgPara.getWidth();
    imgPara.setHeight(ratio*590).setWidth(590).getParent().setAttributes(imageStyle);
  } else {body.appendParagraph("Not Yet Image");}

  body.appendParagraph("  Any Further Information")
  body.appendParagraph("\n\n\n\n(                                                  )                                          (                                                  )\n\n");
  body.appendParagraph("           Marketing Manager                                                        Stor Manager...........................\n");
  body.appendPageBreak();

  // Optionally, you can move the document to a specific folder
  var folder = DriveApp.getFolderById("1A6qG0dN068iZNysJQA0jYuunBzHxWOR_");
  folder.addEditors(["email of anyone allow to edit and execute this script"]);
  
  folder.addViewers(["email of anyone allow to view the project"]);

  var file = DriveApp.getFileById(doc.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  // Participant's Email Controller
  var party = [];

  doc.addEditors(["email of anyone allow to edit and execute this script"]);

  doc.addViewers(["email of anyone allow to view the project"]);
}

/**
* @OnlyCurrentDoc
* @param {Object} e The bookMemo() event object.
*/

function bookMemo(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = sheet.getActiveRange();

  var topic = row.getValues()[0][1];
  var method = row.getValues()[0][3];
  var startDate = row.getValues()[0][5];
  var endDate = row.getValues()[0][6];
  var docLink = row.getValues()[0][8];

  if (endDate == "") var calendList = CalendarApp.getEvents(startDate, startDate);
  else var calendList = CalendarApp.getEvents(startDate, endDate);
  if(calendList.length > 0){
    for(var i = 0; i < calendList.length; i++){
      if(calendList[i].getTitle=="Project " + topic){
        calend = calendList[i];
        break;
      }
    }
    if (!(startDate == ""))
      if (endDate == "") calend = CalendarApp.createAllDayEvent("โครงการ " + topic, startDate);
      else calend = CalendarApp.createEvent("Project " + topic, startDate, endDate);
  } else {
    if (!(startDate == ""))
      if (endDate == "") calend = CalendarApp.createAllDayEvent("โครงการ " + topic, startDate);
      else calend = CalendarApp.createEvent("Project " + topic, startDate, endDate);
    }

  calend.addGuest("Any attendees")
  .addGuest("Any attendees");

  calend.setDescription(method + ", Document Link: " + docLink + ", Folder Link: " + "https://drive.google.com/drive/folders/1A6qG0dN068iZNysJQA0jYuunBzHxWOR_?usp=drive_link")
  .addEmailReminder(9360).addEmailReminder(3600).addEmailReminder(2160);
}
