//// provide editing access to user on the sheets and folders

var semester = "Fall 2024";
var spreadSheetId = "1Nh3tuywXsUbd1ptVohDrbkOO2UfKDbMqJQyj1JNHlRw";   // background processes and student database
var spreadsheet = SpreadsheetApp.openById(spreadSheetId);
var userSheet = SpreadsheetApp.openById("1rvMEg0IZDMoobsX02PBqeFH0xDaBty8zCxDkxe5lQQ0").getSheetByName("User");
var destinationFolderId = "";
var destinationFolder = DriveApp.getFolderById(destinationFolderId);

//// process status showing cell: D5
var processStatusCell = userSheet.getRange("D5");
processStatusCell.setFontColor("Black");
processStatusCell.setValue("Processing please wait ... This will take just a few seconds");
SpreadsheetApp.flush();

//// list of info provided by user, fetched from the User sheet
var infoSheet = spreadsheet.getSheetByName("Info");
var errorMessage = infoSheet.getRange("D1:D3").getValues();
var courseNo = infoSheet.getRange("B1").getValue();
var selectedSection = infoSheet.getRange("B2").getValue();
var facultyIntitial = infoSheet.getRange("B3").getValue();
var facultyFullName = infoSheet.getRange("C3").getValue();
var facultyEmail = infoSheet.getRange("C4").getValue();
var userEmail = Session.getActiveUser();

//// get sorted student data (SL No. Id, & Name)
var processSheet = spreadsheet.getSheetByName("Processing");
var data = processSheet.getRange("L2:N").getValues();
var slNoColumn = processSheet.getRange("L2:L").getValues();

//// gradesheet file name
var fileName = courseNo + "-" + selectedSection + " Final Gradesheet " + semester + " - " +           
  facultyIntitial;

//// template gradesheets
var templateGradesheetUrls = infoSheet.getRange("F1:F8").getValues();
var startingRowOfStudentDataInGradesheet = 11;


//// ##############################################################################################################
function create_gradesheet() {

  //// if course, faculty, or section are incorrect, do not run
  if(errorMessage[0][0].length > 0 || errorMessage[1][0].length > 0 || errorMessage[2][0].length > 0) {
    processStatusCell.setFontColor("Red");
    processStatusCell.setValue("Process terminated for incorrect section/course selection. You are not a faculty of the selected section/course.");
    return;
  }

  //// if user is trying to generate gradesheet for others, do not run
  if(facultyEmail != userEmail) {
    Logger.log("Sorry, you are not logged in as " + userEmail + " but trying to generate for " + facultyEmail);
    processStatusCell.setFontColor("Red");
    processStatusCell.setValue("Sorry, you are logged in as '" + userEmail + "' but trying to generate for '" + facultyFullName + "'. This won't work.");
    return;
  }
  
  //// fetching template sheets according to the course selected
  if(courseNo == "CSE250") {
    templateSheet = SpreadsheetApp.openByUrl(templateGradesheetUrls[0]);
  }
  else if(courseNo == "CSE251") {
    templateSheet = SpreadsheetApp.openByUrl(templateGradesheetUrls[1]);
  }
  else if(courseNo == "CSE350") {
    templateSheet = SpreadsheetApp.openByUrl(templateGradesheetUrls[2]);
  }
  else if(courseNo == "CSE460") {
    templateSheet = SpreadsheetApp.openByUrl(templateGradesheetUrls[3]);
  }
  else if(courseNo == "CSE428") {
    templateSheet = SpreadsheetApp.openByUrl(templateGradesheetUrls[4]);
  }
  else if(courseNo == "CSE481") {
    templateSheet = SpreadsheetApp.openByUrl(templateGradesheetUrls[5]);
  }
  else if(courseNo == "CSE482") {
    templateSheet = SpreadsheetApp.openByUrl(templateGradesheetUrls[6]);
  }
  else if(courseNo == "EEE476") {
    templateSheet = SpreadsheetApp.openByUrl(templateGradesheetUrls[7]);
  }
  else {
    Logger.log("Invalid Course No.")
    return;
  }
  // var newGradesheet = SpreadsheetApp.openById("1rujEgmUgx0dMiIA3Vl33v7t-DXgSGAgULY6BjcJ2JfQ");  // for testing
  
  // deleting all files within the folder. ** if the file ownership is transfered, I cannot delete such files
  // while(destinationFolder.getFiles().hasNext()) {
  //   var file = destinationFolder.getFiles().next();
  //   file.setTrashed(true);
  // }

  //// new gradesheet file created
  var newGradesheetFile = DriveApp.getFileById(templateSheet.getId()).makeCopy(fileName, destinationFolder);
  var newGradesheet = SpreadsheetApp.openById(newGradesheetFile.getId());
  var finalSheet = newGradesheet.getSheetByName("Final GradeSheet");
  var midtermSheet = newGradesheet.getSheetByName("Midterm GradeSheet");

  
  //// set data and header in the gradesheet
  finalSheet.getRange("C3").setValue(semester);
  finalSheet.getRange("C5").setValue(courseNo);
  finalSheet.getRange("C7").setValue(selectedSection);
  finalSheet.getRange("C8").setValue(facultyFullName + " (" + facultyIntitial + ")");
  finalSheet.getRange("A" + startingRowOfStudentDataInGradesheet + ":C" + (data.length + startingRowOfStudentDataInGradesheet - 1)).setValues(data);
  midtermSheet.getRange("A" + startingRowOfStudentDataInGradesheet + ":A"  + (data.length + startingRowOfStudentDataInGradesheet - 1)).setValues(slNoColumn);
  
  //// get the row number of "1st blank row" of the two row gap
  for (i=0; i<data.length; i++) {
    var nameLength1 = data[i].map(w => w.length);
    if(nameLength1[2] < 5) {
      var index = i + 1;
      break;
    }
  }

  /////////////////////////////////// removing content and border of the two row gap
  //// from the 'Final Gradesheet'sheet
  var rowsToOmit = finalSheet.getRange((index+startingRowOfStudentDataInGradesheet-1), 1, 2, finalSheet.getLastColumn());   // student data starts from row 11
  rowsToOmit.clearContent();
  rowsToOmit.setBorder(null, true, null, true, true, true, "white", SpreadsheetApp.BorderStyle.SOLID);
  
  //// from the 'Midterm Gradesheet'sheet
  rowsToOmit = midtermSheet.getRange((index+startingRowOfStudentDataInGradesheet-1), 1, 2, midtermSheet.getLastColumn());   //// student data starts from row 11
  rowsToOmit.clearContent();
  rowsToOmit.setBorder(null, true, null, true, true, true, "white", SpreadsheetApp.BorderStyle.SOLID);
  
  /////////////////////////////////// deleting empty rows in bottom
  /////////////////////////////////// SL Column of gradesheet template must be blank in both Final and Midterm Sheets
  //// from the 'Final Gradesheet'sheet
  var serialNo = finalSheet.getRange("A" + startingRowOfStudentDataInGradesheet + ":A").getValues();
  for(i = serialNo.length-1; i>=0; i--) {
    if(serialNo[i] != "") {
      break;
    }
    else {
      finalSheet.deleteRow((i+startingRowOfStudentDataInGradesheet));
    }
  }

  //// from the 'Midterm Gradesheet'sheet
  serialNo = midtermSheet.getRange("A" + startingRowOfStudentDataInGradesheet + ":A").getValues();
  for(i = serialNo.length-1; i>=0; i--) {
    if(serialNo[i] != "") {
      break;
    }
    else {
      midtermSheet.deleteRow((i+startingRowOfStudentDataInGradesheet));
    }
  }
  
  ////////////////////////////// Transfer ownership of the gradesheet file
  Drive.Permissions.insert(
    {
      'role': 'owner',
      'type': 'user',
      'value': facultyEmail
    },
    newGradesheetFile.getId(),
    {
      'sendNotificationEmails': true
    }
  );
  
  //// removing myself from editing the generated gradesheet, only the faculty should have access
  newGradesheet.removeEditor("purbayan.das@bracu.ac.bd");
  processStatusCell.setFontColor("Black");
  processStatusCell.setValue("Done! Please check your email");
  SpreadsheetApp.flush();
  Utilities.sleep(8000);   // 8 ms delay
  processStatusCell.setValue("");
}
