function getStudentList() {

  // // Check who is running the code. Only PDS can update student list
  var userEmail = Session.getActiveUser();
  if(userEmail != "purbayan.das@bracu.ac.bd") {
    processStatusCell.setFontColor("Red");
    processStatusCell.setValue("Sorry, only the admin can run this");
    return;
  }

  // // asking user which one to update: Before or After Confirmation
  var userStudentListType = ui.prompt("'Before' or 'After'?");
  var studentListType = userStudentListType.getResponseText();

  var folderIdBefore = "1cKFRkRjTl-6NZZ7eeouajMUrev4hgNUK";
  var folderIdAfter = "1V30vSQjf3RkPq8a1taw20Rp6mJLp7k2J";

  // // The following block renames files in 'Before' or 'After' folder, as specified by user. e.g. "CSE250_1_Before"
  var folder;
  if(studentListType == "Before") {
    folder = DriveApp.getFolderById(folderIdBefore);
    renameFiles(folder);
  }
  else if(studentListType == "After") {
    folder = DriveApp.getFolderById(folderIdAfter);
    renameFiles(folder);
  }
  else {
    Logger.log("Invalid input");
    return;
  }

  // // this block clears already existing data for all the courses in all the course-sheets (both before and after)
  var courses = ["CSE250", "CSE251", "CSE350", "CSE460", "CSE428", "CSE481", "CSE482", "EEE476"];
  var sheetName;
  var destinationSheet;
  for (i=0; i<courses.length; i++) {
    sheetName = courses[i] + "_" + studentListType;
    destinationSheet = spreadsheet.getSheetByName(sheetName);
    destinationSheet.getRange("A2:D").clearContent();
  }
  
  var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);  // get all the gsheet files in that folder
  
  while(files.hasNext()) {
    file = files.next();
    var fileName = file.getName();
    var courseNo = fileName.slice(0, 6);
    var sectionNoString = fileName.split(" ").pop();
    var sectionNoNumber = parseInt(sectionNoString);
    Logger.log("Getting data from: " + fileName);

    var fileId = file.getId();
    var sourceSpreadsheet = SpreadsheetApp.openById(fileId);
    var sourceSheet = sourceSpreadsheet.getSheetByName("DynamicReport");

    var sourceRange = "DynamicReport!B4:D";
    if (sourceSheet === null) {
      Logger.log("Sheet 'DynamicReport' not found in file: " + fileId);
      continue;
    }
    
    var sourceData = sourceSheet.getRange(sourceRange).getValues();  // this contains SL, ID, and Name columns w/o header

    // Filter out rows where all cells are empty
    sourceData = sourceData.filter(function (row) {
      return row.some(function (cell) {
        return cell !== "";
      });
    });

    // Add an additional column with the section number at the beginning
    var fileNumberColumn = Array(sourceData.length).fill([sectionNoNumber]);
    sourceData = sourceData.map(function (row, index) {
      return fileNumberColumn[index].concat(row);
    });

    // // dumping data in course-specific sheets
    var sheetName = courseNo + "_" + studentListType;
    var destinationSheet = spreadsheet.getSheetByName(sheetName);
    destinationSheet.getRange(destinationSheet.getLastRow()+1, 1, sourceData.length, sourceData[0].length).setValues(sourceData);
  }

  // // sorting data in all course-specific sheets
  for (i=0; i<courses.length; i++) {
    sheetName = courses[i] + "_" + studentListType;
    destinationSheet = spreadsheet.getSheetByName(sheetName).getRange("A2:D").sort({column: 1, ascending: true});
  }
}
