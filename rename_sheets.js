function renameFiles(folder) {
  var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  var numOfFilesRenamed = 0;

  while (files.hasNext()) {
    numOfFilesRenamed++;
    var individualSheet = files.next();
    var fileID = individualSheet.getId();
    var individualFile = SpreadsheetApp.openById(fileID);
    var individualSheet = individualFile.getSheetByName("DynamicReport");

    if (individualSheet === null) {
      Logger.log("Sheet 'DynamicReport' not found in file: " + fileId);
      numOfFilesRenamed--;
      continue;
    }
    var cellB2Value = individualSheet.getRange("B2").getValue();
    var arrayOfLines = cellB2Value.split("\n");
    var courseNo = arrayOfLines[3].split(":").pop().trim();
    var sectionNo = arrayOfLines[4].split(":").pop().trim();
  
    var oldFileName = individualFile.getName();
    var newFileName = courseNo + " - Section " + sectionNo;
    individualFile.setName(newFileName);
    Logger.log("Renaming " + oldFileName + " to " + newFileName);
  }
  Logger.log("Renamed " + numOfFilesRenamed + " files");
}
