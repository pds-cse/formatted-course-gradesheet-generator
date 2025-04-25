var userSheet = SpreadsheetApp.openById("1rvMEg0IZDMoobsX02PBqeFH0xDaBty8zCxDkxe5lQQ0").getSheetByName("User");
var ui = SpreadsheetApp.getUi();
function onOpen(){
  
  ui.createMenu('GRADESHEET MENU')
      .addItem('GENERATE A GRADESHEET FOR ME', 'create_gradesheet')
      .addItem('Update Student List', 'getStudentList')
      .addToUi();
  
  processStatusCell.setValue("");
  SpreadsheetApp.flush();
}
function onEdit() {
  processStatusCell.setValue("");
  SpreadsheetApp.flush();
}
