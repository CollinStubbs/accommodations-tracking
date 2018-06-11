function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Tracking Sheets')
  .addItem("Create new Tracking Sheets", "newYearSheets")
  .addToUi();
}

function newYearSheets() {
  var files = DriveApp.getFoldersByName("Accommodations Tracking").next().getFiles();
  
  while(files.hasNext()){
    var file = files.next();
    var name = file.getName();
    if(name.indexOf("TEMPLATE") > -1){
      
    }
    else{
      var ss = SpreadsheetApp.open(file);
      var sheets = ss.getSheets();
      var sheet = sheets[sheets.length-1]; 
      // if(sheet.getName().indexOf("Grade :") >-1){
      var gradeSTR = sheet.getRange(2,1).getDisplayValue();
      var index = gradeSTR.indexOf(":");
      var subs = gradeSTR.substring(index+1, index+4).trim();
      var grade = Number(subs);
      //  ss.renameActiveSheet("Grade "+subs);
      if(grade<12){
        var newGrade = grade+1;
        ss.setActiveSheet(sheet);
        var newsheet = ss.duplicateActiveSheet();
        ss.setActiveSheet(newsheet);
        ss.renameActiveSheet("Grade " + newGrade);
        
        newsheet.getRange(6, 1, 70, 15).clearContent();
        newsheet.getRange(2,1).setValue("Grade: "+newGrade);
        
      }
      // }
    }
  }
}