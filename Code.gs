function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Tracking Sheets')
  .addItem("Create new Tracking Sheets", "newYearSheets")
  .addItem("Archive Tracking Sheets", "archiveSheets")
  //.addItem("Remove Last Sheet", "removeLastSheet")//too dangerous
  .addToUi();
}

function archiveSheets(){
  var folders = DriveApp.getFoldersByName("IIP Tracking").next().getFolders();//.next().getFiles();
  var sheetFolder = null;
  var yearFolders = null;
  
  while(folders.hasNext()){
    var folder = folders.next();
    if(folder.getName() == "Accommodations Tracking"){
      sheetFolder = folder.getFiles();
    }
    else if(folder.getName() == "Student Files"){
      yearFolders = folder.getFolders(); 
    }
  }
  //Don't want to create a new studentFolder iterator everytime -> arrays?
  //  while(sheetFolder.hasNext()){
  //    var sheetHolder = sheetFolder.next();
  //    if(sheetHolder.getName().indexOf("TEMPLATE") == -1){
  //      var stdName = standardName(sheetHolder.getName());
  //      
  //    }
  //    
  //  }
  
  //this is not the best option but it'll do
  while(sheetFolder.hasNext()){
    var sheetHolder = sheetFolder.next();
    if(sheetHolder.getName().indexOf("TEMPLATE") == -1){
      
      var studentFolder = DriveApp.getFoldersByName(sheetHolder.getName());
      if(studentFolder.hasNext()){
        
        studentFolder = studentFolder.next();
        var ss = SpreadsheetApp.open(sheetHolder);
        console.log(ss.getName());
        var sheets = ss.getSheets();
        
        hideSheets(sheets);//GAS does not export hidden sheets
        var theBlob = ss.getBlob().getAs('application/pdf').setName(ss.getName()+" "+(new Date).getFullYear());
        var newFile = studentFolder.createFile(theBlob);
        displaySheets(sheets);
      }
      else{
        //new students 
        var ss = SpreadsheetApp.open(sheetHolder);
        var classYear = getClassYear(ss);
        
        while(yearFolders.hasNext()){
          var yearHolder = yearFolders.next();
          if(yearHolder.getName().indexOf(classYear) > -1){
            var newFolder = yearHolder.createFolder(ss.getName());
            var sheets = ss.getSheets();
            
            hideSheets(sheets);//GAS does not export hidden sheets
            var theBlob = ss.getBlob().getAs('application/pdf').setName(ss.getName()+" "+(new Date).getFullYear());
            var newFile = newFolder.createFile(theBlob);
            displaySheets(sheets);
            
          }
        }
        
      }
      //console.log(studentFolder);
    }
  }
  
}

function getClassYear(ss){
  var sheets = ss.getSheets();
  var sheet = sheets[sheets.length-1];
  var gradeSTR = sheet.getRange(2,1).getDisplayValue();
  var index = gradeSTR.indexOf(":");
  var subs = gradeSTR.substring(index+1, index+4).trim();
  var grade = Number(subs);
  
  var curYear = Number((new Date()).getFullYear());
  var classYear = (12-grade)+curYear;
  
  return classYear;
}

function hideSheets(sheets){
  for(var i = 0; i< sheets.length-1; i++){
    sheets[i].hideSheet(); 
  }
}

function displaySheets(sheets){
  for(var i = 0; i< sheets.length-1; i++){
    sheets[i].showSheet(); 
  }
}

function standardName(name){
  var std = "";
  
  name = name.toLowerCase();
  var splitter = name.split(',');
  var first = splitter[1].trim();
  var last = splitter[0].trim();
  
  std = first+last;
  
  return std;
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

function removeLastSheet(){
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
      ss.deleteSheet(sheet);
      // }
    }
  }
}