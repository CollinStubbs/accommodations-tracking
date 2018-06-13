function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Tracking Sheets')
  .addItem("Create new Tracking Sheets", "newYearSheets")
  .addItem("Archive Tracking Sheets", "archiveSheets")
  .addItem("Archive IIPs", "archiveIIPs")
  .addItem("Archive Student Tracking", "archiveST")
  //.addItem("Remove Last Sheet", "removeLastSheet")//too dangerous
  .addToUi();
}

function archiveST(){
  var yearFolders = DriveApp.getFolderById("0B1FKC1RnG8Dpb1pLZFQtcWEzLU0").getFiles();
  while(yearFolders.hasNext()){
    var yearHolder = yearFolders.next();
    
    var names = getArrayNames(yearHolder);
    if(names[0] != "${Student_Name}"){
      var studentArchiveFolder = findArchiveFolder(names, getStudentYearST(yearHolder.getName()), false);
      var theBlob = yearHolder.getBlob().getAs('application/pdf').setName(yearHolder.getName());
      var newFile = studentArchiveFolder.createFile(theBlob);
      yearHolder.setTrashed(true);
    }
  }
}


function archiveIIPs(){
  var yearFolders = DriveApp.getFolderById("0B1FKC1RnG8DpNGNmYV9FWEt6eWc").getFolders();
  while(yearFolders.hasNext()){
    var yearHolder = yearFolders.next();
    if(yearHolder.getName() != "0 - IIP Template and Resources"){
      var yearFiles = yearHolder.getFiles();
      while(yearFiles.hasNext()){
        var yrFileHolder = yearFiles.next();
        if(yrFileHolder.getName() != "IIP Template"){
          var names = getArrayNames(yrFileHolder);
          var studentArchiveFolder = findArchiveFolder(names, getStudentYear(yearHolder.getName()), true);
          var theBlob = yrFileHolder.getBlob().getAs('application/pdf').setName(yrFileHolder.getName());
          var newFile = studentArchiveFolder.createFile(theBlob);
        }
      }
    }
  }
}

function findArchiveFolder(names, studentYear, isIIP){
  var archYearFolders = DriveApp.getFolderById("0B1FKC1RnG8DpVnBUNUdIeklsVmM").getFolders();
  while(archYearFolders.hasNext()){
    var folderHolder = archYearFolders.next();
    var x = folderHolder.getName();
    if(folderHolder.getName().indexOf(studentYear) > -1){
      var studentFolders = folderHolder.getFolders();
      while(studentFolders.hasNext()){
        var studentFolder = studentFolders.next();
        var y = studentFolder.getName();
        if(studentFolder.getName().indexOf(names[0]) > -1 && studentFolder.getName().indexOf(names[1]) > -1){
          var infoFolders = studentFolder.getFolders();
          while(infoFolders.hasNext()){
            var infoHolder = infoFolders.next();
            if(isIIP){
              if(infoHolder.getName().indexOf("Archived IIP") > -1){
                return infoHolder;
              }
              else{
               return studentFolder; 
              }
            }
            else{
               return studentFolder; 

            }
          }
          return studentFolder; 
        }
        else{
         console.log(names);
        }
      }
    }
   
  }
  
}

function getStudentYear(yearString){
  var year = yearString.substring(9,13);
  return year;
}
function getStudentYearST(yearString){
  var check = yearString.substring(0,1);
  if(Number(check) == 1){
   check = yearString.substring(0,2);
  }
  var year = Number((new Date()).getFullYear()) + (12 - Number(check));
  return year;
}

function getArrayNames(yearFile){
  var fileName = yearFile.getName();
  var index = fileName.indexOf("-");
  var index2 = fileName.indexOf("-", index+1);
  var name = fileName.substring(index+2, index2);
  name = name.split(" ");
  return name;
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