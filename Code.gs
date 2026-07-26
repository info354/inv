

function nextClient(){
  var a1 = s.getRange('A1');
  var c1 = s.getRange('C1');
  c1.setValue(a1.getValue());
}

function inv_header_1003() {
  for (var i = 0; i < d.length; i++) {
    var name = d[i];
    var sheet = ss.getSheetByName(name);
    var range = sheet.getRange('A1:C4');
    var tempSheet = ss.getSheetByName('1003');
    var srcRange = tempSheet.getRange('A1:C4');
    srcRange.copyTo(range);
    var range1 = sheet.getRange(1, 3, 4, 1);
    var range2 = sheet.getRange(1, 3, 4, sheet.getLastColumn() - 2);
    range1.copyTo(range2, { formatOnly: true });

    var bg = cell.offset(i, 0).getBackground();
    if (bg == 'lightyellow') {
      var bg = 'lightgreen';
    } else { var bg = 'lightyellow'; }
    cell.offset(i, 0).setBackground(bg);

  }
}

function getfx() {
  for (var i = 0; i < d.length; i++) {

    var sheet = ss.getSheetByName(d[i]);
    var col2 = sheet.getRange('6:6').getValues()[0].indexOf('R.Bsc') + 1;
    var cell2 = sheet.getRange(9, col2);
    var fx = cell2.getFormulaR1C1();

    cell.offset(i, 1).setBackground('yellow').setValue(fx.toString().replace('=', ''));

    Logger.log(d[i] + ' | ' + fx);
  }
}

function formatDate() {
  var src = ss.getSheetByName('1001').getRange('A9');
  for (i = 0; i < d.length; i++) {
    var desSheet = ss.getSheetByName(d[i]);
    var desRange = desSheet.getRange(9, 1, desSheet.getMaxRows() - 8, 1);
    src.copyTo(desRange, { formatOnly: true });
    desSheet.setColumnWidth(1, 90);
  }
}

function renameSheets() {
  for (i = 0; i < d.length; i++) {
    var name1 = d2[i][0].toString();
    var name2 = d2[i][1].toString();
    Logger.log(name1);
    Logger.log(name2);
    ss.getSheetByName(name1).setName(name2);
  }
}

//Thay đổi từ PSVN INV 2105



function removeTabColor() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();
  s.setTabColor(null);
}

function removeTabColorAll() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();
  var s2 = ss.getSheets();
  for (i = s.getIndex(); i < s2.length; i++) {
    s2[i].setTabColor(null);
  }

}

function removeTabColorList() {
  for (i = 0; i < d.length; i++) {
    var sh = ss.getSheetByName(d[i].toString());
    sh.setTabColor(null);
    cell.offset(i, 0).setBackground('lightgreen');
    SpreadsheetApp.flush();
  }

}

function toArchiveList() {  // Dùng ok
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var destination = SpreadsheetApp.openById('1ucme3trAD1mbsi7t9Z8Wxc6uzojbAzYskOCVtHmqfSw');

  for (i = 0; i < d.length; i++) {
    var sheet = ss.getSheetByName(d[i]);
    var sheetName = sheet.getName();
    var newSheet = sheet.copyTo(destination);
    newSheet.setName(sheetName);
    ss.deleteSheet(sheet);
    cell.offset(i, 0).setBackground('lightgreen');
    SpreadsheetApp.flush();
  }
}

function getSheetsFromArchiveSelected() {
  var ui = SpreadsheetApp.getUi();
  var source = SpreadsheetApp.openById('1ucme3trAD1mbsi7t9Z8Wxc6uzojbAzYskOCVtHmqfSw');
  for (i = 0; i < d.length; i++) {
    try {
      var sheet = source.getSheetByName(d[i]);
      var newSheet = sheet.copyTo(ss);
      newSheet.setName(d[i]);
      source.deleteSheet(sheet);
      cell.offset(i, 0).setBackground('lightgreen');
    } catch {
      ui.alert(d[i] + ' not exist!!!');
      cell.offset(i, 0).setBackground('lightyellow');
      continue;
    }


  }
}

function moveSheetsToArchive() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var destination = SpreadsheetApp.openById('1ucme3trAD1mbsi7t9Z8Wxc6uzojbAzYskOCVtHmqfSw');
  var index = ss.getActiveSheet().getIndex() - 2;
  var sheetName = ss.getActiveSheet().getName();

  ss.getActiveSheet().copyTo(destination);
  ss.deleteSheet(ss.getActiveSheet());
  SpreadsheetApp.setActiveSheet(ss.getSheets()[index]);

  var newSheet = destination.getSheetByName('Copy of ' + sheetName);
  newSheet.setName(sheetName);

}

function getSheetsFromArchive() {
  var sheetName = cell.getValue();
  var src = SpreadsheetApp.openById('1ucme3trAD1mbsi7t9Z8Wxc6uzojbAzYskOCVtHmqfSw');
  src.getSheetByName(sheetName).copyTo(ss);
  var newSheet = ss.getSheetByName('Copy of ' + cell.getValue());
  newSheet.setName(sheetName);
  ss.setActiveSheet(newSheet)
  ss.moveActiveSheet(s.getIndex() + 1);
  src.deleteSheet(src.getSheetByName(sheetName));

}

function moveHiddenSheetToArchive() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();
  var sheet = ss.getSheets();
  var index = s.getIndex();

  for (i = index; i < sheet.length; i++) {
    ss.setActiveSheet(sheet[i])
    var sheetName = sheet[i].getName();
    if (sheet[i].isSheetHidden() == true) {
      var destination = SpreadsheetApp.openById('1ucme3trAD1mbsi7t9Z8Wxc6uzojbAzYskOCVtHmqfSw');
      sheet[i].copyTo(destination);
      var newSheet = destination.getSheetByName('Copy of ' + sheetName);
      newSheet.setName(sheetName);
      newSheet.showSheet();
      ss.deleteSheet(sheet[i]);
    }
    SpreadsheetApp.flush();
  }
}

function moveEmptySheetToArchive() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();
  var sheet = ss.getSheets();
  var index = s.getIndex();
  var count = 0;
  for (i = index - 1; i < sheet.length; i++) {
    var old_name = sheet[i].getName();
    ss.toast(old_name + ' processing..');
    // ss.setActiveSheet(sheet[i])
    var sheetName = sheet[i].getName();
    if (sheet[i].getRange('A9').isBlank()) {

      count++;
      var destination = SpreadsheetApp.openById('1ucme3trAD1mbsi7t9Z8Wxc6uzojbAzYskOCVtHmqfSw');
      sheet[i].copyTo(destination);
      var newSheet = destination.getSheetByName('Copy of ' + sheetName);
      newSheet.setName(sheetName);
      newSheet.showSheet();
      destination.setActiveSheet(newSheet);
      destination.moveActiveSheet(2);
      ss.deleteSheet(sheet[i]);
      ss.toast(old_name + ' archiving..');
    }
    SpreadsheetApp.flush();
  }
  Browser.msgBox(old_name + ' archived!');
}

function empty_inv_list() {
  for (i = 0; i < d.length; i++) {
    var sheet = ss.getSheetByName(d[i]);
    var range = sheet.getRange(9, 1, sheet.getLastRow() - 8, sheet.getLastColumn() - 1);
    range.clearContent();
    cell.offset(i, 0).setBackground('lightgreen');
    SpreadsheetApp.flush();
  }
}

function getSheetLink() {
  for (i = 0; i < d.length; i++) {
    var row2 = row + i;
    var sName = d[i];
    var no = s.getRange(row2, col);
    var link = no.offset(0, 1);


    if (ss.getSheetByName(sName)) {
      var gid = ss.getSheetByName(sName).getSheetId();
      var sUrl = 'https://docs.google.com/spreadsheets/d/1oZ0E2Y0Uh19ZdPlbmdbBeo0zjPvH5CKaGA_hNHHoUDA/edit#gid=' + gid;
    } else { var sUrl = 'No Link'; }



    // try {
    //   var gid = ss.getSheetByName(sName).getSheetId();
    //   var sUrl = 'https://docs.google.com/spreadsheets/d/1oZ0E2Y0Uh19ZdPlbmdbBeo0zjPvH5CKaGA_hNHHoUDA/edit#gid=' + gid;

    // } catch (e) { var sUrl = 'No Link'; continue;  }

    link.setValue(sUrl);

  }
}
function clearRetouchConfirm() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (j = ss.getActiveSheet().getIndex() - 1; j < sheets.length; j++) {
    var s = sheets[j]
    if (s.isSheetHidden() != true) {
      ss.setActiveSheet(s);
      for (i = 5; i < s.getMaxColumns(); i++) {
        var basic = s.getRange(6, i);
        if (basic.getValue() == 'R.Bsc') {
          var retouchSumBackup = s.getRange(1, i, 1, 4);
          retouchSumBackup.clear();
          break;
        }
      }
      SpreadsheetApp.flush();
    }
    SpreadsheetApp.flush();
  }
  SpreadsheetApp.flush();
}

function fillRetouchList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();

  var fomular = '=iferror(if(sumifs(JW!AH:AH,JW!$A:$A,$A9,JW!$D:$D,$B9)=0,,sumifs(JW!AH:AH,JW!$A:$A,$A9,JW!$D:$D,$B9)),)';
  for (j = 0; j < d.length; j++) {
    var s = ss.getSheetByName(d[j]);
    if (s.isSheetHidden() != true) {
      for (i = 5; i < s.getMaxColumns(); i++) {
        var basic = s.getRange(6, i);
        if (basic.getValue() == 'R.Bsc') {
          var basic1 = basic.offset(3, 0);
          basic1.setValue(fomular);
          var retouchRange = s.getRange('G9:H15');
          var row = 9;
          var col = i;
          var rowNums = s.getLastRow() - 8;
          var colNums = 4;
          var retouchRange = s.getRange(row, col, rowNums, colNums);
          basic1.copyTo(retouchRange);
          var retouchSum = s.getRange(7, col, 1, 4);
          var retouchSumBackup = s.getRange(1, i, 1, 4);
          retouchSumBackup.setValues(retouchSum.getValues());
          break;
          cell.offset(1, 0).setBackground('lightgreen');
        }
      }
      SpreadsheetApp.flush();
    }
    SpreadsheetApp.flush();
  }
  SpreadsheetApp.flush();
}


function fillRetouchAll() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var fomular = '=iferror(if(sumifs(JW!AJ:AJ,JW!$A:$A,$A9,JW!$D:$D,$B9)=0,,sumifs(JW!AJ:AJ,JW!$A:$A,$A9,JW!$D:$D,$B9)),)';
  for (j = ss.getActiveSheet().getIndex(); j < sheets.length; j++) {
    var s = sheets[j]
    if (s.isSheetHidden() != true) {
      ss.setActiveSheet(s);
      for (i = 5; i < s.getMaxColumns(); i++) {
        var basic = s.getRange(6, i);
        if (basic.getValue() == 'R.Bsc') {
          var basic1 = basic.offset(3, 0);
          basic1.setValue(fomular);
          var retouchRange = s.getRange('G9:H15');
          var row = 9;
          var col = i;
          var rowNums = s.getLastRow() - 8;
          var colNums = 4;
          var retouchRange = s.getRange(row, col, rowNums, colNums);
          basic1.copyTo(retouchRange);
          var retouchSum = s.getRange(7, col, 1, 4);
          var retouchSumBackup = s.getRange(1, i, 1, 4);
          retouchSumBackup.setValues(retouchSum.getValues());
          break;
        }
      }
      SpreadsheetApp.flush();
    }
    SpreadsheetApp.flush();
  }
  SpreadsheetApp.flush();
}


function fillRetouchList2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var fomular = '=iferror(if(sumifs(JW!AH:AH,JW!$A:$A,$A9,JW!$D:$D,$B9)=0,,sumifs(JW!AH:AH,JW!$A:$A,$A9,JW!$D:$D,$B9)),)';
  for (j = 0; j < d.length; j++) {
    var s = ss.getSheetByName(d[j]);
    if (s.isSheetHidden() != true) {
      for (i = 5; i < s.getMaxColumns(); i++) {
        var basic = s.getRange(6, i);
        if (basic.getValue() == 'R.Bsc') {
          var basic1 = basic.offset(3, 0);
          basic1.setValue(fomular);
          var retouchRange = s.getRange('G9:H15');
          var row = 9;
          var col = i;
          var rowNums = s.getLastRow() - 8;
          var colNums = 4;
          var retouchRange = s.getRange(row, col, rowNums, colNums);
          basic1.copyTo(retouchRange);
          var retouchSum = s.getRange(7, col, 1, 4);
          var retouchSumBackup = s.getRange(1, i, 1, 4);
          retouchSumBackup.setValues(retouchSum.getValues());
          break;
        }
      }
      SpreadsheetApp.flush();
    }
    SpreadsheetApp.flush();
  }
  SpreadsheetApp.flush();
}

function fillRetouchAllValue() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var fomular = '=iferror(if(sumifs(JW!AH:AH,JW!$A:$A,$A9,JW!$D:$D,$B9)=0,,sumifs(JW!AH:AH,JW!$A:$A,$A9,JW!$D:$D,$B9)),)';
  for (j = ss.getActiveSheet().getIndex(); j < sheets.length; j++) {
    var s = sheets[j]
    if (s.isSheetHidden() != true) {
      ss.setActiveSheet(s);
      for (i = 5; i < s.getMaxColumns(); i++) {
        var basic = s.getRange(6, i);
        if (basic.getValue() == 'R.Bsc') {
          var basic1 = basic.offset(3, 0);
          basic1.setValue(fomular);
          var retouchRange = s.getRange('G9:H15');
          var row = 9;
          var col = i;
          var rowNums = s.getLastRow() - 8;
          var colNums = 4;
          var retouchRange = s.getRange(row, col, rowNums, colNums);

          retouchRange.setValues(retouchRange.getValues());
          break;
        }
      }
      SpreadsheetApp.flush();
    }
    SpreadsheetApp.flush();
  }
  SpreadsheetApp.flush();
}
function changeVhsPrice() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  s = ss.getActiveSheet();
  var sheets = ss.getSheets();

  for (i = s.getIndex(); i < sheets.length; i++) {
    ss.setActiveSheet(ss.getSheets()[i]);
    for (j = 4; j < ss.getSheets()[i].getLastColumn(); j++) {

      var VHS = ss.getSheets()[i].getRange(6, j);
      var VHS_price = VHS.offset(2, 0);

      if (VHS.getValue() == 'VHS' && VHS_price.getValue() == 29) {
        VHS_price.setValue(27);
        break;
      }
    }
    SpreadsheetApp.flush();
  }
  SpreadsheetApp.flush();
}
function copySum() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aS = ss.getActiveSheet();
  var s = ss.getSheets();

  for (i = aS.getIndex(); 1 < s.length; i++) {
    if (ss.getSheets()[i].isSheetHidden() != true) {
      ss.setActiveSheet(ss.getSheets()[i]);
      var src = ss.getSheets()[i].getRange(7, 4, 1, 1);
      var des = ss.getSheets()[i].getRange(7, 4, 1, ss.getSheets()[i].getLastColumn() - 4);
      src.setValue('=if(sum(D9:D)=0,"-",sum(D9:D))');
      src.copyTo(des);
    }
    SpreadsheetApp.flush();
  }
}
function moveSheetByIndex() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();
  var sheetName = s.getName();
  //  var toIndex = s.getRange('D1');

  var toIndex = SpreadsheetApp.getUi().prompt('Enter your input').getResponseText();


  ss.moveActiveSheet(toIndex);

}
function renameSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();
  var sheets = ss.getSheets();
  var r = s.getActiveRange();
  var d = r.getValues();

  for (i = s.getIndex(); i < sheets.length; i++) {
    sheets[i].setName(sheets[i].getName().toString().replace(/-.+/, ""));
  }

}
function getIndex() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();
  var index = s.getIndex();


  Browser.msgBox('The index of current Sheet', index, Browser.Buttons.OK);

}
function showSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheets();
  for (i = 0; i < s.length; i++) {
    ss.getSheets()[i].showSheet();
  }
}
function sheetnames() { // ahab facit 2011
  var out = new Array()
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i = 0; i < sheets.length; i++) out.push([sheets[i].getName()])
  return out
}
function sortSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetNameArray = [];
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    sheetNameArray.push(sheets[i].getName());
  }
  sheetNameArray.sort();
  for (var j = 0; j < sheets.length; j++) {
    ss.setActiveSheet(ss.getSheetByName(sheetNameArray[j]));
    ss.moveActiveSheet(j + 1);
  }
}
function getLinks() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();
  var cell = s.getActiveCell();
  var row = cell.getRow();
  var col = cell.getColumn();
  var r = s.getActiveRange();
  var d = r.getValues();
  for (i = 0; i < d.length; i++) {
    var name = s.getRange(row + i, col);
    var url = name.offset(0, 1);
    url.setValue(DriveApp.getFilesByName(name.getValue()).next().getUrl().replace('?usp=drivesdk', '#'));
  }
}

function generatePdf_1() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var invFolder = DriveApp.getFolderById('1B0Z7G2wzRYE8nXsHgs8TaZvOIleEbzSP');
  var month = ss.getName().split(' ').pop();
  
  var monthFolder = invFolder.getFoldersByName(month);
  if(monthFolder.hasNext()){
    var targetFolder = monthFolder.next();
  }else{
    var targetFolder = invFolder.createFolder(month);
  }

  

  // exit();



  var originalSpreadsheet = SpreadsheetApp.getActive();
  originalSpreadsheet.setActiveSheet(originalSpreadsheet.getActiveSheet());
  var name = originalSpreadsheet.getName() + ' - ' + originalSpreadsheet.getActiveSheet().getName();
  var newSpreadsheet = SpreadsheetApp.create(name);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet = originalSpreadsheet.getActiveSheet();
  sheet.copyTo(newSpreadsheet);
  newSpreadsheet.deleteSheet(newSpreadsheet.getSheetByName('Sheet1'));
  //Save the desired sheet as pdf
  var pdf = DriveApp.getFileById(newSpreadsheet.getId()).getAs('application/pdf');
  var saveCopy = DriveApp.getFolderById(targetFolder.getId()).createFile(pdf);
  //Delete temporary spreadsheet
  DriveApp.getFilesByName(name).next().setTrashed(true);
  SpreadsheetApp.flush();
}

function generatePdf_all() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aS = ss.getActiveSheet();
  var s = ss.getSheets();
  var A9 = aS.getRange('A9');
  for (i = aS.getIndex() - 1; i < s.length; i++) {
    if (ss.getSheets()[i].isSheetHidden() != true) {
      var originalSpreadsheet = SpreadsheetApp.getActive();
      originalSpreadsheet.setActiveSheet(originalSpreadsheet.getSheets()[i]);
      var name = originalSpreadsheet.getName() + ' - ' + originalSpreadsheet.getSheets()[i].getName();
      var newSpreadsheet = SpreadsheetApp.create(name);
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      sheet = originalSpreadsheet.getActiveSheet();
      sheet.copyTo(newSpreadsheet);
      newSpreadsheet.getSheets()[1].showSheet();
      newSpreadsheet.deleteSheet(newSpreadsheet.getSheetByName('Sheet1'));
      //Save the desired sheet as pdf
      var pdf = DriveApp.getFileById(newSpreadsheet.getId()).getAs('application/pdf');
      var saveCopy = DriveApp.createFile(pdf);
      //Delete temporary spreadsheet
      DriveApp.getFilesByName(name).next().setTrashed(true);
    }
    SpreadsheetApp.flush();
  }
}

function generatePdfList() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var invFolder = DriveApp.getFolderById('1B0Z7G2wzRYE8nXsHgs8TaZvOIleEbzSP');
  var month = ss.getName().split(' ').pop();
  
  var monthFolder = invFolder.getFoldersByName(month);
  if(monthFolder.hasNext()){
    var targetFolder = monthFolder.next();
  }else{
    var targetFolder = invFolder.createFolder(month);
  }





  for (i = 0; i < d.length; i++) {
    var name = ss.getName() + ' - ' + d[i];
    var newSpreadsheet = SpreadsheetApp.create(name);
    var sheet = ss.getSheetByName(d[i]);
    sheet.copyTo(newSpreadsheet);
    // newSpreadsheet.getSheets()[1].showSheet();
    newSpreadsheet.deleteSheet(newSpreadsheet.getSheetByName('Sheet1'));
    //Save the desired sheet as pdf
    var pdf = DriveApp.getFileById(newSpreadsheet.getId()).getAs('application/pdf');
    var saveCopy = DriveApp.getFolderById(targetFolder.getId()).createFile(pdf);
    //Delete temporary spreadsheet
    DriveApp.getFilesByName(name).next().setTrashed(true);
    cell.offset(i, 0).setBackground('lightgreen');
    SpreadsheetApp.flush();
  }
}

function generatePdfList_2() {






  for (i = 0; i < d.length; i++) {
    var name = ss.getName() + ' - ' + d[i];
    var newSpreadsheet = SpreadsheetApp.create(name);
    var sheet = ss.getSheetByName(d[i]);
    sheet.copyTo(newSpreadsheet);
    // newSpreadsheet.getSheets()[1].showSheet();
    newSpreadsheet.deleteSheet(newSpreadsheet.getSheetByName('Sheet1'));
    //Save the desired sheet as pdf
    var pdf = DriveApp.getFileById(newSpreadsheet.getId()).getAs('application/pdf');
    var saveCopy = DriveApp.createFile(pdf);
    //Delete temporary spreadsheet
    DriveApp.getFilesByName(name).next().setTrashed(true);
    cell.offset(i, 0).setBackground('lightgreen');
    SpreadsheetApp.flush();
  }
}

function fixJW() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();
  var r = s.getRange(1, 1, s.getMaxRows(), s.getMaxColumns());
  r.copyTo(r, { contentsOnly: true });
}

function moveThisSheetNameHere() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var s = ss.getActiveSheet();
    var index = s.getIndex();
    var name = s.getActiveCell().getValue();
    var target_sheet = ss.getSheetByName(name);
    target_sheet.activate();
    target_sheet.getRange('A9').activate();
    ss.moveActiveSheet(index);
  }
  catch (e) {
    var cell = SpreadsheetApp.getActiveSpreadsheet().getActiveCell();
    var name = Browser.inputBox('Enter Tab Name:', '', Browser.Buttons.OK_CANCEL);
    if (name) {
      try {
        //create a regex out of "name"
        regexp = new RegExp(name, 'g');
        //get all the sheets
        var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
        //go through all the sheets
        for (var i = 0; i < sheets.length; i++) {
          //if sheet name matches the user entry then activate and go out of the loop
          if (sheets[i].getName().match(regexp)) {
            sheets[i].activate();
            
            ss.moveActiveSheet(index);
            break;
          }
        }
      } catch (e) {
        Browser.msgBox('Sheet named: "' + name + '" does not exists!');
      }
    }
  };
}


function goToSheetByName() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var s = ss.getActiveSheet();
    var name = s.getActiveCell().getValue();
    var target_sheet = ss.getSheetByName(name);
    target_sheet.activate();
  }
  catch (e) {
    var cell = SpreadsheetApp.getActiveSpreadsheet().getActiveCell();
    var name = Browser.inputBox('Enter Tab Name:', '', Browser.Buttons.OK_CANCEL);
    if (name) {
      try {
        //create a regex out of "name"
        regexp = new RegExp(name, 'g');
        //get all the sheets
        var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
        //go through all the sheets
        for (var i = 0; i < sheets.length; i++) {
          //if sheet name matches the user entry then activate and go out of the loop
          if (sheets[i].getName().match(regexp)) {
            sheets[i].activate();
            break;
          }
        }
      } catch (e) {
        Browser.msgBox('Sheet named: "' + name + '" does not exists!');
      }
    }
  };
}



function newClientList() {
  var copy = ss.getSheetByName('_Copy');

  for (i = 0; i < d.length; i++) {
    var clientNo = d[i];
    try {
      var newSheet = copy.copyTo(ss);
      newSheet.setName(clientNo);
    } catch (e) {
      Browser.msgBox('Đã có lỗi!!!', d[i] + ' đã tồn tại. Chọn tên khác', Browser.Buttons.OK);
      ss.deleteSheet(newSheet)
    }
    cell.offset(i, 0).setBackground('lightgreen');
  }

}

function fillRetouch1() {
  removeTabColor();
  unHideColumns();
  removeMoney();
  removeEmptyRows();
  money();
  copyFormat();
  remove_header();
  // Fill retouch    
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();


  var fomular = '=iferror(if(sumifs(JW!AH:AH,JW!$C:$C,$B9,JW!$B:$B,$A9)=0,,sumifs(JW!AH:AH,JW!$C:$C,$B9,JW!$B:$B,$A9)))';
  for (i = 4; i < s.getMaxColumns(); i++) {
    var basic = s.getRange(6, i);
    if (basic.getValue() == 'R.Bsc') {
      var basic1 = basic.offset(3, 0);
      basic1.setValue(fomular);
      var row = 9;
      var col = i;
      var rowNums = s.getLastRow() - 8;
      var colNums = 4;
      var retouchRange = s.getRange(row, col, rowNums, colNums);
      basic1.copyTo(retouchRange);
      break;
    }
  }

  hideEmptyCols();

  SpreadsheetApp.flush();
}

function retouch_fill_all() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();
  var sheet = ss.getSheets();
  for (j = s.getIndex() - 1; j < sheet.length; j++) {
    ss.setActiveSheet(sheet[j]);
    fillRetouch1();
    SpreadsheetApp.flush();
  }
}

function retouch_remove_1() {
  // Fill retouch    
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();

  for (i = 4; i < s.getMaxColumns(); i++) {
    var basic = s.getRange(6, i);
    if (basic.getValue() == 'R.Bsc') {
      var basic1 = basic.offset(3, 0);
      var row = 9;
      var col = i;
      var rowNums = s.getLastRow() - 8;
      var colNums = 4;
      var retouchRange = s.getRange(row, col, rowNums, colNums);
      retouchRange.clearContent();
      break;
    }
  }
}

function retouch_remove_all() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();
  var sheet = ss.getSheets();
  for (j = s.getIndex() - 1; j < sheet.length; j++) {
    ss.setActiveSheet(sheet[j]);
    retouch_remove_1();
    SpreadsheetApp.flush();
  }
}



function note2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var url = ss.getUrl();
  var s = ss.getActiveSheet();
  var cell = s.getActiveCell();
  var note2 = ss.getSheetByName('note2');
  var sName = s.getName();
  var gid = s.getSheetId().toString();

  var sUrl = 'https://docs.google.com/spreadsheets/d/1oZ0E2Y0Uh19ZdPlbmdbBeo0zjPvH5CKaGA_hNHHoUDA/edit#gid=' + gid;
  var input = 'Giá ' + cell.getValue();

  if (input == '') { exit(); }
  note2.appendRow(['', sName, sUrl, input, 'Phong']);
  ss.setActiveSheet(note2);

}

function moveNext() {
  var index = s.getIndex() + 1;
  ss.moveActiveSheet(index);
}

function clearINV() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();
  var sheet = ss.getSheets();

  for (i = s.getIndex(); i < sheet.length; i++) {
    if (sheet[i].isSheetHidden() != true) {
      ss.setActiveSheet(sheet[i])
      var row = 9;
      var col = 1;
      var numRow = sheet[i].getLastRow() - 8;
      var numCol = sheet[i].getLastColumn() - 1;
      sheet[i].getRange(row, col, numRow, numCol).clearContent();
    }
  }
}


function copyFormat() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();

  if (s.getMaxRows() > 9) {
    var srcRange = s.getRange(9, 1, 1, s.getMaxColumns());
    var desRange = s.getRange(10, 1, s.getMaxRows() - 9, s.getMaxColumns());
    srcRange.copyTo(desRange, { formatOnly: true });
  }
  s.getRange(1, 1, 4, s.getMaxColumns()).setBorder(true, true, true, true, false, false, 'silver', null)
}

function removeEmptyRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();


  if (s.getMaxRows() > 9) {
    if (s.getRange(s.getMaxRows(), 3).getValue() == '') {
      var rowPosition = s.getLastRow() + 1;
      var howMany = s.getMaxRows() - s.getLastRow();

      //Logger.log(rowPosition+' | '+howMany);
      s.deleteRows(rowPosition, howMany);
    }
  }
}

function removeMoney() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();

  if (s.getMaxRows() > 9) {
    var rangeToRemove = s.getRange(10, s.getMaxColumns(), s.getMaxRows() - 9, 1);
    rangeToRemove.clearContent();
  }
}

function unHideColumns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();

  //unhide Column
  var row = 1;
  var col = 1;
  var rowNums = 1;
  var colNums = s.getMaxColumns();
  s.unhideColumn(s.getRange(row, col, rowNums, colNums));
}

function copyMoney() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();

  Logger.log(s.getMaxRows());
  if (s.getMaxRows() > 9) {
    var firstMoney = s.getRange(9, s.getMaxColumns(), 1, 1);
    var nextMoneys = s.getRange(10, s.getMaxColumns(), s.getMaxRows(), 1);
    firstMoney.copyTo(nextMoneys);
  }
}

function fillRetouch1a() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();

  // Fill retouch  
  var fomular = '=iferror(if(sumifs(JW!AH:AH,JW!$A:$A,$A9,JW!$D:$D,$B9)=0,,sumifs(JW!AH:AH,JW!$A:$A,$A9,JW!$D:$D,$B9)),)';

  for (i = 5; i < s.getMaxColumns(); i++) {
    var basic = s.getRange(6, i);

    if (basic.getValue() == 'R.Bsc') {
      var basic1 = basic.offset(3, 0);
      basic1.setValue(fomular);

      var row = 9;
      var col = i;
      var rowNums = s.getLastRow() - 8;
      var colNums = 4;

      var retouchRange = s.getRange(row, col, rowNums, colNums);
      basic1.copyTo(retouchRange);
      break;
    }

  }
}

function hideEmptySheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();
  for (i = s.getIndex(); i < ss.getSheets().length; i++) {
    ss.setActiveSheet(ss.getSheets()[i]);

    if (ss.getSheets()[i].getRange('A9').getValue() == '') {
      ss.getSheets()[i].hideSheet();
    }
    SpreadsheetApp.flush();
  }
}

function fillRetouchAllValue() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var fomular = '=iferror(if(sumifs(0JW!AH:AH,0JW!$A:$A,$A9,0JW!$D:$D,$B9)=0,,sumifs(0JW!AH:AH,0JW!$A:$A,$A9,0JW!$D:$D,$B9)),)';
  for (j = ss.getActiveSheet().getIndex() - 1; j < sheets.length; j++) {
    var s = sheets[j]
    if (s.isSheetHidden() != true) {
      ss.setActiveSheet(s);
      for (i = 5; i < s.getMaxColumns(); i++) {
        s.getRange(1, 1, 4, s.getMaxColumns()).setBorder(true, true, true, true, false, false, 'silver', null)
        var basic = s.getRange(6, i);
        if (basic.getValue() == 'R.Bsc') {
          var basic1 = basic.offset(3, 0);
          basic1.setValue(fomular);

          var row = 9;
          var col = i;
          var rowNums = s.getLastRow() - 8;
          var colNums = 4;
          var retouchRange = s.getRange(row, col, rowNums, colNums);
          retouchRange.setValues(retouchRange.getValues());
          break;
        }
      }
      SpreadsheetApp.flush();
    }
    SpreadsheetApp.flush();
  }
  SpreadsheetApp.flush();
}

function retouchConvertToValue() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();
  var row = 9;
  var col = 4;
  var numRows = s.getLastRow() - 8;
  var numCols = s.getLastColumn() - 4;
  var convertRange = s.getRange(row, col, numRows, numCols);
  convertRange.setValues(convertRange.getValues());
  s.getRange(1, 4, 1, s.getLastColumn() - 3).clearContent();
}

function retouchConvertToValueBatch() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();
  var sheet = ss.getSheets();

  for (i = s.getIndex() - 1; i < sheet.length; i++) {
    if (sheet[i].isSheetHidden() != true) {
      ss.setActiveSheet(sheet[i]);
      retouchConvertToValue();
    }
    SpreadsheetApp.flush();
  }
}

function hideEmptyCols() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var s = ss.getActiveSheet();
  for (var i = 4; i < s.getLastColumn(); i++) {
    if (s.getRange(7, i).getValue() == '-') {
      s.hideColumns(i);
    }
  }
}

function hideEmptyColsBatch2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var s = ss.getActiveSheet();
  var sheet = ss.getSheets();
  for (var i = s.getIndex() - 1; i < sheet.length; i++) {
    if (sheet[i].isSheetHidden() != true) {
      ss.setActiveSheet(sheet[i]);
      hideEmptyCols();
    }
    SpreadsheetApp.flush();
  }
}

function unHideEmptyCols() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var s = ss.getActiveSheet();
  for (var i = 4; i < s.getLastColumn(); i++) {
    if (s.getRange(7, i).getValue() == '-') {
      s.showColumns(i);
    }
  }
}

function unHideEmptyCol_all() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var s = ss.getActiveSheet();
  var sheet = ss.getSheets();
  for (var i = s.getIndex() - 1; i < sheet.length; i++) {
    if (sheet[i].isSheetHidden() != true) {
      ss.setActiveSheet(sheet[i]);
      unHideEmptyCols();
    }
    SpreadsheetApp.flush();
  }
}

function money() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();

  var a = s.getRange(1, s.getLastColumn() - 1);
  a.setValue('=left(address(ROW(),COLUMN(),4),1)');
  var b = a.getValue();
  var a1 = s.getRange(1, s.getLastColumn());
  a1.setValue('=left(address(ROW(),COLUMN(),4),1)');
  var b1 = a1.getValue();
  var cell = s.getRange(9, s.getLastColumn());
  cell.setValue('=if(sumproduct($D$8:$' + b + '$8,D9:' + b + '9)=0,"-",sumproduct($D$8:$' + b + '$8,D9:' + b + '9))');
  var range = s.getRange(9, s.getLastColumn(), s.getLastRow() - 8, 1);
  cell.copyTo(range);
  s.getRange(7, s.getLastColumn()).setValue('=sum(' + b1 + '9:' + b1 + ')');

  var sum1 = s.getRange('D7');
  sum1.setValue('=if(sum(D9:D)=0,"-",sum(D9:D))');
  //var sumRange = s.getRange('D7:'+a.getValue()+'7');
  var sumRange = s.getRange('D7:' + b + '7');
  //sum1.copyTo(sumRange);
  sumRange.setFormula(sum1.getFormula());

  a.clearContent();
  a1.clearContent();


}

function moneyBatch() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();
  var sheet = ss.getSheets();

  for (i = s.getIndex(); i < sheet.length; i++) {
    if (sheet[i].isSheetHidden() != true) {
      ss.setActiveSheet(sheet[i]);
      money();
    }
    SpreadsheetApp.flush();
  }
}

function copyJW2() {
  copyFomula();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();
  var toJW = ss.getSheetByName('JW');
  var a1 = s.getRange('A1');
  var config = ss.getSheetByName('Config');
  var url_combine = config.getRange('A1').getValue();
  var url_inv = config.getRange('A2').getValue();

  var src_ss = SpreadsheetApp.openByUrl(url_combine);
  var src_s = src_ss.getSheetByName('JW2');
  var src_r = src_s.getRange(2, 1, src_s.getLastRow(), src_s.getLastColumn());

  var des_ss = SpreadsheetApp.openByUrl(url_inv);
  var des_s = des_ss.getSheetByName('JW');
  var des_r = des_s.getRange(2, 1, src_s.getLastRow(), src_s.getLastColumn());

  des_r.setValues(src_r.getValues());


  SpreadsheetApp.flush();

}

function copyFomula() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName('JW');

  var AD1 = s.getRange('AD1');
  AD1.setFormula(AD1.getNote());

  var r = s.getRange('AD1:AK1');
  r.setFormula(AD1.getFormula());
}

function refreshData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName('JW');
  s.getRange('A2').setFormula(s.getRange('A2').getNote());
  var range = s.getRange(2, 1, s.getLastRow(), s.getLastColumn());
  range.setValues(range.getValues());
  var AD1 = s.getRange('AD1');
  AD1.setFormula(AD1.getNote());

  var r = s.getRange('AD1:AK1');
  r.setFormula(AD1.getFormula());
  SpreadsheetApp.flush();

}

function goToSheetName() {

  var ui = SpreadsheetApp.getUi();

  var input = ui.prompt('Điền tên Sheet', ui.ButtonSet.OK);
  var name = input.getResponseText();
  Logger.log(name);
  if (input.getSelectedButton() == ui.Button.OK) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.setActiveSheet(ss.getSheetByName(name));
  }

}

function pdfLink() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName('PDF');
  var range = s.getActiveRange();
  var data = range.getValues();
  var cell = s.getActiveCell();


  for (i = 0; i < data.length; i++) {
    var link = s.getRange(i + cell.getRow(), 5);
    var id = s.getRange(i + cell.getRow(), 6);

    var name = data[i][0];
    var url = DriveApp.getFilesByName(name).next().getUrl();
    var key = DriveApp.getFilesByName(name).next().getId();

    link.setValue(url);
    id.setValue(key);

    SpreadsheetApp.flush();
  }
}

function pdfLink2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName('PDF');
  var cell = s.getActiveCell();
  var row = cell.getRow();
  var range = s.getActiveRange();
  var numRows = range.getNumRows() + row;
  var data = range.getValues();

  for (i = row; i < numRows; i++) {
    var link = s.getRange(i, 5);
    var id = s.getRange(i, 6);

    var name = data[i][0];
    var url = DriveApp.getFilesByName(name).next().getUrl();
    var key = DriveApp.getFilesByName(name).next().getId();

    link.setValue(url);
    id.setValue(key);

    SpreadsheetApp.flush();
  }
}

function test() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName('PDF');
  var file = DriveApp.getFilesByName('PSVN INV 2003 - 001-prDi.pdf');
  var url = file.next().getUrl();
  s.getRange('E2').setValue(url);
}

function color() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();
  var sheet = ss.getSheets();
  s.setTabColor(null);
  for (i = s.getIndex() - 1; i < sheet.length; i++) {
    sheet[i].setTabColor(null);
  }
}

function myFunction() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();

  var startColumn = 4;
  var numColumns = s.getMaxColumns() - startColumn + 1;
  s.autoResizeColumns(startColumn, numColumns);

}

function copyNewJW() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();
  var jw = ss.getSheetByName('JW');

  // Điền JW
  var A2 = jw.getRange('A2');
  //  A2.setFormula(A2.getNote());

  //  // Đợi 10s
  //  var milliseconds = 1000*10
  //  Utilities.sleep(milliseconds);

  // Xác định Range JW 
  var row, column, numRows, numColumns
  row = 2;
  column = 1;
  numRows = jw.getLastRow();
  numColumns = jw.getLastColumn();
  var range = jw.getRange(row, column, numRows, numColumns);

  // Convert Formula to Contents
  if (A2.isBlank() == false) { range.copyTo(range, { contentsOnly: true }); }


  // Điền công thức Retouch
  var AF1 = jw.getRange('AF1')
  AF1.setFormula(AF1.getNote());
  var AF1_AM1 = jw.getRange('AF1:AM1');
  AF1_AM1.setFormula(AF1.getFormula());


  SpreadsheetApp.flush();
}

function removeSlash() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();
  var cell = s.getActiveCell();
  cell.setValue('=iferror(right(C9,len(C9)-find(" / ",C9)-2),C9)');
}

function remove_header() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();
  var cell = s.getActiveCell();
  var row = cell.getRow();
  var col = cell.getColumn();
  var r = s.getActiveRange();
  var d = r.getValues();
  var range = s.getRange(1, 4, 1, s.getLastColumn() - 4);
  range.clearContent();
}

function batch() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();
  var cell = s.getActiveCell();
  var row = cell.getRow();
  var col = cell.getColumn();
  var r = s.getActiveRange();
  var d = r.getValues();
  var sheets = ss.getSheets();

  for (i = s.getIndex() - 1; i < sheets.length; i++) {
    ss.setActiveSheet(sheets[i]);
    remove_header();
  }
}

function grab() {
  var ui = SpreadsheetApp.getUi();
  var arr = '';
  var title = ['Timestamp', 'ID', 'SubID', 'No', 'Code', 'Qty', 'Style', '0Make', '1Edit', '3QA', '5DC'];
  var arr = d2[0][head.indexOf('JobName')] + '<br>';

  for (i = 0; i < title.length; i++) {
    if (title[i] == 'Timestamp') {
      arr += Utilities.formatDate(d2[0][head.indexOf(title[i])], "GMT+7", "MMM-dd") + ' | ';
    } else {
      arr += d2[0][head.indexOf(title[i])] + ' | ';
    }

  }
  Logger.log(arr);
  var htmlOutput = HtmlService.createHtmlOutput(arr);
  ui.showModalDialog(htmlOutput, 'Job Info');
}

var ss = SpreadsheetApp.getActiveSpreadsheet();
var s = ss.getActiveSheet();
var cell = s.getActiveCell();
var row = cell.getRow();
var col = cell.getColumn();
var r = s.getActiveRange();
var d = r.getValues();
var r2 = s.getRange(1, 1, r.getNumRows(), s.getLastColumn());
var d2 = r2.getValues();

