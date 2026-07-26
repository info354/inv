function editQuick() {

  if(r < 12 || s.getRange(row,40).isBlank()){exit();}
  //Xác định url remote
  var sheet = ss.getSheetByName('JW');
  var range = sheet.getRange('BC12:BE17');
  var data = range.getValues();

  var urls = data.filter(row => row[0] == true);
  var url = urls[0][2];

  // ghi dữ liệu
  var header = s.getRange('11:11').getValues()[0];
    var jw2 = SpreadsheetApp.openByUrl(url).getSheetByName('JW');
    

      var data = s.getRange(row,1,r.getNumRows(),40).getValues();
    for ( i = 0 ; i < d.length ; i++){
      var dataRow = data[i];
      var row2 = dataRow[header.indexOf('R')];
      var col2 = header.indexOf('Style')+1;
      var style = dataRow[header.indexOf('1860')];
      Logger.log(data);
      Logger.log(row2);
      Logger.log(col2);
      Logger.log(style);
      
      jw2.getRange(row2,col2).setValue(style);


      
    }
    


    
  }
  
