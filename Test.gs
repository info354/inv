function quick_copy_style2() {
  //Xác định url
  var sheet = ss.getSheetByName('JW');
  var range = sheet.getRange('BC11:BG17');
  var data = range.getValues();

  var value = data.filter(row => {return row[0] == true}
  );

  Logger.log(value[0][2]);
}