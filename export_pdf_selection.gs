/**
 * Hàm xuất các Sheet được chọn ra file PDF trên Google Drive
 */
function export_pdf_selection() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeRange = ss.getActiveRange();
  
  if (!activeRange) {
    SpreadsheetApp.getUi().alert("Vui lòng chọn dải ô chứa danh sách tên Sheet cần xuất PDF!");
    return;
  }
  
  // 1. Lấy danh sách tên sheet từ vùng chọn (loại bỏ ô trống)
  var values = activeRange.getValues();
  var sheetNames = [];
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      var val = String(values[i][j]).trim();
      if (val !== "") {
        sheetNames.push(val);
      }
    }
  }
  
  if (sheetNames.length === 0) {
    SpreadsheetApp.getUi().alert("Không tìm thấy tên Sheet nào trong dải ô đã chọn!");
    return;
  }
  
  // 2. Định danh Thư mục chính và Tính toán Thư mục Tháng (Lùi 1 tháng)
  var mainFolderId = "1B0Z7G2wzRYE8nXsHgs8TaZvOIleEbzSP"; // ID thư mục chính
  var mainFolder = DriveApp.getFolderById(mainFolderId);
  
  // Lấy ngày hiện tại và lùi lại 1 tháng
  var targetDate = new Date();
  targetDate.setMonth(targetDate.getMonth() - 1); // Lùi 1 tháng
  
  // Định dạng YYMM (ví dụ: Tháng 7/2026 lùi 1 tháng thành Tháng 6/2026 -> 2606)
  var folderNameYYMM = Utilities.formatDate(targetDate, ss.getSpreadsheetTimeZone(), "yyMM");
  
  // Kiểm tra thư mục tháng (YYMM), nếu chưa có thì tạo mới
  var subFolders = mainFolder.getFoldersByName(folderNameYYMM);
  var targetFolder;
  if (subFolders.hasNext()) {
    targetFolder = subFolders.next();
  } else {
    targetFolder = mainFolder.createFolder(folderNameYYMM);
  }
  
  var ssId = ss.getId();
  var successCount = 0;
  var notFoundSheets = [];

  // 3. Vòng lặp duyệt qua từng tên sheet được chọn
  sheetNames.forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      notFoundSheets.push(sheetName);
      return;
    }
    
    var sheetId = sheet.getSheetId();
    var lastRow = sheet.getLastRow();
    
    if (lastRow < 1) {
      lastRow = 1; // Đảm bảo dòng tối thiểu là 1 nếu sheet trống
    }
    
    // 4. Tìm tên file PDF từ dải ô N:O của sheet hiện tại
    var pdfFileName = getPdfFileNameFromSheet(sheet);
    if (!pdfFileName) {
      pdfFileName = "Invoice_" + sheetName; // Tên dự phòng nếu không tìm thấy key inv_pdf_name
    }
    if (!pdfFileName.toLowerCase().endsWith(".pdf")) {
      pdfFileName += ".pdf";
    }
    
    // 5. Cấu hình Url Export cho dải B:L (Cột B = index 1, Cột L = index 11)
    var url = ss.getUrl().replace(/edit$/, '') + 'export?';
    var exportOptions = {
      exportFormat: 'pdf',
      format: 'pdf',
      size: 'A4',             // Kích thước trang A4
      portrait: 'true',       // Khung dọc
      fitw: 'true',           // Tự động vừa chiều rộng
      gridlines: 'false',     // Không hiển thị đường lưới
      printtitle: 'false',
      sheetnames: 'false',
      fch: 'false',
      fzr: 'false',
      gid: sheetId,           // ID của sheet cần xuất
      c1: 1,                  // Cột B (Index 1)
      r1: 0,                  // Dòng 1 (Index 0)
      c2: 12,                 // Hết cột L (Index 12 exclusive)
      r2: lastRow             // Dòng cuối cùng có dữ liệu (getLastRow)
    };
    
    var urlParts = [];
    for (var key in exportOptions) {
      urlParts.push(key + '=' + exportOptions[key]);
    }
    var exportUrl = url + urlParts.join('&');
    
    // 6. Fetch dữ liệu PDF từ URL export
    var response = UrlFetchApp.fetch(exportUrl, {
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
      },
      muteHttpExceptions: true
    });
    
    // 7. Lưu file vào thư mục Drive tương ứng
    var blob = response.getBlob().setName(pdfFileName);
    targetFolder.createFile(blob);
    successCount++;
  });

  // 8. Thông báo kết quả
  var message = "Đã xuất thành công " + successCount + " file PDF vào thư mục: " + folderNameYYMM;
  if (notFoundSheets.length > 0) {
    message += "\n⚠️ Không tìm thấy các Sheet: " + notFoundSheets.join(", ");
  }
  SpreadsheetApp.getUi().alert(message);
}

/**
 * Hàm phụ trợ: Quét cột N và O để lấy giá trị chi tiết tương ứng với key 'inv_pdf_name'
 */
function getPdfFileNameFromSheet(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return null;
  
  // Đọc dải dữ liệu từ cột N đến cột O
  var rangeValues = sheet.getRange("N1:O" + lastRow).getValues();
  
  for (var i = 0; i < rangeValues.length; i++) {
    var key = String(rangeValues[i][0]).trim();
    if (key === "inv_pdf_name") {
      var detailValue = String(rangeValues[i][1]).trim();
      return detailValue !== "" ? detailValue : null;
    }
  }
  return null;
}