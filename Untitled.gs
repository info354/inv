function fillDataToLinks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheet = ss.getActiveSheet();
  
  // 1. Lấy TẤT CẢ những gì đang hiển thị trên màn hình tại vùng BC12:BE
  // Dùng getDisplayValues() để lấy đúng chữ bạn nhìn thấy
  const lastRow = currentSheet.getLastRow();
  if (lastRow < 12) return; 
  
  const configDisplayData = currentSheet.getRange("BC12:BE" + lastRow).getDisplayValues();
  const validUrls = [];
  
  console.log("Đang kiểm tra dữ liệu hiển thị...");

  for (let r = 0; r < configDisplayData.length; r++) {
    let checkboxDisplay = configDisplayData[r].toString().toUpperCase().trim(); // Cột BC
    let urlDisplay = configDisplayData[r].toString().trim();                   // Cột BE
    
    // Kiểm tra các dấu hiệu của tích chọn: "TRUE", "CHECKED", "1", hoặc ký tự đặc biệt của checkbox
    // getDisplayValues thường trả về "TRUE" hoặc "FALSE" với checkbox công thức
    if (checkboxDisplay === "TRUE" || checkboxDisplay === "CHECKED" || checkboxDisplay === "1") {
      if (urlDisplay.startsWith("http")) {
        validUrls.push(urlDisplay);
        console.log("Dòng " + (r + 12) + " hợp lệ: " + urlDisplay);
      }
    }
  }

  if (validUrls.length === 0) {
    SpreadsheetApp.getUi().alert("Vẫn không nhận diện được dấu tích.\nHãy thử gõ tay chữ TRUE (viết hoa) vào ô BC14 để test xem công thức của bạn đang trả về giá trị ẩn gì.");
    return;
  }

  // 2. Lấy vùng bôi đen (Active Range)
  const activeRange = ss.getActiveRange();
  const startRow = activeRange.getRow();
  const startCol = activeRange.getColumn();
  const numRows = activeRange.getNumRows();
  const numCols = activeRange.getNumColumns();

  // 3. Chạy vòng lặp điền dữ liệu
  for (let i = 0; i < numRows; i++) {
    for (let j = 0; j < numCols; j++) {
      let currentRow = startRow + i;
      let currentCol = startCol + j;

      // Lấy ID ở cột C (Cột số 3)
      let idToVerify = currentSheet.getRange(currentRow, 3).getDisplayValue().trim();
      // Lấy giá trị cần điền ở cột AN (Cột 40)
      let valueToFill = currentSheet.getRange(currentRow, 40).getValue();

      if (!idToVerify) continue;

      validUrls.forEach(url => {
        try {
          let targetSs = SpreadsheetApp.openByUrl(url);
          let targetSheet = targetSs.getSheets();
          
          // Quét cột C file đích
          let targetData = targetSheet.getRange("C1:C" + targetSheet.getLastRow()).getDisplayValues();
          let foundRow = -1;

          for (let k = 0; k < targetData.length; k++) {
            if (targetData[k].trim() === idToVerify) {
              foundRow = k + 1;
              break;
            }
          }

          if (foundRow !== -1) {
            targetSheet.getRange(foundRow, currentCol).setValue(valueToFill);
          }
        } catch (e) {
          console.error("Lỗi URL " + url + ": " + e.message);
        }
      });
    }
  }
  ss.toast("Đã xử lý xong!", "Thành công");
}