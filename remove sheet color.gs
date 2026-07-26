/**
 * Reset màu tab về mặc định cho các sheet có tên nằm trong vùng chọn (range).
 */
function resetSelectedSheetTabColors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const range = ss.getActiveRange();
  
  if (!range) {
    ss.toast("Vui lòng chọn các ô chứa tên sheet trước khi chạy script!", "Thông báo", 3);
    return;
  }
  
  const values = range.getValues();
  let successCount = 0;
  let notFoundList = [];

  // Duyệt qua từng ô trong vùng đã chọn
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const sheetName = String(values[r][c]).trim();
      
      if (sheetName !== "") {
        const targetSheet = ss.getSheetByName(sheetName);
        if (targetSheet) {
          // Đặt tabColor về null để quay lại màu mặc định
          targetSheet.setTabColor(null);
          successCount++;
        } else {
          notFoundList.push(sheetName);
        }
      }
    }
  }

  // Hiển thị thông báo kết quả
  let message = `Đã reset màu tab cho ${successCount} sheet.`;
  if (notFoundList.length > 0) {
    message += `\nKhông tìm thấy ${notFoundList.length} sheet: ${notFoundList.join(", ")}`;
  }
  
  ss.toast(message, "Hoàn tất", 5);
}