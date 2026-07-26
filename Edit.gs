

function edit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const activeRange = sheet.getActiveRange();

  const startRow = activeRange.getRow();
  const endRow = startRow + activeRange.getNumRows() - 1;
  const startCol = activeRange.getColumn();
  const endCol = startCol + activeRange.getNumColumns() - 1;

  if (startRow < 11 || startCol < 1 || endCol > 39) {
    SpreadsheetApp.getUi().alert("Vui lòng chọn vùng trong A11:AM...");
    return;
  }

  const headers = sheet.getRange(11, 1, 1, 39).getValues()[0];
  const dateColIndex = headers.indexOf("Date");
  const idColIndex = headers.indexOf("ID"); // ← ✅ THÊM DÒNG NÀY

  if (dateColIndex === -1) {
    SpreadsheetApp.getUi().alert("Không tìm thấy cột 'Date'.");
    return;
  }
  if (idColIndex === -1) {
    SpreadsheetApp.getUi().alert("Không tìm thấy cột 'ID'.");
    return;
  }

  // 🔥 Kiểm tra hàng 7: lấy giá trị nhanh theo từng cột
  const quickFillValues = sheet.getRange(7, startCol, 1, endCol - startCol + 1).getValues()[0];
  const hasQuickFill = quickFillValues.some(val => val !== "" && val != null);

  let useQuickFill = false;
  let singleValue = "";

  if (hasQuickFill) {
    useQuickFill = true;
  } else {
    const ui = SpreadsheetApp.getUi();
    const prompt = ui.prompt(
      "Nhập giá trị để áp dụng cho tất cả các ô đã chọn",
      "Giá trị:",
      ui.ButtonSet.OK_CANCEL
    );
    if (prompt.getSelectedButton() !== ui.Button.OK) return;
    singleValue = prompt.getResponseText();
  }

  // === Lấy dữ liệu URL một lần ===
  const urlData = sheet.getRange("BD11:BE50").getValues();
  let successCount = 0;
  let errorCount = 0;

  for (let r = startRow; r <= endRow; r++) {
    const targetRowNumber = sheet.getRange(r, 1).getValue();
    const currentId = sheet.getRange(r, idColIndex + 1).getValue();
    const dateValue = sheet.getRange(r, dateColIndex + 1).getValue();

    if (typeof targetRowNumber !== 'number' || targetRowNumber < 1) {
      errorCount++;
      continue;
    }
    if (!(dateValue instanceof Date)) {
      errorCount++;
      continue;
    }

    const yymm = Utilities.formatDate(dateValue, Session.getScriptTimeZone(), "yyMM");

    // Phân loại URL
    const todayUrls = [];
    const otherUrls = [];
    for (const [name, link] of urlData) {
      if (!link || typeof name !== 'string') continue;
      if (name.includes("JW_Day_")) {
        todayUrls.push([name, link]);
      } else if (name.includes(yymm)) {
        otherUrls.push([name, link]);
      }
    }

    // Hàm thử cập nhật trong danh sách URLs
    const tryUpdateInUrls = (urlList) => {
      for (const [_, remoteUrl] of urlList) {
        let remoteSS;
        try {
          remoteSS = SpreadsheetApp.openByUrl(remoteUrl);
        } catch (e) {
          continue;
        }

        const remoteSheet = remoteSS.getSheetByName("JW");
        if (!remoteSheet) continue;

        const remoteLastCol = Math.min(remoteSheet.getLastColumn(), 39);
        const remoteHeaders = remoteSheet.getRange(11, 1, 1, remoteLastCol).getValues()[0];
        const remoteIdColIndex = remoteHeaders.indexOf("ID");
        if (remoteIdColIndex === -1) continue;

        const remoteRowCount = Math.max(1, remoteSheet.getLastRow() - 10);
        const fetchColCount = Math.max(2, remoteIdColIndex + 1);
        const remoteData = remoteSheet.getRange(11, 1, remoteRowCount, fetchColCount).getValues();

        let targetRemoteRow = -1;
        for (let i = 0; i < remoteData.length; i++) {
          const rVal = remoteData[i][0];
          const idVal = remoteData[i][remoteIdColIndex];
          if (rVal === targetRowNumber && String(idVal) === String(currentId)) {
            targetRemoteRow = i + 11;
            break;
          }
        }

        if (targetRemoteRow === -1) continue;

        // Cập nhật các cột được chọn (trừ R và ID)
        for (let c = startCol; c <= endCol; c++) {
          const headerName = headers[c - 1];
          if (!headerName || headerName === "R" || headerName === "ID") continue;

          const targetColIndex = remoteHeaders.indexOf(headerName);
          if (targetColIndex === -1) continue;

          const valueToSet = useQuickFill ? quickFillValues[c - startCol] : singleValue;

          try {
            remoteSheet.getRange(targetRemoteRow, targetColIndex + 1).setValue(valueToSet);
            successCount++;
          } catch (e) {
            // im lặng hoặc ghi log nếu cần
          }
        }

        return true; // thành công → dừng
      }
      return false;
    };

    // Ưu tiên bảng hôm nay
    if (!tryUpdateInUrls(todayUrls) && !tryUpdateInUrls(otherUrls)) {
      errorCount++;
    }
  }

  // === Thông báo kết quả (tùy chọn) ===
  // const ui = SpreadsheetApp.getUi();
  // let msg = `✅ Đã cập nhật ${successCount} ô`;
  // if (errorCount > 0) msg += `\n⚠️ ${errorCount} lỗi/bỏ qua`;
  // ui.alert(msg);
}

function edit2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const activeRange = sheet.getActiveRange();

  const startRow = activeRange.getRow();
  const endRow = startRow + activeRange.getNumRows() - 1;
  const startCol = activeRange.getColumn();
  const endCol = startCol + activeRange.getNumColumns() - 1;

  if (startRow < 11 || startCol < 1 || endCol > 39) {
    SpreadsheetApp.getUi().alert("Vui lòng chọn vùng trong A11:AM...");
    return;
  }

  const headers = sheet.getRange(11, 1, 1, 39).getValues()[0];
  const dateColIndex = headers.indexOf("Date");
  if (dateColIndex === -1) {
    SpreadsheetApp.getUi().alert("Không tìm thấy cột 'Date'.");
    return;
  }

  // 🔥 Kiểm tra hàng 7: lấy giá trị nhanh theo từng cột
  const quickFillValues = sheet.getRange(7, startCol, 1, endCol - startCol + 1).getValues()[0];
  const hasQuickFill = quickFillValues.some(val => val !== "" && val != null);

  let useQuickFill = false;
  let singleValue = "";

  if (hasQuickFill) {
    // Dùng giá trị hàng 7 theo từng cột
    useQuickFill = true;
  } else {
    // Không có dữ liệu hàng → popup nhập 1 giá trị chung
    const ui = SpreadsheetApp.getUi();
    const prompt = ui.prompt(
      "Nhập giá trị để áp dụng cho tất cả các ô đã chọn",
      "Giá trị:",
      ui.ButtonSet.OK_CANCEL
    );
    if (prompt.getSelectedButton() !== ui.Button.OK) return;
    singleValue = prompt.getResponseText();
  }

  // === Bắt đầu xử lý ===
  let successCount = 0;
  let errorCount = 0;

  for (let r = startRow; r <= endRow; r++) {
    const targetRowNumber = sheet.getRange(r, 1).getValue();
    if (typeof targetRowNumber !== 'number' || targetRowNumber < 1) {
      errorCount++;
      continue;
    }

    const dateValue = sheet.getRange(r, dateColIndex + 1).getValue();
    if (!(dateValue instanceof Date)) {
      errorCount++;
      continue;
    }

    const yymm = Utilities.formatDate(dateValue, Session.getScriptTimeZone(), "yyMM");

    const urlData = sheet.getRange("BD11:BE50").getValues();
    let remoteUrl = null;
    for (const [name, link] of urlData) {
      if (typeof name === 'string' && name.includes(yymm)) {
        remoteUrl = link;
        break;
      }
    }

    if (!remoteUrl) {
      errorCount++;
      continue;
    }

    let remoteSS;
    try {
      remoteSS = SpreadsheetApp.openByUrl(remoteUrl);
    } catch (e) {
      errorCount++;
      continue;
    }

    const remoteSheet = remoteSS.getSheetByName("JW");
    if (!remoteSheet) {
      errorCount++;
      continue;
    }

    const remoteLastCol = Math.min(remoteSheet.getLastColumn(), 39);
    const remoteHeaders = remoteSheet.getRange(11, 1, 1, remoteLastCol).getValues()[0];

    for (let c = startCol; c <= endCol; c++) {
      const headerName = headers[c - 1];
      if (!headerName || headerName === "R") continue;

      const targetColIndex = remoteHeaders.indexOf(headerName);
      if (targetColIndex === -1) {
        errorCount++;
        continue;
      }

      // Tìm hàng remote theo R
      const remoteRData = remoteSheet.getRange(11, 1, Math.max(1, remoteSheet.getLastRow() - 10), 1).getValues();
      let targetRemoteRow = -1;
      for (let i = 0; i < remoteRData.length; i++) {
        if (remoteRData[i][0] === targetRowNumber) {
          targetRemoteRow = i + 11;
          break;
        }
      }
      if (targetRemoteRow === -1) {
        errorCount++;
        continue;
      }

      // 🔥 Xác định giá trị cần ghi
      let valueToSet;
      if (useQuickFill) {
        valueToSet = quickFillValues[c - startCol];
        // Nếu ô hàng 10 trống → có thể bỏ qua hoặc ghi ""
        // Ở đây: vẫn ghi giá trị (kể cả rỗng)
      } else {
        valueToSet = singleValue;
      }

      try {
        remoteSheet.getRange(targetRemoteRow, targetColIndex + 1).setValue(valueToSet);
        successCount++;
      } catch (e) {
        errorCount++;
      }
    }
  }

  // const ui = SpreadsheetApp.getUi();
  // let msg = `✅ Đã cập nhật ${successCount} ô`;
  // if (errorCount > 0) msg += `\n⚠️ ${errorCount} lỗi/bỏ qua`;
  // ui.alert(msg);
}