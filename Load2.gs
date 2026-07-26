function parseDate(val) {
  if (!val) return new Date(NaN);
  if (val instanceof Date) return val;
  if (typeof val === "number") {
    // Google Sheets lưu ngày dưới dạng số (serial number)
    // Ngày 30/12/1899 là ngày gốc trong Sheets -> chuyển sang JS Date
    return new Date(Math.round((val - 25569) * 86400 * 1000));
  }
  if (typeof val === "string") {
    const trimmed = val.trim();
    if (!trimmed) return new Date(NaN);
    // Hỗ trợ các định dạng: dd/mm/yyyy, d/m/yyyy, dd-mm-yyyy, v.v.
    const parts = trimmed.split(/\D+/).filter(p => p !== "");
    if (parts.length >= 3) {
      const d = parseInt(parts[0], 10);
      const m = parseInt(parts[1], 10);
      let y = parseInt(parts[2], 10);
      if (y < 100) y += 2000;
      return new Date(y, m - 1, d);
    }
  }
  return new Date(NaN);
}

function load() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const selectedRange = ss.getSelection().getActiveRange();
  if (!selectedRange) {
    Logger.log("❌ Vui lòng chọn vùng chứa tên các sheet (giá trị CID).");
    return;
  }
  const sheetNames = selectedRange.getValues().flat().filter(String);
  if (sheetNames.length === 0) {
    Logger.log("❌ Không có tên sheet nào trong vùng chọn.");
    return;
  }
  const statsSheet = ss.getSheetByName("-Stats");
  if (!statsSheet) {
    Logger.log('❌ Sheet "Stats" không tồn tại!');
    return;
  }
  const statsDataRange = statsSheet.getRange("A1").getDataRegion();
  const allValues = statsDataRange.getValues();
  if (allValues.length === 0) {
    Logger.log('❌ Sheet "Stats" không có dữ liệu!');
    return;
  }
  const headers = allValues[0];
  const numCols = headers.length;
  const findColIndex = (name) => {
    const lowerHeaders = headers.map(h => String(h).trim().toLowerCase());
    return lowerHeaders.indexOf(name.toLowerCase());
  };
  const cidCol = findColIndex("CID");
  const idCol = findColIndex("ID");
  const projectCol = findColIndex("PROJECT");
  const dateCol = findColIndex("DATE");
  const amountCol = findColIndex("AMOUNT");
  if ([cidCol, idCol, projectCol, dateCol].some(idx => idx === -1)) {
    Logger.log("❌ Thiếu một trong các cột: CID, ID, PROJECT, DATE.");
    return;
  }
  const colorEven = "#ffffff";
  const colorOdd = "#f5f5f5";

  sheetNames.forEach(sheetName => {
    const targetSheet = ss.getSheetByName(sheetName);
    if (!targetSheet) {
      Logger.log(`⚠️ Sheet "${sheetName}" không tồn tại.`);
      return;
    }
    let dataRows = allValues.slice(1).filter(row => {
      const cellValue = row[cidCol];
      const cellStr = String(cellValue).trim();
      const sheetStr = String(sheetName).trim();
      if (cellStr === sheetStr) return true;
      const num1 = Number(cellStr);
      const num2 = Number(sheetStr);
      if (!isNaN(num1) && !isNaN(num2) && num1 === num2) return true;
      return false;
    });
    if (dataRows.length === 0) {
      Logger.log(`ℹ️ Không có dữ liệu cho CID = "${sheetName}".`);
      return;
    }

    // ✅ SẮP XẾP THEO NGÀY THÁNG ĐÚNG
    dataRows.sort((a, b) => {
      const cidA = String(a[cidCol] || '').trim();
      const cidB = String(b[cidCol] || '').trim();
      if (cidA !== cidB) return cidA.localeCompare(cidB);

      const idA = String(a[idCol] || '').trim();
      const idB = String(b[idCol] || '').trim();
      if (idA !== idB) return idA.localeCompare(idB);

      const projA = String(a[projectCol] || '').trim();
      const projB = String(b[projectCol] || '').trim();
      if (projA !== projB) return projA.localeCompare(projB);

      // ✅ So sánh DATE theo thời gian thực
      const dateA = parseDate(a[dateCol]);
      const dateB = parseDate(b[dateCol]);
      if (isNaN(dateA) && isNaN(dateB)) return 0;
      if (isNaN(dateA)) return 1;
      if (isNaN(dateB)) return -1;
      return dateA - dateB;
    });

    const compactRows = [];
    let prevKey = null;
    for (let row of dataRows) {
      const currentKey = [
        String(row[cidCol] || "").trim(),
        String(row[idCol] || "").trim(),
        String(row[projectCol] || "").trim(),
        String(row[dateCol] || "").trim()
      ].join("||");
      const newRow = [...row];
      if (currentKey === prevKey) {
        newRow[cidCol] = "";
        newRow[idCol] = "";
        newRow[projectCol] = "";
        newRow[dateCol] = "";
      }
      compactRows.push(newRow);
      prevKey = currentKey;
    }

    // ✅ Tính tổng AMOUNT theo ID và chèn cột "TOTAL" sau PROJECT
    let newCompactRows = compactRows;
    let newNumCols = numCols;
    let newHeaders = [...headers];
    if (amountCol !== -1) {
      const idTotals = {};
      dataRows.forEach(row => {
        const idVal = String(row[idCol] || "").trim();
        if (idVal) {
          const amount = parseFloat(row[amountCol]) || 0;
          idTotals[idVal] = (idTotals[idVal] || 0) + amount;
        }
      });
      const projectColIndex = projectCol;
      const newHeadersTemp = [];
      for (let i = 0; i < headers.length; i++) {
        if (i === projectColIndex) {
          newHeadersTemp.push(headers[i]);
          newHeadersTemp.push("TOTAL");
        } else {
          newHeadersTemp.push(headers[i]);
        }
      }
      newHeaders = newHeadersTemp;
      newCompactRows = compactRows.map(row => {
        const idVal = String(row[idCol] || "").trim();
        const total = idVal ? idTotals[idVal] : "";
        const newRow = [];
        for (let i = 0; i < row.length; i++) {
          if (i === projectColIndex) {
            newRow.push(row[i]);
            newRow.push(total);
          } else {
            newRow.push(row[i]);
          }
        }
        return newRow;
      });
      newNumCols = newHeaders.length;
    }

    // Ghi dữ liệu
    const startRow = 22;
    targetSheet.getRange(startRow + 1, 1, 1000, newNumCols).clearContent();
    targetSheet.getRange(startRow, 1, 1, newNumCols).setValues([newHeaders]);
    if (newCompactRows.length > 0) {
      const dataRange = targetSheet.getRange(startRow + 1, 1, newCompactRows.length, newNumCols);
      dataRange.setValues(newCompactRows);
      const backgrounds = [];
      for (let i = 0; i < newCompactRows.length; i++) {
        const bgColor = (i % 2 === 0) ? colorOdd : colorEven;
        backgrounds.push(new Array(newNumCols).fill(bgColor));
      }
      dataRange.setBackgrounds(backgrounds);
    }

    // Xóa hàng trống thừa
    const lastDataRow = startRow + newCompactRows.length;
    const maxSheetRows = targetSheet.getMaxRows();
    if (lastDataRow < maxSheetRows) {
      targetSheet.deleteRows(lastDataRow + 1, maxSheetRows - lastDataRow);
    }

    // Kẻ viền cam từ cột 2
    if (newCompactRows.length > 0) {
      const borderRange = targetSheet.getRange(startRow + 1, 2, newCompactRows.length, newNumCols - 1);
      borderRange.setBorder(
        true, true, true, true, false, false,
        "#ff9900",
        SpreadsheetApp.BorderStyle.SOLID
      );
    }

    Logger.log(`✅ Đã xử lý sheet "${sheetName}" với ${newCompactRows.length} dòng.`);
  });
  Logger.log(`✅ Hoàn thành! Đã xử lý ${sheetNames.length} sheet.`);
}