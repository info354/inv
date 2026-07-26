/**
 * THIẾT KẾ HTML CSS V3 - GOOGLE APPS SCRIPT
 * Chuẩn hóa đọc Metadata N:O & Xử lý theo mảng bôi đen
 */

// --- 1. LẤY DANH SÁCH TÊN SHEET TỪ VÙNG BÔI ĐEN (AN TOÀN 100%) ---
function getSelectedSheetNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const range = ss.getActiveRange();
  
  if (!range) {
    throw new Error("Vui lòng bôi đen các ô chứa tên Sheet (VD: 1639, 2068) trên Google Sheet!");
  }
  
  const values = range.getDisplayValues();
  const sheetNames = [];
  
  values.forEach(row => {
    row.forEach(val => {
      const cleanName = (val || "").toString().trim();
      if (cleanName !== "" && ss.getSheetByName(cleanName)) {
        sheetNames.push(cleanName);
      }
    });
  });
  
  if (sheetNames.length === 0) {
    throw new Error("Không tìm thấy tên Sheet hợp lệ nào trong vùng ô bạn đang bôi đen! Hãy kiểm tra lại tên Sheet.");
  }
  
  return sheetNames;
}

// --- 2. ĐỌC METADATA DẢI N:O CỦA MỘT SHEET ---
function getSheetMetadata(sheet) {
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  }
  if (!sheet) return {};

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  
  const rangeValues = sheet.getRange(`N2:O${lastRow}`).getDisplayValues();
  const meta = {};
  
  rangeValues.forEach(([key, val]) => {
    if (key && key.trim() !== '') {
      meta[key.trim().toLowerCase()] = (val || "").toString().trim();
    }
  });
  
  return meta;
}

// --- 3. DỰNG DỮ LIỆU EMAIL TỪ SHEET CHỈ ĐỊNH ---
function buildEmailDataForSheet(sheet) {
  if (!sheet) {
    throw new Error("Sheet không tồn tại hoặc bị undefined!");
  }

  const meta = getSheetMetadata(sheet);

  const recipientEmail = meta['client_email'] || meta['email_cc'];
  const ccEmail = meta['email_cc'] || "";
  const recipientName = meta['client_name'] || "Valued Customer";
  const invPdfName = meta['inv_pdf_name'] || "invoice";
  const invId = meta['inv_id'] || "";

  if (!recipientEmail || !/\S+@\S+\.\S+/.test(recipientEmail)) {
    throw new Error(`Sheet "${sheet.getName()}": Thiếu hoặc sai định dạng client_email ở dải N:O.`);
  }

  const timeZone = Session.getScriptTimeZone();
  const timeString = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd HH:mm");
  const timeShort = Utilities.formatDate(new Date(), timeZone, "HH:mm");
  const cleanInvoiceRef = invId.replace(/:/g, "").trim();
  const subject = `PSVN ${cleanInvoiceRef} [${timeShort}]`;

  const htmlTemplate = HtmlService.createTemplateFromFile('emailTemplate');
  htmlTemplate.meta = meta;
  htmlTemplate.timeString = timeString;
  const htmlContent = htmlTemplate.evaluate().getContent();

  return { meta, recipientEmail, recipientName, ccEmail, invPdfName, cleanInvoiceRef, subject, htmlContent };
}

// --- 4. XỬ LÝ GỬI EMAIL CHO 1 SHEET CHỈ ĐỊNH ---
function processSingleSheetSend(ss, sheet) {
  const emailData = buildEmailDataForSheet(sheet);

  const startColIndex = 2; // B
  const endColIndex = 12;  // L
  const startRow = 1;

  function rowHasContent(sh, rowIndex, startCol, endCol) {
    const values = sh.getRange(rowIndex, startCol, 1, endCol - startCol + 1).getValues()[0];
    return values.some(val => val !== null && val !== "" && val.toString().trim() !== "");
  }

  let firstDataRow = startRow;
  if (!rowHasContent(sheet, startRow, startColIndex, endColIndex)) {
    for (let r = startRow + 1; r <= sheet.getMaxRows(); r++) {
      if (rowHasContent(sheet, r, startColIndex, endColIndex)) {
        firstDataRow = r;
        break;
      }
    }
  }

  let lastDataRow = firstDataRow;
  for (let r = sheet.getMaxRows(); r >= firstDataRow; r--) {
    if (rowHasContent(sheet, r, startColIndex, endColIndex)) {
      lastDataRow = r;
      break;
    }
  }

  if (lastDataRow < firstDataRow) {
    throw new Error(`Không tìm thấy dữ liệu hóa đơn (cột B:L) trên Sheet: ${sheet.getName()}`);
  }

  const colStart = String.fromCharCode(64 + startColIndex);
  const colEnd   = String.fromCharCode(64 + endColIndex);
  const rangeStr = `${colStart}${firstDataRow}:${colEnd}${lastDataRow}`;

  const noteColIndex = endColIndex;
  const originalWidth = sheet.getColumnWidth(noteColIndex);
  const noteVals = sheet.getRange(firstDataRow, noteColIndex, lastDataRow - firstDataRow + 1, 1).getValues();
  let maxLength = 0;
  for (let i = 0; i < noteVals.length; i++) {
    maxLength = Math.max(maxLength, (noteVals[i][0] || "").toString().length);
  }
  sheet.setColumnWidth(noteColIndex, Math.min(350, Math.max(150, maxLength * 6.5)));

  sheet.hideColumns(14, 2); // Ẩn N:O khi xuất PDF

  const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?`;
  const params = {
    exportFormat: 'pdf', format: 'pdf', size: 'letter', portrait: true,
    fitw: true, gridlines: false, printtitle: false, sheetnames: false,
    pagenum: 'UNDEFINED', attachment: true, gid: sheet.getSheetId(), range: rangeStr
  };

  const qs = Object.entries(params).map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v)}`).join('&');
  const invoicePdfBlob = UrlFetchApp.fetch(url + qs, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
  }).getBlob().setName(`${emailData.invPdfName}.pdf`);

  sheet.showColumns(14, 2);
  sheet.setColumnWidth(noteColIndex, originalWidth);

  const attachments = [invoicePdfBlob];
  const payGateLink = emailData.meta['pay_gate_link'];

  if (payGateLink && payGateLink.startsWith("http")) {
    try {
      let downloadUrl = payGateLink;
      const driveMatch = payGateLink.match(/\/d\/([a-zA-Z0-9_-]+)/);
      if (driveMatch && driveMatch[1]) {
        downloadUrl = `https://drive.google.com/uc?export=download&id=${driveMatch[1]}`;
      }
      const payGatePdfBlob = UrlFetchApp.fetch(downloadUrl).getBlob().setName("Payment_Guide.pdf");
      attachments.push(payGatePdfBlob);
    } catch (err) {
      Logger.log("⚠️ Không thể tải file PDF từ pay_gate_link: " + err.message);
    }
  }

  GmailApp.sendEmail(emailData.recipientEmail, emailData.subject, '', {
    htmlBody: emailData.htmlContent,
    cc: emailData.ccEmail,
    attachments: attachments,
    name: "PSVN Team"
  });

  Logger.log(`📧 Gửi thành công cho Sheet ${sheet.getName()} tới ${emailData.recipientEmail}`);
}

// --- 5. HÀM MỞ PREVIEW TỪ VÙNG BÔI ĐEN ---
function openPreviewFromSelection() {
  const selectedSheetNames = getSelectedSheetNames();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const firstSheetName = selectedSheetNames[0];
  const firstSheet = ss.getSheetByName(firstSheetName);

  if (!firstSheet) {
    throw new Error(`Không tìm thấy Sheet tên: "${firstSheetName}"`);
  }

  const emailData = buildEmailDataForSheet(firstSheet);
  const base64Html = Utilities.base64Encode(emailData.htmlContent, Utilities.Charset.UTF_8);

  const dialogTemplate = HtmlService.createTemplateFromFile('previewDialog');
  dialogTemplate.recipientEmail = emailData.recipientEmail;
  dialogTemplate.ccEmail = emailData.ccEmail;
  dialogTemplate.subject = emailData.subject;
  dialogTemplate.base64Html = base64Html;
  dialogTemplate.totalSheets = selectedSheetNames.length;
  dialogTemplate.sheetNamesJson = JSON.stringify(selectedSheetNames);

  const htmlOutput = dialogTemplate.evaluate()
    .setWidth(680)
    .setHeight(580);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, `📧 Preview (Sheet: ${firstSheetName}) - PSVN Control Center`);
}

// Fallback tương thích cũ cho Sidebar
function openEmailPreviewModal() {
  openPreviewFromSelection();
}

// --- 6. HÀM GỬI EMAIL TỪ VÙNG BÔI ĐEN ---
function sendBatchFromSelection() {
  const selectedSheetNames = getSelectedSheetNames();
  return sendBatchBySheetNames(selectedSheetNames);
}

function sendBatchBySheetNames(sheetNamesArray) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let successCount = 0;
  let errorCount = 0;

  sheetNamesArray.forEach((sheetName, index) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`⚠️ Bỏ qua, không tìm thấy Sheet: ${sheetName}`);
      return;
    }

    try {
      processSingleSheetSend(ss, sheet);
      successCount++;
      if (index < sheetNamesArray.length - 1) {
        Utilities.sleep(2500); // Hoãn 2.5s
      }
    } catch (err) {
      errorCount++;
      Logger.log(`❌ Lỗi tại Sheet ${sheetName}: ${err.message}`);
    }
  });

  return `Đã gửi thành công ${successCount}/${sheetNamesArray.length} email! (Lỗi: ${errorCount})`;
}

function executeActualSendFromPreview(sheetNamesJson) {
  const sheetNamesArray = JSON.parse(sheetNamesJson);
  return sendBatchBySheetNames(sheetNamesArray);
}

function sendInvDirectly() {
  return sendBatchFromSelection();
}

// --- 7. HÀM GỬI CHO ACTIVE SHEET ---
function sendInv() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheet = ss.getActiveSheet();
  
  const meta = getSheetMetadata(currentSheet);
  if (!meta['client_email']) {
    throw new Error(`Sheet hiện tại "${currentSheet.getName()}" không có client_email. Vui lòng bôi đen chọn các ô tên Sheet cần gửi (VD: 1639, 2068) trên bảng tính!`);
  }
  
  processSingleSheetSend(ss, currentSheet);
}