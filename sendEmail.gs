/**
 * THIẾT KẾ HTML CSS V3 - GOOGLE APPS SCRIPT
 */

function getSheetMetadata(sheet) {
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

function buildEmailData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  const meta = getSheetMetadata(sheet);

  const recipientEmail = meta['client_email'] || meta['email_cc'];
  const ccEmail = meta['email_cc'] || "";
  const recipientName = meta['client_name'] || "Valued Customer";
  const invPdfName = meta['inv_pdf_name'] || "invoice";
  const invId = meta['inv_id'] || "";

  if (!recipientEmail || !/\S+@\S+\.\S+/.test(recipientEmail)) {
    throw new Error("Invalid or missing client_email in N:O range.");
  }

  const timeZone = Session.getScriptTimeZone();
  const timeString = Utilities.formatDate(new Date(), timeZone, "HH:mm");
  const cleanInvoiceRef = invId.replace(/:/g, "").trim();
  const subject = `PSVN ${cleanInvoiceRef} [${timeString}]`;

  const htmlTemplate = HtmlService.createTemplateFromFile('emailTemplate');
  htmlTemplate.meta = meta;
  htmlTemplate.timeString = timeString;
  const htmlContent = htmlTemplate.evaluate().getContent();

  return {
    meta,
    recipientEmail,
    recipientName,
    ccEmail,
    invPdfName,
    cleanInvoiceRef,
    subject,
    htmlContent
  };
}

function sendInv() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  const emailData = buildEmailData();

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
    throw new Error("No data found in columns B:L.");
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

  sheet.hideColumns(14, 2);

  // === EXPORT INVOICE PDF (FILE 1) ===
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

  // === CHUẨN BỊ ATTACHMENTS (TẠO LINK DOWNLOAD TRỰC TIẾP TỪ GOOGLE DRIVE) ===
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

  // === GỬI EMAIL THỰC TẾ ===
  GmailApp.sendEmail(emailData.recipientEmail, emailData.subject, '', {
    htmlBody: emailData.htmlContent,
    cc: emailData.ccEmail,
    attachments: attachments,
    name: "PSVN Team"
  });

  Logger.log(`📧 Email sent successfully to ${emailData.recipientEmail} with ${attachments.length} PDFs.`);
}

function openEmailPreviewModal() {
  const emailData = buildEmailData();
  const base64Html = Utilities.base64Encode(emailData.htmlContent, Utilities.Charset.UTF_8);

  const dialogTemplate = HtmlService.createTemplateFromFile('previewDialog');
  dialogTemplate.recipientEmail = emailData.recipientEmail;
  dialogTemplate.ccEmail = emailData.ccEmail;
  dialogTemplate.subject = emailData.subject;
  dialogTemplate.base64Html = base64Html;

  const htmlOutput = dialogTemplate.evaluate()
    .setWidth(680)
    .setHeight(580);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '📧 Email Preview - PSVN Control Center');
}

function executeActualSend() {
  sendInv();
}

function sendInvDirectly() {
  sendInv();
}