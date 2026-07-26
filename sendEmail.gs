function sendInv() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var abc = '';
  // update 2

  // update 3

  // --- READ METADATA ---
  const invoiceRef = sheet.getRange("L6").getDisplayValue(); // e.g., "Invoice No : inst-1860-260721"
  const recipientEmail = (sheet.getRange("P9").getValue() || "").toString().trim();
  const ccEmail = (sheet.getRange("P18").getValue() || "").toString().trim();
  const recipientName = (sheet.getRange("P7").getDisplayValue() || "").trim() || "Valued Customer";
  const invName = (sheet.getRange("P8").getValue() || "").toString().trim() || "invoice";
  const invName2 = (sheet.getRange("P8").getValue() || "").toString().trim() || "invoice";

  if (!recipientEmail || !/\S+@\S+\.\S+/.test(recipientEmail)) {
    throw new Error("Invalid or missing email in P9.");
  }

    // --- ICT TIME (FIXED: use script's time zone reliably) ---
  const timeZone = Session.getScriptTimeZone(); // e.g., "Asia/Ho_Chi_Minh"
  const now = new Date();
  const timeString = Utilities.formatDate(now, timeZone, "HH:mm"); // 24h format, e.g., "14:34"

  // Clean invoice ref for subject
  let cleanInvoiceRef = invoiceRef.replace(/:/g, "").trim();
  const subject = `PSVN ${cleanInvoiceRef} [${timeString}]`;

  // // --- ICT TIME ---
  // const now = new Date();
  // const ictTime = new Date(now.getTime() + 7 * 60 * 60 * 1000);
  // const timeString = `${String(ictTime.getHours()).padStart(2,'0')}:${String(ictTime.getMinutes()).padStart(2,'0')}`;
  // const cleanInvoiceRef = invoiceRef.replace(/:/g, "").trim();
  // const subject = `PSVN ${cleanInvoiceRef} [${timeString}]`;

  // === 🎯 CORE: DEFINE PRINT RANGE = COLUMNS B to L (index 2 to 12) ===
  const startColIndex = 2; // B
  const endColIndex = 12;  // L
  const startRow = 1;      // assuming row 1 is header (adjust if needed, e.g., 2)

  // Helper: check if a row has ANY non-empty cell in columns B:L
  function rowHasContent(sheet, rowIndex, startCol, endCol) {
    const values = sheet.getRange(rowIndex, startCol, 1, endCol - startCol + 1).getValues()[0];
    return values.some(val => val !== null && val !== "" && val.toString().trim() !== "");
  }

  // Find first data row (skip header if row 1 is header and empty in B:L)
  let firstDataRow = startRow;
  if (rowHasContent(sheet, startRow, startColIndex, endColIndex)) {
    firstDataRow = startRow;
  } else {
    for (let r = startRow + 1; r <= sheet.getMaxRows(); r++) {
      if (rowHasContent(sheet, r, startColIndex, endColIndex)) {
        firstDataRow = r;
        break;
      }
    }
  }

  // Find last data row (scan from bottom up)
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

  // 👉 Build exact print range: B{firstDataRow}:L{lastDataRow}
  const colStart = String.fromCharCode(64 + startColIndex); // B
  const colEnd   = String.fromCharCode(64 + endColIndex);   // L
  const rangeStr = `${colStart}${firstDataRow}:${colEnd}${lastDataRow}`;

  Logger.log(`✅ Printing range: ${rangeStr} | Rows: ${firstDataRow} → ${lastDataRow}`);

  // --- OPTIONAL: widen column L (NOTE) safely ---
  const noteColIndex = endColIndex; // L
  const originalWidth = sheet.getColumnWidth(noteColIndex);
  const noteVals = sheet.getRange(firstDataRow, noteColIndex, lastDataRow - firstDataRow + 1, 1).getValues();
  let maxLength = 0;
  for (let i = 0; i < noteVals.length; i++) {
    const txt = (noteVals[i][0] || "").toString();
    maxLength = Math.max(maxLength, txt.length);
  }
  sheet.setColumnWidth(noteColIndex, Math.min(350, Math.max(150, maxLength * 6.5)));

  // --- EXPORT PDF ---
  const url = `https://docs.google.com/spreadsheets/d/${SpreadsheetApp.getActiveSpreadsheet().getId()}/export?`;
  const params = {
    exportFormat: 'pdf',
    format: 'pdf',
    size: 'letter',
    portrait: true,
    fitw: true,
    gridlines: false,
    printtitle: false,
    sheetnames: false,
    pagenum: 'UNDEFINED',
    attachment: true,
    gid: sheet.getSheetId(),
    range: rangeStr
  };

  const qs = Object.entries(params)
    .map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v)}`)
    .join('&');

  const pdfBlob = UrlFetchApp.fetch(url + qs, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
  }).getBlob().setName(`${invName}.pdf`);

  // --- EMAIL ---
  const htmlBody = `
    <div style="font-family: Arial; max-width:600px; margin:0 auto; padding:20px;">
      <div style="border:1px solid #ddd; background:#f8f8f8; padding:20px; border-radius:6px;">
        <p>Dear ${recipientName},</p>
        <p>We have attached the invoice for your review.</p>
        <p style="margin-top:20px; font-size:14px; color:#666; border-top:1px solid #ddd; padding-top:15px;">
          <a href="https://forms.gle/1BdEaBRPT3AowyZU7" style="color:#FF6B35; font-weight:bold;">Change billing details</a>
        </p>
      </div>
      <p style="margin-top:20px; font-weight:bold;">Best regards,<br>The PSVN Team</p>
    </div>
  `;

  GmailApp.sendEmail(recipientEmail, subject, '', {
    htmlBody,
    cc: ccEmail,
    attachments: [pdfBlob],
    name: "PSVN Team"
  });

  // Cleanup
  sheet.setColumnWidth(noteColIndex, originalWidth);
  Logger.log(`📧 Email sent with PDF: ${invName}.pdf | Range: ${rangeStr}`);
}

