// =============================================================
//  LAYERISM — Google Apps Script Backend  (Code.gs)
//  Updated to match the redesigned order form:
//    - Product is now a dropdown (stored in 'description' field)
//    - Colour is now a dropdown (stored in 'color' field)
//    - Material field removed
// =============================================================

// ── CONFIGURATION ─────────────────────────────────────────────
// Replace these three values before deploying.

const SHEET_ID     = 'YOUR_GOOGLE_SHEET_ID';   // From Sheet URL: /spreadsheets/d/THIS_PART/edit
const FOLDER_ID    = 'YOUR_DRIVE_FOLDER_ID';   // From Drive folder URL: /folders/THIS_PART
const NOTIFY_EMAIL = 'your@email.com';         // Your email — receives a notification per order

// ── COLUMN HEADERS ────────────────────────────────────────────
const HEADERS = [
  'Timestamp', 'Name', 'Email', 'WhatsApp',
  'Product', 'Colour', 'Quantity',
  'Notes', 'Payment Proof Link', 'Status'
];

// =============================================================
//  doPost — called automatically when the webpage submits a form
// =============================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // 1. Open the Google Sheet and ensure headers exist
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const sheet       = spreadsheet.getActiveSheet();

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);

      // Style the header row
      const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
      headerRange.setFontWeight('bold')
                 .setBackground('#1C3560')   // Layerism navy
                 .setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);

      // Set sensible column widths
      sheet.setColumnWidth(1, 160);  // Timestamp
      sheet.setColumnWidth(2, 140);  // Name
      sheet.setColumnWidth(3, 180);  // Email
      sheet.setColumnWidth(4, 130);  // WhatsApp
      sheet.setColumnWidth(5, 220);  // Product
      sheet.setColumnWidth(6, 160);  // Colour
      sheet.setColumnWidth(7, 80);   // Quantity
      sheet.setColumnWidth(8, 260);  // Notes
      sheet.setColumnWidth(9, 260);  // Payment Proof Link
      sheet.setColumnWidth(10, 100); // Status
    }

    // 2. Save the payment proof image to Google Drive (if uploaded)
    let fileLink = '(not uploaded)';

    if (data.paymentImageBase64 && data.paymentImageBase64.length > 0) {
      const folder       = DriveApp.getFolderById(FOLDER_ID);
      const decodedBytes = Utilities.base64Decode(data.paymentImageBase64);
      const mimeType     = data.paymentFileType || 'image/png';
      const safeDate     = new Date().toISOString().replace(/[:.]/g, '-');
      const fileName     = data.paymentFileName || ('payment_' + safeDate);

      const blob = Utilities.newBlob(decodedBytes, mimeType, fileName);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileLink = file.getUrl();
    }

    // 3. Append order row
    sheet.appendRow([
      new Date(data.timestamp || Date.now()),  // Timestamp
      data.name        || '',                  // Name
      data.email       || '',                  // Email
      data.phone       || '',                  // WhatsApp
      data.description || '',                  // Product (from dropdown)
      data.color       || '',                  // Colour (from dropdown)
      data.quantity    || 1,                   // Quantity
      data.notes       || '',                  // Notes
      fileLink,                                // Payment proof Drive link
      'New',                                   // Status — update manually as order progresses
    ]);

    // 4. Send email notification
    const orderNum = sheet.getLastRow() - 1; // Subtract header row
    const subject  = `🖨️ Layerism — New order #${orderNum} from ${data.name}`;
    const body = [
      `New order received!\n`,
      `Order #:     ${orderNum}`,
      `Name:        ${data.name}`,
      `Email:       ${data.email}`,
      `WhatsApp:    ${data.phone || '—'}`,
      `\nProduct:     ${data.description}`,
      `Colour:      ${data.color}`,
      `Quantity:    ${data.quantity || 1}`,
      `\nNotes:       ${data.notes || '—'}`,
      `\nPayment proof: ${fileLink}`,
      `\nView all orders:\nhttps://docs.google.com/spreadsheets/d/${SHEET_ID}`,
    ].join('\n');

    MailApp.sendEmail({ to: NOTIFY_EMAIL, subject, body });

    return buildResponse({ success: true });

  } catch (err) {
    Logger.log('Error in doPost: ' + err.message);
    return buildResponse({ success: false, error: err.message });
  }
}

// =============================================================
//  doGet — lets you confirm the script is live by visiting the URL
// =============================================================
function doGet() {
  return ContentService
    .createTextOutput('Layerism backend is running. ✓')
    .setMimeType(ContentService.MimeType.TEXT);
}

// =============================================================
//  buildResponse — returns JSON with CORS-compatible headers
// =============================================================
function buildResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

// =============================================================
//  testSetup — run once manually to verify your IDs are correct.
//  In the editor: Run → testSetup
// =============================================================
function testSetup() {
  try {
    const sheet  = SpreadsheetApp.openById(SHEET_ID);
    const folder = DriveApp.getFolderById(FOLDER_ID);
    Logger.log('✓ Sheet found:  ' + sheet.getName());
    Logger.log('✓ Folder found: ' + folder.getName());
    Logger.log('✓ Notify email: ' + NOTIFY_EMAIL);
    Logger.log('✓ All good — ready to deploy!');
  } catch (err) {
    Logger.log('✗ Error: ' + err.message);
    Logger.log('  Double-check your SHEET_ID and FOLDER_ID values.');
  }
}
