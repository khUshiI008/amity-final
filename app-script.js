/**
 * Amity Online LP — lead capture to Google Sheets
 * Paste this whole file into Extensions > Apps Script of your Google Sheet,
 * then Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 */
const SHEET_NAME = 'Leads';
const TIMEZONE = 'Asia/Kolkata';
const HEADERS = [
  'Timestamp (IST)', 'Full name', 'Email', 'Mobile', 'State', 'Programme', 'Consent',
  'Page URL', 'UTM source', 'UTM medium', 'UTM campaign', 'UTM term', 'UTM content', 'Device'
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const p = (e && e.parameter) || {};
    if (p.company_website) return respond_({ ok: true });          // spam trap

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#092b50').setFontColor('#ffffff');
    }

    const row = sheet.getLastRow() + 1;
    sheet.getRange(row, 4).setNumberFormat('@');                     // keep mobile as text
    sheet.getRange(row, 1, 1, HEADERS.length).setValues([[
      Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss'),
      p.fullName || '', p.email || '', p.phone || '', p.state || '', p.programme || '', p.consent || '',
      p.pageUrl || '', p.utm_source || '', p.utm_medium || '', p.utm_campaign || '',
      p.utm_term || '', p.utm_content || '', p.device || ''
    ]]);
    return respond_({ ok: true });
  } catch (err) {
    return respond_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

// Open the Web App URL in a browser to check it is live.
function doGet() {
  return respond_({ ok: true, message: 'Lead endpoint is live' });
}

function respond_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}