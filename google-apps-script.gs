/**
 * Panoptes Landing Page → Google Sheets
 * Receives lead form submissions and appends them as rows.
 *
 * DEPLOYMENT (one time):
 *  1. Open the target sheet:
 *     https://docs.google.com/spreadsheets/d/1Deh9A-KW11OCgYytn9ISb5Ka631BmzV6Kd0u-pHBDy0/edit
 *  2. Extensions → Apps Script
 *  3. Paste this entire file, replacing whatever is there
 *  4. Save (disk icon) and name the project "Panoptes Leads"
 *  5. Deploy → New deployment → Type: Web app
 *       - Description: Panoptes leads webhook
 *       - Execute as: Me
 *       - Who has access: Anyone
 *  6. Authorize when prompted
 *  7. Copy the Web app URL (ends with /exec)
 *  8. Paste it into index.html where it says PASTE_YOUR_APPS_SCRIPT_URL_HERE
 *
 * To update the script later: Deploy → Manage deployments → Edit → New version → Deploy
 */

const SHEET_ID  = '1Deh9A-KW11OCgYytn9ISb5Ka631BmzV6Kd0u-pHBDy0';
const SHEET_TAB = 'Sheet1'; // change if your tab has a different name

// Column order in the sheet. Adding a new key here = new column.
const COLUMNS = [
  'timestamp',
  'client_name',
  'client_homepage',
  'ai_overview_location',
  'categories',
  'competitors',
  'package_type',
  'is_demo',
  'max_prompts',
  'frequency',
  'source',
  'referrer',
  'user_agent',
  'full_payload_json'
];

function doPost(e) {
  try {
    // Accepts both a raw JSON body (e.postData.contents) and a form-encoded field
    // named "payload" (e.parameter.payload). The form-encoded path is used by the
    // hidden-iframe technique which bypasses browser CORS restrictions on redirects.
    let raw;
    if (e.parameter && e.parameter.payload) {
      raw = e.parameter.payload;
    } else {
      raw = e.postData.contents;
    }
    const data = JSON.parse(raw);
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_TAB);

    // Add header row on first run
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(COLUMNS);
      sheet.getRange(1, 1, 1, COLUMNS.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    // Build the row in column order; unknown keys become blank
    const row = COLUMNS.map(col => data[col] !== undefined ? data[col] : '');
    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Optional: smoke-test in the Apps Script editor (Run > doTest)
function doTest() {
  doPost({ parameter: { payload: JSON.stringify({
    timestamp: new Date().toISOString(),
    client_name: 'Test Co',
    client_homepage: 'https://example.com',
    categories: 'cat1, cat2',
    competitors: 'https://comp1.com',
    package_type: 'demo',
    is_demo: true,
    max_prompts: '40',
    frequency: 'daily',
    source: 'test',
    referrer: 'manual',
    user_agent: 'test-runner',
    full_payload_json: '{}'
  }) } });
}
