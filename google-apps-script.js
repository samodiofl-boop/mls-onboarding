/**
 * Beycome MLS Training — Google Apps Script
 *
 * SETUP INSTRUCTIONS:
 * 1. Go to sheets.google.com → create a new sheet → name it "MLS Training Records"
 * 2. In that sheet, click Extensions → Apps Script
 * 3. Delete everything in the editor and paste this entire file
 * 4. Click Save (Ctrl+S)
 * 5. Click Deploy → New deployment
 *    - Type: Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 6. Click Deploy → copy the Web app URL
 * 7. Open /index.html and replace 'YOUR_APPS_SCRIPT_URL_HERE' with that URL
 * 8. Commit and push index.html
 *
 * The sheet will auto-create headers on the first submission.
 */

const SHEET_NAME = 'Records'; // You can rename the tab in Sheets to "Records"

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    // Create sheet with headers if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        'Timestamp',
        'Name',
        'Event',
        'Section #',
        'Section Title',
        'Score',
        'Passed',
        'Attempt #',
        'Training Run #'
      ]);
      // Bold the header row
      sheet.getRange(1, 1, 1, 9).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    const data = JSON.parse(e.postData.contents);

    sheet.appendRow([
      new Date(data.timestamp).toLocaleString('en-US', { timeZone: 'America/New_York' }),
      data.name        || '',
      data.event       || '',
      data.sectionNum  || '',
      data.sectionTitle|| '',
      data.score       || '',
      data.passed      || '',
      data.attemptNum  || '',
      data.trainingRun || ''
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Test this function manually in the Apps Script editor to verify the sheet works
function testLog() {
  const mockEvent = {
    postData: {
      contents: JSON.stringify({
        timestamp: new Date().toISOString(),
        name: 'Test User',
        event: 'quiz_attempt',
        sectionNum: 3,
        sectionTitle: 'Photo Rules',
        score: '4/4',
        passed: 'Yes',
        attemptNum: 1,
        trainingRun: 0
      })
    }
  };
  const result = doPost(mockEvent);
  Logger.log(result.getContent());
}
