// ================================================================
// CLASS 12 CS QUIZ — Google Apps Script
// ================================================================
// SETUP INSTRUCTIONS:
// 1. Go to https://script.google.com → New project
// 2. Delete all existing code and paste this entire file
// 3. Click Save (Ctrl+S), name it "CS Quiz Receiver"
// 4. Click Deploy → New deployment
// 5. Type: Web App
// 6. Execute as: Me
// 7. Who has access: Anyone
// 8. Click Deploy → Copy the Web App URL
// 9. Paste that URL into index.html where it says SHEET_URL
// ================================================================

const SHEET_ID = ''; // Optional: paste your Google Sheet ID here
                     // If blank, a new sheet is auto-created on first run

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getOrCreateSheet();
    appendRow(sheet, data);
    return jsonResponse({status: 'ok', message: 'Result recorded'});
  } catch(err) {
    return jsonResponse({status: 'error', message: err.message});
  }
}

function doGet(e) {
  return jsonResponse({status: 'ok', message: 'CS Quiz API is running'});
}

function getOrCreateSheet() {
  let ss;
  if (SHEET_ID && SHEET_ID.length > 0) {
    ss = SpreadsheetApp.openById(SHEET_ID);
  } else {
    // Find existing sheet named "CS Quiz Results" in Drive
    const files = DriveApp.getFilesByName('CS Quiz Results');
    if (files.hasNext()) {
      ss = SpreadsheetApp.open(files.next());
    } else {
      ss = SpreadsheetApp.create('CS Quiz Results');
    }
  }

  let sheet = ss.getSheetByName('Responses');
  if (!sheet) {
    sheet = ss.insertSheet('Responses');
    // Add headers
    const headers = [
      'Timestamp', 'Student Name', 'Roll No', 'Section',
      'Chapter', 'Difficulty', 'Score', 'Total Questions',
      'Percentage (%)', 'Avg Time per Q (s)', 'Grade'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Style header row
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1a1a2e');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setFontSize(11);
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, headers.length, 160);
  }
  return sheet;
}

function appendRow(sheet, data) {
  const pct = parseFloat(data.percentage) || 0;
  const grade = pct >= 90 ? 'A+' : pct >= 80 ? 'A' : pct >= 70 ? 'B+' :
                pct >= 60 ? 'B' : pct >= 50 ? 'C' : 'D';

  const row = [
    data.timestamp || new Date().toLocaleString('en-IN'),
    data.name || '',
    data.roll || '',
    data.section || '',
    data.chapter || '',
    data.difficulty || '',
    data.score || 0,
    data.total || 0,
    pct,
    data.avgTime || 0,
    grade
  ];

  sheet.appendRow(row);

  // Color-code the grade cell
  const lastRow = sheet.getLastRow();
  const gradeCell = sheet.getRange(lastRow, 11);
  const colors = {'A+':'#d1fae5','A':'#bbf7d0','B+':'#fef9c3','B':'#fde68a','C':'#fed7aa','D':'#fee2e2'};
  gradeCell.setBackground(colors[grade] || '#f5f5f5');

  // Alternate row shading
  if (lastRow % 2 === 0) {
    sheet.getRange(lastRow, 1, 1, 11).setBackground('#f8fafc');
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
