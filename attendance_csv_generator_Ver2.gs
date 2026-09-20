/**
 * ATTENDANCE CSV GENERATOR
 * -------------------------------------------------------------
 * Adds a menu button to your Google Sheet that reads the current
 * attendance grid, cross-references the subject_id mapping tab,
 * and downloads a CSV in the format:
 *   student_email, subject_id, attend_date, status, attend_time
 *
 * SETUP:
 * 1. Open your Google Sheet.
 * 2. Extensions -> Apps Script.
 * 3. Delete any starter code in Code.gs, paste this whole file in.
 * 4. Edit the CONFIGURATION block below -- set MAPPING_SHEET_NAME to
 *    match your subject mapping tab.
 * 5. Save (the disk icon). Close the Apps Script tab.
 * 6. Reload your Google Sheet. A new menu "Attendance Tools" will
 *    appear at the top.
 * 7. Click on the SECTION TAB you want to export first (e.g.
 *    "Section - A"), THEN click Attendance Tools -> Generate CSV.
 *    The script reads whichever tab is open/active, so the same
 *    script works for every section -- just switch tabs and re-run.
 * 8. The FIRST time you run it, Google will ask you to authorize
 *    the script (click through "Advanced" -> "Go to project (unsafe)"
 *    -- this warning is normal for your own scripts). This is a
 *    one-time step.
 * -------------------------------------------------------------
 */

// ====================== CONFIGURATION ======================

// Tab (sheet) names -- EDIT THIS to match your mapping tab's exact name.
// The ATTENDANCE tab is no longer hardcoded -- the script uses whichever
// tab (section) is open/active when you click "Generate CSV", so the
// same script works for Section - A, Section - B, etc. without editing.
var MAPPING_SHEET_NAME    = 'Subject Mapping';  // the separate subject_id tab

// If you have non-section tabs that should never be run by mistake
// (e.g. "Count", "Summary"), list their exact names here.
var NON_ATTENDANCE_SHEET_NAMES = ['Count', 'Summary'];

// Mapping tab layout -- EDIT if your columns are in a different order
// Assumes row 1 is a header row, data starts row 2.
var MAPPING_CLASS_CODE_COL = 1; // column A: class code (e.g. U26AIMLB102)
var MAPPING_SUBJECT_ID_COL = 2; // column B: subject_id

// Attendance tab layout -- matches the screenshot you shared.
// Only change these if your sheet's structure differs.
var SNO_COL   = 1; // A
var NAME_COL  = 2; // B
var PHONE_COL = 3; // C
var EMAIL_COL = 5; // E  <-- the email column you said to use

var DATE_ROW       = 1; // row with the date (merged across a day's classes)
var CLASS_CODE_ROW = 2; // row with class code
var TIMING_ROW      = 3; // row with class timing (e.g. 9:30-10:30)
var SUBJECT_ROW     = 4; // row with subject name (not used in output, FYI only)

var FIRST_DATA_ROW  = 6; // first row of actual student data
var FIRST_CLASS_COL = 6; // column F = first class/attendance column

// =============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Attendance Tools')
    .addItem('Generate CSV', 'generateAttendanceCSV')
    .addToUi();
}

function generateAttendanceCSV() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet(); // whichever tab/section is open right now
  var mappingSheet = ss.getSheetByName(MAPPING_SHEET_NAME);
  var ui = SpreadsheetApp.getUi();

  if (sheet.getName() === MAPPING_SHEET_NAME || NON_ATTENDANCE_SHEET_NAMES.indexOf(sheet.getName()) !== -1) {
    ui.alert('"' + sheet.getName() + '" doesn\'t look like a section/attendance tab. Open the section tab you want to export (e.g. "Section - A"), then click Generate CSV again.');
    return;
  }
  if (!mappingSheet) {
    ui.alert('Could not find a tab named "' + MAPPING_SHEET_NAME + '". Check MAPPING_SHEET_NAME in the script.');
    return;
  }
  if (sheet.getLastRow() < FIRST_DATA_ROW) {
    ui.alert('No student data found below row ' + FIRST_DATA_ROW + ' on "' + sheet.getName() + '".');
    return;
  }

  // ---- Ask which date to export ----
  var promptResponse = ui.prompt(
    'Which date?',
    'Exporting "' + sheet.getName() + '". Enter the date to export, in YYYY-MM-DD format (e.g. 2026-08-31):',
    ui.ButtonSet.OK_CANCEL
  );
  if (promptResponse.getSelectedButton() !== ui.Button.OK) return; // cancelled

  var targetDate = promptResponse.getResponseText().trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(targetDate)) {
    ui.alert('Please enter the date as YYYY-MM-DD, e.g. 2026-08-31.');
    return;
  }

  // ---- Build class code -> subject_id lookup ----
  var mapData = mappingSheet.getDataRange().getValues();
  var codeToSubjectId = {};
  for (var i = 1; i < mapData.length; i++) { // skip header row
    var code = String(mapData[i][MAPPING_CLASS_CODE_COL - 1]).trim();
    var subjectId = mapData[i][MAPPING_SUBJECT_ID_COL - 1];
    if (code) codeToSubjectId[code] = subjectId;
  }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  var dateRow      = sheet.getRange(DATE_ROW, 1, 1, lastCol).getValues()[0];
  var classCodeRow = sheet.getRange(CLASS_CODE_ROW, 1, 1, lastCol).getValues()[0];
  var timingRow    = sheet.getRange(TIMING_ROW, 1, 1, lastCol).getValues()[0];

  // Merged date cells only carry a value in their top-left cell;
  // forward-fill so every class column knows its date.
  var filledDates = dateRow.slice();
  for (var c = FIRST_CLASS_COL - 1; c < lastCol; c++) {
    if ((filledDates[c] === '' || filledDates[c] === null) && c > 0) {
      filledDates[c] = filledDates[c - 1];
    }
  }

  var dataRange = sheet.getRange(FIRST_DATA_ROW, 1, lastRow - FIRST_DATA_ROW + 1, lastCol).getValues();

  var output = [['student_email', 'subject_id', 'attend_date', 'status', 'attend_time']];
  var skippedCodes = {}; // class codes with no subject_id mapping

  // Outer loop = class columns (class-by-class), inner loop = students --
  // so the CSV is grouped by class first, then all students for that
  // class, then the next class.
  for (var c = FIRST_CLASS_COL - 1; c < lastCol; c++) {
    var classCode = String(classCodeRow[c]).trim();
    if (!classCode) continue; // not a class column

    var attendDate = formatDate(filledDates[c]);
    if (attendDate !== targetDate) continue; // not the date we're exporting

    var subjectId = codeToSubjectId[classCode];
    if (subjectId === undefined || subjectId === '') {
      skippedCodes[classCode] = true;
      continue; // no mapping found -- skip, don't guess
    }

    var attendTime = extractStartTime(String(timingRow[c]));

    for (var r = 0; r < dataRange.length; r++) {
      var row = dataRange[r];
      var email = String(row[EMAIL_COL - 1]).trim();
      if (!email || email.toUpperCase() === '#N/A') continue; // only students with a real email

      var statusRaw = String(row[c]).trim().toLowerCase();
      var status = '';
      if (statusRaw.indexOf('present') !== -1) status = 'present';
      else if (statusRaw.indexOf('absent') !== -1) status = 'absent';
      else if (statusRaw.indexOf('late') !== -1) status = 'late';
      else continue; // blank / unrecognized cell -- skip

      output.push([email, subjectId, attendDate, status, attendTime]);
    }
  }

  if (output.length === 1) {
    ui.alert('No matching classes found for ' + targetDate + ' on "' + sheet.getName() + '". Double-check the sheet has classes on that date.');
    return;
  }

  var missing = Object.keys(skippedCodes);
  if (missing.length > 0) {
    var proceed = ui.alert(
      'Heads up',
      'These class codes on ' + targetDate + ' have no subject_id in "' + MAPPING_SHEET_NAME + '" and were skipped:\n' + missing.join(', ') + '\n\nContinue and download anyway?',
      ui.ButtonSet.OK_CANCEL
    );
    if (proceed !== ui.Button.OK) return;
  }

  var csvContent = output.map(function(row) {
    return row.map(csvEscape).join(',');
  }).join('\n');

  downloadCsv(csvContent, sheet.getName(), targetDate);
}

function downloadCsv(csvContent, sectionName, targetDate) {
  var safeSection = sectionName.replace(/[^a-zA-Z0-9]+/g, '_');
  var fileName = 'attendance_' + safeSection + '_' + targetDate + '.csv';
  var base64 = Utilities.base64Encode(csvContent, Utilities.Charset.UTF_8);
  var html = HtmlService.createHtmlOutput(
    '<html><body style="font-family:sans-serif;padding:12px;">' +
    '<p>Your download should start automatically.</p>' +
    '<a id="dl" download="' + fileName + '" href="data:text/csv;charset=utf-8;base64,' + base64 + '">Click here if it does not</a>' +
    '<script>document.getElementById("dl").click();</script>' +
    '</body></html>'
  ).setWidth(320).setHeight(90);
  SpreadsheetApp.getUi().showModalDialog(html, 'Downloading CSV...');
}

function formatDate(rawDate) {
  if (Object.prototype.toString.call(rawDate) === '[object Date]') {
    return Utilities.formatDate(rawDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var s = String(rawDate).trim();
  if (!s) return '';
  // Handle strings like "31-Aug" by assuming the current year
  var parsed = new Date(s + ' ' + new Date().getFullYear());
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return s; // fallback: leave as-is so you can spot it in the output
}

function extractStartTime(timing) {
  // "9:30-10:30" -> "09:30"
  var match = timing.match(/(\d{1,2}):(\d{2})/);
  if (!match) return '';
  var hh = ('0' + match[1]).slice(-2);
  return hh + ':' + match[2];
}

function csvEscape(val) {
  val = String(val);
  if (val.indexOf(',') !== -1 || val.indexOf('"') !== -1 || val.indexOf('\n') !== -1) {
    val = '"' + val.replace(/"/g, '""') + '"';
  }
  return val;
}
