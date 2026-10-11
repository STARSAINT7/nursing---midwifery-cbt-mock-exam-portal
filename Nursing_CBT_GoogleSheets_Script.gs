/**
 * NURSING CBT PORTAL -> GOOGLE SHEETS
 * ------------------------------------
 * Paste this whole file into Extensions > Apps Script of your Google Sheet,
 * then Deploy > New deployment > Web app
 *   Execute as: Me
 *   Who has access: Anyone
 * and copy the Web app URL (ends with /exec) into SHEETS_WEBHOOK_URL
 * inside Nursing.html.
 *
 * One row is kept per exam attempt:
 *   - when a student STARTS the exam  -> a row appears with Status "In Progress"
 *   - when the student SUBMITS        -> the same row is updated with the score
 */

var SHEET_NAME = "Exam Log";

var HEADERS = [
  "Attempt ID", "Reg No", "Full Name", "Gender", "Course", "Status",
  "Started At", "Submitted At", "Time Taken (min)",
  "Questions Answered", "Total Questions", "Correct", "Score (/100)", "Submit Type"
];

function doGet() {
  return ContentService.createTextOutput("Nursing CBT sheet endpoint is running.");
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = getSheet_();

    if (!data.attemptId) {
      return ContentService.createTextOutput("ignored");
    }

    var rowIndex = findRow_(sheet, data.attemptId);

    if (data.event === "started") {
      if (rowIndex === -1) {
        sheet.appendRow(buildRow_(data, "In Progress"));
      } else {
        // The submit message arrived first (e.g. after being offline): only fill start info.
        var existing = sheet.getRange(rowIndex, 1, 1, HEADERS.length).getValues()[0];
        if (!existing[6]) sheet.getRange(rowIndex, 7).setValue(toDate_(data.startedAt));
      }
    } else if (data.event === "submitted") {
      var status = "Submitted";
      var row = buildRow_(data, status);
      if (rowIndex === -1) {
        sheet.appendRow(row);
      } else {
        sheet.getRange(rowIndex, 1, 1, HEADERS.length).setValues([row]);
      }
    }

    SpreadsheetApp.flush();
    return ContentService.createTextOutput("ok");
  } catch (err) {
    return ContentService.createTextOutput("error: " + err);
  } finally {
    lock.releaseLock();
  }
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold").setBackground("#dbeafe");
    sheet.setFrozenRows(1);
    // Keep Reg No as plain text so values like 12/04 are never turned into dates.
    sheet.getRange("B:B").setNumberFormat("@");
    // Show times in a readable format.
    sheet.getRange("G:H").setNumberFormat("dd-mmm-yyyy hh:mm:ss");
  }
  return sheet;
}

function findRow_(sheet, attemptId) {
  var last = sheet.getLastRow();
  if (last < 2) return -1;
  var ids = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(attemptId)) return i + 2;
  }
  return -1;
}

function buildRow_(d, status) {
  var submitted = d.event === "submitted";
  var minutes = (submitted && d.durationSec !== "" && d.durationSec != null)
    ? Math.round((Number(d.durationSec) / 60) * 10) / 10 : "";
  var finalStatus = status;
  if (submitted && d.submitType && String(d.submitType).indexOf("Auto") === 0) {
    finalStatus = "Submitted (time up)";
  }
  return [
    d.attemptId,
    String((d.regNo === undefined || d.regNo === null) ? "" : d.regNo).substring(0, 100),
    clean_(d.name),
    clean_(d.gender),
    clean_(d.course),
    finalStatus,
    toDate_(d.startedAt),
    submitted ? toDate_(d.submittedAt) : "",
    minutes,
    submitted ? d.answered : "",
    d.totalQuestions || "",
    submitted ? d.correct : "",
    submitted ? d.score : "",
    submitted ? clean_(d.submitType) : ""
  ];
}

// Stops text typed by students from being treated as a spreadsheet formula.
function clean_(v) {
  v = (v === undefined || v === null) ? "" : String(v);
  if (/^[=+\-@]/.test(v)) v = "'" + v;
  return v.substring(0, 200);
}

function toDate_(iso) {
  if (!iso) return "";
  var d = new Date(iso);
  return isNaN(d.getTime()) ? "" : d;
}
