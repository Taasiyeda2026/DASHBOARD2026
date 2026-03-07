/**
 * Google Apps Script Web App for Zoom assignments.
 * Deploy settings:
 *  - Execute as: Me
 *  - Who has access: Anyone
 */

const SHEETS = {
  SETTINGS: 'SETTINGS',
  EMPLOYEES: 'EMPLOYEES',
  COURSES: 'COURSES',
  ZOOM_ASSIGNMENTS: 'ZOOM_ASSIGNMENTS'
};

const ZOOM_ASSIGNMENT_HEADERS = [
  'CourseId',
  'Date',
  'Program',
  'Employee',
  'EmployeeID',
  'StartTime',
  'EndTime',
  'ZoomAccount',
  'Notes'
];

const COURSES_HEADERS = [
  'Id',
  'Date',
  'Authority',
  'School',
  'Program',
  'Employee',
  'EmployeeID',
  'StartTime',
  'EndTime',
  'Notes'
];

function doGet(e) {
  const type = (e && e.parameter && e.parameter.type) || 'assignments';
  if (type === 'courses') {
    const rows = getSheetRowsAsObjects_(SHEETS.COURSES, COURSES_HEADERS);
    return jsonResponse_(rows);
  }
  const rows = getSheetRowsAsObjects_(SHEETS.ZOOM_ASSIGNMENTS, ZOOM_ASSIGNMENT_HEADERS);
  return jsonResponse_(rows);
}

function doPost(e) {
  const payload = parsePostData_(e);
  const normalized = {
    CourseId: String(payload.CourseId || '').trim(),
    Date: String(payload.Date || '').trim(),
    Program: String(payload.Program || '').trim(),
    Employee: String(payload.Employee || '').trim(),
    EmployeeID: String(payload.EmployeeID || '').trim(),
    StartTime: String(payload.StartTime || '').trim(),
    EndTime: String(payload.EndTime || '').trim(),
    ZoomAccount: String(payload.ZoomAccount || '').trim(),
    Notes: String(payload.Notes || '').trim()
  };

  if (!normalized.CourseId) {
    return jsonResponse_({ ok: false, error: 'CourseId is required' });
  }

  upsertZoomAssignmentByCourseId_(normalized);
  return jsonResponse_({ ok: true, CourseId: normalized.CourseId });
}

function upsertZoomAssignmentByCourseId_(assignment) {
  const sheet = getSheetByName_(SHEETS.ZOOM_ASSIGNMENTS);
  ensureHeaders_(sheet, ZOOM_ASSIGNMENT_HEADERS);

  const values = sheet.getDataRange().getValues();
  const header = values[0] || ZOOM_ASSIGNMENT_HEADERS;
  const courseIdCol = header.indexOf('CourseId') + 1;

  if (courseIdCol <= 0) {
    throw new Error('CourseId column is missing in ZOOM_ASSIGNMENTS');
  }

  let targetRow = -1;
  for (let row = 1; row < values.length; row++) {
    const existingCourseId = String(values[row][courseIdCol - 1] || '').trim();
    if (existingCourseId === assignment.CourseId) {
      targetRow = row + 1;
      break;
    }
  }

  const rowData = ZOOM_ASSIGNMENT_HEADERS.map((key) => assignment[key] || '');

  if (targetRow > 0) {
    sheet.getRange(targetRow, 1, 1, ZOOM_ASSIGNMENT_HEADERS.length).setValues([rowData]);
    return;
  }

  sheet.appendRow(rowData);
}

function getSheetRowsAsObjects_(sheetName, expectedHeaders) {
  const sheet = getSheetByName_(sheetName);
  ensureHeaders_(sheet, expectedHeaders);

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];

  const header = values[0];
  return values.slice(1).map((row) => {
    const obj = {};
    header.forEach((key, index) => {
      obj[key] = row[index] ?? '';
    });
    return obj;
  });
}

function ensureHeaders_(sheet, headers) {
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  const existing = headerRange.getValues()[0].map((cell) => String(cell || '').trim());
  const same = headers.every((key, index) => existing[index] === key);
  if (same) return;

  headerRange.setValues([headers]);
}

function getSheetByName_(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Missing sheet: ${sheetName}`);
  }
  return sheet;
}

function parsePostData_(e) {
  // URLSearchParams (application/x-www-form-urlencoded) → e.parameter
  if (e && e.parameter && Object.keys(e.parameter).length > 0) {
    return e.parameter;
  }
  // Fallback: JSON body
  if (e && e.postData && e.postData.contents) {
    try {
      return JSON.parse(e.postData.contents);
    } catch (err) {
      return {};
    }
  }
  return {};
}

function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
