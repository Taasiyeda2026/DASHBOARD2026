const SHEETS = {
  SETTINGS: 'SETTINGS',
  EMPLOYEES: 'EMPLOYEES',
  COURSES: 'COURSES',
  ZOOM_ASSIGNMENTS: 'ZOOM_ASSIGNMENTS'
};

const ZOOM_ASSIGNMENT_HEADERS = [
  'CourseId',
  'Date',
  'Authority',
  'School',
  'Program',
  'Employee',
  'EmployeeID',
  'StartTime',
  'EndTime',
  'ZoomAccount',
  'Notes',
  'UpdatedAt'
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
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const payload = parsePostData_(e);

    const normalized = {
      CourseId: String(payload.CourseId || '').trim(),
      Date: String(payload.Date || '').trim(),
      Authority: String(payload.Authority || '').trim(),
      School: String(payload.School || '').trim(),
      Program: String(payload.Program || '').trim(),
      Employee: String(payload.Employee || '').trim(),
      EmployeeID: String(payload.EmployeeID || '').trim(),
      StartTime: String(payload.StartTime || '').trim(),
      EndTime: String(payload.EndTime || '').trim(),
      ZoomAccount: String(payload.ZoomAccount || '').trim(),
      Notes: String(payload.Notes || '').trim(),
      UpdatedAt: String(payload.UpdatedAt || new Date().toISOString()).trim()
    };

    if (!normalized.CourseId) {
      return jsonResponse_({ ok: false, error: 'CourseId is required' });
    }

    upsertZoomAssignmentByCourseId_(normalized);
    return jsonResponse_({ ok: true, CourseId: normalized.CourseId });
  } finally {
    lock.releaseLock();
  }
}

function upsertZoomAssignmentByCourseId_(assignment) {
  const sheet = getSheetByName_(SHEETS.ZOOM_ASSIGNMENTS);
  ensureHeaders_(sheet, ZOOM_ASSIGNMENT_HEADERS);

  const values = sheet.getDataRange().getValues();
  const header = values[0] || ZOOM_ASSIGNMENT_HEADERS;

  const headerMap = {};
  header.forEach((name, idx) => headerMap[name] = idx);

  const courseIdIdx = headerMap['CourseId'];
  const updatedAtIdx = headerMap['UpdatedAt'];

  if (courseIdIdx === undefined) {
    throw new Error('CourseId column is missing in ZOOM_ASSIGNMENTS');
  }

  let targetRow = -1;
  for (let row = 1; row < values.length; row++) {
    const existingCourseId = String(values[row][courseIdIdx] || '').trim();
    if (existingCourseId === assignment.CourseId) {
      targetRow = row + 1;
      break;
    }
  }

  const newUpdatedAt = new Date(assignment.UpdatedAt).getTime() || Date.now();

  if (targetRow > 0) {
    const existingRow = values[targetRow - 1];
    const existingUpdatedAtRaw = existingRow[updatedAtIdx];
    const existingUpdatedAt = new Date(existingUpdatedAtRaw).getTime() || 0;

    if (existingUpdatedAt > newUpdatedAt) {
      return;
    }

    const mergedRow = header.map((key, idx) => {
      if (assignment[key] !== undefined) return assignment[key];
      return existingRow[idx];
    });

    sheet.getRange(targetRow, 1, 1, header.length).setValues([mergedRow]);
    return;
  }

  const rowData = header.map((key) => assignment[key] || '');
  sheet.appendRow(rowData);
}

function getSheetRowsAsObjects_(sheetName, expectedHeaders) {
  const sheet = getSheetByName_(sheetName);
  ensureHeaders_(sheet, expectedHeaders);

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];

  const timeZone =
    SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() ||
    Session.getScriptTimeZone() ||
    'Asia/Jerusalem';

  const header = values[0];

  return values
    .slice(1)
    .filter(row => row.some(cell => String(cell).trim() !== ''))
    .map((row) => {
      const obj = {};

      header.forEach((key, index) => {
        obj[key] = formatCellByHeader_(row[index], key, timeZone);
      });

      return obj;
    });
}

function formatCellByHeader_(value, header, timeZone) {
  if (value === null || value === undefined || value === '') return '';

  const key = String(header || '').trim();

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    if (key === 'Date') {
      return Utilities.formatDate(value, timeZone, 'yyyy-MM-dd');
    }

    if (key === 'StartTime' || key === 'EndTime') {
      return Utilities.formatDate(value, timeZone, 'HH:mm');
    }

    return Utilities.formatDate(value, timeZone, "yyyy-MM-dd'T'HH:mm:ss");
  }

  return value;
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
  if (e && e.postData && e.postData.contents) {
    const type = (e.postData.type || '').toLowerCase();

    if (type.indexOf('application/json') !== -1) {
      try {
        return JSON.parse(e.postData.contents);
      } catch (err) {
        return {};
      }
    }

    if (type.indexOf('application/x-www-form-urlencoded') !== -1) {
      const out = {};
      e.postData.contents.split('&').forEach(pair => {
        const parts = pair.split('=');
        const key = decodeURIComponent((parts[0] || '').replace(/\+/g, ' '));
        const value = decodeURIComponent((parts[1] || '').replace(/\+/g, ' '));
        out[key] = value;
      });
      return out;
    }
  }

  if (e && e.parameter && Object.keys(e.parameter).length > 0) {
    return e.parameter;
  }

  return {};
}

function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
