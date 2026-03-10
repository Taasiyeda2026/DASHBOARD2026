function getSheetByNameOrCreate_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }

  return sheet;
}

function getCoursesSheet_() {
  return getSheetByNameOrCreate_(
    "COURSES",
    ["Id", "Date", "Authority", "School", "Program", "Employee", "EmployeeID", "StartTime", "EndTime"]
  );
}

function getAssignmentsSheet_() {
  return getSheetByNameOrCreate_(
    "ZOOM_ASSIGNMENTS",
    ["CourseId", "Date", "Program", "Employee", "EmployeeID", "StartTime", "EndTime", "ZoomAccount", "Notes"]
  );
}

function sheetToObjects_(sheet) {
  const values = sheet.getDataRange().getDisplayValues();
  if (!values || values.length === 0) return [];

  const headers = values[0];
  const rows = values.slice(1);

  return rows
    .filter(row => row.some(cell => String(cell).trim() !== ""))
    .map(row => {
      const obj = {};
      headers.forEach((header, i) => {
        obj[header] = row[i] || "";
      });
      return obj;
    });
}

function doGet(e) {
  const type = String((e && e.parameter && e.parameter.type) || "").trim().toLowerCase();

  let data = [];

  if (type === "courses") {
    data = sheetToObjects_(getCoursesSheet_());
  } else if (type === "assignments") {
    data = sheetToObjects_(getAssignmentsSheet_());
  } else {
    data = {
      status: "error",
      message: "Missing or invalid 'type' parameter. Use ?type=courses or ?type=assignments"
    };
  }

  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const sheet = getAssignmentsSheet_();

  const courseId = String((e.parameter.CourseId || "")).trim();
  const date = String((e.parameter.Date || "")).trim();
  const program = String((e.parameter.Program || "")).trim();
  const employee = String((e.parameter.Employee || "")).trim();
  const employeeID = String((e.parameter.EmployeeID || "")).trim();
  const startTime = String((e.parameter.StartTime || "")).trim();
  const endTime = String((e.parameter.EndTime || "")).trim();
  const zoomAccount = String((e.parameter.ZoomAccount || "")).trim();
  const notes = String((e.parameter.Notes || "")).trim();

  if (!courseId || !date || !employee || !employeeID || !startTime || !endTime) {
    return ContentService
      .createTextOutput(JSON.stringify({
        status: "error",
        message: "Missing required fields"
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const values = sheet.getDataRange().getDisplayValues();
  let foundRow = -1;

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === courseId) {
      foundRow = i + 1;
      break;
    }
  }

  const rowData = [
    courseId,
    date,
    program,
    employee,
    employeeID,
    startTime,
    endTime,
    zoomAccount,
    notes
  ];

  if (foundRow === -1) {
    sheet.appendRow(rowData);
  } else {
    sheet.getRange(foundRow, 1, 1, rowData.length).setValues([rowData]);
  }

  return ContentService
    .createTextOutput(JSON.stringify({
      status: "ok",
      CourseId: courseId
    }))
    .setMimeType(ContentService.MimeType.JSON);
}
