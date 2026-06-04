const SPREADSHEET_ID = '1K222wkaRfCgE1PBvrJF7Dj9UHe701-Zt_U1rnmMl_cc';
const SHEET_NAME    = 'Availability';

// ----------------------------------------------------------------
//  ONE-TIME SETUP — run this once from the Apps Script editor
// ----------------------------------------------------------------
function setupSheet() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  let   sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  const headers = [
    'Timestamp', 'Name', 'Department', 'WeekOf',
    'Saturday', 'Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);

  // Style header row
  const hRange = sheet.getRange(1, 1, 1, headers.length);
  hRange.setBackground('#1B2A4A');
  hRange.setFontColor('#FFFFFF');
  hRange.setFontWeight('bold');
  hRange.setFontSize(11);

  // Column widths
  sheet.setColumnWidth(1, 200); // Timestamp
  sheet.setColumnWidth(2, 160); // Name
  sheet.setColumnWidth(3, 170); // Department
  sheet.setColumnWidth(4, 110); // WeekOf
  for (let i = 5; i <= 11; i++) sheet.setColumnWidth(i, 200);

  SpreadsheetApp.flush();
  Logger.log('Sheet setup complete! Ready to receive submissions.');
}


// ================================================================
//  MAIN REQUEST HANDLER
// ================================================================
function doGet(e) {
  let result;

  try {
    const action = (e.parameter && e.parameter.action) || '';

    if      (action === 'submit') { result = handleSubmit(e); }
    else if (action === 'fetch')  { result = handleFetch(e);  }
    else {
      result = {
        success : false,
        error   : 'Unknown action. Use ?action=submit&data=... or ?action=fetch&weekOf=YYYY-MM-DD'
      };
    }
  } catch (err) {
    result = { success: false, error: err.toString() };
    Logger.log('Unhandled error: ' + err.toString());
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}


// ================================================================
//  SUBMIT — save (or update) an employee's availability
// ================================================================
function handleSubmit(e) {
  const dataParam = e.parameter && e.parameter.data;
  if (!dataParam) {
    return { success: false, error: 'Missing "data" query parameter.' };
  }

  let data;
  try {
    data = JSON.parse(decodeURIComponent(dataParam));
  } catch (parseErr) {
    return { success: false, error: 'Could not parse data JSON: ' + parseErr.message };
  }

  const { name, department, weekOf, days } = data;

  if (!name || !department || !weekOf || !days) {
    return { success: false, error: 'Missing required fields: name, department, weekOf, days.' };
  }

  if (department !== 'Passenger Service' && department !== 'Ramp') {
    return { success: false, error: 'Department must be "Passenger Service" or "Ramp".' };
  }

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) {
    return { success: false, error: 'Sheet "' + SHEET_NAME + '" not found — run setupSheet() first.' };
  }

  // Check for an existing submission (same name + same week) to update instead of duplicate
  const allData       = sheet.getDataRange().getValues();
  let   existingRow   = -1;
  const nameLower     = name.trim().toLowerCase();

  for (let i = 1; i < allData.length; i++) {
    const rowName = allData[i][1] ? allData[i][1].toString().trim().toLowerCase() : '';
    const rowWeek = allData[i][3] ? allData[i][3].toString() : '';
    if (rowName === nameLower && rowWeek === weekOf) {
      existingRow = i + 1; // convert to 1-indexed sheet row
      break;
    }
  }

  const rowData = [
    new Date().toISOString(),
    name.trim(),
    department,
    weekOf,
    JSON.stringify(days.saturday  || []),
    JSON.stringify(days.sunday    || []),
    JSON.stringify(days.monday    || []),
    JSON.stringify(days.tuesday   || []),
    JSON.stringify(days.wednesday || []),
    JSON.stringify(days.thursday  || []),
    JSON.stringify(days.friday    || [])
  ];

  const isUpdate = existingRow > 0;
  if (isUpdate) {
    sheet.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }

  SpreadsheetApp.flush();
  Logger.log((isUpdate ? 'Updated' : 'New') + ' submission — ' + name + ' / week of ' + weekOf);

  return {
    success  : true,
    isUpdate : isUpdate,
    message  : (isUpdate ? 'Updated' : 'Saved') + ' availability for ' + name + ' (week of ' + weekOf + ')'
  };
}


// ================================================================
//  FETCH — return all submissions for a given week
// ================================================================
function handleFetch(e) {
  const weekOf = e.parameter && e.parameter.weekOf;
  if (!weekOf) {
    return { success: false, error: 'Missing "weekOf" parameter (format: YYYY-MM-DD).' };
  }

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) {
    return { success: false, error: 'Sheet "' + SHEET_NAME + '" not found — run setupSheet() first.' };
  }

  const allData = sheet.getDataRange().getValues();
  const results = [];

  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    if (!row[0] || !row[1]) continue; // skip blank rows

    if (row[3] && row[3].toString() === weekOf) {
      results.push({
        timestamp  : row[0].toString(),
        name       : row[1].toString(),
        department : row[2].toString(),
        weekOf     : row[3].toString(),
        days       : {
          saturday  : safeParseJSON(row[4]),
          sunday    : safeParseJSON(row[5]),
          monday    : safeParseJSON(row[6]),
          tuesday   : safeParseJSON(row[7]),
          wednesday : safeParseJSON(row[8]),
          thursday  : safeParseJSON(row[9]),
          friday    : safeParseJSON(row[10])
        }
      });
    }
  }

  Logger.log('Fetched ' + results.length + ' submission(s) for week of ' + weekOf);
  return { success: true, count: results.length, data: results };
}


// ================================================================
//  HELPER
// ================================================================
function safeParseJSON(val) {
  if (!val) return [];
  try {
    const parsed = JSON.parse(val.toString());
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    return [];
  }
}
