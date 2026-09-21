function doGet(e) {
  return handleRequest(e && e.parameter ? e.parameter : {});
}

function doPost(e) {
  let payload = {};

  try {
    payload = JSON.parse(e && e.postData && e.postData.contents ? e.postData.contents : '{}');
  } catch (err) {
    payload = {};
  }

  if (!payload || Object.keys(payload).length === 0) {
    payload = e && e.parameter ? e.parameter : {};
  }

  return handleRequest(payload);
}

function handleRequest(params) {
  const action = String(params.action || '').trim();
  const passId = String(params.passId || params.pass_id || params.passid || '').trim();
  const number = String(
    params.number ||
    params.mobile ||
    params.phone ||
    params.whatsapp ||
    params.mobileNumber ||
    ''
  ).trim();

  const sheet = getTargetSheet();

  if (!sheet) {
    return jsonResponse({
      success: false,
      status: 'error',
      error: 'Google Sheet is not configured. Set SPREADSHEET_ID in Script Properties.'
    });
  }

  const normalizedAction = action.toLowerCase();

  switch (normalizedAction) {
    case 'checkpass':
      return jsonResponse(checkPass(number, sheet));

    case 'getpassbyid':
    case 'lookuppassbyid':
    case 'getpass':
    case 'lookuppass':
    case 'findpass':
    case 'fetchpass':
      return jsonResponse(getPassById(passId, sheet));

    case 'markattendance':
    case 'updateattendance':
    case 'confirmattendance':
    case 'attendancemark':
    case 'markattend':
      return jsonResponse(markAttendance(passId, sheet));

    default:
      return jsonResponse({
        success: false,
        status: 'error',
        error: 'Invalid action: ' + action
      });
  }
}

function getTargetSheet() {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');

  if (spreadsheetId) {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheetName = PropertiesService.getScriptProperties().getProperty('SHEET_NAME') || 'Registrations';
    return ss.getSheetByName(sheetName) || ss.getSheets()[0];
  }

  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!activeSpreadsheet) return null;

  const sheetName = PropertiesService.getScriptProperties().getProperty('SHEET_NAME') || 'Registrations';
  return activeSpreadsheet.getSheetByName(sheetName) || activeSpreadsheet.getSheets()[0];
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function normalizeValue(value) {
  return String(value == null ? '' : value).trim().toLowerCase();
}

function pickValue(row, keys) {
  for (const key of keys) {
    if (row[key] !== undefined && row[key] !== null && String(row[key]).trim() !== '') {
      return row[key];
    }
  }
  return '';
}

function canonicalizeRow(row) {
  const fallbackPassId = pickValue(row, ['passId', 'pass_id', 'passid', 'Pass ID', 'PassID', 'Pass_Id']);

  return {
    fullName: pickValue(row, ['fullName', 'full_name', 'Full_Name', 'name', 'Name', 'Applicant Name', 'Volunteer Name']),
    mobile: pickValue(row, ['mobile', 'phone', 'Mobile', 'Phone', 'mobileNumber', 'Mobile Number', 'WhatsApp Number', 'whatsapp']),
    whatsapp: pickValue(row, ['whatsapp', 'WhatsApp', 'Whatsapp', 'WhatsApp Number', 'mobile', 'phone']),
    district: pickValue(row, ['district', 'District']),
    constituency: pickValue(row, ['constituency', 'Constituency', 'Constituency Name']),
    passId: fallbackPassId || '',
    permissionStatus: pickValue(row, ['Permission_Status', 'permissionStatus', 'permission_status', 'Permission Status', 'Approval Status', 'approval_status']),
    attendanceStatus: pickValue(row, ['Attendance', 'Attendance_Status', 'Attendance Status', 'attendance', 'attendance_status', 'Attended', 'Status'])
  };
}

function getSheetRows(sheet) {
  const values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) return [];

  const headers = values[0].map(function(h) {
    return String(h || '').trim();
  });

  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const row = {};
    for (let j = 0; j < headers.length; j++) {
      row[headers[j]] = values[i][j];
    }
    rows.push(row);
  }
  return rows;
}

function findRowIndex(sheet, matcherFn) {
  const values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) return -1;

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (matcherFn(row, i, values[0])) return i;
  }

  return -1;
}

function findHeaderIndex(headers, candidates) {
  for (let i = 0; i < headers.length; i++) {
    const header = String(headers[i] || '').trim();
    const normalizedHeader = normalizeValue(header);

    for (const candidate of candidates) {
      if (normalizedHeader === normalizeValue(candidate)) {
        return i;
      }
    }
  }
  return -1;
}

function ensurePassId(value) {
  return String(value || '').trim();
}

function checkPass(number, sheet) {
  if (!number) {
    return {
      success: false,
      status: 'error',
      error: 'Phone number is required.'
    };
  }

  const targetNumber = normalizeValue(number);
  const rows = getSheetRows(sheet);

  for (const row of rows) {
    const rowData = canonicalizeRow(row);
    const mobile = normalizeValue(rowData.mobile || row['Mobile'] || row['Phone'] || row['WhatsApp Number'] || row['whatsapp']);
    const whatsapp = normalizeValue(rowData.whatsapp || row['WhatsApp'] || row['Whatsapp'] || row['WhatsApp Number']);

    if (mobile === targetNumber || whatsapp === targetNumber) {
      const permissionStatus = normalizeValue(
        pickValue(row, ['Permission_Status', 'permissionStatus', 'permission_status', 'Permission Status', 'Approval Status', 'approval_status'])
      );

      if (permissionStatus === 'approved') {
        return {
          success: true,
          status: 'approved',
          data: canonicalizeRow(row)
        };
      }

      if (permissionStatus === 'not_approved') {
        return {
          success: false,
          status: 'not_approved',
          error: 'Your pass has not been approved yet.'
        };
      }

      return {
        success: false,
        status: 'not_found',
        error: 'No registration found with this mobile or WhatsApp number.'
      };
    }
  }

  return {
    success: false,
    status: 'not_found',
    error: 'No registration found with this mobile or WhatsApp number.'
  };
}

function getPassById(passId, sheet) {
  const targetPassId = ensurePassId(passId);

  if (!targetPassId) {
    return {
      success: false,
      status: 'error',
      error: 'Pass ID is required.'
    };
  }

  const rows = getSheetRows(sheet);

  for (const row of rows) {
    const rowPassId = ensurePassId(
      pickValue(row, ['passId', 'pass_id', 'passid', 'Pass ID', 'PassID', 'Pass_Id'])
    );

    if (normalizeValue(rowPassId) === normalizeValue(targetPassId)) {
      return {
        success: true,
        status: 'found',
        data: canonicalizeRow(row)
      };
    }
  }

  return {
    success: false,
    status: 'not_found',
    error: 'Pass ID not found.'
  };
}

function markAttendance(passId, sheet) {
  const targetPassId = ensurePassId(passId);

  if (!targetPassId) {
    return {
      success: false,
      status: 'error',
      error: 'Pass ID is required.'
    };
  }

  const rowIndex = findRowIndex(sheet, function(row) {
    const rowPassId = ensurePassId(
      pickValue(row, ['passId', 'pass_id', 'passid', 'Pass ID', 'PassID', 'Pass_Id'])
    );
    return normalizeValue(rowPassId) === normalizeValue(targetPassId);
  });

  if (rowIndex === -1) {
    return {
      success: false,
      status: 'not_found',
      error: 'Pass ID not found.'
    };
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const attendanceSet = ['Attendance', 'Attendance_Status', 'Attendance Status', 'attendance', 'attendance_status', 'Attended', 'Status'];
  const timeSet = ['Marked At', 'Attendance Time', 'Attended At', 'Updated At', 'Timestamp', 'Date Marked'];

  let didUpdate = false;

  for (const headerName of attendanceSet) {
    const idx = findHeaderIndex(headers, [headerName]);
    if (idx !== -1) {
      sheet.getRange(rowIndex + 1, idx + 1).setValue('Present');
      didUpdate = true;
    }
  }

  for (const headerName of timeSet) {
    const idx = findHeaderIndex(headers, [headerName]);
    if (idx !== -1) {
      sheet.getRange(rowIndex + 1, idx + 1).setValue(new Date().toISOString());
    }
  }

  if (!didUpdate) {
    const appendCandidate = findHeaderIndex(headers, ['Attendance', 'Attendance Status', 'Status']);
    if (appendCandidate === -1) {
      const lastColumn = sheet.getLastColumn();
      sheet.getRange(1, lastColumn + 1).setValue('Attendance');
      sheet.getRange(rowIndex + 1, lastColumn + 1).setValue('Present');
      sheet.getRange(1, lastColumn + 2).setValue('Marked At');
      sheet.getRange(rowIndex + 1, lastColumn + 2).setValue(new Date().toISOString());
    }
  }

  return {
    success: true,
    status: 'updated',
    message: 'Attendance marked successfully.',
    passId: targetPassId
  };
}

function setScriptProperties() {
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', 'PASTE_SPREADSHEET_ID_HERE');
  PropertiesService.getScriptProperties().setProperty('SHEET_NAME', 'Registrations');
}
