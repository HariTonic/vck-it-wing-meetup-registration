const DEFAULT_SPREADSHEET_ID = '1hg25ZSwthYldZPwtD3ZmhbUuxm9R7neXiTHhuo4dXVE';
const DEFAULT_SHEET_NAME = 'SeptMeetResponse';

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

  if (!normalizedAction && params && (params.fullName || params.mobile || params.email || params.district)) {
    return jsonResponse(saveRegistrationRecord(params, sheet));
  }

  switch (normalizedAction) {
    case 'submitregistration':
    case 'submitapplication':
    case 'register':
    case 'applypass':
    case 'addregistration':
      return jsonResponse(saveRegistrationRecord(params, sheet));

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
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || DEFAULT_SPREADSHEET_ID;

  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheetName = PropertiesService.getScriptProperties().getProperty('SHEET_NAME') || DEFAULT_SHEET_NAME;
    return ss.getSheetByName(sheetName) || ss.getSheets()[0];
  } catch (error) {
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!activeSpreadsheet) return null;

    const sheetName = PropertiesService.getScriptProperties().getProperty('SHEET_NAME') || DEFAULT_SHEET_NAME;
    return activeSpreadsheet.getSheetByName(sheetName) || activeSpreadsheet.getSheets()[0];
  }
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function normalizeValue(value) {
  return String(value == null ? '' : value).trim().toLowerCase();
}

function normalizeHeaderName(value) {
  return String(value == null ? '' : value).trim().replace(/\s+/g, ' ').toLowerCase();
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
  const fallbackPassId = pickValue(row, ['passId', 'pass_id', 'passid', 'Pass ID', 'PassID', 'Pass_Id', 'PassID']);

  return {
    fullName: pickValue(row, ['fullName', 'full_name', 'Full_Name', 'name', 'Name', 'Applicant Name', 'Volunteer Name', 'Full Name']),
    mobile: pickValue(row, ['mobile', 'phone', 'Mobile', 'Phone', 'mobileNumber', 'Mobile Number', 'WhatsApp Number', 'whatsapp', 'WhatsApp', 'Whatsapp']),
    whatsapp: pickValue(row, ['whatsapp', 'WhatsApp', 'Whatsapp', 'WhatsApp Number', 'mobile', 'phone', 'Mobile', 'Phone']),
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
    const normalizedHeader = normalizeHeaderName(header);

    for (const candidate of candidates) {
      if (normalizedHeader === normalizeHeaderName(candidate)) {
        return i;
      }
    }
  }
  return -1;
}

function ensurePassId(value) {
  return String(value || '').trim();
}

function generatePassId() {
  return 'vckitwingpass-sept-' + String(Date.now()).slice(-4).padStart(4, '0');
}

function coerceAttendanceValue(value) {
  if (typeof value === 'boolean') return value;

  const normalized = String(value == null ? '' : value).trim().toLowerCase();
  if (normalized === 'true' || normalized === 'yes' || normalized === 'present' || normalized === '1' || normalized === 'y') return true;
  if (normalized === 'false' || normalized === 'no' || normalized === 'absent' || normalized === '0' || normalized === 'n' || normalized === '') return false;

  return Boolean(normalized);
}

function getHeaderList(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) return [];
  return sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(function(header) {
    return String(header || '').trim().replace(/\s+/g, ' ');
  });
}

function toSafeCellValue(value) {
  if (value === undefined || value === null) return '';
  if (Array.isArray(value)) return value.join(', ');
  return String(value);
}

function saveRegistrationRecord(data, sheet) {
  const payload = data || {};
  const fullName = toSafeCellValue(payload['Full Name'] || payload.fullName || payload.name || '').trim();
  const mobile = String(payload.Mobile || payload.mobile || payload.phone || '').trim();
  const whatsapp = String(payload.WhatsApp || payload.whatsapp || '').trim();
  const email = toSafeCellValue(payload.Email || payload.email || '').trim();
  const age = toSafeCellValue(payload.Age || payload.age || '').trim();
  const gender = toSafeCellValue(payload.Gender || payload.gender || '').trim();
  const district = toSafeCellValue(payload.District || payload.district || '').trim();
  const constituency = toSafeCellValue(payload.Constituency || payload.constituency || '').trim();
  const townCity = toSafeCellValue(payload['Town/City'] || payload.townCity || payload.city || '').trim();
  const areaWard = toSafeCellValue(payload['Area/Ward'] || payload.areaWard || payload.zipCode || '').trim();
  const skills = toSafeCellValue(payload.Skills || payload.skills || '').trim();
  const preferredActivity = toSafeCellValue(payload['Preferred Activity'] || payload.preferredActivity || '').trim();
  const foodType = toSafeCellValue(payload['Food Type'] || payload.foodType || '').trim();
  const additionalInfo = toSafeCellValue(payload['Additional Information'] || payload.additionalInfo || '').trim();
  const timestamp = toSafeCellValue(payload.Timestamp || payload.submittedAt || new Date().toISOString()).trim();
  const passId = toSafeCellValue(payload.PassID || payload.passId || generatePassId()).trim();
  const permissionStatus = toSafeCellValue(payload.Permission_Status || payload.permissionStatus || 'pending').trim();
  const attendanceStatus = coerceAttendanceValue(payload.Attendance ?? payload.attendanceStatus ?? payload.attendance ?? false);

  if (!fullName || !mobile) {
    return {
      success: false,
      status: 'error',
      error: 'Name and mobile are required.'
    };
  }

  const headers = getHeaderList(sheet);
  const requiredHeaders = [
    'Timestamp', 'Full Name', 'Mobile', 'WhatsApp', 'Email', 'Age', 'Gender', 'District', 'Constituency', 'Town/City', 'Area/Ward',
    'Skills', 'Food Type', 'Preferred Activity', 'Additional Information', 'PassID', 'Permission_Status', 'Attendance'
  ];

  if (headers.length === 0) {
    sheet.appendRow(requiredHeaders);
  } else {
    for (const name of requiredHeaders) {
      const exists = headers.some(function(header) {
        return normalizeHeaderName(header) === normalizeHeaderName(name);
      });

      if (!exists) {
        const lastColumn = sheet.getLastColumn();
        sheet.getRange(1, lastColumn + 1).setValue(name);
      }
    }
  }

  const finalHeaders = getHeaderList(sheet);
  const rowMap = {};
  for (let i = 0; i < finalHeaders.length; i++) {
    rowMap[normalizeHeaderName(finalHeaders[i])] = i;
  }

  const row = [];
  for (let i = 0; i < finalHeaders.length; i++) {
    row.push('');
  }

  const values = {
    'Timestamp': timestamp,
    'Full Name': fullName,
    'Mobile': mobile,
    'WhatsApp': whatsapp,
    'Email': email,
    'Age': age,
    'Gender': gender,
    'District': district,
    'Constituency': constituency,
    'Town/City': townCity,
    'Area/Ward': areaWard,
    'Skills': skills,
    'Food Type': foodType,
    'Preferred Activity': preferredActivity,
    'Additional Information': additionalInfo,
    'PassID': passId,
    'Permission_Status': permissionStatus,
    'Attendance': attendanceStatus
  };

  Object.keys(values).forEach(function(header) {
    const normalizedHeader = normalizeHeaderName(header);
    if (rowMap[normalizedHeader] !== undefined) {
      row[rowMap[normalizedHeader]] = values[header];
    }
  });

  sheet.appendRow(row);

  return {
    success: true,
    status: 'submitted',
    message: 'Application submitted successfully.',
    data: {
      fullName: fullName,
      mobile: mobile,
      whatsapp: whatsapp,
      passId: passId,
      permissionStatus: permissionStatus,
      timestamp: timestamp
    }
  };
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

  const values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) {
    return {
      success: false,
      status: 'not_found',
      error: 'Pass ID not found.'
    };
  }

  const headers = values[0].map(function(header) {
    return String(header || '').trim();
  });

  const passIdColumnIndex = findHeaderIndex(headers, ['PassID', 'Pass ID', 'Pass_Id', 'passId', 'pass_id', 'passid'])
  let rowIndex = -1;

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const cellValue = passIdColumnIndex === -1 ? '' : row[passIdColumnIndex];
    const rowPassId = ensurePassId(cellValue);
    const matchedByRaw = row.some(function(cell) {
      return normalizeValue(cell) === normalizeValue(targetPassId);
    });

    if (matchedByRaw || normalizeValue(rowPassId) === normalizeValue(targetPassId)) {
      rowIndex = i;
      break;
    }
  }

  if (rowIndex === -1) {
    return {
      success: false,
      status: 'not_found',
      error: 'Pass ID not found.'
    };
  }

  const attendanceCol = findHeaderIndex(headers, ['Attendance', 'Attendance_Status', 'Attendance Status', 'attendance', 'attendance_status', 'Attended', 'Status']);
  const timeCol = findHeaderIndex(headers, ['Marked At', 'Attendance Time', 'Attended At', 'Updated At', 'Timestamp', 'Date Marked']);

  const finalAttendanceCol = attendanceCol === -1 ? headers.length : attendanceCol;
  if (attendanceCol === -1) {
    sheet.getRange(1, finalAttendanceCol + 1).setValue('Attendance');
  }

  sheet.getRange(rowIndex + 1, finalAttendanceCol + 1).setValue(true);

  if (timeCol === -1) {
    sheet.getRange(1, finalAttendanceCol + 2).setValue('Marked At');
    sheet.getRange(rowIndex + 1, finalAttendanceCol + 2).setValue(new Date().toISOString());
  } else {
    sheet.getRange(rowIndex + 1, timeCol + 1).setValue(new Date().toISOString());
  }

  return {
    success: true,
    status: 'updated',
    message: 'Attendance marked successfully.',
    passId: targetPassId
  };
}

function setScriptProperties() {
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', DEFAULT_SPREADSHEET_ID);
  PropertiesService.getScriptProperties().setProperty('SHEET_NAME', DEFAULT_SHEET_NAME);
}
