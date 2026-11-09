const EXTERNAL_SUMMARY_CONFIG = {
  SPREADSHEET_ID: '1gF1SXav2UYAfygX9ZHR2dTteU5_Pawuoev-5MXtQmqA',
  TAB_NAME: 'Daily_Summary_Records'
};

/**
 * HURIX DIGITAL - ATTENDANCE & PRODUCTIVITY HUB
 * Complete Backend Server Architecture (Code.gs)
 */

const SHEET_NAMES = {
  ATTENDANCE: 'Attendance_Shift_Logs',
  PROJECT_LOGS: 'Project_Task_Logs',
  USER_MASTER: 'User_Master',
  CLIENT_MASTER: 'Client_Project_Master',
  BREAK_LOGS: 'Break_Logs'
};

const PROJECT_LOGS_COLUMNS = [
  'task_log_id', 'log_id', 'emp_id', 'email', 'active_client', 'active_project',
  'task_status', 'completion_pct', 'work_start_time', 'work_end_time', 'task_count',
  'aht_benchmark_mins', 'actual_worked_hours', 'productive_hours', 'speed_efficiency_pct', 'performance_leakage_hours', 'user_remarks', 'created_at',
  'activity_type', 'idle_hours', 'bench_hours', 'blocked_hours'
];

/**
 * HELPER: Auto-upgrades Project_Task_Logs header row if columns are missing
 */
function ensureProjectLogsHeaders(sheet) {
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow === 0) {
    sheet.appendRow(PROJECT_LOGS_COLUMNS);
    return;
  }
  var lastCol = sheet.getLastColumn();
  var headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  var headerLowerMap = headers.map(function(h) { return String(h).trim().toLowerCase(); });
  var missingCols = [];
  PROJECT_LOGS_COLUMNS.forEach(function(colName) {
    if (headerLowerMap.indexOf(colName.toLowerCase()) === -1) {
      missingCols.push(colName);
    }
  });
  if (missingCols.length > 0) {
    var curCols = sheet.getLastColumn();
    missingCols.forEach(function(colName, idx) {
      sheet.getRange(1, curCols + idx + 1).setValue(colName);
    });
  }
  if (sheet.getMaxColumns() < PROJECT_LOGS_COLUMNS.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), PROJECT_LOGS_COLUMNS.length - sheet.getMaxColumns());
  }
}

/**
 * HELPER: Normalizes any date value (Date object, ISO string, DD-MM-YYYY, DD/MM/YYYY, YYYY/MM/DD) into 'yyyy-MM-dd' ISO string format.
 */
function normalizeDateStr(raw, timeZone) {
  if (!raw) return '';
  timeZone = timeZone || Session.getScriptTimeZone();
  if (raw instanceof Date) {
    return Utilities.formatDate(raw, timeZone, 'yyyy-MM-dd');
  }
  var s = String(raw).trim();
  if (!s) return '';

  // Case 1: YYYY-MM-DD or YYYY/MM/DD (e.g., "2026-08-04" or "2026/08/04 10:30:00")
  var ymdMatch = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (ymdMatch) {
    var y = ymdMatch[1];
    var m = ymdMatch[2].length === 1 ? '0' + ymdMatch[2] : ymdMatch[2];
    var d = ymdMatch[3].length === 1 ? '0' + ymdMatch[3] : ymdMatch[3];
    return y + '-' + m + '-' + d;
  }

  // Case 2: DD-MM-YYYY or DD/MM/YYYY (e.g., "04-08-2026" or "4/8/2026")
  var dmyMatch = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);
  if (dmyMatch) {
    var d = dmyMatch[1].length === 1 ? '0' + dmyMatch[1] : dmyMatch[1];
    var m = dmyMatch[2].length === 1 ? '0' + dmyMatch[2] : dmyMatch[2];
    var y = dmyMatch[3];
    return y + '-' + m + '-' + d;
  }

  // Case 3: Standard JS Date string parse fallback
  try {
    var dt = new Date(s);
    if (!isNaN(dt.getTime())) {
      return Utilities.formatDate(dt, timeZone, 'yyyy-MM-dd');
    }
  } catch(e) {}

  return s.length >= 10 ? s.substring(0, 10) : '';
}

/**
 * EXACT WORK ACTIVITIES REQUIRED ACROSS SYSTEM (Includes Support)
 */
const ALLOWED_ACTIVITIES = [
  'Production',
  'Support',
  'Client Meeting',
  'Internal Meeting',
  'Training (Material Go-Through)',
  'Idle',
  'Down Time'
];

/**
 * Normalizes any activity string to one of the allowed activities
 */
function normalizeActivityType(raw) {
  if (!raw) return 'Production';
  var s = String(raw).trim();
  for (var i = 0; i < ALLOWED_ACTIVITIES.length; i++) {
    if (ALLOWED_ACTIVITIES[i].toLowerCase() === s.toLowerCase()) {
      return ALLOWED_ACTIVITIES[i];
    }
  }
  var lower = s.toLowerCase();
  if (lower.indexOf('support') !== -1) return 'Support';
  if (lower.indexOf('client meet') !== -1) return 'Client Meeting';
  if (lower.indexOf('internal meet') !== -1 || lower.indexOf('meeting') !== -1) return 'Internal Meeting';
  if (lower.indexOf('train') !== -1 || lower.indexOf('upskill') !== -1 || lower.indexOf('material') !== -1 || lower.indexOf('bench') !== -1) return 'Training (Material Go-Through)';
  if (lower.indexOf('idle') !== -1 || lower.indexOf('waiting') !== -1 || lower.indexOf('allocation') !== -1) return 'Idle';
  if (lower.indexOf('downtime') !== -1 || lower.indexOf('down time') !== -1) return 'Down Time';
  return 'Production';
}

/**
 * Calculates start and end dates and label for day, week, and month periods
 */
function getDateRangeForPeriod(targetDateStr, period, timeZone) {
  timeZone = timeZone || Session.getScriptTimeZone();
  var normalized = normalizeDateStr(targetDateStr, timeZone) || Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd');
  var parts = normalized.split('-');
  var y = parseInt(parts[0], 10);
  var m = parseInt(parts[1], 10) - 1;
  var d = parseInt(parts[2], 10);
  var dt = new Date(y, m, d);

  if (period === 'week') {
    var dayOfWeek = dt.getDay(); // 0 is Sunday, 1 is Monday
    var diffToMonday = (dayOfWeek === 0 ? -6 : 1) - dayOfWeek;
    var monday = new Date(y, m, d + diffToMonday);
    var sunday = new Date(y, m, d + diffToMonday + 6);
    return {
      startDate: Utilities.formatDate(monday, timeZone, 'yyyy-MM-dd'),
      endDate: Utilities.formatDate(sunday, timeZone, 'yyyy-MM-dd'),
      label: 'Week: ' + Utilities.formatDate(monday, timeZone, 'dd MMM') + ' - ' + Utilities.formatDate(sunday, timeZone, 'dd MMM yyyy')
    };
  } else if (period === 'month') {
    var firstDay = new Date(y, m, 1);
    var lastDay = new Date(y, m + 1, 0);
    return {
      startDate: Utilities.formatDate(firstDay, timeZone, 'yyyy-MM-dd'),
      endDate: Utilities.formatDate(lastDay, timeZone, 'yyyy-MM-dd'),
      label: Utilities.formatDate(firstDay, timeZone, 'MMMM yyyy')
    };
  } else {
    // 'day'
    return {
      startDate: normalized,
      endDate: normalized,
      label: Utilities.formatDate(dt, timeZone, 'EEE, dd MMM yyyy')
    };
  }
}

/**
 * Returns the active Spreadsheet object
 */
function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Serves the HTML Web App interface
 */
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  template.userRole = 'User';
  template.userEmail = '';
  template.empName = '';
  template.empId = '';

  return template.evaluate()
    .setTitle('Hurix Digital - Attendance & Productivity Hub')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Helper to include HTML sub-files cleanly
 */
function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (err) {
    Logger.log('Include file missing: ' + filename + '. Error: ' + err.toString());
    return '<!-- Sub-file ' + filename + ' not found in Apps Script project -->';
  }
}

/**
 * USER LOOKUP & AUTHENTICATION ENGINES
 */
function getUserByEmail(email) {
  if (!email) return null;
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.USER_MASTER);
  if (!sheet) return null;

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return null;

  var headers = data[0];
  var emailIdx = -1;
  for (var h = 0; h < headers.length; h++) {
    if (String(headers[h]).trim().toLowerCase() === 'email') emailIdx = h;
  }
  if (emailIdx === -1) return null;

  for (var i = 1; i < data.length; i++) {
    var sheetEmail = String(data[i][emailIdx]).trim().toLowerCase();
    var currentEmail = String(email).trim().toLowerCase();

    if (sheetEmail === currentEmail) {
      var userObj = {};
      headers.forEach(function(h, idx) {
        var cleanHeader = String(h).trim();
        var val = data[i][idx];
        if (typeof val === 'string') val = val.trim();
        userObj[cleanHeader] = val;
        var lowerKey = cleanHeader.toLowerCase();
        userObj[lowerKey] = val;
        userObj[lowerKey.replace(/\s+/g, '_')] = val;
      });
      return userObj;
    }
  }
  return null;
}

/**
 * REFRESH USER PROFILE DIRECTLY FROM USER_MASTER TAB
 */
function refreshUserProfile(identifier) {
  if (!identifier) return { success: false, notFound: true, message: 'No user identifier provided.' };
  var user = getUserByEmail(identifier) || getUserByEmpId(identifier);
  if (user) {
    var activeVal = String(user.is_active || 'TRUE').trim().toUpperCase();
    var isActive = (activeVal === 'TRUE' || activeVal === '1' || activeVal === 'ACTIVE' || activeVal === 'YES' || user.is_active === true);
    if (!isActive) {
      return { success: false, inactive: true, message: 'Your account has been deactivated. Contact Admin.' };
    }
    return { success: true, user: user };
  }
  return { success: false, notFound: true, message: 'Your account has been removed from User_Master. Access revoked.' };
}


/**
 * INITIAL APP DATA LOADER
 */
function getInitialAppData() {
  var activeEmail = Session.getActiveUser().getEmail();
  var user = null;

  if (activeEmail) {
    user = getUserByEmail(activeEmail);
  }

  return {
    success: true,
    user: user,
    activeEmail: activeEmail,
    clientMaster: getClientMasterData()
  };
}

/**
 * AUTHENTICATION: Email & Password
 */
function authenticateUserWithPassword(email, password) {
  if (!email || !password) {
    return { success: false, message: 'Please enter both Email and Password.' };
  }

  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.USER_MASTER);
  if (!sheet) return { success: false, message: 'User database not initialized.' };

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { success: false, message: 'No registered employees found.' };

  var headers = data[0];
  var emailIdx = -1;
  var passIdx = -1;
  var activeIdx = -1;

  for (var h = 0; h < headers.length; h++) {
    var hn = String(headers[h]).trim().toLowerCase();
    if (hn === 'email') emailIdx = h;
    if (hn === 'password') passIdx = h;
    if (hn === 'is_active') activeIdx = h;
  }

  if (emailIdx === -1 || passIdx === -1) {
    return { success: false, message: 'User database headers corrupted.' };
  }

  var cleanInputEmail = String(email).trim().toLowerCase();
  var cleanInputPass = String(password).trim();

  for (var i = 1; i < data.length; i++) {
    var rowEmail = String(data[i][emailIdx]).trim().toLowerCase();
    var rowPass = String(data[i][passIdx]).trim();

    if (rowEmail === cleanInputEmail) {
      if (activeIdx !== -1) {
        var activeVal = String(data[i][activeIdx] || 'TRUE').trim().toUpperCase();
        var isActive = (activeVal === 'TRUE' || activeVal === '1' || activeVal === 'ACTIVE' || activeVal === 'YES' || data[i][activeIdx] === true);
        if (!isActive) {
          return { success: false, message: 'Access Denied: Your account is deactivated. Contact Administrator.' };
        }
      }

      if (rowPass === cleanInputPass) {
        var userObj = {};
        headers.forEach(function(h, idx) {
          var cleanHeader = String(h).trim();
          var val = data[i][idx];
          if (typeof val === 'string') val = val.trim();
          userObj[cleanHeader] = val;
          var lowerKey = cleanHeader.toLowerCase();
          userObj[lowerKey] = val;
          userObj[lowerKey.replace(/\s+/g, '_')] = val;
        });

        return {
          success: true,
          user: userObj,
          message: 'Authentication successful.'
        };
      } else {
        return { success: false, message: 'Invalid password. Please verify and try again.' };
      }
    }
  }

  return { success: false, message: 'Email address not found in system directory.' };
}

/**
 * CLIENT & PROJECT MASTER DATA ENGINE
 */
function getClientMasterData() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.CLIENT_MASTER);
  var list = [];
  if (!sheet) return list;

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return list;

  var headers = data[0];
  var activeIdx = headers.indexOf('is_active');

  for (var i = 1; i < data.length; i++) {
    if (activeIdx !== -1) {
      var actVal = String(data[i][activeIdx] || 'TRUE').trim().toUpperCase();
      if (actVal !== 'TRUE' && actVal !== '1' && actVal !== 'ACTIVE' && actVal !== 'YES' && data[i][activeIdx] !== true) continue;
    }

    var obj = {};
    headers.forEach(function(h, idx) {
      var val = data[i][idx];
      if (typeof val === 'string') val = val.trim();
      obj[h] = val;
    });

    if (obj.client_name && obj.project_name) {
      list.push(obj);
    }
  }
  return list;
}

/**
 * SHIFT DEFINITIONS & PRODUCTION HOUR RULES
 * General: 8 hours
 * 1st, Afternoon, 2nd: 7 hours
 */
var SHIFT_TYPES = {
  GENERAL: 'General',
  SHIFT_1ST: '1st',
  AFTERNOON: 'Afternoon',
  SHIFT_2ND: '2nd'
};

var ALLOWED_SHIFTS = ['General', '1st', 'Afternoon', '2nd'];

function normalizeShiftType(shift) {
  if (!shift) return 'General';
  var s = String(shift).trim().toLowerCase();
  if (s === '1st' || s === '1' || s.indexOf('1st') !== -1) return '1st';
  if (s === 'afternoon' || s.indexOf('afternoon') !== -1) return 'Afternoon';
  if (s === '2nd' || s === '2' || s.indexOf('2nd') !== -1) return '2nd';
  return 'General';
}

function getShiftStandardHours(shiftType) {
  var s = normalizeShiftType(shiftType);
  if (s === '1st' || s === 'Afternoon' || s === '2nd') return 7.0;
  return 8.0; // General is 8 hours
}

function ensureAttendanceHeaders(sheet) {
  if (!sheet) return [];
  var headers = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
  var needed = ['shift_type', 'leave_hours'];
  var added = false;
  needed.forEach(function(col) {
    if (headers.indexOf(col) === -1) {
      sheet.getRange(1, headers.length + 1).setValue(col);
      headers.push(col);
      added = true;
    }
  });
  if (added) SpreadsheetApp.flush();
  return headers;
}

/**
 * HELPER: Fetch Today's Shift State for Employee
 */
function getActiveUserShiftState(empId) {
  if (!empId) return null;
  autoClosePreviousOpenShiftsForUser(empId);
  var ss = getSpreadsheet();
  var timeZone = Session.getScriptTimeZone();
  var todayStr = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd');
  var attSheet = ss.getSheetByName(SHEET_NAMES.ATTENDANCE);
  if (!attSheet) return null;

  ensureAttendanceHeaders(attSheet);
  var aData = attSheet.getDataRange().getValues();
  if (aData.length <= 1) return null;

  var aHeaders = aData[0];
  var empIdx = -1, dateIdx = -1, emailIdx = -1;

  for (var h = 0; h < aHeaders.length; h++) {
    var hn = String(aHeaders[h]).trim().toLowerCase().replace(/_/g, '');
    if (hn === 'empid' || hn === 'employeeid') empIdx = h;
    if (hn === 'date') dateIdx = h;
    if (hn === 'email') emailIdx = h;
  }

  if (empIdx === -1 && emailIdx === -1) return null;
  if (dateIdx === -1) return null;

  var userCanonical = getUserByEmpId(empId) || getUserByEmail(empId);
  var targetEmpId = (userCanonical && userCanonical.emp_id ? userCanonical.emp_id : empId).toString().trim().toLowerCase();
  var targetEmail = (userCanonical && userCanonical.email ? userCanonical.email : empId).toString().trim().toLowerCase();

  for (var j = aData.length - 1; j >= 1; j--) {
    var rowDate = normalizeDateStr(aData[j][dateIdx], timeZone);
    var rowEmp = empIdx !== -1 ? String(aData[j][empIdx]).trim().toLowerCase() : '';
    var rowEmail = emailIdx !== -1 ? String(aData[j][emailIdx]).trim().toLowerCase() : '';

    var matchEmp = (rowEmp === targetEmpId || (rowEmail && rowEmail === targetEmail) || (rowEmp && rowEmp === targetEmail));

    if (matchEmp && rowDate === todayStr) {
      var shiftObj = {};
      aHeaders.forEach(function(h, idx) {
        var val = aData[j][idx];
        if (val instanceof Date) {
          if (h === 'date' || h === 'created_at') {
            shiftObj[h] = Utilities.formatDate(val, timeZone, 'yyyy-MM-dd');
          } else {
            shiftObj[h] = Utilities.formatDate(val, timeZone, 'HH:mm:ss');
          }
        } else {
          shiftObj[h] = val;
        }
      });

      // Fixed 1-hour daily break applied automatically for every user
      shiftObj.total_break_minutes = 60;
      shiftObj.shift_type = normalizeShiftType(shiftObj.shift_type || 'General');
      shiftObj.leave_hours = Number(shiftObj.leave_hours) || 0;
      return shiftObj;
    }
  }
  return null;
}

/**
 * FETCH USER TODAY TASKS
 */
function getUserTodayTasks(empId, logId) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.PROJECT_LOGS);
  var tasks = [];
  if (!sheet) return tasks;

  ensureProjectLogsHeaders(sheet);

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return tasks;

  var headers = data[0];
  var empIdx = -1, emailIdx = -1, logIdIdx = -1, createdAtIdx = -1, wsIdx = -1;

  for (var h = 0; h < headers.length; h++) {
    var hn = String(headers[h]).trim().toLowerCase().replace(/_/g, '');
    if (hn === 'empid' || hn === 'employeeid') empIdx = h;
    if (hn === 'email') emailIdx = h;
    if (hn === 'logid') logIdIdx = h;
    if (hn === 'createdat') createdAtIdx = h;
    if (hn === 'workstarttime' || hn === 'starttime') wsIdx = h;
  }

  var timeZone = Session.getScriptTimeZone();
  var todayStr = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd');

  var userCanonical = getUserByEmpId(empId) || getUserByEmail(empId);
  var targetEmpId = (userCanonical && userCanonical.emp_id ? userCanonical.emp_id : empId || '').toString().trim().toLowerCase();
  var targetEmail = (userCanonical && userCanonical.email ? userCanonical.email : empId || '').toString().trim().toLowerCase();

  for (var i = 1; i < data.length; i++) {
    var rowEmpId = empIdx !== -1 ? String(data[i][empIdx]).trim().toLowerCase() : '';
    var rowEmail = emailIdx !== -1 ? String(data[i][emailIdx]).trim().toLowerCase() : '';
    var rowLogId = logIdIdx !== -1 ? String(data[i][logIdIdx]).trim() : '';

    var matchEmp = (targetEmpId && rowEmpId === targetEmpId) || (targetEmail && rowEmail === targetEmail) || (targetEmail && rowEmpId === targetEmail);
    var matchLog = logId && rowLogId && rowLogId === String(logId).trim();

    var rawDateVal = wsIdx !== -1 ? data[i][wsIdx] : (createdAtIdx !== -1 ? data[i][createdAtIdx] : '');
    var rowDateStr = normalizeDateStr(rawDateVal, timeZone);
    if (!rowDateStr && createdAtIdx !== -1) {
      rowDateStr = normalizeDateStr(data[i][createdAtIdx], timeZone);
    }

    var isTodayTask = matchEmp && (!rowDateStr || rowDateStr === todayStr);

    if ((matchEmp && matchLog) || isTodayTask) {
      var task = {};
      for (var col = 0; col < headers.length; col++) {
        var hName = headers[col];
        var val = data[i][col];
        if (val instanceof Date) {
          task[hName] = Utilities.formatDate(val, timeZone, 'yyyy-MM-dd HH:mm:ss');
        } else {
          task[hName] = val !== null && val !== undefined ? val : '';
        }
      }

      task.activity_type = normalizeActivityType(task.activity_type || data[i][18]);
      task.actual_worked_hours = Number(task.actual_worked_hours) || 0;

      tasks.push(task);
    }
  }
  return tasks;
}

/**
 * ATTENDANCE & SHIFT CLOCK-IN ENGINE
 */
function clockInUser(payload) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) {
    return { success: false, message: 'Server busy. Please try again.' };
  }

  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAMES.ATTENDANCE);
    var email = payload.email || Session.getActiveUser().getEmail();
    var user = getUserByEmail(email) || getUserByEmpId(payload.emp_id);
    if (!user) {
      return { success: false, message: 'Access Denied: Your account has been removed from User_Master. Contact Admin.' };
    }
    var activeVal = String(user.is_active || 'TRUE').trim().toUpperCase();
    var isActive = (activeVal === 'TRUE' || activeVal === '1' || activeVal === 'ACTIVE' || activeVal === 'YES' || user.is_active === true);
    if (!isActive) {
      return { success: false, message: 'Access Denied: Your account is currently inactive. Contact Admin.' };
    }

    autoClosePreviousOpenShiftsForUser(user.emp_id);

    ensureAttendanceHeaders(sheet);
    var timeZone = Session.getScriptTimeZone();
    var todayStr = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd');
    var data = sheet.getDataRange().getValues();
    var headers = data[0];

    var empIdx = headers.indexOf('emp_id');
    var dateIdx = headers.indexOf('date');
    var statusIdx = headers.indexOf('shift_status');
    var startIdx = headers.indexOf('shift_start_time');
    var endIdx = headers.indexOf('shift_end_time');
    var shiftTypeIdx = headers.indexOf('shift_type');

    var shiftType = normalizeShiftType(payload.shift_type || payload.shift || user.default_shift || 'General');
    var userEmpIdLower = String(user.emp_id || '').trim().toLowerCase();

    for (var i = 1; i < data.length; i++) {
      var rowDate = normalizeDateStr(data[i][dateIdx], timeZone);
      var rowEmp = empIdx !== -1 ? String(data[i][empIdx]).trim().toLowerCase() : '';

      if (rowEmp === userEmpIdLower && rowDate === todayStr) {
        var statusVal = String(data[i][statusIdx]).toUpperCase();
        var rowIndex = i + 1;

        if (statusVal === 'CLOSED' || statusVal === 'LEAVE_CLOSED' || statusVal === 'AUTO_CLOSED' || statusVal === 'NOT_CLOCKED_IN') {
          var rowValues = data[i];
          rowValues[statusIdx] = 'OPEN';
          if (endIdx !== -1) rowValues[endIdx] = '';
          
          var existingStart = data[i][startIdx];
          var startTimeVal = existingStart ? (existingStart instanceof Date ? Utilities.formatDate(existingStart, timeZone, 'HH:mm:ss') : String(existingStart)) : Utilities.formatDate(new Date(), timeZone, 'HH:mm:ss');
          if (!existingStart && startIdx !== -1) {
            rowValues[startIdx] = startTimeVal;
          }

          var permHrsIdx = headers.indexOf('permission_hours');
          var existingPermHrs = permHrsIdx !== -1 ? (Number(data[i][permHrsIdx]) || 0) : 0;
          var attIdx = headers.indexOf('attendance_status');
          var attStatus = attIdx !== -1 ? String(data[i][attIdx]) : 'Present';
          if (attIdx !== -1 && (attStatus === 'NOT_CLOCKED_IN' || !attStatus)) {
            rowValues[attIdx] = 'Present';
            attStatus = 'Present';
          }

          if (shiftTypeIdx !== -1 && (payload.shift_type || payload.shift)) {
            rowValues[shiftTypeIdx] = shiftType;
          }

          sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
          SpreadsheetApp.flush();
          syncExternalDailySummaryRecords(todayStr, user.emp_id);

          var activeShiftName = (shiftTypeIdx !== -1 && rowValues[shiftTypeIdx]) ? rowValues[shiftTypeIdx] : shiftType;
          return {
            success: true,
            log_id: data[i][headers.indexOf('log_id')],
            shift_start_time: startTimeVal,
            shift_status: 'OPEN',
            shift_type: activeShiftName,
            permission_hours: existingPermHrs,
            attendance_status: attStatus,
            total_break_minutes: 60
          };
        }

        var currentShiftName = (shiftTypeIdx !== -1 && data[i][shiftTypeIdx]) ? data[i][shiftTypeIdx] : shiftType;
        return {
          success: true,
          already_open: true,
          log_id: data[i][headers.indexOf('log_id')],
          shift_start_time: data[i][startIdx] instanceof Date ? Utilities.formatDate(data[i][startIdx], timeZone, 'HH:mm:ss') : String(data[i][startIdx]),
          shift_status: statusVal,
          shift_type: currentShiftName,
          total_break_minutes: 60
        };
      }
    }

    var now = new Date();
    var timeStr = Utilities.formatDate(now, timeZone, 'HH:mm:ss');
    var cleanEmpId = String(user.emp_id || '').trim().replace(/[^a-zA-Z0-9]/g, '');
    var logId = 'ATT-' + todayStr.replace(/-/g, '') + (cleanEmpId ? '-' + cleanEmpId : '') + '-' + Math.floor(100000 + Math.random() * 900000);

    var newRow = new Array(headers.length);
    for (var h = 0; h < headers.length; h++) newRow[h] = '';

    if (headers.indexOf('log_id') !== -1) newRow[headers.indexOf('log_id')] = logId;
    if (headers.indexOf('date') !== -1) newRow[headers.indexOf('date')] = todayStr;
    if (headers.indexOf('emp_id') !== -1) newRow[headers.indexOf('emp_id')] = user.emp_id;
    if (headers.indexOf('emp_name') !== -1) newRow[headers.indexOf('emp_name')] = user.emp_name;
    if (headers.indexOf('attendance_status') !== -1) newRow[headers.indexOf('attendance_status')] = 'Present';
    if (headers.indexOf('permission_hours') !== -1) newRow[headers.indexOf('permission_hours')] = 0;
    if (headers.indexOf('shift_start_time') !== -1) newRow[headers.indexOf('shift_start_time')] = timeStr;
    if (headers.indexOf('shift_end_time') !== -1) newRow[headers.indexOf('shift_end_time')] = '';
    if (headers.indexOf('total_shift_hours') !== -1) newRow[headers.indexOf('total_shift_hours')] = 0;
    if (headers.indexOf('shrinkage_hours') !== -1) newRow[headers.indexOf('shrinkage_hours')] = 0;
    if (headers.indexOf('proof_url') !== -1) newRow[headers.indexOf('proof_url')] = '';
    if (headers.indexOf('shift_status') !== -1) newRow[headers.indexOf('shift_status')] = 'OPEN';
    if (headers.indexOf('created_at') !== -1) newRow[headers.indexOf('created_at')] = new Date();
    if (headers.indexOf('break_start_time') !== -1) newRow[headers.indexOf('break_start_time')] = '';
    if (headers.indexOf('total_break_minutes') !== -1) newRow[headers.indexOf('total_break_minutes')] = 60;
    if (headers.indexOf('active_client') !== -1) newRow[headers.indexOf('active_client')] = user.default_client || 'iMerit';
    if (headers.indexOf('active_project') !== -1) newRow[headers.indexOf('active_project')] = user.default_project || 'Pravah';
    if (headers.indexOf('shift_type') !== -1) newRow[headers.indexOf('shift_type')] = shiftType;
    if (headers.indexOf('leave_hours') !== -1) newRow[headers.indexOf('leave_hours')] = 0;

    sheet.appendRow(newRow);
    SpreadsheetApp.flush();
    syncExternalDailySummaryRecords(todayStr, user.emp_id);
    return { success: true, log_id: logId, shift_start_time: timeStr, shift_status: 'OPEN', shift_type: shiftType, total_break_minutes: 60 };
  } catch (err) {
    return { success: false, message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * SAVE / UPDATE USER'S SHIFT SELECTION FOR TODAY (General, 1st, Afternoon, 2nd)
 */
function saveUserTodayShift(payload) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) {
    return { success: false, message: 'Server busy. Please try again.' };
  }

  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAMES.ATTENDANCE);
    if (!sheet) return { success: false, message: 'Attendance sheet not found.' };

    var user = getUserByEmpId(payload.emp_id) || getUserByEmail(payload.email || Session.getActiveUser().getEmail());
    if (!user) return { success: false, message: 'User not found in User_Master.' };

    var timeZone = Session.getScriptTimeZone();
    var todayStr = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd');
    var shiftType = normalizeShiftType(payload.shift_type || payload.shift || 'General');

    ensureAttendanceHeaders(sheet);
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var empIdx = headers.indexOf('emp_id');
    var dateIdx = headers.indexOf('date');
    var shiftTypeIdx = headers.indexOf('shift_type');

    var userEmpLower = String(user.emp_id || '').trim().toLowerCase();
    var found = false;

    for (var i = 1; i < data.length; i++) {
      var rDate = normalizeDateStr(data[i][dateIdx], timeZone);
      var rEmp = empIdx !== -1 ? String(data[i][empIdx]).trim().toLowerCase() : '';
      if (rEmp === userEmpLower && rDate === todayStr) {
        if (shiftTypeIdx !== -1) {
          sheet.getRange(i + 1, shiftTypeIdx + 1).setValue(shiftType);
        }
        found = true;
        break;
      }
    }

    if (!found) {
      var now = new Date();
      var cleanEmpId = String(user.emp_id || '').trim().replace(/[^a-zA-Z0-9]/g, '');
      var logId = 'ATT-' + todayStr.replace(/-/g, '') + (cleanEmpId ? '-' + cleanEmpId : '') + '-' + Math.floor(100000 + Math.random() * 900000);
      var newRow = new Array(headers.length);
      for (var h = 0; h < headers.length; h++) newRow[h] = '';
      if (headers.indexOf('log_id') !== -1) newRow[headers.indexOf('log_id')] = logId;
      if (headers.indexOf('date') !== -1) newRow[headers.indexOf('date')] = todayStr;
      if (headers.indexOf('emp_id') !== -1) newRow[headers.indexOf('emp_id')] = user.emp_id;
      if (headers.indexOf('emp_name') !== -1) newRow[headers.indexOf('emp_name')] = user.emp_name;
      if (headers.indexOf('attendance_status') !== -1) newRow[headers.indexOf('attendance_status')] = 'NOT_CLOCKED_IN';
      if (headers.indexOf('shift_status') !== -1) newRow[headers.indexOf('shift_status')] = 'NOT_CLOCKED_IN';
      if (headers.indexOf('total_break_minutes') !== -1) newRow[headers.indexOf('total_break_minutes')] = 60;
      if (shiftTypeIdx !== -1) newRow[shiftTypeIdx] = shiftType;
      if (headers.indexOf('created_at') !== -1) newRow[headers.indexOf('created_at')] = now;
      sheet.appendRow(newRow);
    }

    SpreadsheetApp.flush();
    return { success: true, shift_type: shiftType };
  } catch (err) {
    return { success: false, message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * HELPER: Parses time strings, datetime-local strings, or Date objects into a valid Date object
 */
function parseTimeStringToDate(timeVal, dateStr, timeZone) {
  if (!timeVal) return null;
  timeZone = timeZone || Session.getScriptTimeZone();
  var normalizedDate = normalizeDateStr(dateStr, timeZone) || Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd');
  var dParts = normalizedDate.split('-');
  if (dParts.length !== 3) return null;
  var y = parseInt(dParts[0], 10);
  var m = parseInt(dParts[1], 10) - 1;
  var d = parseInt(dParts[2], 10);

  if (timeVal instanceof Date) {
    if (!isNaN(timeVal.getTime())) {
      var hh = timeVal.getHours();
      var mm = timeVal.getMinutes();
      var ss = timeVal.getSeconds();
      return new Date(y, m, d, hh, mm, ss);
    }
    return null;
  }

  var str = String(timeVal).trim();
  if (!str) return null;

  // Search for any HH:mm:ss or HH:mm pattern in the string
  var tmMatch = str.match(/(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?/);
  if (tmMatch) {
    var hh = parseInt(tmMatch[1], 10);
    var mm = parseInt(tmMatch[2], 10);
    var ss = tmMatch[3] ? parseInt(tmMatch[3], 10) : 0;
    var dt = new Date(y, m, d, hh, mm, ss);
    if (!isNaN(dt.getTime())) return dt;
  }

  return null;
}

/**
 * PAUSE / RESUME SHIFT (DEPRECATED - Fixed 1-hour break is automatically accounted for)
 */
function toggleShiftPause(payload) {
  return {
    success: true,
    shift_status: 'OPEN',
    total_break_minutes: 60,
    message: 'Break-time logging is removed. A fixed 1-hour break per day is automatically accounted for.'
  };
}

/**
 * HELPER: Calculate and Sync Break Durations
 * Fixed 1-hour (60 minutes) daily break applied automatically.
 */
function calculateAndSyncBreakDurations(ss, logId) {
  return 60;
}

/**
 * HELPER: Fetch Itemized Break Logs for Employee / Log ID
 * Deprecated: itemized break logs are removed.
 */
function getEmpBreakLogs(logId, empId, targetDate) {
  return [];
}

/**
 * CLOCK OUT ENGINE
 */
function clockOutUser(payload) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) {
    return { success: false, message: 'Server busy.' };
  }

  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAMES.ATTENDANCE);
    var data = sheet.getDataRange().getValues();
    var headers = data[0];

    var logIdIdx = headers.indexOf('log_id');
    var rowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][logIdIdx]).trim() === String(payload.log_id).trim()) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) return { success: false, message: 'Active shift log not found.' };


    var now = new Date();
    var timeZone = Session.getScriptTimeZone();
    var endStr = Utilities.formatDate(now, timeZone, 'HH:mm:ss');
    var todayStr = Utilities.formatDate(now, timeZone, 'yyyy-MM-dd');
    var startIdx = headers.indexOf('shift_start_time');
    var startRaw = data[rowIndex - 1][startIdx];
    
    var startStr = startRaw ? (startRaw instanceof Date ? Utilities.formatDate(startRaw, timeZone, 'HH:mm:ss') : String(startRaw)) : '';
    
    // Default 1-hour break applied automatically
    var totalBreakMins = 60;
    var breakHours = 1.0;

    var grossHours = 0;
    if (startStr) {
      var startTime = parseTimeStringToDate(startStr, todayStr, timeZone);
      if (startTime && !isNaN(startTime.getTime())) {
        grossHours = Math.max(0, ((now.getTime() - startTime.getTime()) / (1000 * 60 * 60)));
      }
    }

    var netWorkedHours = Math.max(0, grossHours - breakHours).toFixed(2);
    
    var permHrsInput = Number(payload.permission_hours);
    var existingPermIdx = headers.indexOf('permission_hours');
    var existingPermHrs = (existingPermIdx !== -1 && data[rowIndex - 1][existingPermIdx] !== '') ? (Number(data[rowIndex - 1][existingPermIdx]) || 0) : 0;
    var permHrs = (!isNaN(permHrsInput) && permHrsInput > 0) ? permHrsInput : existingPermHrs;

    var rowValues = data[rowIndex - 1];
    if (headers.indexOf('total_break_minutes') !== -1) rowValues[headers.indexOf('total_break_minutes')] = totalBreakMins;
    if (headers.indexOf('shift_end_time') !== -1) rowValues[headers.indexOf('shift_end_time')] = endStr;
    if (headers.indexOf('total_shift_hours') !== -1) rowValues[headers.indexOf('total_shift_hours')] = Number(netWorkedHours);
    if (headers.indexOf('permission_hours') !== -1) rowValues[headers.indexOf('permission_hours')] = permHrs;
    if (headers.indexOf('shrinkage_hours') !== -1) rowValues[headers.indexOf('shrinkage_hours')] = permHrs + breakHours;
    if (payload.proof_url && headers.indexOf('proof_url') !== -1) rowValues[headers.indexOf('proof_url')] = payload.proof_url;
    if (headers.indexOf('shift_status') !== -1) rowValues[headers.indexOf('shift_status')] = 'CLOSED';

    sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
    SpreadsheetApp.flush();

    var empIdVal = data[rowIndex - 1][headers.indexOf('emp_id')];
    syncExternalDailySummaryRecords(todayStr, empIdVal);
    return {
      success: true,
      total_shift_hours: netWorkedHours,
      shift_start_time: startStr,
      shift_end_time: endStr,
      shift_status: 'CLOSED'
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * LEAVE & PERMISSION APPLICATION ENGINE
 */
function applyLeaveOrPermission(payload) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) {
    return { success: false, message: 'Server busy.' };
  }

  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAMES.ATTENDANCE);
    var email = payload.email || Session.getActiveUser().getEmail();
    var user = getUserByEmail(email) || getUserByEmpId(payload.emp_id);
    if (!user) {
      return { success: false, message: 'Access Denied: Your account has been removed from User_Master. Contact Admin.' };
    }
    var activeVal = String(user.is_active || 'TRUE').trim().toUpperCase();
    var isActive = (activeVal === 'TRUE' || activeVal === '1' || activeVal === 'ACTIVE' || activeVal === 'YES' || user.is_active === true);
    if (!isActive) {
      return { success: false, message: 'Access Denied: Your account is currently inactive. Contact Admin.' };
    }

    var timeZone = Session.getScriptTimeZone();
    var targetDate = normalizeDateStr(payload.date, timeZone) || Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd');
    var data = sheet.getDataRange().getValues();
    var headers = data[0];

    var empIdx = headers.indexOf('emp_id');
    var dateIdx = headers.indexOf('date');

    var rawLeaveCode = String(payload.leave_type || payload.category || 'PL').trim().toUpperCase();
    var leaveCode = rawLeaveCode;
    if (leaveCode === 'WEEK OFF' || leaveCode === 'WEEKOFF') leaveCode = 'WO';
    if (leaveCode === 'HOLIDAY') leaveCode = 'HOLIDAY';
    var isPresentOverride = (leaveCode === 'PRESENT');
    var isScheduleOff = !isPresentOverride && (leaveCode === 'WO' || leaveCode === 'HOLIDAY');
    var isFullDayLeave = !isScheduleOff && !isPresentOverride && (leaveCode === 'PL' || leaveCode === 'UL' || leaveCode === 'ABSENT');
    var isShortPerm = !isScheduleOff && !isPresentOverride && (leaveCode === 'PERM' || leaveCode === 'PERMISSION' || leaveCode === 'OD');
    var permHrs = isShortPerm ? (Number(payload.permission_hours) || 2.0) : 0.0;
    var shiftType = normalizeShiftType(payload.shift_type || payload.shift || user.default_shift || 'General');
    var stdShiftHours = getShiftStandardHours(shiftType);
    var leaveHours = (isScheduleOff || isPresentOverride) ? 0 : (isFullDayLeave ? stdShiftHours : (leaveCode === 'HL' ? (stdShiftHours / 2.0) : permHrs));
    var shrinkageHrs = leaveHours;
    var proofUrl = payload.proof_url || '';
    var clientName = payload.active_client || user.default_client || 'iMerit';
    var projName = payload.active_project || user.default_project || 'Pravah';
    var remarks = payload.remarks || (isPresentOverride ? 'Manager manual shift override: Present' : (isScheduleOff ? (leaveCode === 'WO' ? 'Week Off' : 'Holiday') : ''));

    var userEmpLower = String(user.emp_id || '').trim().toLowerCase();

    // Check if an attendance row already exists for target date
    for (var i = 1; i < data.length; i++) {
      var rowDate = normalizeDateStr(data[i][dateIdx], timeZone);
      var rowEmp = empIdx !== -1 ? String(data[i][empIdx]).trim().toLowerCase() : '';

      if (rowEmp === userEmpLower && rowDate === targetDate) {
        var rowIndex = i + 1;
        var existingStatusIdx = headers.indexOf('shift_status');
        var existingStatus = existingStatusIdx !== -1 ? String(data[i][existingStatusIdx]).toUpperCase() : '';

        var newShiftStatus = isPresentOverride ? 'CLOSED' : (isScheduleOff ? leaveCode : (isFullDayLeave ? 'LEAVE_CLOSED' : (existingStatus && existingStatus !== 'NOT_CLOCKED_IN' ? existingStatus : 'NOT_CLOCKED_IN')));
        var attStatus = isPresentOverride ? 'Present' : (isScheduleOff ? leaveCode : (isFullDayLeave ? leaveCode : (leaveCode === 'HL' ? 'HL' : (existingStatus === 'OPEN' || existingStatus === 'CLOSED' ? 'Present' : leaveCode))));

        var rowValues = data[i];
        if (headers.indexOf('attendance_status') !== -1) rowValues[headers.indexOf('attendance_status')] = attStatus;
        if (headers.indexOf('permission_hours') !== -1) rowValues[headers.indexOf('permission_hours')] = permHrs;
        if (headers.indexOf('shrinkage_hours') !== -1) rowValues[headers.indexOf('shrinkage_hours')] = shrinkageHrs;
        if (headers.indexOf('leave_hours') !== -1) rowValues[headers.indexOf('leave_hours')] = leaveHours;
        if (headers.indexOf('shift_type') !== -1) rowValues[headers.indexOf('shift_type')] = shiftType;
        if (headers.indexOf('proof_url') !== -1 && proofUrl) rowValues[headers.indexOf('proof_url')] = proofUrl;
        if (headers.indexOf('shift_status') !== -1) rowValues[headers.indexOf('shift_status')] = newShiftStatus;
        if (isPresentOverride) {
          if (headers.indexOf('shift_start_time') !== -1) rowValues[headers.indexOf('shift_start_time')] = '09:00:00';
          if (headers.indexOf('shift_end_time') !== -1) rowValues[headers.indexOf('shift_end_time')] = '18:00:00';
          if (headers.indexOf('total_shift_hours') !== -1) rowValues[headers.indexOf('total_shift_hours')] = stdShiftHours;
        } else if (isFullDayLeave || isScheduleOff) {
          if (headers.indexOf('shift_start_time') !== -1) rowValues[headers.indexOf('shift_start_time')] = '';
          if (headers.indexOf('shift_end_time') !== -1) rowValues[headers.indexOf('shift_end_time')] = '';
          if (headers.indexOf('total_shift_hours') !== -1) rowValues[headers.indexOf('total_shift_hours')] = 0;
        }
        if (headers.indexOf('active_client') !== -1 && clientName) rowValues[headers.indexOf('active_client')] = clientName;
        if (headers.indexOf('active_project') !== -1 && projName) rowValues[headers.indexOf('active_project')] = projName;
        if (headers.indexOf('user_remarks') !== -1 && remarks) {
          var existingUserRem = String(data[i][headers.indexOf('user_remarks')] || '').trim();
          var updatedUserRem = existingUserRem ? (existingUserRem + ' | Update: ' + remarks) : remarks;
          rowValues[headers.indexOf('user_remarks')] = updatedUserRem;
        }

        sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
        SpreadsheetApp.flush();
        syncExternalDailySummaryRecords(targetDate, user.emp_id);
        return {
          success: true,
          shift_status: newShiftStatus,
          attendance_status: attStatus,
          permission_hours: permHrs,
          leave_hours: leaveHours,
          shift_type: shiftType,
          message: isPresentOverride ? ('Updated attendance: Present (' + stdShiftHours + ' hrs shift recorded)') : ('Updated attendance record: ' + leaveCode + ' (' + leaveHours + ' hrs leave/permission recorded)')
        };
      }
    }

    // New Row Creation
    var cleanEmpIdLeave = String(user.emp_id || '').trim().replace(/[^a-zA-Z0-9]/g, '');
    var logId = 'ATT-' + targetDate.replace(/-/g, '') + (cleanEmpIdLeave ? '-' + cleanEmpIdLeave : '') + '-' + Math.floor(100000 + Math.random() * 900000);
    var newRow = new Array(headers.length);
    for (var h = 0; h < headers.length; h++) newRow[h] = '';

    var newShiftStatus = isPresentOverride ? 'CLOSED' : (isScheduleOff ? leaveCode : (isFullDayLeave ? 'LEAVE_CLOSED' : 'NOT_CLOCKED_IN'));
    var attStatus = isPresentOverride ? 'Present' : (isScheduleOff ? leaveCode : (isFullDayLeave ? leaveCode : (leaveCode === 'HL' ? 'HL' : 'Present')));

    if (headers.indexOf('log_id') !== -1) newRow[headers.indexOf('log_id')] = logId;
    if (headers.indexOf('date') !== -1) newRow[headers.indexOf('date')] = targetDate;
    if (headers.indexOf('emp_id') !== -1) newRow[headers.indexOf('emp_id')] = user.emp_id;
    if (headers.indexOf('emp_name') !== -1) newRow[headers.indexOf('emp_name')] = user.emp_name;
    if (headers.indexOf('attendance_status') !== -1) newRow[headers.indexOf('attendance_status')] = attStatus;
    if (headers.indexOf('permission_hours') !== -1) newRow[headers.indexOf('permission_hours')] = permHrs;
    if (headers.indexOf('leave_hours') !== -1) newRow[headers.indexOf('leave_hours')] = leaveHours;
    if (headers.indexOf('shift_type') !== -1) newRow[headers.indexOf('shift_type')] = shiftType;
    if (headers.indexOf('shift_start_time') !== -1) newRow[headers.indexOf('shift_start_time')] = isPresentOverride ? '09:00:00' : '';
    if (headers.indexOf('shift_end_time') !== -1) newRow[headers.indexOf('shift_end_time')] = isPresentOverride ? '18:00:00' : '';
    if (headers.indexOf('total_shift_hours') !== -1) newRow[headers.indexOf('total_shift_hours')] = isPresentOverride ? stdShiftHours : 0;
    if (headers.indexOf('shrinkage_hours') !== -1) newRow[headers.indexOf('shrinkage_hours')] = shrinkageHrs;
    if (headers.indexOf('proof_url') !== -1) newRow[headers.indexOf('proof_url')] = proofUrl;
    if (headers.indexOf('shift_status') !== -1) newRow[headers.indexOf('shift_status')] = newShiftStatus;
    if (headers.indexOf('created_at') !== -1) newRow[headers.indexOf('created_at')] = new Date();
    if (headers.indexOf('active_client') !== -1) newRow[headers.indexOf('active_client')] = clientName;
    if (headers.indexOf('active_project') !== -1) newRow[headers.indexOf('active_project')] = projName;
    if (headers.indexOf('user_remarks') !== -1) newRow[headers.indexOf('user_remarks')] = remarks;

    sheet.appendRow(newRow);
    SpreadsheetApp.flush();
    syncExternalDailySummaryRecords(targetDate, user.emp_id);
    return {
      success: true,
      log_id: logId,
      shift_status: newShiftStatus,
      attendance_status: attStatus,
      permission_hours: permHrs,
      leave_hours: leaveHours,
      shift_type: shiftType,
      message: isPresentOverride ? ('Recorded Present shift (' + stdShiftHours + ' hrs)') : ('Submitted ' + leaveCode + ' (' + leaveHours + ' hrs leave/permission recorded)')
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * SET EMPLOYEE SCHEDULE STATUS (Single-Click Week Off / Holiday for Managers)
 * Sets WO, HOLIDAY, or CLEAR for a specific employee and date.
 */
function setEmployeeScheduleStatus(payload) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) {
    return { success: false, message: 'Server busy. Please try again.' };
  }

  try {
    var empId = payload ? String(payload.emp_id || '').trim() : '';
    var dateStr = payload ? String(payload.date || '').trim() : '';
    var status = payload ? String(payload.status || '').trim().toUpperCase() : '';

    if (!empId) return { success: false, message: 'Missing employee ID.' };
    if (!dateStr) return { success: false, message: 'Missing target date.' };

    var user = getUserByEmpId(empId);
    if (!user) return { success: false, message: 'Employee not found: ' + empId };

    var ss = getSpreadsheet();
    var attSheet = ss.getSheetByName(SHEET_NAMES.ATTENDANCE);
    if (!attSheet) return { success: false, message: 'Attendance sheet not found.' };

    ensureAttendanceHeaders(attSheet);
    var timeZone = Session.getScriptTimeZone();
    var normalizedTargetDate = normalizeDateStr(dateStr, timeZone) || dateStr;

    var data = attSheet.getDataRange().getValues();
    var headers = data[0];
    var empIdx = headers.indexOf('emp_id');
    var dateIdx = headers.indexOf('date');
    var attStatusIdx = headers.indexOf('attendance_status');
    var shiftStatusIdx = headers.indexOf('shift_status');
    var shiftHrsIdx = headers.indexOf('total_shift_hours');
    var leaveHrsIdx = headers.indexOf('leave_hours');
    var remarksIdx = headers.indexOf('user_remarks');
    var shiftTypeIdx = headers.indexOf('shift_type');

    var foundRow = -1;
    var empLower = empId.toLowerCase();

    for (var r = 1; r < data.length; r++) {
      var rEmp = empIdx !== -1 ? String(data[r][empIdx]).trim().toLowerCase() : '';
      var rDate = normalizeDateStr(data[r][dateIdx], timeZone);
      if (rEmp === empLower && rDate === normalizedTargetDate) {
        foundRow = r + 1;
        break;
      }
    }

    var defaultShift = normalizeShiftType(user.default_shift || 'General');

    if (status === 'CLEAR') {
      if (foundRow !== -1) {
        if (attStatusIdx !== -1) attSheet.getRange(foundRow, attStatusIdx + 1).setValue('NOT_CLOCKED_IN');
        if (shiftStatusIdx !== -1) attSheet.getRange(foundRow, shiftStatusIdx + 1).setValue('NOT_CLOCKED_IN');
        if (shiftHrsIdx !== -1) attSheet.getRange(foundRow, shiftHrsIdx + 1).setValue(0);
        if (leaveHrsIdx !== -1) attSheet.getRange(foundRow, leaveHrsIdx + 1).setValue(0);
        if (remarksIdx !== -1) attSheet.getRange(foundRow, remarksIdx + 1).setValue('');
      }
      SpreadsheetApp.flush();
      return { success: true, emp_id: empId, date: normalizedTargetDate, status: 'NOT_CLOCKED_IN', message: 'Cleared status for ' + user.emp_name };
    }

    var statusLabel = (status === 'WO' || status === 'WEEK_OFF' || status === 'WEEKOFF') ? 'WO' : 'HOLIDAY';
    var remarks = (statusLabel === 'WO') ? 'Week Off set by Manager' : 'Holiday set by Manager';

    if (foundRow !== -1) {
      if (attStatusIdx !== -1) attSheet.getRange(foundRow, attStatusIdx + 1).setValue(statusLabel);
      if (shiftStatusIdx !== -1) attSheet.getRange(foundRow, shiftStatusIdx + 1).setValue(statusLabel);
      if (shiftHrsIdx !== -1) attSheet.getRange(foundRow, shiftHrsIdx + 1).setValue(0);
      if (leaveHrsIdx !== -1) attSheet.getRange(foundRow, leaveHrsIdx + 1).setValue(0);
      if (remarksIdx !== -1) attSheet.getRange(foundRow, remarksIdx + 1).setValue(remarks);
    } else {
      var logId = 'ATT-' + normalizedTargetDate.replace(/-/g, '') + '-' + empId.replace(/[^a-zA-Z0-9]/g, '') + '-' + Math.floor(100000 + Math.random() * 900000);
      var newRow = new Array(headers.length);
      for (var h = 0; h < headers.length; h++) newRow[h] = '';
      if (headers.indexOf('log_id') !== -1) newRow[headers.indexOf('log_id')] = logId;
      if (headers.indexOf('date') !== -1) newRow[headers.indexOf('date')] = normalizedTargetDate;
      if (empIdx !== -1) newRow[empIdx] = user.emp_id;
      if (headers.indexOf('emp_name') !== -1) newRow[headers.indexOf('emp_name')] = user.emp_name;
      if (attStatusIdx !== -1) newRow[attStatusIdx] = statusLabel;
      if (shiftStatusIdx !== -1) newRow[shiftStatusIdx] = statusLabel;
      if (shiftTypeIdx !== -1) newRow[shiftTypeIdx] = defaultShift;
      if (shiftHrsIdx !== -1) newRow[shiftHrsIdx] = 0;
      if (leaveHrsIdx !== -1) newRow[leaveHrsIdx] = 0;
      if (headers.indexOf('total_break_minutes') !== -1) newRow[headers.indexOf('total_break_minutes')] = 60;
      if (headers.indexOf('active_client') !== -1) newRow[headers.indexOf('active_client')] = user.default_client || 'iMerit';
      if (headers.indexOf('active_project') !== -1) newRow[headers.indexOf('active_project')] = user.default_project || 'Pravah';
      if (remarksIdx !== -1) newRow[remarksIdx] = remarks;
      if (headers.indexOf('created_at') !== -1) newRow[headers.indexOf('created_at')] = new Date();

      attSheet.appendRow(newRow);
    }

    SpreadsheetApp.flush();
    return {
      success: true,
      emp_id: empId,
      date: normalizedTargetDate,
      status: statusLabel,
      message: 'Set ' + statusLabel + ' for ' + user.emp_name + ' on ' + normalizedTargetDate
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * MANAGER DIRECT MARK LEAVE ENGINE (Supports Today & Past Dates, WO and Holiday)
 */
function markLeaveOnBehalfByManager(payload) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) {
    return { success: false, message: 'Server busy.' };
  }

  try {
    if (!payload || !payload.emp_id) {
      return { success: false, message: 'Employee ID is required.' };
    }

    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAMES.ATTENDANCE);
    if (!sheet) return { success: false, message: 'Attendance sheet not found.' };

    var user = getUserByEmpId(payload.emp_id);
    if (!user) return { success: false, message: 'Employee not found in User_Master.' };

    var timeZone = Session.getScriptTimeZone();
    var targetDate = normalizeDateStr(payload.date, timeZone) || Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd');
    var rawLeaveCode = String(payload.leave_type || 'UL').trim().toUpperCase();
    var leaveCode = rawLeaveCode;
    if (leaveCode === 'WEEK OFF' || leaveCode === 'WEEKOFF') leaveCode = 'WO';
    if (leaveCode === 'HOLIDAY') leaveCode = 'HOLIDAY';

    var isScheduleOff = (leaveCode === 'WO' || leaveCode === 'HOLIDAY');
    var isFullDayLeave = !isScheduleOff && (leaveCode === 'PL' || leaveCode === 'UL' || leaveCode === 'ABSENT');
    var isShortPermAdmin = !isScheduleOff && (leaveCode === 'PERM' || leaveCode === 'PERMISSION' || leaveCode === 'OD');
    var permHrs = isShortPermAdmin ? (Number(payload.permission_hours) || 2.0) : 0.0;
    var shiftType = normalizeShiftType(payload.shift_type || payload.shift || user.default_shift || 'General');
    var stdShiftHours = getShiftStandardHours(shiftType);
    var leaveHours = isScheduleOff ? 0 : (isFullDayLeave ? stdShiftHours : (leaveCode === 'HL' ? (stdShiftHours / 2.0) : permHrs));
    var shrinkageHrs = leaveHours;
    var remarks = payload.remarks ? ('Manager Note: ' + payload.remarks) : (isScheduleOff ? (leaveCode === 'WO' ? 'Week Off set by Manager' : 'Holiday set by Manager') : 'Marked by Manager');

    ensureAttendanceHeaders(sheet);
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var empIdx = headers.indexOf('emp_id');
    var dateIdx = headers.indexOf('date');

    var targetEmpLower = String(user.emp_id || '').trim().toLowerCase();

    for (var i = 1; i < data.length; i++) {
      var rowDate = normalizeDateStr(data[i][dateIdx], timeZone);
      var rowEmp = empIdx !== -1 ? String(data[i][empIdx]).trim().toLowerCase() : '';

      if (rowEmp === targetEmpLower && rowDate === targetDate) {
        var rowIndex = i + 1;
        var newShiftStatus = isScheduleOff ? leaveCode : (isFullDayLeave ? 'LEAVE_CLOSED' : 'NOT_CLOCKED_IN');

        var rowValues = data[i];
        if (headers.indexOf('attendance_status') !== -1) rowValues[headers.indexOf('attendance_status')] = leaveCode;
        if (headers.indexOf('permission_hours') !== -1) rowValues[headers.indexOf('permission_hours')] = permHrs;
        if (headers.indexOf('shrinkage_hours') !== -1) rowValues[headers.indexOf('shrinkage_hours')] = shrinkageHrs;
        if (headers.indexOf('leave_hours') !== -1) rowValues[headers.indexOf('leave_hours')] = leaveHours;
        if (headers.indexOf('shift_type') !== -1) rowValues[headers.indexOf('shift_type')] = shiftType;
        if (headers.indexOf('shift_status') !== -1) rowValues[headers.indexOf('shift_status')] = newShiftStatus;
        if (isFullDayLeave || isScheduleOff) {
          if (headers.indexOf('shift_start_time') !== -1) rowValues[headers.indexOf('shift_start_time')] = '';
          if (headers.indexOf('shift_end_time') !== -1) rowValues[headers.indexOf('shift_end_time')] = '';
          if (headers.indexOf('total_shift_hours') !== -1) rowValues[headers.indexOf('total_shift_hours')] = 0;
        }
        if (headers.indexOf('user_remarks') !== -1) {
          var existingRem = String(data[i][headers.indexOf('user_remarks')] || '').trim();
          var combinedRem = existingRem ? (existingRem + ' | ' + remarks) : remarks;
          rowValues[headers.indexOf('user_remarks')] = combinedRem;
        }

        sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
        SpreadsheetApp.flush();
        syncExternalDailySummaryRecords(targetDate, user.emp_id);
        return {
          success: true,
          shift_status: newShiftStatus,
          attendance_status: leaveCode,
          leave_hours: leaveHours,
          shift_type: shiftType,
          message: 'Successfully updated status for ' + user.emp_name + ' on ' + targetDate + ' as ' + leaveCode
        };
      }
    }

    // New Row Creation for target date
    var cleanEmpIdAdmin = String(user.emp_id || '').trim().replace(/[^a-zA-Z0-9]/g, '');
    var logId = 'ATT-' + targetDate.replace(/-/g, '') + (cleanEmpIdAdmin ? '-' + cleanEmpIdAdmin : '') + '-' + Math.floor(100000 + Math.random() * 900000);
    var newRow = new Array(headers.length);
    for (var h = 0; h < headers.length; h++) newRow[h] = '';

    var newShiftStatus = isScheduleOff ? leaveCode : (isFullDayLeave ? 'LEAVE_CLOSED' : 'NOT_CLOCKED_IN');

    if (headers.indexOf('log_id') !== -1) newRow[headers.indexOf('log_id')] = logId;
    if (headers.indexOf('date') !== -1) newRow[headers.indexOf('date')] = targetDate;
    if (headers.indexOf('emp_id') !== -1) newRow[headers.indexOf('emp_id')] = user.emp_id;
    if (headers.indexOf('emp_name') !== -1) newRow[headers.indexOf('emp_name')] = user.emp_name;
    if (headers.indexOf('attendance_status') !== -1) newRow[headers.indexOf('attendance_status')] = leaveCode;
    if (headers.indexOf('permission_hours') !== -1) newRow[headers.indexOf('permission_hours')] = permHrs;
    if (headers.indexOf('leave_hours') !== -1) newRow[headers.indexOf('leave_hours')] = leaveHours;
    if (headers.indexOf('shift_type') !== -1) newRow[headers.indexOf('shift_type')] = shiftType;
    if (headers.indexOf('shift_start_time') !== -1) newRow[headers.indexOf('shift_start_time')] = '';
    if (headers.indexOf('shift_end_time') !== -1) newRow[headers.indexOf('shift_end_time')] = '';
    if (headers.indexOf('total_shift_hours') !== -1) newRow[headers.indexOf('total_shift_hours')] = 0;
    if (headers.indexOf('shrinkage_hours') !== -1) newRow[headers.indexOf('shrinkage_hours')] = shrinkageHrs;
    if (headers.indexOf('shift_status') !== -1) newRow[headers.indexOf('shift_status')] = newShiftStatus;
    if (headers.indexOf('created_at') !== -1) newRow[headers.indexOf('created_at')] = new Date();
    if (headers.indexOf('active_client') !== -1) newRow[headers.indexOf('active_client')] = user.default_client || 'iMerit';
    if (headers.indexOf('active_project') !== -1) newRow[headers.indexOf('active_project')] = user.default_project || 'Pravah';
    if (headers.indexOf('user_remarks') !== -1) newRow[headers.indexOf('user_remarks')] = remarks;

    sheet.appendRow(newRow);
    SpreadsheetApp.flush();
    syncExternalDailySummaryRecords(targetDate, user.emp_id);

    return {
      success: true,
      log_id: logId,
      shift_status: newShiftStatus,
      attendance_status: leaveCode,
      leave_hours: leaveHours,
      shift_type: shiftType,
      message: 'Successfully marked ' + leaveCode + ' for ' + user.emp_name + ' on ' + targetDate
    };

  } catch (err) {
    return { success: false, message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * DAILY AUTO-ATTENDANCE TRIGGER
 * If staff has not clocked in by 11:00 AM, auto-marks as UL (Unplanned Leave)
 */
function autoMarkUnappliedAbsences() {
  try {
    var ss = getSpreadsheet();
    var uSheet = ss.getSheetByName(SHEET_NAMES.USER_MASTER);
    var attSheet = ss.getSheetByName(SHEET_NAMES.ATTENDANCE);
    if (!uSheet || !attSheet) return;

    var timeZone = Session.getScriptTimeZone();
    var todayStr = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd');

    var aData = attSheet.getDataRange().getValues();
    var aHeaders = aData[0];
    var aDateIdx = aHeaders.indexOf('date');
    var aEmpIdx = aHeaders.indexOf('emp_id');

    var loggedMap = {};
    for (var i = 1; i < aData.length; i++) {
      var rowDate = normalizeDateStr(aData[i][aDateIdx], timeZone);
      if (rowDate === todayStr) {
        var empIdVal = aEmpIdx !== -1 ? String(aData[i][aEmpIdx]).trim().toLowerCase() : '';
        if (empIdVal) loggedMap[empIdVal] = true;
      }
    }

    var uData = uSheet.getDataRange().getValues();
    var uHeaders = uData[0];
    var uEmpIdx = uHeaders.indexOf('emp_id');
    var uNameIdx = uHeaders.indexOf('emp_name');
    var uActiveIdx = uHeaders.indexOf('is_active');
    var clientIdx = uHeaders.indexOf('default_client');
    var projIdx = uHeaders.indexOf('default_project');
    var roleIdx = uHeaders.indexOf('app_role');

    for (var u = 1; u < uData.length; u++) {
      var empId = uEmpIdx !== -1 ? String(uData[u][uEmpIdx]).trim() : '';
      var empName = uNameIdx !== -1 ? String(uData[u][uNameIdx]).trim() : '';
      var actVal = uActiveIdx !== -1 ? String(uData[u][uActiveIdx] || 'TRUE').trim().toUpperCase() : 'TRUE';
      var isActive = (actVal === 'TRUE' || actVal === '1' || actVal === 'ACTIVE' || actVal === 'YES' || uData[u][uActiveIdx] === true);
      var clientName = clientIdx !== -1 ? String(uData[u][clientIdx] || 'iMerit') : 'iMerit';
      var projName = projIdx !== -1 ? String(uData[u][projIdx] || 'Pravah') : 'Pravah';
      var uRole = roleIdx !== -1 ? String(uData[u][roleIdx] || '').trim().toLowerCase() : '';

      if (uRole === 'management admin' || uRole === 'management_admin' || uRole === 'management' || uRole === 'super admin' || uRole === 'superadmin') {
        continue; // Skip Management Admins
      }

      if (isActive && empId && !loggedMap[empId.toLowerCase()]) {
        var cleanEmpIdAuto = String(empId || '').trim().replace(/[^a-zA-Z0-9]/g, '');
        var logId = 'ATT-' + todayStr.replace(/-/g, '') + (cleanEmpIdAuto ? '-' + cleanEmpIdAuto : '') + '-' + Math.floor(100000 + Math.random() * 900000);
        var newRow = new Array(aHeaders.length);
        for (var h = 0; h < aHeaders.length; h++) newRow[h] = '';

        if (aHeaders.indexOf('log_id') !== -1) newRow[aHeaders.indexOf('log_id')] = logId;
        if (aHeaders.indexOf('date') !== -1) newRow[aHeaders.indexOf('date')] = todayStr;
        if (aHeaders.indexOf('emp_id') !== -1) newRow[aHeaders.indexOf('emp_id')] = empId;
        if (aHeaders.indexOf('emp_name') !== -1) newRow[aHeaders.indexOf('emp_name')] = empName;
        if (aHeaders.indexOf('attendance_status') !== -1) newRow[aHeaders.indexOf('attendance_status')] = 'UL';
        if (aHeaders.indexOf('permission_hours') !== -1) newRow[aHeaders.indexOf('permission_hours')] = 0.0;
        if (aHeaders.indexOf('shift_start_time') !== -1) newRow[aHeaders.indexOf('shift_start_time')] = '';
        if (aHeaders.indexOf('shift_end_time') !== -1) newRow[aHeaders.indexOf('shift_end_time')] = '';
        if (aHeaders.indexOf('total_shift_hours') !== -1) newRow[aHeaders.indexOf('total_shift_hours')] = 0;
        if (aHeaders.indexOf('shrinkage_hours') !== -1) newRow[aHeaders.indexOf('shrinkage_hours')] = 8.0;
        if (aHeaders.indexOf('shift_status') !== -1) newRow[aHeaders.indexOf('shift_status')] = 'LEAVE_CLOSED';
        if (aHeaders.indexOf('created_at') !== -1) newRow[aHeaders.indexOf('created_at')] = new Date();
        if (aHeaders.indexOf('active_client') !== -1) newRow[aHeaders.indexOf('active_client')] = clientName;
        if (aHeaders.indexOf('active_project') !== -1) newRow[aHeaders.indexOf('active_project')] = projName;
        if (aHeaders.indexOf('user_remarks') !== -1) newRow[aHeaders.indexOf('user_remarks')] = 'Automated System Marking: Unapplied Leave';

        attSheet.appendRow(newRow);
        SpreadsheetApp.flush();
        syncExternalDailySummaryRecords(todayStr, empId);
      }
    }
  } catch (err) {
    Logger.log('autoMarkUnappliedAbsences Error: ' + err.toString());
  }
}

/**
 * TASK LOGGING ENGINE (Cleaned: No AHT calculations, captures Client, Project, Activity, Hours Worked)
 */
function logTaskEntry(payload) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) {
    return { success: false, message: 'Server busy.' };
  }

  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAMES.PROJECT_LOGS) || ss.insertSheet(SHEET_NAMES.PROJECT_LOGS);
    ensureProjectLogsHeaders(sheet);

    var email = payload.email || Session.getActiveUser().getEmail();
    var user = getUserByEmail(email) || getUserByEmpId(payload.emp_id);
    if (!user) {
      return { success: false, message: 'Access Denied: Your account has been removed from User_Master. Contact Admin.' };
    }
    var activeVal = String(user.is_active || 'TRUE').trim().toUpperCase();
    var isActive = (activeVal === 'TRUE' || activeVal === '1' || activeVal === 'ACTIVE' || activeVal === 'YES' || user.is_active === true);
    if (!isActive) {
      return { success: false, message: 'Access Denied: Your account is currently inactive. Contact Admin.' };
    }

    var timeZone = Session.getScriptTimeZone();
    var todayStr = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd');
    var cleanEmpIdTask = String(user.emp_id || '').trim().replace(/[^a-zA-Z0-9]/g, '');
    var taskLogId = 'TASK-' + todayStr.replace(/-/g, '') + (cleanEmpIdTask ? '-' + cleanEmpIdTask : '') + '-' + Math.floor(100000 + Math.random() * 900000);

    var actType = normalizeActivityType(payload.activity_type);
    var clientName = String(payload.active_client || '').trim();
    var projectName = String(payload.active_project || '').trim();

    var actualLoggedHrs = 0;
    if (payload.actual_worked_hours !== undefined && payload.actual_worked_hours !== null && payload.actual_worked_hours !== '') {
      actualLoggedHrs = Number(payload.actual_worked_hours) || 0;
    } else if (payload.hours_worked !== undefined && payload.hours_worked !== null && payload.hours_worked !== '') {
      actualLoggedHrs = Number(payload.hours_worked) || 0;
    }

    var startTime = parseTimeStringToDate(payload.work_start_time, todayStr, timeZone);
    var endTime = parseTimeStringToDate(payload.work_end_time, todayStr, timeZone);
    var now = new Date();

    if (!endTime) endTime = now;
    if (!startTime) {
      startTime = new Date(endTime.getTime() - Math.max(1000, actualLoggedHrs * 3600 * 1000));
    }

    if (actualLoggedHrs <= 0) {
      actualLoggedHrs = Math.max(0.01, (endTime.getTime() - startTime.getTime()) / (1000 * 60 * 60));
    }
    actualLoggedHrs = Number(actualLoggedHrs.toFixed(2));

    var idleHours = (actType === 'Idle') ? actualLoggedHrs : 0.00;
    var benchHours = (actType === 'Training (Material Go-Through)') ? actualLoggedHrs : 0.00;
    var blockedHours = (actType === 'Down Time') ? actualLoggedHrs : 0.00;

    var isEditRow = false;
    if (payload.task_log_id) {
      var dVals = sheet.getDataRange().getValues();
      var hList = dVals[0];
      var idIdx = -1;
      for (var hi = 0; hi < hList.length; hi++) {
        if (String(hList[hi]).trim().toLowerCase().replace(/_/g, '') === 'tasklogid') idIdx = hi;
      }
      if (idIdx !== -1) {
        for (var r = 1; r < dVals.length; r++) {
          if (String(dVals[r][idIdx]).trim() === String(payload.task_log_id).trim()) {
            taskLogId = payload.task_log_id;
            isEditRow = r + 1;
            break;
          }
        }
      }
    }

    var row = [
      taskLogId, payload.log_id || '', user.emp_id, email,
      clientName, projectName,
      payload.task_status || 'COMPLETED', Number(payload.completion_pct) || 100,
      Utilities.formatDate(startTime, timeZone, 'yyyy-MM-dd HH:mm:ss'),
      Utilities.formatDate(endTime, timeZone, 'yyyy-MM-dd HH:mm:ss'),
      0, 0, actualLoggedHrs, actualLoggedHrs, 100, 0,
      payload.user_remarks || '', new Date(),
      actType, idleHours, benchHours, blockedHours
    ];

    if (isEditRow) {
      sheet.getRange(isEditRow, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
    SpreadsheetApp.flush();

    // Trigger sync for external summary
    syncExternalDailySummaryRecords(todayStr, user.emp_id);

    return { success: true, task_log_id: taskLogId };
  } catch (err) {
    return { success: false, message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * USER/ADMIN: Delete Logged Task Entry
 */
function deleteTaskLog(taskLogId, empId) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAMES.PROJECT_LOGS);
    if (!sheet) return { success: false, message: 'Task logs sheet not found.' };

    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var taskLogIdx = headers.indexOf('task_log_id');
    var empIdx = headers.indexOf('emp_id');
    var timeZone = Session.getScriptTimeZone();
    var todayStr = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd');

    for (var i = 1; i < data.length; i++) {
      var rowTaskId = taskLogIdx !== -1 ? String(data[i][taskLogIdx]).trim() : '';
      var rowEmpId = empIdx !== -1 ? String(data[i][empIdx]).trim() : empId;

      if (rowTaskId === String(taskLogId).trim()) {
        sheet.deleteRow(i + 1);
        SpreadsheetApp.flush();
        if (rowEmpId) {
          syncExternalDailySummaryRecords(todayStr, rowEmpId);
        }
        return { success: true, message: 'Task log deleted successfully.' };
      }
    }
    return { success: false, message: 'Task log entry not found.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * MODULE E: Fetch Past Missing Monday-Friday Timesheet Dates
 */
var APP_LAUNCH_DATE = '2026-08-01';

function getMissingTimesheetDates(empId) {
  var missingDates = [];
  if (!empId) return missingDates;

  var user = getUserByEmpId(empId);
  if (user) {
    var uRole = String(user.app_role || user.role || '').trim().toLowerCase();
    if (uRole === 'management admin' || uRole === 'management_admin' || uRole === 'management' || uRole === 'super admin' || uRole === 'superadmin') {
      return missingDates; // Management Admins are exempt from missing timesheets
    }
  }

  try {
    var ss = getSpreadsheet();
    var attSheet = ss.getSheetByName(SHEET_NAMES.ATTENDANCE);
    if (!attSheet) return missingDates;

    var aData = attSheet.getDataRange().getValues();
    var timeZone = Session.getScriptTimeZone();
    var headers = aData[0];
    var empIdx = headers.indexOf('emp_id');
    var dateIdx = headers.indexOf('date');

    var targetEmpLower = String(empId).trim().toLowerCase();
    var loggedDates = {};
    for (var i = 1; i < aData.length; i++) {
      var rowEmp = empIdx !== -1 ? String(aData[i][empIdx]).trim().toLowerCase() : '';
      if (rowEmp === targetEmpLower) {
        var rDate = normalizeDateStr(aData[i][dateIdx], timeZone);
        if (rDate) loggedDates[rDate] = true;
      }
    }

    // Scan past 14 weekdays on or after APP_LAUNCH_DATE (excluding today and weekends)
    var now = new Date();
    for (var d = 1; d <= 14; d++) {
      var checkDate = new Date(now.getTime() - (d * 24 * 60 * 60 * 1000));
      var dayOfWeek = checkDate.getDay(); // 0 = Sun, 6 = Sat
      if (dayOfWeek !== 0 && dayOfWeek !== 6) { // Monday-Friday only
        var dateStr = Utilities.formatDate(checkDate, timeZone, 'yyyy-MM-dd');
        if (dateStr >= APP_LAUNCH_DATE && !loggedDates[dateStr]) {
          missingDates.push(dateStr);
        }
      }
    }
  } catch (e) {
    Logger.log('Missing timesheet check error: ' + e.toString());
  }

  return missingDates;
}

/**
 * MODULE B: Nightly Auto-Close Routine for Unclosed Overnight Shifts
 */
function autoCloseOvernightShifts() {
  try {
    var ss = getSpreadsheet();
    var attSheet = ss.getSheetByName(SHEET_NAMES.ATTENDANCE);
    if (!attSheet) return;

    var dataRange = attSheet.getDataRange();
    var data = dataRange.getValues();
    var headers = data[0];
    var timeZone = Session.getScriptTimeZone();
    var todayStr = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd');

    var statusIdx = headers.indexOf('shift_status');
    var endIdx = headers.indexOf('shift_end_time');
    var hoursIdx = headers.indexOf('total_shift_hours');
    var dateIdx = headers.indexOf('date');

    var modified = false;
    for (var i = 1; i < data.length; i++) {
      var rDate = normalizeDateStr(data[i][dateIdx], timeZone);
      var status = String(data[i][statusIdx]).toUpperCase();
      if ((status === 'OPEN' || status === 'ON_BREAK') && rDate && rDate !== todayStr) {
        data[i][statusIdx] = 'AUTO_CLOSED';
        if (endIdx !== -1) data[i][endIdx] = '20:00:00';
        if (hoursIdx !== -1) data[i][hoursIdx] = 8.00;
        modified = true;
      }
    }

    if (modified) {
      dataRange.setValues(data);
      SpreadsheetApp.flush();
    }
  } catch (e) {
    Logger.log('Auto-close overnight shifts error: ' + e.toString());
  }
}

/**
 * Dynamic Self-Healing Auto-Close: Automatically closes any open or on-break shifts
 * from previous days for a specific employee.
 */
function autoClosePreviousOpenShiftsForUser(empId) {
  if (!empId) return;
  try {
    var ss = getSpreadsheet();
    var attSheet = ss.getSheetByName(SHEET_NAMES.ATTENDANCE);
    if (!attSheet) return;

    var dataRange = attSheet.getDataRange();
    var data = dataRange.getValues();
    var headers = data[0];
    var timeZone = Session.getScriptTimeZone();
    var todayStr = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd');

    var empIdx = headers.indexOf('emp_id');
    var statusIdx = headers.indexOf('shift_status');
    var endIdx = headers.indexOf('shift_end_time');
    var hoursIdx = headers.indexOf('total_shift_hours');
    var dateIdx = headers.indexOf('date');

    if (empIdx === -1 || statusIdx === -1 || dateIdx === -1) return;

    var targetEmpLower = String(empId).trim().toLowerCase();
    var modified = false;
    for (var i = 1; i < data.length; i++) {
      var rowEmp = String(data[i][empIdx]).trim().toLowerCase();
      if (rowEmp === targetEmpLower) {
        var rDate = normalizeDateStr(data[i][dateIdx], timeZone);
        var status = String(data[i][statusIdx]).toUpperCase();
        if ((status === 'OPEN' || status === 'ON_BREAK') && rDate && rDate !== todayStr) {
          data[i][statusIdx] = 'AUTO_CLOSED';
          if (endIdx !== -1) data[i][endIdx] = '20:00:00';
          if (hoursIdx !== -1) data[i][hoursIdx] = 8.00;
          modified = true;
        }
      }
    }

    if (modified) {
      dataRange.setValues(data);
      SpreadsheetApp.flush();
    }
  } catch (e) {
    Logger.log('Dynamic auto-close previous shifts error for ' + empId + ': ' + e.toString());
  }
}

/**
 * CONSOLIDATED ADMIN DASHBOARD ENGINE (Day, Week, Month Periods + Work Log Breakdowns + Fixed 1h Daily Break)
 */
function getAdminDashboardData(filters) {
  try {
    var activeEmail = Session.getActiveUser().getEmail();
    if (activeEmail) {
      var callingUser = getUserByEmail(activeEmail);
      if (callingUser) {
        var uRole = String(callingUser.app_role || callingUser.role || '').trim().toLowerCase();
        if (uRole !== 'admin' && uRole !== 'executive' && uRole !== 'manager' && uRole !== 'management admin' && uRole !== 'management_admin' && uRole !== 'management' && uRole !== 'super admin' && uRole.indexOf('admin') === -1 && uRole.indexOf('exec') === -1 && uRole.indexOf('manager') === -1) {
          return { success: false, message: 'Access Denied: Admin role required to view dashboard metrics.' };
        }
      }
    }

    var ss = getSpreadsheet();
    var timeZone = Session.getScriptTimeZone();
    var todayStr = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd');
    var targetDate = (filters && filters.date) ? normalizeDateStr(filters.date, timeZone) : todayStr;
    if (!targetDate) targetDate = todayStr;

    var period = (filters && filters.period) ? String(filters.period).toLowerCase() : 'day';
    if (['day', 'week', 'month'].indexOf(period) === -1) period = 'day';

    var dateRange = getDateRangeForPeriod(targetDate, period, timeZone);
    var startDate = dateRange.startDate;
    var endDate = dateRange.endDate;

    var selectedClient = (filters && (filters.client || filters.active_client)) ? String(filters.client || filters.active_client).trim() : 'ALL';
    var selectedProject = (filters && (filters.project || filters.active_project)) ? String(filters.project || filters.active_project).trim() : 'ALL';
    var selectedTeam = (filters && filters.team) ? String(filters.team).trim() : 'ALL';
    var selectedShift = (filters && filters.shift) ? String(filters.shift).trim() : 'ALL';
    var selectedSearch = (filters && filters.search) ? String(filters.search).trim().toLowerCase() : '';

    var uSheet = ss.getSheetByName(SHEET_NAMES.USER_MASTER);
    var users = [];
    if (uSheet) {
      var uData = uSheet.getDataRange().getValues();
      if (uData.length > 1) {
        var uHeaders = uData[0];
        var activeIdx = uHeaders.indexOf('is_active');
        for (var i = 1; i < uData.length; i++) {
          if (activeIdx !== -1) {
            var actVal = String(uData[i][activeIdx] || 'TRUE').trim().toUpperCase();
            var isActive = (actVal === 'TRUE' || actVal === '1' || actVal === 'ACTIVE' || actVal === 'YES' || uData[i][activeIdx] === true);
            if (!isActive) continue; // Skip inactive accounts
          }
          var uObj = {};
          for (var col = 0; col < uHeaders.length; col++) {
            var h = uHeaders[col];
            var val = uData[i][col];
            if (val instanceof Date) {
              uObj[h] = Utilities.formatDate(val, timeZone, 'yyyy-MM-dd');
            } else {
              uObj[h] = val !== null && val !== undefined ? val : '';
            }
          }
          users.push(uObj);
        }
      }
    }

    // Extended query range for summary trends if period is month (allows historical Aug, Sep, Oct, Nov progression)
    var queryStartDate = startDate;
    var queryEndDate = endDate;
    if (period === 'month') {
      var tParts = targetDate.split('-');
      var tY = parseInt(tParts[0], 10);
      var tM = parseInt(tParts[1], 10) - 1;
      var startQuarterDate = new Date(tY, Math.max(0, tM - 3), 1);
      var qStartStr = Utilities.formatDate(startQuarterDate, timeZone, 'yyyy-MM-dd');
      if (qStartStr < queryStartDate) {
        queryStartDate = qStartStr;
      }
    }

    // Load Attendance Logs in [queryStartDate, queryEndDate]
    var aSheet = ss.getSheetByName(SHEET_NAMES.ATTENDANCE);
    if (aSheet) ensureAttendanceHeaders(aSheet);
    var attendanceByEmp = {}; // empKey -> array of attendance rows in selected [startDate, endDate]
    var allAttendanceRows = []; // all rows in [queryStartDate, queryEndDate] for summary calculations
    if (aSheet) {
      var aData = aSheet.getDataRange().getValues();
      if (aData.length > 1) {
        var aHeaders = aData[0];
        var aDateIdx = aHeaders.indexOf('date');
        var aEmpIdx = aHeaders.indexOf('emp_id');

        for (var k = 1; k < aData.length; k++) {
          var rDate = normalizeDateStr(aData[k][aDateIdx], timeZone);
          if (rDate && rDate >= queryStartDate && rDate <= queryEndDate) {
            var aObj = {};
            for (var col = 0; col < aHeaders.length; col++) {
              var h = aHeaders[col];
              var val = aData[k][col];
              if (val instanceof Date) {
                if (h === 'date' || h === 'created_at') {
                  aObj[h] = Utilities.formatDate(val, timeZone, 'yyyy-MM-dd');
                } else {
                  aObj[h] = Utilities.formatDate(val, timeZone, 'HH:mm:ss');
                }
              } else {
                aObj[h] = val !== null && val !== undefined ? val : '';
              }
            }
            aObj.normalizedDate = rDate;
            aObj.shift_type = normalizeShiftType(aObj.shift_type || 'General');
            aObj.leave_hours = Number(aObj.leave_hours) || 0;
            allAttendanceRows.push(aObj);

            if (rDate >= startDate && rDate <= endDate) {
              var empKey = aEmpIdx !== -1 ? String(aData[k][aEmpIdx]).trim().toLowerCase() : String(aObj.emp_id || '').trim().toLowerCase();
              if (!attendanceByEmp[empKey]) attendanceByEmp[empKey] = [];
              attendanceByEmp[empKey].push(aObj);
            }
          }
        }
      }
    }

    // Load Project Task Logs in [queryStartDate, queryEndDate]
    var tSheet = ss.getSheetByName(SHEET_NAMES.PROJECT_LOGS);
    var tasksByEmp = {};
    var rawTaskLogs = [];
    var allQueryTasks = [];
    if (tSheet) {
      ensureProjectLogsHeaders(tSheet);
      var tData = tSheet.getDataRange().getValues();
      if (tData.length > 1) {
        var tHeaders = tData[0];
        var workStartIdx = tHeaders.indexOf('work_start_time');
        var createdIdx = tHeaders.indexOf('created_at');
        var tEmpIdx = tHeaders.indexOf('emp_id');

        for (var t = 1; t < tData.length; t++) {
          var rawDateVal = workStartIdx !== -1 ? tData[t][workStartIdx] : (createdIdx !== -1 ? tData[t][createdIdx] : null);
          var tDate = normalizeDateStr(rawDateVal, timeZone);
          if (!tDate && createdIdx !== -1) {
            tDate = normalizeDateStr(tData[t][createdIdx], timeZone);
          }

          if (tDate && tDate >= queryStartDate && tDate <= queryEndDate) {
            var tObj = {};
            for (var col = 0; col < tHeaders.length; col++) {
              var h = tHeaders[col];
              var val = tData[t][col];
              if (val instanceof Date) {
                tObj[h] = Utilities.formatDate(val, timeZone, 'yyyy-MM-dd HH:mm:ss');
              } else {
                tObj[h] = val !== null && val !== undefined ? val : '';
              }
            }

            // Normalize activity type to one of the 6 allowed activities
            tObj.activity_type = normalizeActivityType(tObj.activity_type || tData[t][18]);
            tObj.actual_worked_hours = Number(tObj.actual_worked_hours) || Number(tObj.hours_worked) || 0;
            tObj.normalizedDate = tDate;

            allQueryTasks.push(tObj);

            if (tDate >= startDate && tDate <= endDate) {
              rawTaskLogs.push(tObj);
              var empKey = tEmpIdx !== -1 ? String(tData[t][tEmpIdx]).trim().toLowerCase() : String(tObj.emp_id || '').trim().toLowerCase();
              if (!tasksByEmp[empKey]) tasksByEmp[empKey] = [];
              tasksByEmp[empKey].push(tObj);
            }
          }
        }
      }
    }

    var clientMaster = getClientMasterData();

    // Aggregations
    var employeeStatusList = [];
    var leaveQueue = [];
    var clockedInCount = 0;
    var distinctLeaveStaffCount = 0;

    var totalShiftWorkingHours = 0; // Net Shift Worked Hours (fixed 1h break deducted per day worked)
    var totalLoggedWorkHours = 0;   // Actual Logged Task Hours across all activities
    var totalPeriodLeaveHours = 0;  // Leave Hours kept separate from worked & prod hours

    var activityHoursMap = {};
    ALLOWED_ACTIVITIES.forEach(function(act) { activityHoursMap[act] = 0; });

    var clientHoursMap = {};
    var projectHoursMap = {};

    var prodHeadcountDays = 0;
    var supportHeadcountDays = 0;
    var totalAvailablePeriodProdHours = 0;

    for (var uIdx = 0; uIdx < users.length; uIdx++) {
      var u = users[uIdx];
      var uRole = String(u.app_role || u.role || '').trim().toLowerCase();
      var isMgmtAdmin = (uRole === 'management admin' || uRole === 'management_admin' || uRole === 'management' || uRole === 'super admin' || uRole === 'superadmin' || uRole === 'executive' || uRole === 'mgmt admin');
      if (isMgmtAdmin) continue; // Management Admins are pure management; exclude from shift directory capacity

      var userEmpId = String(u.emp_id || '').trim();
      var userEmpKey = userEmpId.toLowerCase();
      var userEmpName = String(u.emp_name || '');
      var userTeam = String(u.team || 'N/A');

      // Filter by Team
      if (selectedTeam !== 'ALL' && userTeam.toUpperCase() !== selectedTeam.toUpperCase()) continue;

      // Filter by Search string (emp_id, emp_name, team)
      if (selectedSearch) {
        var matchSearch = (userEmpId.toLowerCase().indexOf(selectedSearch) !== -1 ||
          userEmpName.toLowerCase().indexOf(selectedSearch) !== -1 ||
          userTeam.toLowerCase().indexOf(selectedSearch) !== -1);
        if (!matchSearch) continue;
      }

      var allEmpTasks = tasksByEmp[userEmpKey] || [];

      // Filter tasks by Client and Project
      var empFilteredTasks = [];
      for (var st = 0; st < allEmpTasks.length; st++) {
        var tsk = allEmpTasks[st];
        if (selectedClient !== 'ALL' && String(tsk.active_client).toUpperCase() !== selectedClient.toUpperCase()) continue;
        if (selectedProject !== 'ALL' && String(tsk.active_project).toUpperCase() !== selectedProject.toUpperCase()) continue;
        empFilteredTasks.push(tsk);
      }

      // If client/project filter is active and employee has no matching tasks and their default is not matching, check
      if (selectedClient !== 'ALL' || selectedProject !== 'ALL') {
        var userDefClient = String(u.default_client || u.client || '').toUpperCase();
        var userDefProj = String(u.default_project || u.project || '').toUpperCase();
        var matchesDefault = (selectedClient === 'ALL' || userDefClient === selectedClient.toUpperCase()) &&
                             (selectedProject === 'ALL' || userDefProj === selectedProject.toUpperCase());
        if (empFilteredTasks.length === 0 && !matchesDefault) continue;
      }

      var empAttendanceRows = attendanceByEmp[userEmpKey] || [];

      // Determine Employee Primary Shift
      var empShiftType = 'General';
      for (var sIdx = empAttendanceRows.length - 1; sIdx >= 0; sIdx--) {
        if (empAttendanceRows[sIdx].shift_type) {
          empShiftType = normalizeShiftType(empAttendanceRows[sIdx].shift_type);
          break;
        }
      }
      if (!empShiftType) empShiftType = normalizeShiftType(u.default_shift || 'General');

      // Filter by Shift if applicable
      if (selectedShift !== 'ALL' && empShiftType.toUpperCase() !== selectedShift.toUpperCase()) {
        continue;
      }

      // Determine Prod vs Support Headcount
      var userDesig = String(u.job_designation || u.designation || '').toUpperCase();
      var userRoleUpper = String(u.app_role || u.role || '').toUpperCase();
      var userTeamUpper = userTeam.toUpperCase();
      var isSupportStaff = (userDesig && ['TL', 'LEAD', 'MANAGER', 'QA', 'SUPPORT', 'TRAINER', 'SME', 'OPS'].indexOf(userDesig) !== -1) ||
                           (userRoleUpper && ['ADMIN', 'MANAGER', 'LEAD', 'MANAGEMENT'].indexOf(userRoleUpper) !== -1) ||
                           (userTeamUpper === 'MANAGEMENT' || userTeamUpper === 'SUPPORT');
      var isProdStaff = !isSupportStaff;

      // Calculate Net Shift Working Hours and Leave Hours
      var empNetShiftHours = 0;
      var empLeaveHours = 0;
      var workedDaysSet = {};
      var leaveDaysCount = 0;
      var latestShiftStatus = 'NOT_CLOCKED_IN';
      var latestAttStatus = 'NOT_CLOCKED_IN';
      var latestShiftStart = '--:--';
      var latestShiftEnd = '--:--';

      for (var aIdx = 0; aIdx < empAttendanceRows.length; aIdx++) {
        var attRow = empAttendanceRows[aIdx];
        var rowStatus = String(attRow.shift_status || 'NOT_CLOCKED_IN').toUpperCase();
        var rowAttStatus = String(attRow.attendance_status || '').toUpperCase();
        var rowDate = attRow.normalizedDate;
        var rowShiftType = normalizeShiftType(attRow.shift_type || empShiftType);
        var rowStdShiftHrs = getShiftStandardHours(rowShiftType);

        var isScheduleOff = (rowAttStatus === 'WO' || rowStatus === 'WO' || rowAttStatus === 'HOLIDAY' || rowStatus === 'HOLIDAY');

        var isLeave = !isScheduleOff && (['PL', 'UL', 'HL', 'PERMISSION', 'LEAVE_CLOSED'].indexOf(rowAttStatus) !== -1 || rowStatus === 'LEAVE_CLOSED');
        if (isLeave) {
          leaveDaysCount++;

          var rowLeaveHrs = Number(attRow.leave_hours) || 0;
          if (rowLeaveHrs <= 0) {
            if (rowAttStatus === 'PL' || rowAttStatus === 'UL' || rowAttStatus === 'ABSENT' || rowStatus === 'LEAVE_CLOSED') {
              rowLeaveHrs = rowStdShiftHrs;
            } else if (rowAttStatus === 'HL') {
              rowLeaveHrs = rowStdShiftHrs / 2.0;
            } else {
              rowLeaveHrs = Number(attRow.permission_hours) || 0;
            }
          }
          empLeaveHours += rowLeaveHrs;
          totalPeriodLeaveHours += rowLeaveHrs;

          leaveQueue.push({
            emp_id: userEmpId,
            emp_name: userEmpName,
            team: userTeam,
            leave_type: rowAttStatus || 'LEAVE',
            perm_hours: Number(attRow.permission_hours) || 0,
            leave_hours: rowLeaveHrs,
            date: rowDate,
            proof_url: String(attRow.proof_url || '')
          });
        }

        var isWorked = !isScheduleOff && (rowStatus === 'OPEN' || rowStatus === 'CLOSED' || rowAttStatus === 'PRESENT');
        if (isWorked && !isLeave) {
          workedDaysSet[rowDate] = true;
          var rowHours = Number(attRow.total_shift_hours) || 0;

          if (rowStatus === 'OPEN' && rowDate === todayStr && attRow.shift_start_time) {
            try {
              var st = parseTimeStringToDate(attRow.shift_start_time, todayStr, timeZone);
              var nowD = new Date();
              if (st && nowD > st) {
                var grossHrs = (nowD.getTime() - st.getTime()) / (1000 * 60 * 60);
                rowHours = Math.max(0, grossHrs - 1.0); // 1-hour break automatically excluded
              }
            } catch(e) {}
          } else if (rowHours <= 0 && attRow.shift_start_time && attRow.shift_end_time) {
            try {
              var st = parseTimeStringToDate(attRow.shift_start_time, rowDate, timeZone);
              var et = parseTimeStringToDate(attRow.shift_end_time, rowDate, timeZone);
              if (st && et && et > st) {
                var grossHrs = (et.getTime() - st.getTime()) / (1000 * 60 * 60);
                rowHours = Math.max(0, grossHrs - 1.0); // 1-hour break automatically excluded
              }
            } catch(e) {}
          }

          empNetShiftHours += rowHours;
        }

        // Keep track of latest shift status for roster display
        if (aIdx === empAttendanceRows.length - 1 || rowDate === todayStr) {
          latestShiftStatus = isScheduleOff ? ((rowAttStatus === 'WO' || rowStatus === 'WO') ? 'WO' : 'HOLIDAY') : rowStatus;
          latestAttStatus = isScheduleOff ? latestShiftStatus : (rowAttStatus || rowStatus);
          if (attRow.shift_start_time) latestShiftStart = attRow.shift_start_time instanceof Date ? Utilities.formatDate(attRow.shift_start_time, timeZone, 'HH:mm:ss') : String(attRow.shift_start_time);
          if (attRow.shift_end_time) latestShiftEnd = attRow.shift_end_time instanceof Date ? Utilities.formatDate(attRow.shift_end_time, timeZone, 'HH:mm:ss') : String(attRow.shift_end_time);
        }
      }

      if (leaveDaysCount > 0) {
        distinctLeaveStaffCount++;
      }

      var isEmpActiveInPeriod = false;
      if (period === 'day') {
        isEmpActiveInPeriod = (latestShiftStatus === 'OPEN' || latestShiftStatus === 'CLOSED' || latestAttStatus === 'PRESENT');
      } else {
        isEmpActiveInPeriod = (Object.keys(workedDaysSet).length > 0 || empFilteredTasks.length > 0 || latestShiftStatus === 'OPEN' || latestShiftStatus === 'CLOSED');
      }
      if (isEmpActiveInPeriod) {
        clockedInCount++;
      }

      // Aggregate Work Logs for this Employee
      var empActualTaskHrs = 0;
      var empActBreakdown = {};
      ALLOWED_ACTIVITIES.forEach(function(act) { empActBreakdown[act] = 0; });
      var empClientHours = {};
      var empProjHours = {};

      for (var tskIdx = 0; tskIdx < empFilteredTasks.length; tskIdx++) {
        var tsk = empFilteredTasks[tskIdx];
        var hrs = Number(tsk.actual_worked_hours) || 0;
        var act = tsk.activity_type || 'Production';
        var cl = String(tsk.active_client || 'General').trim();
        var pr = String(tsk.active_project || 'Standard').trim();

        empActualTaskHrs += hrs;
        empActBreakdown[act] = (empActBreakdown[act] || 0) + hrs;
        empClientHours[cl] = (empClientHours[cl] || 0) + hrs;
        empProjHours[pr] = (empProjHours[pr] || 0) + hrs;

        // Global breakdowns
        activityHoursMap[act] = (activityHoursMap[act] || 0) + hrs;
        clientHoursMap[cl] = (clientHoursMap[cl] || 0) + hrs;
        var projKey = cl + '|||' + pr;
        projectHoursMap[projKey] = (projectHoursMap[projKey] || 0) + hrs;
      }

      totalShiftWorkingHours += empNetShiftHours;
      totalLoggedWorkHours += empActualTaskHrs;

      // Available Production Hours & Headcount calculation:
      // Support Headcount: calculated from records where activity_type === 'Support' or role is Support.
      // Support headcount is NOT counted toward Production Hours.
      // Production Hours calculated using only employees assigned to Production, Idle, Down Time, Client Meeting, or Internal Meeting.
      // Exclude Support and Training from Production Hours.
      // Resources marked WO or Holiday are excluded from Production Head Count for affected dates.
      var workedDaysCount = Object.keys(workedDaysSet).length;
      if (workedDaysCount === 0 && empFilteredTasks.length > 0) workedDaysCount = 1;

      var hasSupportTask = (empActBreakdown['Support'] > 0);
      var hasProdTask = (empActBreakdown['Production'] > 0 || empActBreakdown['Idle'] > 0 || empActBreakdown['Down Time'] > 0 || empActBreakdown['Client Meeting'] > 0 || empActBreakdown['Internal Meeting'] > 0);
      var onlyTraining = !hasSupportTask && !hasProdTask && (empActBreakdown['Training (Material Go-Through)'] > 0);
      var isScheduleOffEmp = (latestShiftStatus === 'WO' || latestShiftStatus === 'HOLIDAY');

      if (!isScheduleOffEmp && workedDaysCount > 0) {
        if (hasSupportTask) {
          supportHeadcountDays += workedDaysCount;
        }
        if (hasProdTask) {
          prodHeadcountDays += workedDaysCount;
          var stdHours = getShiftStandardHours(empShiftType);
          totalAvailablePeriodProdHours += (stdHours * workedDaysCount);
        }
        // If neither specific task logged, fall back to designation/role:
        if (!hasSupportTask && !hasProdTask && !onlyTraining) {
          if (isSupportStaff) {
            supportHeadcountDays += workedDaysCount;
          } else {
            prodHeadcountDays += workedDaysCount;
            var stdHours = getShiftStandardHours(empShiftType);
            totalAvailablePeriodProdHours += (stdHours * workedDaysCount);
          }
        }
      }

      // Find top client and project for this employee
      var topClient = u.default_client || u.client || 'N/A';
      var topClientHrs = -1;
      for (var c in empClientHours) {
        if (empClientHours[c] > topClientHrs) {
          topClientHrs = empClientHours[c];
          topClient = c;
        }
      }

      var topProj = u.default_project || u.project || 'N/A';
      var topProjHrs = -1;
      for (var p in empProjHours) {
        if (empProjHours[p] > topProjHrs) {
          topProjHrs = empProjHours[p];
          topProj = p;
        }
      }

      employeeStatusList.push({
        emp_id: userEmpId,
        emp_name: userEmpName,
        email: String(u.email || ''),
        team: userTeam,
        work_mode: String(u.work_mode || 'Office'),
        app_role: String(u.app_role || 'User'),
        job_designation: userDesig || (isSupportStaff ? 'Support' : 'Associate'),
        role_type: isSupportStaff ? 'Support' : 'Production',
        shift_type: empShiftType,
        shift_status: latestShiftStatus,
        attendance_status: latestAttStatus || latestShiftStatus,
        shift_start_time: latestShiftStart,
        shift_end_time: latestShiftEnd,
        days_worked: Object.keys(workedDaysSet).length,
        days_leave: leaveDaysCount,
        leave_hours: Number(empLeaveHours.toFixed(2)),
        net_shift_hours: Number(empNetShiftHours.toFixed(2)),
        total_shift_hours: Number(empNetShiftHours.toFixed(2)), // Alias for compatibility
        total_task_hours: Number(empActualTaskHrs.toFixed(2)),
        actual_worked_hours: Number(empActualTaskHrs.toFixed(2)), // Alias for compatibility
        activity_breakdown: empActBreakdown,
        primary_client: topClient,
        primary_project: topProj,
        tasks_count: empFilteredTasks.length,
        tasks: empFilteredTasks
      });
    }

    // Format Breakdown Lists
    var timeByActivity = ALLOWED_ACTIVITIES.map(function(act) {
      var hrs = Number((activityHoursMap[act] || 0).toFixed(2));
      var pct = totalLoggedWorkHours > 0 ? Number(((hrs / totalLoggedWorkHours) * 100).toFixed(1)) : 0;
      return { activity: act, hours: hrs, percentage: pct };
    });

    var timeByClient = Object.keys(clientHoursMap).map(function(cl) {
      var hrs = Number((clientHoursMap[cl] || 0).toFixed(2));
      var pct = totalLoggedWorkHours > 0 ? Number(((hrs / totalLoggedWorkHours) * 100).toFixed(1)) : 0;
      return { client: cl, hours: hrs, percentage: pct };
    }).sort(function(a, b) { return b.hours - a.hours; });

    var timeByProject = Object.keys(projectHoursMap).map(function(cpKey) {
      var parts = cpKey.split('|||');
      var cl = parts[0];
      var pr = parts[1] || '';
      var hrs = Number((projectHoursMap[cpKey] || 0).toFixed(2));
      var pct = totalLoggedWorkHours > 0 ? Number(((hrs / totalLoggedWorkHours) * 100).toFixed(1)) : 0;
      return { client: cl, project: pr, hours: hrs, percentage: pct };
    }).sort(function(a, b) { return b.hours - a.hours; });

    var productionHours = Number((activityHoursMap['Production'] || 0).toFixed(2));
    var meetingHours = Number(((activityHoursMap['Client Meeting'] || 0) + (activityHoursMap['Internal Meeting'] || 0)).toFixed(2));
    var trainingHours = Number((activityHoursMap['Training (Material Go-Through)'] || 0).toFixed(2));
    var downtimeIdleHours = Number(((activityHoursMap['Idle'] || 0) + (activityHoursMap['Down Time'] || 0)).toFixed(2));

    // BUILD EXECUTIVE PRODUCTIVITY & HEADCOUNT SUMMARY TABLE (Matching Spreadsheet Reference)
    var summaryTable = buildProductivitySummaryTable(
      allQueryTasks,
      allAttendanceRows,
      users,
      period,
      targetDate,
      startDate,
      endDate,
      selectedClient,
      selectedProject,
      selectedTeam,
      selectedShift,
      timeZone
    );

    var totalAvailProdHrs = summaryTable.total ? summaryTable.total.totalAvailableProdHrs : Number(totalAvailablePeriodProdHours.toFixed(2));
    var prodUtilizationPct = totalAvailProdHrs > 0 ? Number(((productionHours / totalAvailProdHrs) * 100).toFixed(2)) : 0;

    return {
      success: true,
      targetDate: targetDate,
      period: period,
      dateRange: dateRange,
      selectedShift: selectedShift,
      summary: {
        totalEmployees: employeeStatusList.length,
        clockedInCount: clockedInCount,
        leaveCount: distinctLeaveStaffCount,
        totalLeaveHours: Number(totalPeriodLeaveHours.toFixed(2)),
        totalShiftHours: Number(totalShiftWorkingHours.toFixed(2)),
        totalLoggedHours: Number(totalLoggedWorkHours.toFixed(2)),
        totalTaskHours: Number(totalLoggedWorkHours.toFixed(2)),
        productionHours: productionHours,
        supportHours: Number((activityHoursMap['Support'] || 0).toFixed(2)),
        totalAvailableProductionHours: totalAvailProdHrs,
        productionUtilizationPct: prodUtilizationPct,
        meetingHours: meetingHours,
        trainingHours: trainingHours,
        downtimeIdleHours: downtimeIdleHours,
        avgProdHeadCount: summaryTable.total ? summaryTable.total.prodHeadCount : 0,
        avgSupportHeadCount: summaryTable.total ? summaryTable.total.supportHeadCount : 0,
        // Backward compatibility properties
        notClockedInCount: Math.max(0, employeeStatusList.length - clockedInCount - distinctLeaveStaffCount),
        onBreakCount: 0
      },
      summaryTable: summaryTable,
      timeByActivity: timeByActivity,
      timeByClient: timeByClient,
      timeByProject: timeByProject,
      employeeStatusList: employeeStatusList,
      leaveQueue: leaveQueue,
      clientMaster: clientMaster,
      rawTaskLogs: rawTaskLogs,
      spreadsheetUrl: ss ? ss.getUrl() : ''
    };
  } catch (err) {
    return { success: false, message: 'Admin Data Error: ' + err.toString() };
  }
}

/**
 * BUILDS EXECUTIVE PRODUCTIVITY & HEADCOUNT SUMMARY TABLE
 * Directly modeled from the spreadsheet reference.
 * Formulas:
 * - General Shift: Avg Prod HC * 8 hrs * Working Days
 * - 1st, Afternoon, 2nd: Avg Prod HC * 7 hrs * Working Days
 * - Prod Utilization % = (Active Prod Hrs / Total Available Prod Hrs) * 100
 * - Non Productive % = ((Training + Idle + Down Time + Meeting) / Total Available Prod Hrs) * 100
 */
function buildProductivitySummaryTable(
  allTasks,
  allAttendance,
  users,
  period,
  targetDate,
  startDate,
  endDate,
  selectedClient,
  selectedProject,
  selectedTeam,
  selectedShift,
  timeZone
) {
  timeZone = timeZone || Session.getScriptTimeZone();
  var rows = [];

  // 1. Build lookup for users
  var userMap = {};
  var userList = [];
  for (var u = 0; u < users.length; u++) {
    var usr = users[u];
    var empKey = String(usr.emp_id || '').trim().toLowerCase();
    var desig = String(usr.job_designation || usr.designation || '').toUpperCase();
    var role = String(usr.app_role || usr.role || '').toUpperCase();
    var tm = String(usr.team || '').toUpperCase();

    var isSupport = (desig && ['TL', 'LEAD', 'MANAGER', 'QA', 'SUPPORT', 'TRAINER', 'SME', 'OPS'].indexOf(desig) !== -1) ||
                    (role && ['ADMIN', 'MANAGER', 'LEAD', 'MANAGEMENT'].indexOf(role) !== -1) ||
                    (tm === 'MANAGEMENT' || tm === 'SUPPORT');
    var isProd = !isSupport;

    var shift = normalizeShiftType(usr.default_shift || 'General');

    // Filter by team
    if (selectedTeam !== 'ALL' && tm !== selectedTeam.toUpperCase()) continue;
    // Filter by shift
    if (selectedShift !== 'ALL' && shift.toUpperCase() !== selectedShift.toUpperCase()) continue;

    var uInfo = {
      emp_id: usr.emp_id,
      emp_name: usr.emp_name,
      team: usr.team || 'Annotation',
      isSupport: isSupport,
      isProd: isProd,
      defaultClient: usr.default_client || 'iMerit',
      defaultProject: usr.default_project || 'AMBALL',
      shiftType: shift
    };
    userMap[empKey] = uInfo;
    userList.push(uInfo);
  }

  // 2. Identify slices and attendance mapping
  var sliceMap = {};
  var empAttendanceDaily = {}; // dateStr -> empKey -> attRow

  for (var a = 0; a < allAttendance.length; a++) {
    var att = allAttendance[a];
    var aDate = att.normalizedDate;
    if (!aDate) continue;
    var aEmp = String(att.emp_id || '').trim().toLowerCase();
    if (!userMap[aEmp]) continue;

    if (!empAttendanceDaily[aDate]) empAttendanceDaily[aDate] = {};
    empAttendanceDaily[aDate][aEmp] = att;
  }

  function getSliceInfo(dateStr) {
    var dParts = dateStr.split('-');
    var yr = parseInt(dParts[0], 10);
    var mo = parseInt(dParts[1], 10) - 1;
    var dy = parseInt(dParts[2], 10);
    var dObj = new Date(yr, mo, dy);

    if (period === 'month') {
      var moShort = Utilities.formatDate(dObj, timeZone, 'MMM');
      var moKey = dParts[0] + '-' + dParts[1];
      var label = moShort;
      var todayNormalized = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd');
      if (moKey === todayNormalized.substring(0, 7) && dy > 1) {
        var daySuffix = (dy === 1 || dy === 21 || dy === 31) ? 'st' : ((dy === 2 || dy === 22) ? 'nd' : ((dy === 3 || dy === 23) ? 'rd' : 'th'));
        label = moShort + ' till ' + dy + daySuffix;
      }
      return { key: moKey, label: label, sortOrder: moKey };
    } else if (period === 'week') {
      var wRange = getDateRangeForPeriod(dateStr, 'week', timeZone);
      return { key: wRange.startDate, label: wRange.label, sortOrder: wRange.startDate };
    } else {
      return { key: dateStr, label: Utilities.formatDate(dObj, timeZone, 'dd MMM yyyy'), sortOrder: dateStr };
    }
  }

  function getDatesForSlice(slKey, slLabel, periodType, dtStart, dtEnd) {
    var dates = [];
    if (periodType === 'day') {
      dates.push(slKey);
      return dates;
    }
    if (periodType === 'week') {
      var p = slKey.split('-');
      var baseDate = new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10));
      for (var i = 0; i < 7; i++) {
        var cur = new Date(baseDate.getTime() + (i * 24 * 60 * 60 * 1000));
        dates.push(Utilities.formatDate(cur, timeZone, 'yyyy-MM-dd'));
      }
      return dates;
    }
    if (periodType === 'month') {
      var parts = slKey.split('-');
      var yr = parseInt(parts[0], 10);
      var mo = parseInt(parts[1], 10); // 1-12
      var tillMatch = slLabel && slLabel.match(/till\s+(\d+)/i);
      var maxDays = tillMatch ? parseInt(tillMatch[1], 10) : new Date(yr, mo, 0).getDate();
      for (var d = 1; d <= maxDays; d++) {
        var dStr = (d < 10 ? '0' : '') + d;
        var moStr = (mo < 10 ? '0' : '') + mo;
        dates.push(yr + '-' + moStr + '-' + dStr);
      }
      return dates;
    }
    if (dtStart && dtEnd) {
      var p1 = dtStart.split('-');
      var p2 = dtEnd.split('-');
      var d1 = new Date(parseInt(p1[0], 10), parseInt(p1[1], 10) - 1, parseInt(p1[2], 10));
      var d2 = new Date(parseInt(p2[0], 10), parseInt(p2[1], 10) - 1, parseInt(p2[2], 10));
      var curD = new Date(d1.getTime());
      while (curD <= d2) {
        dates.push(Utilities.formatDate(curD, timeZone, 'yyyy-MM-dd'));
        curD.setTime(curD.getTime() + 24 * 60 * 60 * 1000);
      }
      if (dates.length > 0) return dates;
    }
    dates.push(slKey);
    return dates;
  }

  // Populate tasks into slices
  for (var t = 0; t < allTasks.length; t++) {
    var tsk = allTasks[t];
    var tDate = tsk.normalizedDate;
    if (!tDate) continue;

    var tEmp = String(tsk.emp_id || '').trim().toLowerCase();
    var usr = userMap[tEmp];
    if (!usr) continue;

    var clientName = String(tsk.active_client || usr.defaultClient || 'iMerit').trim();
    var projName = String(tsk.active_project || usr.defaultProject || 'AMBALL').trim();

    if (selectedClient !== 'ALL' && clientName.toUpperCase() !== selectedClient.toUpperCase()) continue;
    if (selectedProject !== 'ALL' && projName.toUpperCase() !== selectedProject.toUpperCase()) continue;

    var projDisplay = clientName + ' - ' + projName;

    var sInfo = getSliceInfo(tDate);
    if (!sliceMap[sInfo.key]) {
      sliceMap[sInfo.key] = { label: sInfo.label, sortOrder: sInfo.sortOrder, projects: {} };
    }
    if (!sliceMap[sInfo.key].projects[projDisplay]) {
      sliceMap[sInfo.key].projects[projDisplay] = {
        client: clientName,
        project: projName,
        tasks: [],
        tasksByDateAndEmp: {}
      };
    }

    var pObj = sliceMap[sInfo.key].projects[projDisplay];
    pObj.tasks.push(tsk);
    if (!pObj.tasksByDateAndEmp[tDate]) pObj.tasksByDateAndEmp[tDate] = {};
    if (!pObj.tasksByDateAndEmp[tDate][tEmp]) pObj.tasksByDateAndEmp[tDate][tEmp] = [];
    pObj.tasksByDateAndEmp[tDate][tEmp].push(tsk);
  }

  // Baseline fallback if sliceMap is empty or projects need registration
  var activeDateStr = targetDate || Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd');
  var activeSlice = getSliceInfo(activeDateStr);
  if (!sliceMap[activeSlice.key]) {
    sliceMap[activeSlice.key] = { label: activeSlice.label, sortOrder: activeSlice.sortOrder, projects: {} };
  }

  var hasFilteredProject = (selectedProject !== 'ALL' || selectedClient !== 'ALL');
  if (hasFilteredProject) {
    for (var uIdx = 0; uIdx < userList.length; uIdx++) {
      var uInf = userList[uIdx];
      var pDisplay = (uInf.defaultClient || 'iMerit') + ' - ' + (uInf.defaultProject || 'AMBALL');
      if (selectedClient !== 'ALL' && uInf.defaultClient.toUpperCase() !== selectedClient.toUpperCase()) continue;
      if (selectedProject !== 'ALL' && uInf.defaultProject.toUpperCase() !== selectedProject.toUpperCase()) continue;

      for (var sKey in sliceMap) {
        if (!sliceMap[sKey].projects[pDisplay]) {
          sliceMap[sKey].projects[pDisplay] = {
            client: uInf.defaultClient || 'iMerit',
            project: uInf.defaultProject || 'AMBALL',
            tasks: [],
            tasksByDateAndEmp: {}
          };
        }
      }
    }
  }

  // Ensure active slice has at least one project
  for (var sKey in sliceMap) {
    if (Object.keys(sliceMap[sKey].projects).length === 0) {
      var defClient = selectedClient !== 'ALL' ? selectedClient : 'iMerit';
      var defProj = selectedProject !== 'ALL' ? selectedProject : 'AMBALL';
      var defKey = defClient + ' - ' + defProj;
      sliceMap[sKey].projects[defKey] = {
        client: defClient,
        project: defProj,
        tasks: [],
        tasksByDateAndEmp: {}
      };
    }
  }

  // 3. Compute metric rows across all slices and projects
  var sortedSliceKeys = Object.keys(sliceMap).sort(function(a, b) {
    return sliceMap[a].sortOrder.localeCompare(sliceMap[b].sortOrder);
  });

  var totalProdHCSum = 0;
  var totalSupportHCSum = 0;
  var totalAvailHoursSum = 0;
  var totalActiveProdHoursSum = 0;
  var totalTrainingHoursSum = 0;
  var totalIdleHoursSum = 0;
  var totalDownTimeHoursSum = 0;
  var totalMeetingHoursSum = 0;
  var totalLeaveHoursSum = 0;

  for (var k = 0; k < sortedSliceKeys.length; k++) {
    var slKey = sortedSliceKeys[k];
    var slObj = sliceMap[slKey];
    var datesInSlice = getDatesForSlice(slKey, slObj.label, period, startDate, endDate);
    var totalDaysCount = datesInSlice.length || 1;

    for (var pName in slObj.projects) {
      var prj = slObj.projects[pName];
      var sumDailyProdHC = 0;
      var sumDailySupportHC = 0;
      var sumDailyAvailProdHrs = 0;
      var prjLeaveHrs = 0;

      // Iterate through EVERY calendar day in the selected range
      // Days with 0 employee logins contribute 0 headcount and are NOT omitted from the average divisor
      for (var d = 0; d < datesInSlice.length; d++) {
        var dtStr = datesInSlice[d];
        var dailyProdHC = 0;
        var dailySupportHC = 0;
        var dailyAvailProdHrs = 0;

        for (var e = 0; e < userList.length; e++) {
          var usr = userList[e];
          var empIdLower = String(usr.emp_id || '').trim().toLowerCase();

          var uTasks = (prj.tasksByDateAndEmp[dtStr] && prj.tasksByDateAndEmp[dtStr][empIdLower]) ? prj.tasksByDateAndEmp[dtStr][empIdLower] : [];
          
          function cleanMatchStr(str) {
            return String(str || '').toUpperCase().replace(/[\s\-_]/g, '');
          }

          var belongsToProject = (uTasks.length > 0) ||
            ((selectedProject === 'ALL' || cleanMatchStr(usr.defaultProject) === cleanMatchStr(prj.project)) &&
             (selectedClient === 'ALL' || cleanMatchStr(usr.defaultClient) === cleanMatchStr(prj.client)));

          if (!belongsToProject) continue;

          var att = (empAttendanceDaily[dtStr] && empAttendanceDaily[dtStr][empIdLower]) ? empAttendanceDaily[dtStr][empIdLower] : null;
          var attStatus = att ? String(att.attendance_status || '').toUpperCase() : '';
          var shStatus = att ? String(att.shift_status || '').toUpperCase() : '';
          var empShift = (att && att.shift_type) ? normalizeShiftType(att.shift_type) : usr.shiftType;
          var stdH = getShiftStandardHours(empShift);

          // 1. Week Off or Holiday: strictly EXCLUDE from Head Count and Production Hours ONLY IF NO WORK LOGS RECORDED
          var isWOorHoliday = (attStatus === 'WO' || shStatus === 'WO' || attStatus === 'HOLIDAY' || shStatus === 'HOLIDAY');
          if (isWOorHoliday && uTasks.length === 0) {
            continue;
          }

          // 2. Full Day Leave: strictly EXCLUDE from active Head Count and Production Hours ONLY IF NO WORK LOGS RECORDED
          var isFullLeave = (uTasks.length === 0) && (['PL', 'UL', 'ABSENT', 'LEAVE_CLOSED'].indexOf(attStatus) !== -1 || shStatus === 'LEAVE_CLOSED');
          if (isFullLeave) {
            var lHrs = (att && Number(att.leave_hours) > 0) ? Number(att.leave_hours) : stdH;
            prjLeaveHrs += lHrs;
            continue;
          }

          // 3. Partial Leave / Overlapping Leave Hours
          if (attStatus === 'HL') {
            prjLeaveHrs += (stdH / 2.0);
          } else if (attStatus === 'PERMISSION' || attStatus === 'PERM') {
            prjLeaveHrs += (att && Number(att.permission_hours) > 0 ? Number(att.permission_hours) : 2.0);
          } else if (att && Number(att.leave_hours) > 0 && uTasks.length > 0) {
            var actualTaskHrs = uTasks.reduce(function(acc, t) { return acc + (Number(t.actual_worked_hours) || 0); }, 0);
            var effectiveLeave = Math.max(0, stdH - actualTaskHrs);
            prjLeaveHrs += effectiveLeave;
          }

          // 4. Activity qualification:
          // Support Head Count: calculated from records where activity_type === 'Support'
          // Support headcount does NOT count toward Production Hours
          // Production Hours calculated using ONLY employees assigned to Production, Idle, Down Time, Client Meeting, or Internal Meeting
          // Exclude Support and Training from Production Hours
          if (uTasks.length > 0) {
            var hasSupport = uTasks.some(function(t) { return t.activity_type === 'Support'; });
            var hasProdQual = uTasks.some(function(t) {
              return ['Production', 'Idle', 'Down Time', 'Client Meeting', 'Internal Meeting'].indexOf(t.activity_type) !== -1;
            });
            var onlyTraining = !hasSupport && !hasProdQual && uTasks.some(function(t) {
              return t.activity_type === 'Training (Material Go-Through)';
            });

            if (hasSupport && !hasProdQual) {
              dailySupportHC++;
            }
            if (hasProdQual) {
              dailyProdHC++;
              dailyAvailProdHrs += stdH;
            }
          } else {
            var isClocked = att && (shStatus === 'OPEN' || shStatus === 'CLOSED' || attStatus === 'PRESENT');
            if (isClocked) {
              if (usr.isSupport) {
                dailySupportHC++;
              } else {
                dailyProdHC++;
                dailyAvailProdHrs += stdH;
              }
            }
          }
        }

        sumDailyProdHC += dailyProdHC;
        sumDailySupportHC += dailySupportHC;
        sumDailyAvailProdHrs += dailyAvailProdHrs;
      }

      var avgProdHC = Number((sumDailyProdHC / totalDaysCount).toFixed(2));
      var avgSupportHC = Number((sumDailySupportHC / totalDaysCount).toFixed(2));
      var availProdHours = Number(sumDailyAvailProdHrs.toFixed(2));

      // Benchmark fallback if employees are scheduled
      if (availProdHours === 0 && avgProdHC > 0) {
        var defaultStdHrs = getShiftStandardHours(selectedShift !== 'ALL' ? selectedShift : 'General');
        availProdHours = Number((avgProdHC * defaultStdHrs * totalDaysCount).toFixed(2));
      }

      var activeProdHrs = 0;
      var trainingHrs = 0;
      var idleHrs = 0;
      var downTimeHrs = 0;
      var meetingHrs = 0;

      for (var tk = 0; tk < prj.tasks.length; tk++) {
        var tTask = prj.tasks[tk];
        var actType = tTask.activity_type || 'Production';
        var actHrs = Number(tTask.actual_worked_hours) || 0;

        if (actType === 'Production') {
          activeProdHrs += actHrs;
        } else if (actType === 'Training (Material Go-Through)') {
          trainingHrs += actHrs;
        } else if (actType === 'Idle') {
          idleHrs += actHrs;
        } else if (actType === 'Down Time') {
          downTimeHrs += actHrs;
        } else if (actType === 'Client Meeting' || actType === 'Internal Meeting') {
          meetingHrs += actHrs;
        }
      }

      activeProdHrs = Number(activeProdHrs.toFixed(2));
      trainingHrs = Number(trainingHrs.toFixed(2));
      idleHrs = Number(idleHrs.toFixed(2));
      downTimeHrs = Number(downTimeHrs.toFixed(2));
      meetingHrs = Number(meetingHrs.toFixed(2));
      var leaveHrs = Number(prjLeaveHrs.toFixed(2));

      // Distinct Task Contributor Safeguard:
      // If qualifying productive work logs exist on this project/date slice, Head Count and Available Prod Hours must NEVER be 0
      if (avgProdHC === 0 && (activeProdHrs > 0 || idleHrs > 0 || meetingHrs > 0 || downTimeHrs > 0)) {
        var distinctEmps = {};
        for (var tk2 = 0; tk2 < prj.tasks.length; tk2++) {
          var tTk2 = prj.tasks[tk2];
          var aTyp2 = tTk2.activity_type || 'Production';
          if (['Production', 'Idle', 'Down Time', 'Client Meeting', 'Internal Meeting'].indexOf(aTyp2) !== -1) {
            distinctEmps[String(tTk2.emp_id || '').trim().toLowerCase()] = true;
          }
        }
        var countDistinct = Object.keys(distinctEmps).length;
        if (countDistinct > 0) {
          avgProdHC = Number((countDistinct / totalDaysCount).toFixed(2));
          var defaultShiftH = getShiftStandardHours(selectedShift !== 'ALL' ? selectedShift : 'General');
          availProdHours = Number((avgProdHC * defaultShiftH * totalDaysCount).toFixed(2));
        }
      }

      var prodUtilPct = availProdHours > 0 ? Number(((activeProdHrs / availProdHours) * 100).toFixed(2)) : 0;
      var nonProdPct = availProdHours > 0 ? Number((((trainingHrs + idleHrs + downTimeHrs + meetingHrs) / availProdHours) * 100).toFixed(2)) : 0;

      // Skip empty/phantom rows if not explicitly filtered and no activity/headcount recorded
      if (selectedProject === 'ALL' && selectedClient === 'ALL' && prj.tasks.length === 0 && avgProdHC === 0 && avgSupportHC === 0 && availProdHours === 0 && activeProdHrs === 0) {
        continue;
      }

      rows.push({
        period: slObj.label,
        projectName: pName,
        prodHeadCount: avgProdHC,
        supportHeadCount: avgSupportHC,
        totalAvailableProdHrs: availProdHours,
        activeProductionHrs: activeProdHrs,
        trainingHrs: trainingHrs,
        idleHrs: idleHrs,
        downTimeHrs: downTimeHrs,
        meetingHrs: meetingHrs,
        leaveHrs: leaveHrs,
        productionUtilizationPct: prodUtilPct,
        nonProductiveHoursPct: nonProdPct
      });

      totalProdHCSum += avgProdHC;
      totalSupportHCSum += avgSupportHC;
      totalAvailHoursSum += availProdHours;
      totalActiveProdHoursSum += activeProdHrs;
      totalTrainingHoursSum += trainingHrs;
      totalIdleHoursSum += idleHrs;
      totalDownTimeHoursSum += downTimeHrs;
      totalMeetingHoursSum += meetingHrs;
      totalLeaveHoursSum += leaveHrs;
    }
  }

  if (rows.length === 0) {
    var defClient = selectedClient !== 'ALL' ? selectedClient : 'iMerit';
    var defProj = selectedProject !== 'ALL' ? selectedProject : 'AMBALL';
    rows.push({
      period: activeSlice.label,
      projectName: defClient + ' - ' + defProj,
      prodHeadCount: 0,
      supportHeadCount: 0,
      totalAvailableProdHrs: 0,
      activeProductionHrs: 0,
      trainingHrs: 0,
      idleHrs: 0,
      downTimeHrs: 0,
      meetingHrs: 0,
      leaveHrs: 0,
      productionUtilizationPct: 0,
      nonProductiveHoursPct: 0
    });
  }

  var rowCount = rows.length || 1;
  var overallAvgProdHC = Number((totalProdHCSum / rowCount).toFixed(2));
  var overallAvgSupportHC = Number((totalSupportHCSum / rowCount).toFixed(2));
  var overallAvailProdHrs = Number(totalAvailHoursSum.toFixed(2));
  var overallActiveProdHrs = Number(totalActiveProdHoursSum.toFixed(2));
  var overallTrainingHrs = Number(totalTrainingHoursSum.toFixed(2));
  var overallIdleHrs = Number(totalIdleHoursSum.toFixed(2));
  var overallDownTimeHrs = Number(totalDownTimeHoursSum.toFixed(2));
  var overallMeetingHrs = Number(totalMeetingHoursSum.toFixed(2));
  var overallLeaveHrs = Number(totalLeaveHoursSum.toFixed(2));

  var overallProdUtilPct = overallAvailProdHrs > 0 ? Number(((overallActiveProdHrs / overallAvailProdHrs) * 100).toFixed(2)) : 0;
  var overallNonProdPct = overallAvailProdHrs > 0 ? Number((((overallTrainingHrs + overallIdleHrs + overallDownTimeHrs + overallMeetingHrs) / overallAvailProdHrs) * 100).toFixed(2)) : 0;

  var totalRow = {
    period: 'Total',
    projectName: '',
    prodHeadCount: overallAvgProdHC,
    supportHeadCount: overallAvgSupportHC,
    totalAvailableProdHrs: overallAvailProdHrs,
    activeProductionHrs: overallActiveProdHrs,
    trainingHrs: overallTrainingHrs,
    idleHrs: overallIdleHrs,
    downTimeHrs: overallDownTimeHrs,
    meetingHrs: overallMeetingHrs,
    leaveHrs: overallLeaveHrs,
    productionUtilizationPct: overallProdUtilPct,
    nonProductiveHoursPct: overallNonProdPct
  };

  return {
    columns: [
      'Month', 'Project Name', 'Prod Head Count', 'Support Head Count',
      'Total Available Production Hrs', 'Active production Hours', 'Training (Client + Internal)',
      'Idle', 'Down Time', 'Meeting (Client + Internal)', 'Leave',
      'Production Utilization %', 'Non Productive Hours %'
    ],
    rows: rows,
    total: totalRow
  };
}

/**
 * ADMIN: Add/Update Client Project Benchmark
 */
function updateClientProjectMaster(payload) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAMES.CLIENT_MASTER);
    if (!sheet) return { success: false, message: 'Client Master sheet not found.' };

    var data = sheet.getDataRange().getValues();
    var headers = data[0];

    var clientIdx = headers.indexOf('client_name');
    var projIdx = headers.indexOf('project_name');
    var actIdx = headers.indexOf('activity_type');
    var ahtIdx = headers.indexOf('aht_minutes');
    var billableIdx = headers.indexOf('is_billable');
    var activeIdx = headers.indexOf('is_active');
    var notesIdx = headers.indexOf('notes');

    var clientVal = String(payload.client_name || '').trim();
    var projVal = String(payload.project_name || '').trim();
    var actVal = String(payload.activity_type || 'Production Work').trim();
    var ahtVal = Number(payload.aht_minutes) || 15;
    var billableVal = payload.is_billable ? true : false;
    var activeVal = payload.is_active !== undefined ? (payload.is_active ? true : false) : true;
    var notesVal = String(payload.notes || '').trim();

    for (var i = 1; i < data.length; i++) {
      var rClient = clientIdx !== -1 ? String(data[i][clientIdx] || '').trim().toLowerCase() : '';
      var rProj = projIdx !== -1 ? String(data[i][projIdx] || '').trim().toLowerCase() : '';
      var rAct = actIdx !== -1 ? String(data[i][actIdx] || '').trim().toLowerCase() : '';

      var matchClientProj = rClient === clientVal.toLowerCase() && rProj === projVal.toLowerCase();
      var matchAct = actIdx === -1 || !rAct || rAct === actVal.toLowerCase();

      if (matchClientProj && matchAct) {
        if (actIdx !== -1) sheet.getRange(i + 1, actIdx + 1).setValue(actVal);
        if (ahtIdx !== -1) sheet.getRange(i + 1, ahtIdx + 1).setValue(ahtVal);
        if (billableIdx !== -1) sheet.getRange(i + 1, billableIdx + 1).setValue(billableVal);
        if (activeIdx !== -1) sheet.getRange(i + 1, activeIdx + 1).setValue(activeVal);
        if (notesIdx !== -1) sheet.getRange(i + 1, notesIdx + 1).setValue(notesVal);
        return { success: true, message: 'Updated existing client project benchmark successfully.' };
      }
    }

    if (actIdx !== -1) {
      sheet.appendRow([clientVal, projVal, actVal, ahtVal, billableVal, activeVal, notesVal]);
    } else {
      sheet.appendRow([clientVal, projVal, ahtVal, billableVal, activeVal, notesVal]);
    }
    return { success: true, message: 'Added new client project benchmark successfully.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * HELPER: Initialize Database Tab Structure & Ensure Break_Logs Exists
 */
function ensureSheetStructureAndSeedData() {
  var ss = getSpreadsheet();

  // 1. User_Master
  var uSheet = ss.getSheetByName(SHEET_NAMES.USER_MASTER) || ss.insertSheet(SHEET_NAMES.USER_MASTER);
  if (uSheet.getLastRow() === 0) {
    uSheet.appendRow([
      'emp_id', 'email', 'password', 'emp_name', 'default_client', 'default_project',
      'app_role', 'is_active', 'team', 'work_mode', 'job_designation', 'reporting_manager', 'project_manager'
    ]);
    uSheet.appendRow(['CC1778', 'arun.prakash@hurix.ai', 'Hurix@123', 'Arunprakash B', 'iMerit', 'Pravah', 'Admin', true, 'Annotation', 'Office', 'Associates', 'Dinesh Kumar Rajagopal', 'Dinesh Kumar Rajagopal']);
    uSheet.appendRow(['CC1770', 'logesh.govindarajan@hurix.ai', 'Hurix@123', 'Logesh G', 'iMerit', 'Pravah', 'User', true, 'Annotation', 'Office', 'Associates', 'Arunprakash B', 'Dinesh Kumar Rajagopal']);
    uSheet.appendRow(['CC1696', 'abdul.basith@hurix.ai', 'Hurix@123', 'Abdul Basith', 'iMerit', 'Pravah', 'User', true, 'Annotation', 'Office', 'Associates', 'Arunprakash B', 'Dinesh Kumar Rajagopal']);
    uSheet.appendRow(['C1232', 'suryakumar.munusamy@hurix.com', 'Hurix@123', 'Surya Kumar', 'Floor', 'Floor', 'Admin', true, 'Management', 'Office', 'TL', 'Dinesh Kumar Rajagopal', 'Dinesh Kumar Rajagopal']);
    uSheet.appendRow(['CC1823', 'saran.karthik@hurix.ai', 'Hurix@123', 'Saran Karthik F', 'iMerit', 'Pravah', 'User', true, 'Annotation', 'Office', 'Associates', 'Arunprakash B', 'Dinesh Kumar Rajagopal']);
  }

  // 2. Attendance_Shift_Logs
  var aSheet = ss.getSheetByName(SHEET_NAMES.ATTENDANCE) || ss.insertSheet(SHEET_NAMES.ATTENDANCE);
  if (aSheet.getLastRow() === 0) {
    aSheet.appendRow([
      'log_id', 'date', 'emp_id', 'emp_name', 'attendance_status', 'permission_hours',
      'shift_start_time', 'shift_end_time', 'total_shift_hours', 'shrinkage_hours',
      'proof_url', 'shift_status', 'created_at', 'total_break_minutes',
      'active_client', 'active_project'
    ]);
  }

  // 3. Project_Task_Logs
  var tSheet = ss.getSheetByName(SHEET_NAMES.PROJECT_LOGS) || ss.insertSheet(SHEET_NAMES.PROJECT_LOGS);
  ensureProjectLogsHeaders(tSheet);

  // 4. Client_Project_Master
  var cSheet = ss.getSheetByName(SHEET_NAMES.CLIENT_MASTER) || ss.insertSheet(SHEET_NAMES.CLIENT_MASTER);
  if (cSheet.getLastRow() === 0) {
    cSheet.appendRow(['client_name', 'project_name', 'activity_type', 'aht_minutes', 'is_billable', 'is_active', 'notes']);
  }

  // 5. Break_Logs
  var bSheet = ss.getSheetByName(SHEET_NAMES.BREAK_LOGS) || ss.insertSheet(SHEET_NAMES.BREAK_LOGS);
  if (bSheet.getLastRow() === 0) {
    bSheet.appendRow([
      'break_id', 'log_id', 'emp_id', 'emp_name', 'date',
      'break_start_time', 'break_end_time', 'duration_minutes', 'break_status', 'created_at'
    ]);
  }
}

/**
 * CROSS-SPREADSHEET LIVE REAL-TIME SYNC ENGINE
 * Syncs every employee's daily metrics to external sheet or master spreadsheet
 * Tab Name: Daily_Summary_Records
 */
function syncExternalDailySummaryRecords(dateStr, empId) {
  if (!dateStr || !empId) return;

  try {
    var masterSs = getSpreadsheet();
    var timeZone = Session.getScriptTimeZone();
    var normalizedDate = normalizeDateStr(dateStr, timeZone);
    if (!normalizedDate) return;
    
    // 1. Fetch User Profile from User_Master
    var user = getUserByEmail(empId) || getUserByEmpId(empId) || { emp_id: empId, emp_name: 'Employee', team: 'Annotation', work_mode: 'Office' };

    // 2. Fetch Shift Log for dateStr + empId from Attendance_Shift_Logs
    var attSheet = masterSs.getSheetByName(SHEET_NAMES.ATTENDANCE);
    var shiftLog = null;
    if (attSheet) {
      var aData = attSheet.getDataRange().getValues();
      if (aData.length > 1) {
        var aHeaders = aData[0];
        var eIdx = aHeaders.indexOf('emp_id');
        var dIdx = aHeaders.indexOf('date');
        var targetEmpLower = String(empId).trim().toLowerCase();

        for (var i = aData.length - 1; i >= 1; i--) {
          var rDate = normalizeDateStr(aData[i][dIdx], timeZone);
          var rEmp = eIdx !== -1 ? String(aData[i][eIdx]).trim().toLowerCase() : '';
          if (rEmp === targetEmpLower && rDate === normalizedDate) {
            shiftLog = {};
            aHeaders.forEach(function(h, idx) { shiftLog[h] = aData[i][idx]; });
            break;
          }
        }
      }
    }

    // 3. Fetch Task Logs for dateStr + empId from Project_Task_Logs
    var tSheet = masterSs.getSheetByName(SHEET_NAMES.PROJECT_LOGS);
    var loggedTaskHrs = 0;
    if (tSheet) {
      var tData = tSheet.getDataRange().getValues();
      if (tData.length > 1) {
        var tHeaders = tData[0];
        var teIdx = tHeaders.indexOf('emp_id');
        var wsIdx = tHeaders.indexOf('work_start_time');
        var createdIdx = tHeaders.indexOf('created_at');
        var actHrsIdx = tHeaders.indexOf('actual_worked_hours');
        var targetEmpLower = String(empId).trim().toLowerCase();

        for (var t = 1; t < tData.length; t++) {
          var rawStart = wsIdx !== -1 ? tData[t][wsIdx] : (createdIdx !== -1 ? tData[t][createdIdx] : null);
          var tDate = normalizeDateStr(rawStart, timeZone);
          var rEmp = teIdx !== -1 ? String(tData[t][teIdx]).trim().toLowerCase() : '';

          if (rEmp === targetEmpLower && tDate === normalizedDate) {
            loggedTaskHrs += actHrsIdx !== -1 ? (Number(tData[t][actHrsIdx]) || 0) : 0;
          }
        }
      }
    }

    // Prepare Metrics
    var shiftStatus = shiftLog ? String(shiftLog.shift_status || 'NOT_CLOCKED_IN').toUpperCase() : 'NOT_CLOCKED_IN';
    var attStatus = shiftLog ? String(shiftLog.attendance_status || 'NOT_CLOCKED_IN').toUpperCase() : 'NOT_CLOCKED_IN';
    var isFullLeaveSync = (attStatus === 'PL' || attStatus === 'UL' || attStatus === 'ABSENT' || shiftStatus === 'LEAVE_CLOSED');
    var isHalfLeaveSync = (attStatus === 'HL');

    var startTimeStr = (!isFullLeaveSync && shiftLog && shiftLog.shift_start_time) ? (shiftLog.shift_start_time instanceof Date ? Utilities.formatDate(shiftLog.shift_start_time, timeZone, 'HH:mm:ss') : String(shiftLog.shift_start_time)) : '--:--';
    var endTimeStr = (!isFullLeaveSync && shiftLog && shiftLog.shift_end_time) ? (shiftLog.shift_end_time instanceof Date ? Utilities.formatDate(shiftLog.shift_end_time, timeZone, 'HH:mm:ss') : String(shiftLog.shift_end_time)) : '--:--';
    var breakMins = 60; // Fixed 1-hour break auto-deducted
    var rawNetWorked = shiftLog ? Number(shiftLog.total_shift_hours) || 0 : 0;
    var netWorkedHrs = isFullLeaveSync ? 0 : (isHalfLeaveSync ? Math.min(4.50, rawNetWorked) : rawNetWorked);
    var proofUrl = shiftLog ? (shiftLog.proof_url || '') : '';
    var teamName = user.team || (shiftLog ? shiftLog.team : '') || 'Annotation';
    var modeName = user.work_mode || (shiftLog ? shiftLog.work_mode : '') || 'Office';

    // 4. Open External Summary Spreadsheet or fallback to Master Spreadsheet
    var extSs = null;
    try {
      if (EXTERNAL_SUMMARY_CONFIG && EXTERNAL_SUMMARY_CONFIG.SPREADSHEET_ID) {
        extSs = SpreadsheetApp.openById(EXTERNAL_SUMMARY_CONFIG.SPREADSHEET_ID);
      }
    } catch (e) {
      Logger.log('Could not open external spreadsheet, falling back to master: ' + e.toString());
    }
    if (!extSs) {
      extSs = masterSs;
    }

    var tabName = (EXTERNAL_SUMMARY_CONFIG && EXTERNAL_SUMMARY_CONFIG.TAB_NAME) || 'Daily_Summary_Records';
    var extSheet = extSs.getSheetByName(tabName);
    if (!extSheet) {
      extSheet = extSs.insertSheet(tabName);
    }

    if (extSheet.getLastRow() === 0) {
      extSheet.appendRow([
        'date', 'emp_id', 'emp_name', 'team', 'work_mode',
        'shift_status', 'attendance_status', 'shift_start_time', 'shift_end_time',
        'break_minutes', 'net_worked_hours', 'earned_hours', 'leakage_hours',
        'proof_url', 'last_synced_at'
      ]);
    }

    var extData = extSheet.getDataRange().getValues();
    var extHeaders = extData[0];

    var exDateIdx = extHeaders.indexOf('date');
    var exEmpIdx = extHeaders.indexOf('emp_id');
    var targetRowIndex = -1;

    // NORMALIZED DATE & EMP ID LOOKUP FOR 100% IN-PLACE UPDATES
    for (var r = 1; r < extData.length; r++) {
      var cellDateStr = normalizeDateStr(extData[r][exDateIdx], timeZone);
      var cellEmpIdStr = String(extData[r][exEmpIdx]).trim().toLowerCase();
      var targetEmpIdStr = String(empId).trim().toLowerCase();

      if (cellDateStr === normalizedDate && cellEmpIdStr === targetEmpIdStr) {
        targetRowIndex = r + 1; // Update existing row in-place!
        break;
      }
    }

    var rowPayload = [
      normalizedDate, user.emp_id, user.emp_name, teamName, modeName,
      shiftStatus, attStatus, startTimeStr, endTimeStr,
      breakMins, netWorkedHrs, Number(loggedTaskHrs.toFixed(2)), 0,
      proofUrl, Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd HH:mm:ss')
    ];

    if (targetRowIndex !== -1) {
      extSheet.getRange(targetRowIndex, 1, 1, rowPayload.length).setValues([rowPayload]);
    } else {
      extSheet.appendRow(rowPayload);
    }
  } catch (err) {
    Logger.log('External Sync Error: ' + err.toString());
  }
}

/**
 * AUTOMATIC DEDUPLICATION & CLEANING ROUTINE
 * Removes duplicate rows in Daily_Summary_Records, keeping 1 clean row per (date + emp_id)
 */
function deduplicateDailySummaryRecords() {
  try {
    var masterSs = getSpreadsheet();
    var extSs = null;
    try {
      if (EXTERNAL_SUMMARY_CONFIG && EXTERNAL_SUMMARY_CONFIG.SPREADSHEET_ID) {
        extSs = SpreadsheetApp.openById(EXTERNAL_SUMMARY_CONFIG.SPREADSHEET_ID);
      }
    } catch(e) {}
    if (!extSs) extSs = masterSs;

    var tabName = (EXTERNAL_SUMMARY_CONFIG && EXTERNAL_SUMMARY_CONFIG.TAB_NAME) || 'Daily_Summary_Records';
    var extSheet = extSs.getSheetByName(tabName);
    if (!extSheet) return { success: false, message: 'Tab not found.' };

    var extData = extSheet.getDataRange().getValues();
    if (extData.length <= 1) return { success: true, message: 'Sheet is empty or has headers only.' };

    var timeZone = Session.getScriptTimeZone();
    var extHeaders = extData[0];
    var exDateIdx = extHeaders.indexOf('date');
    var exEmpIdx = extHeaders.indexOf('emp_id');

    var seenKeys = {};
    var removedCount = 0;

    for (var r = 1; r < extData.length; r++) {
      var cellDateStr = normalizeDateStr(extData[r][exDateIdx], timeZone);
      var cellEmpIdStr = String(extData[r][exEmpIdx]).trim().toLowerCase();
      var key = cellDateStr + '_' + cellEmpIdStr;

      if (!cellDateStr || !cellEmpIdStr || cellEmpIdStr === 'emp_id') {
        continue; // Skip invalid or duplicate headers
      }

      if (seenKeys[key]) {
        // Keeps the latest synced row by replacing
        seenKeys[key] = extData[r];
        removedCount++;
      } else {
        seenKeys[key] = extData[r];
      }
    }

    // Rebuild cleaned sheet
    var cleanedData = [extHeaders];
    Object.keys(seenKeys).forEach(function(k) {
      cleanedData.push(seenKeys[k]);
    });

    extSheet.clearContents();
    extSheet.getRange(1, 1, cleanedData.length, extHeaders.length).setValues(cleanedData);

    return {
      success: true,
      message: 'Cleaned up duplicate rows. Total unique records remaining: ' + (cleanedData.length - 1) + '. Removed duplicates: ' + removedCount,
      uniqueCount: cleanedData.length - 1
    };
  } catch (err) {
    return { success: false, message: 'Deduplication error: ' + err.toString() };
  }
}

/**
 * MASTER SHEET EDIT TRIGGER
 * Reflects direct master spreadsheet edits live to Daily_Summary_Records
 */
function onMasterSheetChange(e) {
  try {
    var ss = getSpreadsheet();
    var activeSheet = ss.getActiveSheet();
    var sheetName = activeSheet.getName();

    if (sheetName === SHEET_NAMES.BREAK_LOGS) {
      var activeRow = activeSheet.getActiveCell().getRow();
      if (activeRow > 1) {
        var logIdVal = activeSheet.getRange(activeRow, 2).getValue();
        if (logIdVal) {
          calculateAndSyncBreakDurations(ss, logIdVal);
        }
      }
      backfillAllHistoricalDailySummaryRecords();
      deduplicateDailySummaryRecords();
    } else if (sheetName === SHEET_NAMES.ATTENDANCE || sheetName === SHEET_NAMES.PROJECT_LOGS || sheetName === SHEET_NAMES.USER_MASTER) {
      backfillAllHistoricalDailySummaryRecords();
      deduplicateDailySummaryRecords();
    }
  } catch (err) {
    Logger.log('Master sheet change sync error: ' + err.toString());
  }
}

function getUserByEmpId(empId) {
  if (!empId) return null;
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.USER_MASTER);
  if (!sheet) return null;

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return null;

  var headers = data[0];
  var empIdx = -1;
  for (var h = 0; h < headers.length; h++) {
    var hn = String(headers[h]).trim().toLowerCase().replace(/_/g, '');
    if (hn === 'empid' || hn === 'emp id' || hn === 'id' || hn === 'employeeid') empIdx = h;
  }
  if (empIdx === -1) return null;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][empIdx]).trim().toLowerCase() === String(empId).trim().toLowerCase()) {
      var userObj = {};
      headers.forEach(function(h, idx) {
        var cleanHeader = String(h).trim();
        var val = data[i][idx];
        if (typeof val === 'string') val = val.trim();
        userObj[cleanHeader] = val;
        var lowerKey = cleanHeader.toLowerCase();
        userObj[lowerKey] = val;
        userObj[lowerKey.replace(/\s+/g, '_')] = val;
      });
      return userObj;
    }
  }
  return null;
}

/**
 * ONE-CLICK HISTORICAL & PRESENT BACKFILL ENGINE
 * Populates all past date records (July/August 2026 to present) into external sheet 1gF1SXav2UYAfygX9ZHR2dTteU5_Pawuoev-5MXtQmqA
 * Tab Name: Daily_Summary_Records
 */
function backfillAllHistoricalDailySummaryRecords() {
  try {
    var masterSs = getSpreadsheet();
    var timeZone = Session.getScriptTimeZone();
    var uSheet = masterSs.getSheetByName(SHEET_NAMES.USER_MASTER);
    if (!uSheet) return { success: false, message: 'User_Master not found.' };

    var uData = uSheet.getDataRange().getValues();
    if (uData.length <= 1) return { success: false, message: 'User_Master is empty.' };

    var uHeaders = uData[0];
    var empIdx = -1;
    for (var h = 0; h < uHeaders.length; h++) {
      var hn = String(uHeaders[h]).trim().toLowerCase();
      if (hn === 'emp_id' || hn === 'emp id' || hn === 'empid') empIdx = h;
    }

    var users = [];
    for (var i = 1; i < uData.length; i++) {
      var eId = empIdx !== -1 ? String(uData[i][empIdx]).trim() : '';
      if (eId) {
        var uObj = {};
        uHeaders.forEach(function(h, idx) { uObj[h] = uData[i][idx]; });
        users.push(uObj);
      }
    }

    // Collect all unique dates from Attendance_Shift_Logs & Project_Task_Logs
    var dateMap = {};
    var aSheet = masterSs.getSheetByName(SHEET_NAMES.ATTENDANCE);
    if (aSheet) {
      var aData = aSheet.getDataRange().getValues();
      if (aData.length > 1) {
        var aHeaders = aData[0];
        var dIdx = -1;
        for (var ah = 0; ah < aHeaders.length; ah++) {
          if (String(aHeaders[ah]).trim().toLowerCase() === 'date') dIdx = ah;
        }
        if (dIdx !== -1) {
          for (var a = 1; a < aData.length; a++) {
            var rDate = normalizeDateStr(aData[a][dIdx], timeZone);
            if (rDate && rDate.length === 10 && rDate.indexOf('-') !== -1) dateMap[rDate] = true;
          }
        }
      }
    }

    var tSheet = masterSs.getSheetByName(SHEET_NAMES.PROJECT_LOGS);
    if (tSheet) {
      var tData = tSheet.getDataRange().getValues();
      if (tData.length > 1) {
        var tHeaders = tData[0];
        var wsIdx = -1;
        for (var th = 0; th < tHeaders.length; th++) {
          var thn = String(tHeaders[th]).trim().toLowerCase();
          if (thn === 'work_start_time' || thn === 'date' || thn === 'created_at') wsIdx = th;
        }
        if (wsIdx !== -1) {
          for (var t = 1; t < tData.length; t++) {
            var rawStart = tData[t][wsIdx];
            var tDate = normalizeDateStr(rawStart, timeZone);
            if (tDate && tDate.length === 10 && tDate.indexOf('-') !== -1) dateMap[tDate] = true;
          }
        }
      }
    }

    // Always include today's date
    var todayStr = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd');
    dateMap[todayStr] = true;

    var sortedDates = Object.keys(dateMap).sort();
    var syncedCount = 0;

    sortedDates.forEach(function(dStr) {
      users.forEach(function(u) {
        var eId = u.emp_id || u['Emp ID'] || u.Emp_ID;
        syncExternalDailySummaryRecords(dStr, eId);
        syncedCount++;
      });
    });

    deduplicateDailySummaryRecords();

    return {
      success: true,
      message: 'Successfully populated ' + syncedCount + ' historical & present records across ' + sortedDates.length + ' dates into Daily_Summary_Records!',
      datesProcessed: sortedDates,
      syncedRecordsCount: syncedCount
    };
  } catch (err) {
    return { success: false, message: 'Backfill Error: ' + err.toString() };
  }
}

/**
 * PROGRAMMATIC INSTALLABLE ONCHANGE TRIGGER SETUP
 * Run this function ONCE in Apps Script to automatically install background change trigger!
 */
function setupAutomaticSheetChangeTrigger() {
  try {
    var ss = getSpreadsheet();
    var triggers = ScriptApp.getUserTriggers(ss);
    
    var changeInstalled = false;
    var nightlyInstalled = false;

    // Check if triggers are already installed
    for (var i = 0; i < triggers.length; i++) {
      var handler = triggers[i].getHandlerFunction();
      if (handler === 'handleAutomaticMasterSheetChange') {
        changeInstalled = true;
      }
      if (handler === 'autoCloseOvernightShifts') {
        nightlyInstalled = true;
      }
    }

    var messageParts = [];

    if (!changeInstalled) {
      ScriptApp.newTrigger('handleAutomaticMasterSheetChange')
        .forSpreadsheet(ss)
        .onChange()
        .create();
      messageParts.push('Automatic background live sync trigger installed.');
    } else {
      messageParts.push('Automatic background live sync trigger already exists.');
    }

    if (!nightlyInstalled) {
      ScriptApp.newTrigger('autoCloseOvernightShifts')
        .timeBased()
        .everyDays(1)
        .atHour(1) // Runs between 1:00 AM and 2:00 AM
        .create();
      messageParts.push('Nightly auto-close time-driven trigger installed.');
    } else {
      messageParts.push('Nightly auto-close time-driven trigger already exists.');
    }

    return { success: true, message: messageParts.join(' | ') };
  } catch (err) {
    return { success: false, message: 'Trigger installation error: ' + err.toString() };
  }
}

/**
 * AUTOMATIC BACKGROUND MASTER SHEET CHANGE ENGINE
 * Fires automatically 24/7 whenever data is edited in Master Sheet or submitted via Web App
 */
function handleAutomaticMasterSheetChange(e) {
  try {
    var ss = getSpreadsheet();
    var activeSheet = ss.getActiveSheet();
    if (!activeSheet) return;

    var sheetName = activeSheet.getName();

    // If change happens in Break Logs, Attendance, Task Logs, or User Master
    if (sheetName === SHEET_NAMES.BREAK_LOGS) {
      var activeRow = activeSheet.getActiveCell().getRow();
      if (activeRow > 1) {
        var logIdVal = activeSheet.getRange(activeRow, 2).getValue();
        if (logIdVal) {
          calculateAndSyncBreakDurations(ss, logIdVal);
        }
      }
      backfillAllHistoricalDailySummaryRecords();
      deduplicateDailySummaryRecords();
    } else if ([SHEET_NAMES.ATTENDANCE, SHEET_NAMES.PROJECT_LOGS, SHEET_NAMES.USER_MASTER].indexOf(sheetName) !== -1) {
      backfillAllHistoricalDailySummaryRecords();
      deduplicateDailySummaryRecords();
    }
  } catch (err) {
    Logger.log('Automatic background change sync error: ' + err.toString());
  }
}

/**
 * Calculates month-relative week string (e.g., 'Week 1', 'Week 2') matching formula:
 * =WEEKNUM(date) - WEEKNUM(EOMONTH(date, -1) + 1) + 1
 */
function getWeekOfMonthStr(dateStr) {
  if (!dateStr) return 'Week 1';
  var parts = String(dateStr).split('-');
  if (parts.length < 3) return 'Week 1';
  var y = parseInt(parts[0], 10);
  var m = parseInt(parts[1], 10) - 1;
  var d = parseInt(parts[2], 10);

  var targetDate = new Date(y, m, d);
  if (isNaN(targetDate.getTime())) return 'Week 1';

  var firstOfYear = new Date(y, 0, 1);
  var getWeekNum = function(dt) {
    var dayOfYear = Math.floor((dt.getTime() - firstOfYear.getTime()) / (1000 * 60 * 60 * 24)) + 1;
    return Math.ceil((dayOfYear + firstOfYear.getDay()) / 7);
  };

  var firstOfMonth = new Date(y, m, 1);
  var w = (getWeekNum(targetDate) - getWeekNum(firstOfMonth)) + 1;
  if (w < 1) w = 1;
  return 'Week ' + w;
}

