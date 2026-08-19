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

      // Calculate total break minutes and active break start time exclusively from Break_Logs
      if (shiftObj.log_id) {
        var totalBreakMins = calculateAndSyncBreakDurations(ss, shiftObj.log_id);
        shiftObj.total_break_minutes = totalBreakMins;

        // If the user is currently ON_BREAK, retrieve the break_start_time
        if (String(shiftObj.shift_status).toUpperCase() === 'ON_BREAK') {
          var breakSheet = ss.getSheetByName(SHEET_NAMES.BREAK_LOGS);
          if (breakSheet) {
            var bData = breakSheet.getDataRange().getValues();
            var bHeaders = bData[0];
            var bLogIdx = bHeaders.indexOf('log_id');
            var bStartIdx = bHeaders.indexOf('break_start_time');
            var bStatusIdx = bHeaders.indexOf('break_status');
            for (var b = 1; b < bData.length; b++) {
              if (bLogIdx !== -1 && String(bData[b][bLogIdx]).trim() === String(shiftObj.log_id).trim()) {
                var status = bStatusIdx !== -1 ? String(bData[b][bStatusIdx]).toUpperCase() : '';
                if (status === 'ACTIVE') {
                  var startVal = bStartIdx !== -1 ? bData[b][bStartIdx] : null;
                  if (startVal) {
                    shiftObj.break_start_time = startVal instanceof Date ? Utilities.formatDate(startVal, timeZone, 'HH:mm:ss') : String(startVal);
                  }
                }
              }
            }
          }
        }
      }

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

      // Ensure fallback properties for standard calculation fields if headers had older format
      if (task.activity_type === undefined && data[i][18] !== undefined) task.activity_type = data[i][18];
      if (task.idle_hours === undefined && data[i][19] !== undefined) task.idle_hours = data[i][19];
      if (task.bench_hours === undefined && data[i][20] !== undefined) task.bench_hours = data[i][20];
      if (task.blocked_hours === undefined && data[i][21] !== undefined) task.blocked_hours = data[i][21];

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

    var timeZone = Session.getScriptTimeZone();
    var todayStr = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd');
    var data = sheet.getDataRange().getValues();
    var headers = data[0];

    var empIdx = headers.indexOf('emp_id');
    var dateIdx = headers.indexOf('date');
    var statusIdx = headers.indexOf('shift_status');
    var startIdx = headers.indexOf('shift_start_time');
    var endIdx = headers.indexOf('shift_end_time');

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

          sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
          SpreadsheetApp.flush();
          syncExternalDailySummaryRecords(todayStr, user.emp_id);

          return {
            success: true,
            log_id: data[i][headers.indexOf('log_id')],
            shift_start_time: startTimeVal,
            shift_status: 'OPEN',
            permission_hours: existingPermHrs,
            attendance_status: attStatus,
            total_break_minutes: 60
          };
        }

        return {
          success: true,
          already_open: true,
          log_id: data[i][headers.indexOf('log_id')],
          shift_start_time: data[i][startIdx] instanceof Date ? Utilities.formatDate(data[i][startIdx], timeZone, 'HH:mm:ss') : String(data[i][startIdx]),
          shift_status: statusVal,
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

    sheet.appendRow(newRow);
    SpreadsheetApp.flush();
    syncExternalDailySummaryRecords(todayStr, user.emp_id);
    return { success: true, log_id: logId, shift_start_time: timeStr, shift_status: 'OPEN', total_break_minutes: 60 };
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
 * PAUSE / RESUME SHIFT & GRANULAR BREAK_LOGS RECORDING
 */
function toggleShiftPause(payload) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) {
    return { success: false, message: 'Server busy.' };
  }

  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAMES.ATTENDANCE);
    var breakSheet = ss.getSheetByName(SHEET_NAMES.BREAK_LOGS) || ss.insertSheet(SHEET_NAMES.BREAK_LOGS);

    var data = sheet.getDataRange().getValues();
    var headers = data[0];

    var logIdIdx = headers.indexOf('log_id');
    var statusIdx = headers.indexOf('shift_status');
    var totalBreakIdx = headers.indexOf('total_break_minutes');
    var breakStartIdx = headers.indexOf('break_start_time');
    var empIdIdx = headers.indexOf('emp_id');
    var empNameIdx = headers.indexOf('emp_name');
    var dateIdx = headers.indexOf('date');

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
    var timeStr = Utilities.formatDate(now, timeZone, 'HH:mm:ss');
    var todayStr = Utilities.formatDate(now, timeZone, 'yyyy-MM-dd');

    var empIdVal = data[rowIndex - 1][empIdIdx];
    var empNameVal = data[rowIndex - 1][empNameIdx];

    if (payload.action === 'PAUSE') {
      // Check if there is already an ACTIVE break for this log_id to prevent double logging
      var bData = breakSheet.getDataRange().getValues();
      if (bData.length > 1) {
        var bHeaders = bData[0];
        var bLogIdx = bHeaders.indexOf('log_id');
        var bStartIdx = bHeaders.indexOf('break_start_time');
        var bStatusIdx = bHeaders.indexOf('break_status');
        var bIdIdx = bHeaders.indexOf('break_id');

        for (var b = 1; b < bData.length; b++) {
          var matchLog = bLogIdx !== -1 && String(bData[b][bLogIdx]).trim() === String(payload.log_id).trim();
          var matchStatus = bStatusIdx !== -1 && String(bData[b][bStatusIdx]).toUpperCase() === 'ACTIVE';

          if (matchLog && matchStatus) {
            var existingStart = bStartIdx !== -1 ? bData[b][bStartIdx] : '';
            var existingStartStr = existingStart ? (existingStart instanceof Date ? Utilities.formatDate(existingStart, timeZone, 'HH:mm:ss') : String(existingStart)) : timeStr;
            var existingBreakId = bIdIdx !== -1 ? String(bData[b][bIdIdx]) : '';
            
            // Ensure status in Attendance sheet is ON_BREAK
            sheet.getRange(rowIndex, statusIdx + 1).setValue('ON_BREAK');
            SpreadsheetApp.flush();
            syncExternalDailySummaryRecords(todayStr, empIdVal);
            
            return { success: true, shift_status: 'ON_BREAK', break_start_time: existingStartStr, break_id: existingBreakId };
          }
        }
      }

      // 1. Update Attendance_Shift_Logs status to ON_BREAK
      sheet.getRange(rowIndex, statusIdx + 1).setValue('ON_BREAK');

      // 2. Append Active Row to Break_Logs matching user sheet headers
      var cleanEmpIdBrk = String(empIdVal || '').trim().replace(/[^a-zA-Z0-9]/g, '');
      var breakId = 'BRK-' + todayStr.replace(/-/g, '') + (cleanEmpIdBrk ? '-' + cleanEmpIdBrk : '') + '-' + Math.floor(100000 + Math.random() * 900000);
      var breakRow = [breakId, payload.log_id, empIdVal, empNameVal, todayStr, timeStr, '', 0, 'ACTIVE', new Date()];
      breakSheet.appendRow(breakRow);
      SpreadsheetApp.flush();
      syncExternalDailySummaryRecords(todayStr, empIdVal);

      return { success: true, shift_status: 'ON_BREAK', break_start_time: timeStr, break_id: breakId };
    } else {
      // RESUME ACTION
      var addedMins = 0;

      // 1. Locate and close ACTIVE break row in Break_Logs
      var bData = breakSheet.getDataRange().getValues();
      var bHeaders = bData[0];

      var bLogIdx = bHeaders.indexOf('log_id');
      var bStartIdx = bHeaders.indexOf('break_start_time');
      var bEndIdx = bHeaders.indexOf('break_end_time');
      
      // Support both singular and plural duration header variants
      var bDurIdx = bHeaders.indexOf('duration_minutes');
      if (bDurIdx === -1) {
        bDurIdx = bHeaders.indexOf('duration_minute');
      }
      
      var bStatusIdx = bHeaders.indexOf('break_status');
      var bCreatedAtIdx = bHeaders.indexOf('created_at');

      for (var b = bData.length - 1; b >= 1; b--) {
        var matchLog = bLogIdx !== -1 && String(bData[b][bLogIdx]).trim() === String(payload.log_id).trim();
        var matchStatus = bStatusIdx !== -1 && String(bData[b][bStatusIdx]).toUpperCase() === 'ACTIVE';

        if (matchLog && matchStatus) {
          var bStartVal = bStartIdx !== -1 ? bData[b][bStartIdx] : null;
          var bCreatedAtVal = bCreatedAtIdx !== -1 ? bData[b][bCreatedAtIdx] : null;

          var startTime = null;
          if (bCreatedAtVal instanceof Date && !isNaN(bCreatedAtVal.getTime())) {
            startTime = bCreatedAtVal;
          } else if (bStartVal) {
            startTime = parseTimeStringToDate(bStartVal, todayStr, timeZone);
          }

          if (startTime && !isNaN(startTime.getTime())) {
            var diffMs = now.getTime() - startTime.getTime();
            addedMins = Math.max(0.01, Math.round((diffMs / (1000 * 60)) * 100) / 100);
          }

          var bRowIndex = b + 1;
          if (bEndIdx !== -1) breakSheet.getRange(bRowIndex, bEndIdx + 1).setValue(timeStr);
          if (bDurIdx !== -1) breakSheet.getRange(bRowIndex, bDurIdx + 1).setValue(addedMins);
          if (bStatusIdx !== -1) breakSheet.getRange(bRowIndex, bStatusIdx + 1).setValue('COMPLETED');
          break;
        }
      }

      // 2. Sum all completed break durations for this log_id from Break_Logs
      var totalBreakMins = calculateAndSyncBreakDurations(ss, payload.log_id);

      sheet.getRange(rowIndex, statusIdx + 1).setValue('OPEN');
      if (totalBreakIdx !== -1) sheet.getRange(rowIndex, totalBreakIdx + 1).setValue(totalBreakMins);
      SpreadsheetApp.flush();
      syncExternalDailySummaryRecords(todayStr, empIdVal);

      return { success: true, shift_status: 'OPEN', total_break_minutes: totalBreakMins };
    }
  } catch (err) {
    return { success: false, message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * HELPER: Calculate and Sync Break Durations
 * Handles both duration_minutes and duration_minute column variants.
 * Calculates elapsed break minutes and syncs them to Attendance and Break_Logs.
 */
function calculateAndSyncBreakDurations(ss, logId) {
  var breakSheet = ss.getSheetByName(SHEET_NAMES.BREAK_LOGS);
  if (!breakSheet) return 0;
  
  var bData = breakSheet.getDataRange().getValues();
  if (bData.length <= 1) return 0;
  
  var bHeaders = bData[0];
  var bLogIdx = bHeaders.indexOf('log_id');
  var bStartIdx = bHeaders.indexOf('break_start_time');
  var bEndIdx = bHeaders.indexOf('break_end_time');
  
  // Support both singular and plural duration header variants
  var bDurIdx = bHeaders.indexOf('duration_minutes');
  if (bDurIdx === -1) {
    bDurIdx = bHeaders.indexOf('duration_minute');
  }
  
  var bStatusIdx = bHeaders.indexOf('break_status');
  var bDateIdx = bHeaders.indexOf('date');
  
  var timeZone = Session.getScriptTimeZone();
  var todayStr = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd');
  
  var totalBreakMins = 0;
  var sheetUpdated = false;
  
  for (var b = 1; b < bData.length; b++) {
    if (bLogIdx !== -1 && String(bData[b][bLogIdx]).trim() === String(logId).trim()) {
      var status = bStatusIdx !== -1 ? String(bData[b][bStatusIdx]).toUpperCase() : '';
      var startVal = bStartIdx !== -1 ? bData[b][bStartIdx] : null;
      var endVal = bEndIdx !== -1 ? bData[b][bEndIdx] : null;
      
      var startTime = null;
      var endTime = null;
      var bDateVal = bDateIdx !== -1 ? bData[b][bDateIdx] : null;
      var bDateStr = normalizeDateStr(bDateVal, timeZone) || todayStr;
      
      if (startVal) startTime = parseTimeStringToDate(startVal, bDateStr, timeZone);
      if (endVal) endTime = parseTimeStringToDate(endVal, bDateStr, timeZone);
      
      if (status === 'COMPLETED' || status === 'CLOSED' || (startTime && endTime)) {
        var duration = bDurIdx !== -1 ? (Number(bData[b][bDurIdx]) || 0) : 0;
        if (duration <= 0 && startTime && endTime && !isNaN(startTime.getTime()) && !isNaN(endTime.getTime())) {
          var diffMs = endTime.getTime() - startTime.getTime();
          if (diffMs > 0) {
            duration = Math.round((diffMs / (1000 * 60)) * 100) / 100;
            
            // Sync/Write back duration to Break_Logs sheet
            var bRowIndex = b + 1;
            if (bDurIdx !== -1) {
              breakSheet.getRange(bRowIndex, bDurIdx + 1).setValue(duration);
              sheetUpdated = true;
            }
            if (bStatusIdx !== -1 && (!status || status === 'ACTIVE' || status === '')) {
              breakSheet.getRange(bRowIndex, bStatusIdx + 1).setValue('COMPLETED');
              sheetUpdated = true;
            }
          }
        }
        totalBreakMins += duration;
      }
    }
  }
  
  if (sheetUpdated) {
    SpreadsheetApp.flush();
  }
  
  // Sync the accumulated value to Attendance sheet
  var attSheet = ss.getSheetByName(SHEET_NAMES.ATTENDANCE);
  if (attSheet) {
    var aData = attSheet.getDataRange().getValues();
    var aHeaders = aData[0];
    var logIdIdx = aHeaders.indexOf('log_id');
    var totalBreakIdx = aHeaders.indexOf('total_break_minutes');
    var shiftStartIdx = aHeaders.indexOf('shift_start_time');
    var shiftEndIdx = aHeaders.indexOf('shift_end_time');
    var totalShiftHoursIdx = aHeaders.indexOf('total_shift_hours');
    var dateIdx = aHeaders.indexOf('date');

    if (logIdIdx !== -1) {
      for (var r = 1; r < aData.length; r++) {
        if (String(aData[r][logIdIdx]).trim() === String(logId).trim()) {
          var aRowIndex = r + 1;
          if (totalBreakIdx !== -1) {
            attSheet.getRange(aRowIndex, totalBreakIdx + 1).setValue(totalBreakMins);
          }

          // Recalculate shift hours if shift is already closed/ended
          if (shiftStartIdx !== -1 && shiftEndIdx !== -1 && totalShiftHoursIdx !== -1) {
            var startStr = aData[r][shiftStartIdx];
            var endStr = aData[r][shiftEndIdx];
            var shiftDateVal = aData[r][dateIdx];
            var shiftDateStr = normalizeDateStr(shiftDateVal, timeZone) || todayStr;

            if (startStr && endStr && String(endStr).trim() !== '--:--' && String(endStr).trim() !== '') {
              var startTime = parseTimeStringToDate(startStr, shiftDateStr, timeZone);
              var endTime = parseTimeStringToDate(endStr, shiftDateStr, timeZone);
              if (startTime && endTime && !isNaN(startTime.getTime()) && !isNaN(endTime.getTime())) {
                var grossMs = endTime.getTime() - startTime.getTime();
                var breakMs = totalBreakMins * 60 * 1000;
                var netWorkedHours = Math.max(0, (grossMs - breakMs) / (1000 * 60 * 60));
                attSheet.getRange(aRowIndex, totalShiftHoursIdx + 1).setValue(Number(netWorkedHours.toFixed(2)));
              }
            }
          }
          break;
        }
      }
    }
  }
  
  return Math.round(totalBreakMins * 100) / 100;
}

/**
 * HELPER: Fetch Itemized Break Logs for Employee / Log ID
 */
function getEmpBreakLogs(logId, empId, targetDate) {
  var ss = getSpreadsheet();
  var breakSheet = ss.getSheetByName(SHEET_NAMES.BREAK_LOGS);
  var breakList = [];
  if (!breakSheet) return breakList;

  var bData = breakSheet.getDataRange().getValues();
  if (bData.length <= 1) return breakList;

  var timeZone = Session.getScriptTimeZone();
  var headers = bData[0];
  var logIdx = headers.indexOf('log_id');
  var empIdx = headers.indexOf('emp_id');
  var dateIdx = headers.indexOf('date');
  var createdAtIdx = headers.indexOf('created_at');

  var targetDateStr = targetDate ? normalizeDateStr(targetDate, timeZone) : '';

  for (var i = 1; i < bData.length; i++) {
    var rowLogId = logIdx !== -1 ? String(bData[i][logIdx]).trim() : '';
    var rowEmpId = empIdx !== -1 ? String(bData[i][empIdx]).trim().toLowerCase() : '';
    var rowDateStr = normalizeDateStr(bData[i][dateIdx], timeZone);
    if (!rowDateStr && createdAtIdx !== -1) {
      rowDateStr = normalizeDateStr(bData[i][createdAtIdx], timeZone);
    }

    var targetEmpLower = empId ? String(empId).trim().toLowerCase() : '';
    var matchLog = logId && rowLogId && rowLogId === String(logId).trim();
    var matchEmp = targetEmpLower && rowEmpId && rowEmpId === targetEmpLower;
    var matchDate = !targetDateStr || (rowDateStr && rowDateStr === targetDateStr);

    if ((matchLog || matchEmp) && matchDate) {
      var bObj = {};
      headers.forEach(function(h, idx) {
        var val = bData[i][idx];
        var key = h;
        // Normalize singular key to plural for frontend compatibility
        if (key === 'duration_minute') key = 'duration_minutes';
        
        if (val instanceof Date) {
          if (key === 'date' || key === 'created_at') {
            bObj[key] = Utilities.formatDate(val, timeZone, 'yyyy-MM-dd HH:mm:ss');
          } else {
            bObj[key] = Utilities.formatDate(val, timeZone, 'HH:mm:ss');
          }
        } else {
          bObj[key] = val !== null && val !== undefined ? String(val) : '';
        }
      });
      breakList.push(bObj);
    }
  }
  return breakList;
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

    var currentStatusIdx = headers.indexOf('shift_status');
    var currentStatus = currentStatusIdx !== -1 ? String(data[rowIndex - 1][currentStatusIdx]).toUpperCase() : '';
    if (currentStatus === 'ON_BREAK') {
      toggleShiftPause({ log_id: payload.log_id, action: 'RESUME' });
      data = sheet.getDataRange().getValues();
    }

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

    var leaveCode = payload.leave_type || 'PL';
    var isFullDayLeave = (leaveCode === 'PL' || leaveCode === 'UL' || leaveCode === 'ABSENT');
    var isShortPerm = (leaveCode === 'PERM' || leaveCode === 'PERMISSION' || leaveCode === 'OD');
    var permHrs = isShortPerm ? (Number(payload.permission_hours) || 2.0) : 0.0;
    var shrinkageHrs = isFullDayLeave ? 8.0 : (leaveCode === 'HL' ? 4.0 : permHrs);
    var proofUrl = payload.proof_url || '';
    var clientName = payload.active_client || user.default_client || 'iMerit';
    var projName = payload.active_project || user.default_project || 'Pravah';
    var remarks = payload.remarks || '';

    var userEmpLower = String(user.emp_id || '').trim().toLowerCase();

    // Check if an attendance row already exists for target date
    for (var i = 1; i < data.length; i++) {
      var rowDate = normalizeDateStr(data[i][dateIdx], timeZone);
      var rowEmp = empIdx !== -1 ? String(data[i][empIdx]).trim().toLowerCase() : '';

      if (rowEmp === userEmpLower && rowDate === targetDate) {
        var rowIndex = i + 1;
        var existingStatusIdx = headers.indexOf('shift_status');
        var existingStatus = existingStatusIdx !== -1 ? String(data[i][existingStatusIdx]).toUpperCase() : '';

        var newShiftStatus = isFullDayLeave ? 'LEAVE_CLOSED' : (existingStatus && existingStatus !== 'NOT_CLOCKED_IN' ? existingStatus : 'NOT_CLOCKED_IN');
        var attStatus = isFullDayLeave ? leaveCode : (leaveCode === 'HL' ? 'HL' : (existingStatus === 'OPEN' || existingStatus === 'CLOSED' ? 'Present' : leaveCode));

        var rowValues = data[i];
        if (headers.indexOf('attendance_status') !== -1) rowValues[headers.indexOf('attendance_status')] = attStatus;
        if (headers.indexOf('permission_hours') !== -1) rowValues[headers.indexOf('permission_hours')] = permHrs;
        if (headers.indexOf('shrinkage_hours') !== -1) rowValues[headers.indexOf('shrinkage_hours')] = shrinkageHrs;
        if (headers.indexOf('proof_url') !== -1 && proofUrl) rowValues[headers.indexOf('proof_url')] = proofUrl;
        if (headers.indexOf('shift_status') !== -1) rowValues[headers.indexOf('shift_status')] = newShiftStatus;
        if (isFullDayLeave) {
          if (headers.indexOf('shift_start_time') !== -1) rowValues[headers.indexOf('shift_start_time')] = '';
          if (headers.indexOf('shift_end_time') !== -1) rowValues[headers.indexOf('shift_end_time')] = '';
          if (headers.indexOf('total_shift_hours') !== -1) rowValues[headers.indexOf('total_shift_hours')] = 0;
        }
        if (headers.indexOf('active_client') !== -1 && clientName) rowValues[headers.indexOf('active_client')] = clientName;
        if (headers.indexOf('active_project') !== -1 && projName) rowValues[headers.indexOf('active_project')] = projName;
        if (headers.indexOf('user_remarks') !== -1 && remarks) {
          var existingUserRem = String(data[i][headers.indexOf('user_remarks')] || '').trim();
          var updatedUserRem = existingUserRem ? (existingUserRem + ' | User Update: ' + remarks) : remarks;
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
          message: 'Updated attendance record: ' + leaveCode + ' (' + permHrs + ' hrs permission/leave recorded)'
        };
      }
    }

    // New Row Creation
    var cleanEmpIdLeave = String(user.emp_id || '').trim().replace(/[^a-zA-Z0-9]/g, '');
    var logId = 'ATT-' + targetDate.replace(/-/g, '') + (cleanEmpIdLeave ? '-' + cleanEmpIdLeave : '') + '-' + Math.floor(100000 + Math.random() * 900000);
    var newRow = new Array(headers.length);
    for (var h = 0; h < headers.length; h++) newRow[h] = '';

    var newShiftStatus = isFullDayLeave ? 'LEAVE_CLOSED' : 'NOT_CLOCKED_IN';
    var attStatus = isFullDayLeave ? leaveCode : (leaveCode === 'HL' ? 'HL' : 'Present');

    if (headers.indexOf('log_id') !== -1) newRow[headers.indexOf('log_id')] = logId;
    if (headers.indexOf('date') !== -1) newRow[headers.indexOf('date')] = targetDate;
    if (headers.indexOf('emp_id') !== -1) newRow[headers.indexOf('emp_id')] = user.emp_id;
    if (headers.indexOf('emp_name') !== -1) newRow[headers.indexOf('emp_name')] = user.emp_name;
    if (headers.indexOf('attendance_status') !== -1) newRow[headers.indexOf('attendance_status')] = attStatus;
    if (headers.indexOf('permission_hours') !== -1) newRow[headers.indexOf('permission_hours')] = permHrs;
    if (headers.indexOf('shift_start_time') !== -1) newRow[headers.indexOf('shift_start_time')] = '';
    if (headers.indexOf('shift_end_time') !== -1) newRow[headers.indexOf('shift_end_time')] = '';
    if (headers.indexOf('total_shift_hours') !== -1) newRow[headers.indexOf('total_shift_hours')] = 0;
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
      message: 'Submitted ' + leaveCode + ' (' + permHrs + ' hrs permission/leave recorded)'
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * MANAGER DIRECT MARK LEAVE ENGINE (Supports Today & Past Dates)
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
    var leaveCode = payload.leave_type || 'UL';
    var isFullDayLeave = (leaveCode === 'PL' || leaveCode === 'UL' || leaveCode === 'ABSENT');
    var isShortPermAdmin = (leaveCode === 'PERM' || leaveCode === 'PERMISSION' || leaveCode === 'OD');
    var permHrs = isShortPermAdmin ? (Number(payload.permission_hours) || 2.0) : 0.0;
    var shrinkageHrs = isFullDayLeave ? 8.0 : (leaveCode === 'HL' ? 4.0 : permHrs);
    var remarks = payload.remarks ? ('Manager Note: ' + payload.remarks) : 'Marked by Manager';

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
        var newShiftStatus = isFullDayLeave ? 'LEAVE_CLOSED' : 'NOT_CLOCKED_IN';

        var rowValues = data[i];
        if (headers.indexOf('attendance_status') !== -1) rowValues[headers.indexOf('attendance_status')] = leaveCode;
        if (headers.indexOf('permission_hours') !== -1) rowValues[headers.indexOf('permission_hours')] = permHrs;
        if (headers.indexOf('shrinkage_hours') !== -1) rowValues[headers.indexOf('shrinkage_hours')] = shrinkageHrs;
        if (headers.indexOf('shift_status') !== -1) rowValues[headers.indexOf('shift_status')] = newShiftStatus;
        if (isFullDayLeave) {
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
          message: 'Successfully updated leave for ' + user.emp_name + ' on ' + targetDate + ' as ' + leaveCode
        };
      }
    }

    // New Row Creation for target date
    var cleanEmpIdAdmin = String(user.emp_id || '').trim().replace(/[^a-zA-Z0-9]/g, '');
    var logId = 'ATT-' + targetDate.replace(/-/g, '') + (cleanEmpIdAdmin ? '-' + cleanEmpIdAdmin : '') + '-' + Math.floor(100000 + Math.random() * 900000);
    var newRow = new Array(headers.length);
    for (var h = 0; h < headers.length; h++) newRow[h] = '';

    var newShiftStatus = isFullDayLeave ? 'LEAVE_CLOSED' : 'NOT_CLOCKED_IN';

    if (headers.indexOf('log_id') !== -1) newRow[headers.indexOf('log_id')] = logId;
    if (headers.indexOf('date') !== -1) newRow[headers.indexOf('date')] = targetDate;
    if (headers.indexOf('emp_id') !== -1) newRow[headers.indexOf('emp_id')] = user.emp_id;
    if (headers.indexOf('emp_name') !== -1) newRow[headers.indexOf('emp_name')] = user.emp_name;
    if (headers.indexOf('attendance_status') !== -1) newRow[headers.indexOf('attendance_status')] = leaveCode;
    if (headers.indexOf('permission_hours') !== -1) newRow[headers.indexOf('permission_hours')] = permHrs;
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
      message: 'Successfully marked leave for ' + user.emp_name + ' on ' + targetDate + ' as ' + leaveCode
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
 * TASK LOGGING ENGINE (With Module A Non-Billable Exemptions & Mandatory Remarks)
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

    var actType = String(payload.activity_type || 'Production Work').trim();
    var actTypeLower = actType.toLowerCase();
    var clientLower = String(payload.active_client || '').toLowerCase();
    var projLower = String(payload.active_project || '').toLowerCase();

    var isIdle = (actTypeLower.indexOf('downtime') !== -1 || actTypeLower.indexOf('allocation') !== -1 || actTypeLower.indexOf('waiting') !== -1 || projLower.indexOf('downtime') !== -1);
    var isBench = (actTypeLower.indexOf('bench') !== -1 || actTypeLower.indexOf('training') !== -1 || actTypeLower.indexOf('upskilling') !== -1
      || projLower.indexOf('bench') !== -1 || projLower.indexOf('unallocated') !== -1 || projLower.indexOf('floor management') !== -1 || projLower.indexOf('training') !== -1 || projLower.indexOf('upskilling') !== -1
      || clientLower.indexOf('bench') !== -1 || clientLower.indexOf('floor management') !== -1 || clientLower.indexOf('unallocated') !== -1);

    var isBillable = true;
    if (payload.is_billable !== undefined && payload.is_billable !== null) {
      isBillable = (payload.is_billable === true || String(payload.is_billable).toUpperCase() === 'TRUE');
    } else {
      var clientMaster = getClientMasterData();
      var matchProj = clientMaster.find(function(cp) {
        return String(cp.client_name).toUpperCase() === String(payload.active_client).toUpperCase() &&
               String(cp.project_name).toUpperCase() === String(payload.active_project).toUpperCase() &&
               String(cp.activity_type || '').toUpperCase() === actType.toUpperCase();
      });
      if (!matchProj) {
        matchProj = clientMaster.find(function(cp) {
          return String(cp.client_name).toUpperCase() === String(payload.active_client).toUpperCase() &&
                 String(cp.project_name).toUpperCase() === String(payload.active_project).toUpperCase();
        });
      }
      if (matchProj && matchProj.is_billable !== undefined) {
        isBillable = (matchProj.is_billable === true || String(matchProj.is_billable).toUpperCase() === 'TRUE');
      }
    }

    if (!isBillable || isIdle || isBench) {
      isBillable = false;
    }

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

    var taskCount = Number(payload.task_count) || 0;
    var ahtBenchMins = Number(payload.aht_benchmark_mins);
    if (isNaN(ahtBenchMins) || ahtBenchMins <= 0) {
      var clientMasterDataList = getClientMasterData();
      var matchProjAht = clientMasterDataList.find(function(cp) {
        return String(cp.client_name).toUpperCase() === String(payload.active_client).toUpperCase() &&
               String(cp.project_name).toUpperCase() === String(payload.active_project).toUpperCase() &&
               String(cp.activity_type || '').toUpperCase() === actType.toUpperCase();
      });
      if (!matchProjAht) {
        matchProjAht = clientMasterDataList.find(function(cp) {
          return String(cp.client_name).toUpperCase() === String(payload.active_client).toUpperCase() &&
                 String(cp.project_name).toUpperCase() === String(payload.active_project).toUpperCase();
        });
      }
      ahtBenchMins = matchProjAht ? (Number(matchProjAht.aht_minutes) || 60) : 60;
    }

    var calculatedEarnedHrs = Number(((taskCount * ahtBenchMins) / 60).toFixed(2));

    var productiveHours = 0.00; // Billable Earned Hours
    var trainingEarnedHours = 0.00; // Bench / Training Earned Hours
    var speedEfficiencyPct = 100;

    if (isBench) {
      trainingEarnedHours = 0.00;
      speedEfficiencyPct = 100;
    } else if (isIdle) {
      productiveHours = 0.00;
      trainingEarnedHours = 0.00;
      speedEfficiencyPct = 0;
    } else {
      productiveHours = calculatedEarnedHrs;
      speedEfficiencyPct = actualLoggedHrs > 0 ? Math.round((productiveHours / actualLoggedHrs) * 100) : 100;
    }

    var blockedMins = Number(payload.blocked_mins || payload.blocked_duration_mins) || 0;
    var blockedHours = Number((blockedMins / 60).toFixed(2));

    var performanceLeakageHrs = 0.00;
    if (isBillable && !isBench && !isIdle) {
      var grossLeak = Math.max(0, actualLoggedHrs - productiveHours);
      performanceLeakageHrs = Number(Math.max(0, grossLeak - blockedHours).toFixed(2));
    }

    var idleHours = 0.00;
    var benchHours = 0.00;

    if (isBench) {
      benchHours = actualLoggedHrs;
      idleHours = 0.00;
    } else if (isIdle) {
      idleHours = actualLoggedHrs;
      benchHours = 0.00;
    } else if (!isBillable) {
      benchHours = actualLoggedHrs;
      idleHours = 0.00;
    }

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
      payload.active_client || '', payload.active_project || '',
      payload.task_status || 'COMPLETED', Number(payload.completion_pct) || 100,
      Utilities.formatDate(startTime, timeZone, 'yyyy-MM-dd HH:mm:ss'),
      Utilities.formatDate(endTime, timeZone, 'yyyy-MM-dd HH:mm:ss'),
      taskCount, ahtBenchMins, actualLoggedHrs, productiveHours, speedEfficiencyPct, performanceLeakageHrs,
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
 * COMPLETE ADMIN DASHBOARD ENGINE
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

    var selectedClient = (filters && (filters.client || filters.active_client)) ? String(filters.client || filters.active_client).trim() : 'ALL';
    var selectedProject = (filters && (filters.project || filters.active_project)) ? String(filters.project || filters.active_project).trim() : 'ALL';
    var selectedTeam = (filters && filters.team) ? String(filters.team).trim() : 'ALL';
    var selectedMode = (filters && filters.work_mode) ? String(filters.work_mode).trim() : 'ALL';

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

    var aSheet = ss.getSheetByName(SHEET_NAMES.ATTENDANCE);
    var attendanceByEmp = {};
    if (aSheet) {
      var aData = aSheet.getDataRange().getValues();
      if (aData.length > 1) {
        var aHeaders = aData[0];
        var aDateIdx = aHeaders.indexOf('date');
        var aEmpIdx = aHeaders.indexOf('emp_id');

        for (var k = 1; k < aData.length; k++) {
          var rDate = normalizeDateStr(aData[k][aDateIdx], timeZone);

          if (rDate === targetDate || !targetDate) {
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
            var empKey = aEmpIdx !== -1 ? String(aData[k][aEmpIdx]).trim().toLowerCase() : String(aObj.emp_id || '').trim().toLowerCase();
            attendanceByEmp[empKey] = aObj;
          }
        }
      }
    }

    var tSheet = ss.getSheetByName(SHEET_NAMES.PROJECT_LOGS);
    var tasksByEmp = {};
    var taskLogs = [];
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

          if (tDate === targetDate || !targetDate) {
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

            // Fallback for missing column names
            if (tObj.activity_type === undefined && tData[t][18] !== undefined) tObj.activity_type = tData[t][18];
            if (tObj.idle_hours === undefined && tData[t][19] !== undefined) tObj.idle_hours = tData[t][19];
            if (tObj.bench_hours === undefined && tData[t][20] !== undefined) tObj.bench_hours = tData[t][20];
            if (tObj.blocked_hours === undefined && tData[t][21] !== undefined) tObj.blocked_hours = tData[t][21];

            taskLogs.push(tObj);
            var empKey = tEmpIdx !== -1 ? String(tData[t][tEmpIdx]).trim().toLowerCase() : String(tObj.emp_id || '').trim().toLowerCase();
            if (!tasksByEmp[empKey]) tasksByEmp[empKey] = [];
            tasksByEmp[empKey].push(tObj);
          }
        }
      }
    }

    var clientMaster = getClientMasterData();
    var adminBillableMap = {};
    if (clientMaster && clientMaster.length > 0) {
      clientMaster.forEach(function(cp) {
        var key = String(cp.client_name).toUpperCase() + '|||' + String(cp.project_name).toUpperCase();
        adminBillableMap[key] = (cp.is_billable === true || String(cp.is_billable).toUpperCase() === 'TRUE');
      });
    }

    var employeeStatusList = [];
    var leaveQueue = [];
    var totalEmployees = users.length;
    var clockedInCount = 0;
    var onBreakCount = 0;
    var leaveCount = 0;
    var notClockedInCount = 0;
    var plCount = 0, ulCount = 0, hlCount = 0, permCount = 0;
    var totalPermissionHours = 0;

    var totalLoggedHours = 0; // Net Shift Worked Hours
    var totalTaskHours = 0;   // Actual Logged Task Hours
    var totalEarnedHours = 0;
    var totalLeakageHours = 0;
    var totalIdleHours = 0;
    var totalBenchHours = 0;
    var totalBlockedHours = 0;

    for (var uIdx = 0; uIdx < users.length; uIdx++) {
      var u = users[uIdx];
      var uRole = String(u.app_role || u.role || '').trim().toLowerCase();
      var isMgmtAdmin = (uRole === 'management admin' || uRole === 'management_admin' || uRole === 'management' || uRole === 'super admin' || uRole === 'superadmin' || uRole === 'executive' || uRole === 'mgmt admin');
      if (isMgmtAdmin) continue; // Management Admins are pure management; exclude from shift directory capacity

      var userEmpId = String(u.emp_id || '').trim();
      var userEmpKey = userEmpId.toLowerCase();
      var allEmpTasks = tasksByEmp[userEmpKey] || [];
      var userClient = u.default_client || u.client || '';
      var userProj = u.default_project || u.project || '';

      if (selectedClient !== 'ALL' && userClient.toUpperCase() !== selectedClient.toUpperCase()) {
        var hasClientTask = false;
        for (var st = 0; st < allEmpTasks.length; st++) {
          if (String(allEmpTasks[st].active_client).toUpperCase() === selectedClient.toUpperCase()) {
            hasClientTask = true;
            break;
          }
        }
        if (!hasClientTask) continue;
      }
      if (selectedProject !== 'ALL' && userProj.toUpperCase() !== selectedProject.toUpperCase()) {
        var hasProjTask = false;
        for (var st = 0; st < allEmpTasks.length; st++) {
          if (String(allEmpTasks[st].active_project).toUpperCase() === selectedProject.toUpperCase()) {
            hasProjTask = true;
            break;
          }
        }
        if (!hasProjTask) continue;
      }
      if (selectedTeam !== 'ALL' && String(u.team).toUpperCase() !== selectedTeam.toUpperCase()) continue;
      if (selectedMode !== 'ALL' && String(u.work_mode).toUpperCase() !== selectedMode.toUpperCase()) continue;

      var att = attendanceByEmp[userEmpKey];
      var empTasks = [];
      for (var st = 0; st < allEmpTasks.length; st++) {
        var tsk = allEmpTasks[st];
        if (selectedClient !== 'ALL' && String(tsk.active_client).toUpperCase() !== selectedClient.toUpperCase()) continue;
        if (selectedProject !== 'ALL' && String(tsk.active_project).toUpperCase() !== selectedProject.toUpperCase()) continue;
        empTasks.push(tsk);
      }

      var empEarnedHrs = 0;
      var empLeakageHrs = 0;
      var empActualTaskHrs = 0;
      var empIdleHrs = 0;
      var empBenchHrs = 0;
      var empBlockedHrs = 0;

      for (var tskIdx = 0; tskIdx < empTasks.length; tskIdx++) {
        var tsk = empTasks[tskIdx];
        var tClient = String(tsk.active_client || '').trim();
        var tProject = String(tsk.active_project || '').trim();
        var bKey = tClient.toUpperCase() + '|||' + tProject.toUpperCase();
        var isTaskBillable = adminBillableMap[bKey] !== undefined ? adminBillableMap[bKey] : true;

        var actCategory = String(tsk.activity_category || tsk.activity_type || '').trim().toLowerCase();
        var tProjLower = tProject.toLowerCase();
        var tClientLower = tClient.toLowerCase();

        var isTaskBench = (tProjLower.indexOf('bench') !== -1 || tProjLower.indexOf('floor management') !== -1 || tProjLower.indexOf('unallocated') !== -1 || tProjLower.indexOf('training') !== -1 || tProjLower.indexOf('upskilling') !== -1)
          || (tClientLower.indexOf('bench') !== -1 || tClientLower.indexOf('floor management') !== -1 || tClientLower.indexOf('unallocated') !== -1)
          || (actCategory.indexOf('bench') !== -1 || actCategory.indexOf('training') !== -1 || actCategory.indexOf('upskilling') !== -1);

        var isTaskIdle = (actCategory.indexOf('downtime') !== -1 || actCategory.indexOf('allocation') !== -1 || actCategory.indexOf('waiting') !== -1 || tProjLower.indexOf('downtime') !== -1);

        var tskActual = Number(tsk.actual_worked_hours);
        if (isNaN(tskActual) || tskActual <= 0) {
          if (tsk.work_start_time && tsk.work_end_time) {
            try {
              var st = new Date(tsk.work_start_time).getTime();
              var et = new Date(tsk.work_end_time).getTime();
              if (et > st) tskActual = (et - st) / (1000 * 60 * 60);
            } catch(e) { tskActual = 0; }
          }
        }
        tskActual = Number(tskActual) || 0;
        empActualTaskHrs += tskActual;

        var taskIdle = 0;
        var taskBench = 0;
        var taskBlocked = Number(tsk.blocked_hours || tsk.blocked_duration_hours) || (Number(tsk.blocked_mins) ? Number((Number(tsk.blocked_mins)/60).toFixed(2)) : 0);

        if (isTaskBench) {
          taskBench = Number(tsk.bench_hours) || tskActual;
          taskIdle = 0;
        } else if (isTaskIdle) {
          taskIdle = Number(tsk.idle_hours) || tskActual;
          taskBench = 0;
        } else if (!isTaskBillable) {
          taskBench = Number(tsk.bench_hours) || tskActual;
          taskIdle = 0;
        } else {
          taskIdle = Number(tsk.idle_hours) || 0;
          taskBench = Number(tsk.bench_hours) || 0;
        }

        var taskProd = (isTaskBillable && !isTaskIdle && !isTaskBench) ? (Number(tsk.productive_hours) || 0) : 0;
        var taskLeak = (isTaskBillable && !isTaskIdle && !isTaskBench) ? (Number(tsk.performance_leakage_hours) || 0) : 0;

        empEarnedHrs += taskProd;
        empLeakageHrs += taskLeak;
        empIdleHrs += taskIdle;
        empBenchHrs += taskBench;
        empBlockedHrs += taskBlocked;
      }

      var shiftStatus = att ? String(att.shift_status || 'NOT_CLOCKED_IN').toUpperCase() : 'NOT_CLOCKED_IN';
      var attStatus = att ? String(att.attendance_status || 'NOT_CLOCKED_IN').toUpperCase() : 'NOT_CLOCKED_IN';
      var breakMins = att ? Number(att.total_break_minutes) || 0 : 0;
      var permHrs = att ? (Number(att.permission_hours) || 0) : 0;

      var isFullDayLeave = (attStatus === 'PL' || attStatus === 'UL' || attStatus === 'ABSENT' || shiftStatus === 'LEAVE_CLOSED');
      var isHalfDayLeave = (attStatus === 'HL');
      var isFullOrHalfLeave = (isFullDayLeave || isHalfDayLeave);
      var realPermHrs = isFullOrHalfLeave ? 0 : permHrs;

      // STRICT ZERO-LEAVE INTEGRITY RULE: Full-Day Leave MUST ALWAYS BE 0.00 hrs
      var rawWorkedHours = att ? Number(att.total_shift_hours) || 0 : 0;
      if (rawWorkedHours <= 0 && att && att.shift_start_time && (shiftStatus === 'OPEN' || shiftStatus === 'ON_BREAK') && targetDate === todayStr) {
        try {
          var st = parseTimeStringToDate(att.shift_start_time, targetDate, timeZone);
          var nowD = new Date();
          if (st && nowD > st) {
            rawWorkedHours = Math.max(0, ((nowD.getTime() - st.getTime()) / (1000 * 60 * 60)) - (breakMins / 60));
          }
        } catch(err) {}
      } else if (rawWorkedHours <= 0 && att && att.shift_start_time && att.shift_end_time) {
        try {
          var st = parseTimeStringToDate(att.shift_start_time, targetDate, timeZone);
          var et = parseTimeStringToDate(att.shift_end_time, targetDate, timeZone);
          if (st && et && et > st) {
            rawWorkedHours = Math.max(0, ((et.getTime() - st.getTime()) / (1000 * 60 * 60)) - (breakMins / 60));
          }
        } catch(err) {}
      }
      var workedHours = isFullDayLeave ? 0.00 : (isHalfDayLeave ? Math.min(4.50, rawWorkedHours) : rawWorkedHours);

      if (attStatus === 'PL') {
        plCount++;
      } else if (attStatus === 'UL') {
        ulCount++;
      } else if (attStatus === 'HL') {
        hlCount++;
      }

      if (!isFullOrHalfLeave && (realPermHrs > 0 || attStatus === 'PERMISSION' || attStatus === 'PERM' || attStatus.indexOf('PERM') !== -1)) {
        permCount++;
        totalPermissionHours += realPermHrs;
      }

      if (shiftStatus === 'OPEN' || shiftStatus === 'ON_BREAK' || shiftStatus === 'CLOSED' || attStatus === 'PRESENT') {
        clockedInCount++;
      }

      if (shiftStatus === 'ON_BREAK') {
        onBreakCount++;
      } else if (['PL', 'UL', 'HL', 'PERMISSION', 'LEAVE_CLOSED'].indexOf(attStatus) !== -1 || shiftStatus === 'LEAVE_CLOSED') {
        leaveCount++;

        leaveQueue.push({
          emp_id: String(u.emp_id || ''),
          emp_name: String(u.emp_name || ''),
          team: String(u.team || 'N/A'),
          leave_type: String(attStatus),
          perm_hours: permHrs,
          proof_url: att ? String(att.proof_url || '') : '',
          active_client: att ? String(att.active_client || u.default_client || 'iMerit') : String(u.default_client || 'iMerit'),
          active_project: att ? String(att.active_project || u.default_project || 'Pravah') : String(u.default_project || 'Pravah')
        });
      } else if (shiftStatus === 'NOT_CLOCKED_IN' || (!shiftStatus && !isFullOrHalfLeave)) {
        notClockedInCount++;
      }

      totalLoggedHours += workedHours; // Net Shift Worked Hours
      totalTaskHours += empActualTaskHrs;
      totalEarnedHours += empEarnedHrs;
      totalLeakageHours += empLeakageHrs;
      totalIdleHours += empIdleHrs;
      totalBenchHours += empBenchHrs;
      totalBlockedHours += empBlockedHrs;

      var shiftStartStr = '--:--';
      var shiftEndStr = '--:--';
      if (!isFullDayLeave) {
        if (att && att.shift_start_time) {
          shiftStartStr = att.shift_start_time instanceof Date ? Utilities.formatDate(att.shift_start_time, timeZone, 'HH:mm:ss') : String(att.shift_start_time);
        }
        if (att && att.shift_end_time) {
          shiftEndStr = att.shift_end_time instanceof Date ? Utilities.formatDate(att.shift_end_time, timeZone, 'HH:mm:ss') : String(att.shift_end_time);
        }
      }

      employeeStatusList.push({
        emp_id: String(u.emp_id || ''),
        emp_name: String(u.emp_name || ''),
        email: String(u.email || ''),
        team: String(u.team || 'N/A'),
        work_mode: String(u.work_mode || 'Office'),
        app_role: String(u.app_role || 'User'),
        shift_status: String(shiftStatus),
        attendance_status: String(attStatus),
        log_id: att ? String(att.log_id || '') : '',
        shift_start_time: shiftStartStr,
        shift_end_time: shiftEndStr,
        total_break_minutes: Number(breakMins),
        permission_hours: Number(permHrs),
        total_shift_hours: Number(workedHours.toFixed ? workedHours.toFixed(2) : workedHours),
        actual_worked_hours: Number(empActualTaskHrs.toFixed(2)),
        earned_hours: Number(empEarnedHrs.toFixed(2)),
        leakage_hours: Number(empLeakageHrs.toFixed(2)),
        idle_hours: Number(empIdleHrs.toFixed(2)),
        bench_hours: Number(empBenchHrs.toFixed(2)),
        blocked_hours: Number(empBlockedHrs.toFixed(2)),
        proof_url: att ? String(att.proof_url || '') : '',
        tasks_count: empTasks.length,
        tasks: empTasks
      });
    }

    var totalEmployees = employeeStatusList.length;
    var shrinkageRate = totalEmployees > 0 ? Number(((leaveCount / totalEmployees) * 100).toFixed(2)) : 0;
    var overallSpeedPct = totalTaskHours > 0 ? Math.round((totalEarnedHours / totalTaskHours) * 100) : 100;

    return {
      success: true,
      targetDate: targetDate,
      summary: {
        totalEmployees: totalEmployees,
        clockedInCount: clockedInCount,
        onBreakCount: onBreakCount,
        leaveCount: leaveCount,
        notClockedInCount: notClockedInCount,
        plCount: plCount,
        ulCount: ulCount,
        hlCount: hlCount,
        permCount: permCount,
        totalPermissionHours: Number(totalPermissionHours.toFixed(2)),
        shrinkageRate: shrinkageRate,
        totalLoggedHours: Number(totalLoggedHours.toFixed(2)),
        totalTaskHours: Number(totalTaskHours.toFixed(2)),
        totalEarnedHours: Number(totalEarnedHours.toFixed(2)),
        totalLeakageHours: Number(totalLeakageHours.toFixed(2)),
        totalIdleHours: Number(totalIdleHours.toFixed(2)),
        totalBenchHours: Number(totalBenchHours.toFixed(2)),
        totalBlockedHours: Number(totalBlockedHours.toFixed(2)),
        overallSpeedPct: overallSpeedPct
      },
      employeeStatusList: employeeStatusList,
      leaveQueue: leaveQueue,
      clientMaster: clientMaster,
      rawTaskLogs: taskLogs
    };
  } catch (err) {
    return { success: false, message: 'Admin Data Error: ' + err.toString() };
  }
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
    var earnedHrs = 0;
    var leakageHrs = 0;
    if (tSheet) {
      var tData = tSheet.getDataRange().getValues();
      if (tData.length > 1) {
        var tHeaders = tData[0];
        var teIdx = tHeaders.indexOf('emp_id');
        var wsIdx = tHeaders.indexOf('work_start_time');
        var createdIdx = tHeaders.indexOf('created_at');
        var prodIdx = tHeaders.indexOf('productive_hours');
        var leakIdx = tHeaders.indexOf('performance_leakage_hours');
        var targetEmpLower = String(empId).trim().toLowerCase();

        for (var t = 1; t < tData.length; t++) {
          var rawStart = wsIdx !== -1 ? tData[t][wsIdx] : (createdIdx !== -1 ? tData[t][createdIdx] : null);
          var tDate = normalizeDateStr(rawStart, timeZone);
          var rEmp = teIdx !== -1 ? String(tData[t][teIdx]).trim().toLowerCase() : '';

          if (rEmp === targetEmpLower && tDate === normalizedDate) {
            earnedHrs += Number(tData[t][prodIdx]) || 0;
            leakageHrs += Number(tData[t][leakIdx]) || 0;
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
    var breakMins = shiftLog ? Number(shiftLog.total_break_minutes) || 0 : 0;
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
      breakMins, netWorkedHrs, Number(earnedHrs.toFixed(2)), Number(leakageHrs.toFixed(2)),
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

/**
 * EXECUTIVE ANALYTICS BACKEND AGGREGATOR
 * Returns Attendance, Utilization, Shrinkage, Client/Team/PM Summaries & Employee Period Performance Monitor
 */
function getExecutiveAnalyticsSummary(filterPayload) {
  try {
    var activeEmail = Session.getActiveUser().getEmail();
    if (activeEmail) {
      var callingUser = getUserByEmail(activeEmail);
      if (callingUser) {
        var uRole = String(callingUser.app_role || callingUser.role || callingUser.user_role || '').trim().toLowerCase();
        if (uRole && uRole !== 'admin' && uRole !== 'executive' && uRole !== 'manager' && uRole !== 'super admin' && uRole.indexOf('admin') === -1 && uRole.indexOf('exec') === -1 && uRole.indexOf('manager') === -1) {
          return { success: false, message: 'Access Denied: Admin or Executive role required to view executive analytics.' };
        }
      }
    }

    filterPayload = filterPayload || {};
    var masterSs = getSpreadsheet();
    var timeZone = Session.getScriptTimeZone();

    var todayStr = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd');
    var monthFilter = filterPayload.month || 'ALL';
    var weekFilter = filterPayload.week || 'ALL';
    var teamFilter = filterPayload.team || 'ALL';
    var clientFilter = filterPayload.client || 'ALL';
    var startDateStr = filterPayload.startDate ? normalizeDateStr(filterPayload.startDate, timeZone) : '';
    var endDateStr = filterPayload.endDate ? normalizeDateStr(filterPayload.endDate, timeZone) : '';

    // Convert Month string (e.g., 'Aug-2026' -> '2026-08', 'Jul-2026' -> '2026-07')
    var monthPrefix = '';
    if (monthFilter && monthFilter !== 'ALL') {
      var monthMap = { 'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04', 'May': '05', 'Jun': '06', 'Jul': '07', 'Aug': '08', 'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12' };
      var mParts = monthFilter.split('-');
      if (mParts.length === 2 && monthMap[mParts[0]]) {
        monthPrefix = mParts[1] + '-' + monthMap[mParts[0]];
      }
    }

    // Default to TODAY if no month or week or date range is explicitly selected
    if (!monthPrefix && (!weekFilter || weekFilter === 'ALL') && !startDateStr && !endDateStr) {
      startDateStr = todayStr;
      endDateStr = todayStr;
    }

    // 1. Fetch User_Master
    var uSheet = masterSs.getSheetByName(SHEET_NAMES.USER_MASTER);
    if (!uSheet) return { success: false, message: 'User_Master sheet not found.' };

    var uData = uSheet.getDataRange().getValues();
    if (uData.length <= 1) return { success: false, message: 'User_Master sheet is empty.' };

    var uHeaders = uData[0];
    var empIdx = -1, nameIdx = -1, teamIdx = -1, modeIdx = -1, pmIdx = -1, clientIdx = -1, roleIdx = -1;

    for (var h = 0; h < uHeaders.length; h++) {
      var hn = String(uHeaders[h]).trim().toLowerCase().replace(/_/g, ' ');
      if (hn === 'emp id' || hn === 'empid' || hn === 'employee id' || hn === 'id') empIdx = h;
      else if (hn === 'emp name' || hn === 'name' || hn === 'employee name') nameIdx = h;
      else if (hn === 'team' || hn === 'department' || hn === 'dept' || hn === 'team name' || hn === 'dept name') teamIdx = h;
      else if (hn === 'work mode' || hn === 'mode') modeIdx = h;
      else if (hn === 'project manager' || hn === 'pm' || hn === 'manager' || hn === 'manager name' || hn === 'pm name') pmIdx = h;
      else if (hn === 'default client' || hn === 'client' || hn === 'active client' || hn === 'client name') clientIdx = h;
      else if (hn === 'app role' || hn === 'approle' || hn === 'role' || hn === 'user role') roleIdx = h;
    }

    var activeIdx = uHeaders.indexOf('is_active');

    var usersMap = {};
    for (var i = 1; i < uData.length; i++) {
      if (activeIdx !== -1) {
        var actVal = String(uData[i][activeIdx] || 'TRUE').trim().toUpperCase();
        var isActive = (actVal === 'TRUE' || actVal === '1' || actVal === 'ACTIVE' || actVal === 'YES' || uData[i][activeIdx] === true);
        if (!isActive) continue; // Skip inactive accounts
      }

      var uRole = roleIdx !== -1 ? String(uData[i][roleIdx] || '').trim().toLowerCase() : '';
      var isMgmtAdmin = (uRole === 'management admin' || uRole === 'management_admin' || uRole === 'management' || uRole === 'super admin' || uRole === 'superadmin' || uRole === 'executive' || uRole === 'mgmt admin');
      if (isMgmtAdmin) continue; // Exclude Management Admins from Executive Analytics

      var eId = empIdx !== -1 ? String(uData[i][empIdx]).trim() : '';
      if (eId) {
        var rawTeam = teamIdx !== -1 ? String(uData[i][teamIdx] || '').trim() : '';
        var rawClient = clientIdx !== -1 ? String(uData[i][clientIdx] || '').trim() : '';
        var rawPm = pmIdx !== -1 ? String(uData[i][pmIdx] || '').trim() : '';
        var uObj = {
          emp_id: eId,
          emp_name: nameIdx !== -1 ? String(uData[i][nameIdx] || 'Employee').trim() : 'Employee',
          team: (rawTeam && rawTeam !== 'None' && rawTeam !== 'null') ? rawTeam : 'Annotation',
          work_mode: modeIdx !== -1 ? String(uData[i][modeIdx] || 'Office').trim() : 'Office',
          pm: (rawPm && rawPm !== 'None' && rawPm !== 'null') ? rawPm : 'Dinesh Kumar Rajagopal',
          client: (rawClient && rawClient !== 'None' && rawClient !== 'null') ? rawClient : 'iMerit'
        };
        usersMap[eId] = uObj;
        usersMap[eId.toUpperCase()] = uObj;
      }
    }

    // 2. Fetch Attendance_Shift_Logs
    var attSheet = masterSs.getSheetByName(SHEET_NAMES.ATTENDANCE);
    var attData = attSheet ? attSheet.getDataRange().getValues() : [];
    var attLogs = [];

    var availableMonthsMap = {};
    var monthWeeksMap = {};
    var monthNamesList = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

    if (attData.length > 1) {
      var aHeaders = attData[0];
      var dIdx = -1, aeIdx = -1, sStatusIdx = -1, aStatusIdx = -1, stIdx = -1, etIdx = -1, thIdx = -1, brkIdx = -1, permIdx = -1;
      for (var ah = 0; ah < aHeaders.length; ah++) {
        var ahn = String(aHeaders[ah]).trim().toLowerCase().replace(/_/g, ' ');
        if (ahn === 'date' || ahn === 'shift date' || ahn === 'log date' || ahn === 'created at' || ahn === 'created date' || ahn === 'timestamp') dIdx = ah;
        else if (ahn === 'emp id' || ahn === 'empid' || ahn === 'employee id') aeIdx = ah;
        else if (ahn === 'shift status' || ahn === 'status') sStatusIdx = ah;
        else if (ahn === 'attendance status' || ahn === 'att status') aStatusIdx = ah;
        else if (ahn === 'shift start time' || ahn === 'start time' || ahn === 'clock in') stIdx = ah;
        else if (ahn === 'shift end time' || ahn === 'end time' || ahn === 'clock out') etIdx = ah;
        else if (ahn === 'total shift hours' || ahn === 'shift hours' || ahn === 'total hours' || ahn === 'net worked hours' || ahn === 'net worked') thIdx = ah;
        else if (ahn === 'total break minutes' || ahn === 'break minutes' || ahn === 'break mins') brkIdx = ah;
        else if (ahn === 'permission hours' || ahn === 'permission hrs' || ahn === 'permission') permIdx = ah;
      }
      if (dIdx === -1 && stIdx !== -1) dIdx = stIdx;

      for (var a = 1; a < attData.length; a++) {
        var rawD = dIdx !== -1 ? attData[a][dIdx] : '';
        if (!rawD && stIdx !== -1) rawD = attData[a][stIdx];
        if (!rawD) continue;
        var rowDate = normalizeDateStr(rawD, timeZone);
        var rEmp = aeIdx !== -1 ? String(attData[a][aeIdx]).trim() : '';

        // Track Month and Week metadata from Attendance_Shift_Logs for filter dropdowns
        if (rowDate && rowDate.length >= 7) {
          var parts = rowDate.split('-');
          var yNum = parseInt(parts[0], 10);
          var mNum = parseInt(parts[1], 10);
          if (mNum >= 1 && mNum <= 12) {
            var mKey = monthNamesList[mNum - 1] + '-' + yNum;
            availableMonthsMap[mKey] = true;
            if (!monthWeeksMap[mKey]) monthWeeksMap[mKey] = {};
            var wKey = getWeekOfMonthStr(rowDate);
            monthWeeksMap[mKey][wKey] = true;
          }
        }

        if (rEmp) {
          var user = usersMap[rEmp] || usersMap[rEmp.toUpperCase()] || {};
          if (teamFilter !== 'ALL' && user.team && user.team !== teamFilter) continue;
          if (clientFilter !== 'ALL' && user.client && user.client !== clientFilter) continue;
          if (monthPrefix && rowDate && rowDate.indexOf(monthPrefix) !== 0) continue;
          if (!monthPrefix && startDateStr && rowDate && rowDate < startDateStr) continue;
          if (!monthPrefix && endDateStr && rowDate && rowDate > endDateStr) continue;
          if (weekFilter && weekFilter !== 'ALL' && rowDate && getWeekOfMonthStr(rowDate) !== weekFilter) continue;

          attLogs.push({
            date: rowDate,
            emp_id: rEmp,
            shift_status: sStatusIdx !== -1 ? String(attData[a][sStatusIdx] || 'CLOSED') : 'CLOSED',
            attendance_status: aStatusIdx !== -1 ? String(attData[a][aStatusIdx] || 'Present') : 'Present',
            start_time: stIdx !== -1 ? (attData[a][stIdx] instanceof Date ? Utilities.formatDate(attData[a][stIdx], timeZone, 'HH:mm:ss') : String(attData[a][stIdx] || '')) : '',
            end_time: etIdx !== -1 ? (attData[a][etIdx] instanceof Date ? Utilities.formatDate(attData[a][etIdx], timeZone, 'HH:mm:ss') : String(attData[a][etIdx] || '')) : '',
            total_shift_hours: thIdx !== -1 ? (Number(attData[a][thIdx]) || 0) : 0,
            break_minutes: brkIdx !== -1 ? (Number(attData[a][brkIdx]) || 0) : 0,
            permission_hours: permIdx !== -1 ? (Number(attData[a][permIdx]) || 0) : 0
          });
        }
      }
    }

    var availableMonths = Object.keys(availableMonthsMap);

    // Format monthWeeksMap to arrays of week strings
    var formattedMonthWeeksMap = {};
    var mwKeys = Object.keys(monthWeeksMap);
    for (var mw = 0; mw < mwKeys.length; mw++) {
      var mKey = mwKeys[mw];
      formattedMonthWeeksMap[mKey] = Object.keys(monthWeeksMap[mKey]).sort();
    }

    // 3. Fetch Project_Task_Logs
    var tSheet = masterSs.getSheetByName(SHEET_NAMES.PROJECT_LOGS);
    var tData = tSheet ? tSheet.getDataRange().getValues() : [];
    var taskLogs = [];

    if (tData.length > 1) {
      var tHeaders = tData[0];
      var teIdx = -1, cIdx = -1, pIdx = -1, cntIdx = -1, prodIdx = -1, leakIdx = -1, wsIdx = -1, weIdx = -1, createdIdx = -1, actualWorkedIdx = -1;
      for (var th = 0; th < tHeaders.length; th++) {
        var thn = String(tHeaders[th]).trim().toLowerCase().replace(/_/g, ' ');
        if (thn === 'emp id' || thn === 'empid' || thn === 'employee id') teIdx = th;
        else if (thn === 'active client' || thn === 'client' || thn === 'default client') cIdx = th;
        else if (thn === 'active project' || thn === 'project' || thn === 'default project') pIdx = th;
        else if (thn === 'task count' || thn === 'units' || thn === 'count') cntIdx = th;
        else if (thn === 'productive hours' || thn === 'earned hours' || thn === 'productive hrs') prodIdx = th;
        else if (thn === 'performance leakage hours' || thn === 'leakage hours' || thn === 'leakage hrs') leakIdx = th;
        else if (thn === 'actual worked hours' || thn === 'actual hours' || thn === 'actual hrs' || thn === 'worked hours') actualWorkedIdx = th;
        else if (thn === 'created at' || thn === 'created date' || thn === 'date' || thn === 'task date' || thn === 'log date') createdIdx = th;
        else if (thn === 'work start time' || thn === 'start time') wsIdx = th;
        else if (thn === 'work end time' || thn === 'end time') weIdx = th;
      }

      var clientMasterList = getClientMasterData();
      var billableMap = {};
      if (clientMasterList && clientMasterList.length > 0) {
        clientMasterList.forEach(function(cp) {
          var key = String(cp.client_name).toUpperCase() + '|||' + String(cp.project_name).toUpperCase();
          billableMap[key] = (cp.is_billable === true || String(cp.is_billable).toUpperCase() === 'TRUE');
        });
      }

      for (var t = 1; t < tData.length; t++) {
        var tEmp = teIdx !== -1 ? String(tData[t][teIdx]).trim() : '';
        if (tEmp) {
          var wsRaw = wsIdx !== -1 ? tData[t][wsIdx] : '';
          var createdRaw = createdIdx !== -1 ? tData[t][createdIdx] : '';

          var rawDateVal = createdRaw || wsRaw;
          var tDate = normalizeDateStr(rawDateVal, timeZone);
          if (tDate && !tDate.match(/^\d{4}-\d{2}-\d{2}$/)) {
            tDate = '';
          }

          var user = usersMap[tEmp] || usersMap[tEmp.toUpperCase()] || {};
          var tClientCheck = cIdx !== -1 ? String(tData[t][cIdx] || user.client || 'iMerit') : (user.client || 'iMerit');
          if (teamFilter !== 'ALL' && user.team && user.team !== teamFilter) continue;
          if (clientFilter !== 'ALL' && (user.client !== clientFilter && tClientCheck !== clientFilter)) continue;
          if (monthPrefix && tDate && tDate.indexOf(monthPrefix) !== 0) continue;
          if (!monthPrefix && startDateStr && tDate && tDate < startDateStr) continue;
          if (!monthPrefix && endDateStr && tDate && tDate > endDateStr) continue;
          if (weekFilter && weekFilter !== 'ALL' && tDate && getWeekOfMonthStr(tDate) !== weekFilter) continue;

          var actualWorked = actualWorkedIdx !== -1 ? (Number(tData[t][actualWorkedIdx]) || 0) : 0;
          if (actualWorked <= 0 && wsRaw && weIdx !== -1 && tData[t][weIdx]) {
            try {
              var st = new Date(wsRaw).getTime();
              var et = new Date(tData[t][weIdx]).getTime();
              if (et > st) actualWorked = (et - st) / (1000 * 60 * 60);
            } catch(e) {}
          }

          var weRaw = weIdx !== -1 ? tData[t][weIdx] : '';

          var wsStr = wsRaw ? (wsRaw instanceof Date ? Utilities.formatDate(wsRaw, timeZone, 'yyyy-MM-dd HH:mm:ss') : String(wsRaw)) : '';
          var weStr = weRaw ? (weRaw instanceof Date ? Utilities.formatDate(weRaw, timeZone, 'yyyy-MM-dd HH:mm:ss') : String(weRaw)) : '';

          var tClient = tClientCheck;
          var tProject = pIdx !== -1 ? String(tData[t][pIdx] || 'Pravah') : 'Pravah';
          var bKey = tClient.toUpperCase() + '|||' + tProject.toUpperCase();
          var isTaskBillable = billableMap[bKey] !== undefined ? billableMap[bKey] : true;

          var prodHrs = (isTaskBillable && prodIdx !== -1) ? (Number(tData[t][prodIdx]) || 0) : 0;
          var leakHrs = isTaskBillable ? (leakIdx !== -1 ? (Number(tData[t][leakIdx]) || 0) : 0) : 0;
          var speedPct = actualWorked > 0 ? Math.round((prodHrs / actualWorked) * 100) : 100;

          var actTypeVal = 'Production Work';
          for (var thk = 0; thk < tHeaders.length; thk++) {
            var thkn = String(tHeaders[thk]).trim().toLowerCase().replace(/_/g, ' ');
            if (thkn === 'activity type' || thkn === 'activity') actTypeVal = String(tData[t][thk] || 'Production Work').trim();
          }

          var ahtMinsVal = 60;
          for (var thk = 0; thk < tHeaders.length; thk++) {
            var thkn = String(tHeaders[thk]).trim().toLowerCase().replace(/_/g, ' ');
            if (thkn === 'aht benchmark mins' || thkn === 'aht minutes' || thkn === 'aht') ahtMinsVal = Number(tData[t][thk]) || 60;
          }

          var idleHrsVal = 0;
          var benchHrsVal = 0;
          var blockedHrsVal = 0;
          var blockedMinsVal = 0;

          for (var thk = 0; thk < tHeaders.length; thk++) {
            var thkn = String(tHeaders[thk]).trim().toLowerCase().replace(/_/g, ' ');
            if (thkn === 'idle hours' || thkn === 'idle hrs') idleHrsVal = Number(tData[t][thk]) || 0;
            else if (thkn === 'bench hours' || thkn === 'bench hrs') benchHrsVal = Number(tData[t][thk]) || 0;
            else if (thkn === 'blocked hours' || thkn === 'blocked hrs') blockedHrsVal = Number(tData[t][thk]) || 0;
            else if (thkn === 'blocked mins' || thkn === 'blocked duration mins') blockedMinsVal = Number(tData[t][thk]) || 0;
          }

          var remIdx = tHeaders.indexOf('user_remarks');
          var remVal = remIdx !== -1 ? String(tData[t][remIdx] || '') : '';

          var tProjLower = tProject.toLowerCase();
          var tClientLower = tClient.toLowerCase();
          var actLower = actTypeVal.toLowerCase();

          var isTaskBench = (tProjLower.indexOf('bench') !== -1 || tProjLower.indexOf('floor management') !== -1 || tProjLower.indexOf('unallocated') !== -1 || tProjLower.indexOf('training') !== -1 || tProjLower.indexOf('upskilling') !== -1)
            || (tClientLower.indexOf('bench') !== -1 || tClientLower.indexOf('floor management') !== -1 || tClientLower.indexOf('unallocated') !== -1)
            || (actLower.indexOf('bench') !== -1 || actLower.indexOf('training') !== -1 || actLower.indexOf('upskilling') !== -1);

          var isTaskIdle = (actLower.indexOf('downtime') !== -1 || actLower.indexOf('allocation') !== -1 || actLower.indexOf('waiting') !== -1 || tProjLower.indexOf('downtime') !== -1);

          if (isTaskBench) {
            benchHrsVal = benchHrsVal || actualWorked;
            idleHrsVal = 0;
          } else if (isTaskIdle) {
            idleHrsVal = idleHrsVal || actualWorked;
            benchHrsVal = 0;
          } else if (!isTaskBillable) {
            benchHrsVal = benchHrsVal || actualWorked;
            idleHrsVal = 0;
          }

          taskLogs.push({
            date: tDate,
            emp_id: tEmp,
            client: tClient,
            project: tProject,
            activity_type: actTypeVal,
            aht_benchmark_mins: ahtMinsVal,
            is_billable: isTaskBillable,
            task_count: cntIdx !== -1 ? (Number(tData[t][cntIdx]) || 0) : 0,
            productive_hours: prodHrs,
            leakage_hours: leakHrs,
            actual_worked_hours: actualWorked,
            speed_efficiency_pct: speedPct,
            idle_hours: idleHrsVal,
            bench_hours: benchHrsVal,
            blocked_hours: blockedHrsVal,
            blocked_mins: blockedMinsVal || (blockedHrsVal * 60),
            work_start_time: wsStr,
            work_end_time: weStr,
            user_remarks: remVal
          });
        }
      }
    }

    var filteredAttLogs = attLogs;
    var filteredTaskLogs = taskLogs;

    // Collect all unique dates present in filtered logs
    var periodDatesMap = {};
    for (var rIdx = 0; rIdx < filteredAttLogs.length; rIdx++) {
      var r = filteredAttLogs[rIdx];
      if (r.date) periodDatesMap[r.date] = true;
    }
    for (var rIdx = 0; rIdx < filteredTaskLogs.length; rIdx++) {
      var r = filteredTaskLogs[rIdx];
      if (r.date) periodDatesMap[r.date] = true;
    }

    var uniqueDatesList = Object.keys(periodDatesMap);
    if (uniqueDatesList.length === 0) {
      uniqueDatesList = [Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd')];
    }
    var periodDaysCount = uniqueDatesList.length;

    // Build Employee Period Performance Monitor Object
    var empMonitorMap = {};
    var empIdLookupMap = {}; // Maps uppercase emp_id -> primary emp_id

    var uKeys = Object.keys(usersMap);
    for (var uk = 0; uk < uKeys.length; uk++) {
      var eId = uKeys[uk];
      var u = usersMap[eId];
      var normId = String(eId).trim().toUpperCase();
      if (empMonitorMap[normId]) continue; // Skip duplicate normalized keys
      if (teamFilter !== 'ALL' && u.team !== teamFilter) continue;
      if (clientFilter !== 'ALL' && u.client !== clientFilter) continue;

      empIdLookupMap[normId] = u.emp_id || eId;

      empMonitorMap[normId] = {
        emp_id: u.emp_id || eId,
        emp_name: u.emp_name,
        team: u.team,
        work_mode: u.work_mode,
        client: u.client,
        pm: u.pm,
        days_present: 0,
        days_leave: 0,
        days_not_clocked: 0,
        pl_count: 0,
        ul_count: 0,
        hl_count: 0,
        perm_count: 0,
        total_permission_hours: 0,
        total_break_minutes: 0,
        total_shift_hours: 0,
        total_net_worked: 0,
        total_task_hours: 0,
        total_actual_worked: 0,
        total_earned_hours: 0,
        total_leakage_hours: 0,
        total_idle_hours: 0,
        total_bench_hours: 0,
        total_blocked_hours: 0,
        tasks_count: 0,
        tasks_list: []
      };
    }

    // Deduplicate attendance logs per employee per date so each employee is counted at most once per date
    var dailyAttMap = {}; // Key: normId + '|||' + log.date
    for (var aIdx = 0; aIdx < filteredAttLogs.length; aIdx++) {
      var log = filteredAttLogs[aIdx];
      var rawId = log.emp_id ? String(log.emp_id).trim() : '';
      var normId = rawId.toUpperCase();
      if (!empMonitorMap[normId]) continue;

      var dateKey = normId + '|||' + log.date;
      if (!dailyAttMap[dateKey]) {
        dailyAttMap[dateKey] = log;
      } else {
        var existing = dailyAttMap[dateKey];
        var newHrs = Number(log.total_shift_hours) || 0;
        var existHrs = Number(existing.total_shift_hours) || 0;
        var newShift = String(log.shift_status || '').toUpperCase();
        var existShift = String(existing.shift_status || '').toUpperCase();

        if (newHrs > existHrs || (newShift === 'CLOSED' && existShift !== 'CLOSED') || (newShift === 'OPEN' && existShift === 'NOT_CLOCKED_IN')) {
          dailyAttMap[dateKey] = log;
        }
      }
    }

    var dailyAttKeys = Object.keys(dailyAttMap);
    for (var dk = 0; dk < dailyAttKeys.length; dk++) {
      var log = dailyAttMap[dailyAttKeys[dk]];
      var rawId = log.emp_id ? String(log.emp_id).trim() : '';
      var normId = rawId.toUpperCase();
      var empObj = empMonitorMap[normId];
      if (!empObj) continue;

      var attStat = String(log.attendance_status || '').toUpperCase();
      var shiftStat = String(log.shift_status || '').toUpperCase();
      var permHrs = Number(log.permission_hours) || 0;
      var isLeaveExec = (attStat === 'PL' || attStat === 'UL' || attStat === 'HL' || attStat === 'ABSENT' || shiftStat === 'LEAVE_CLOSED');
      var realPermExec = isLeaveExec ? 0 : permHrs;

      if (!isLeaveExec && (realPermExec > 0 || attStat.indexOf('PERM') !== -1)) {
        empObj.perm_count++;
        empObj.total_permission_hours += realPermExec;
      }

      var isFullDayLeave = (attStat === 'PL' || attStat === 'UL' || attStat === 'ABSENT' || shiftStat === 'LEAVE_CLOSED');
      var isHalfDayLeave = (attStat === 'HL');
      var isFutureDate = (log.date && log.date > todayStr);

      if (attStat === 'PRESENT' || shiftStat === 'OPEN' || shiftStat === 'ON_BREAK' || shiftStat === 'CLOSED') {
        if (!isFullDayLeave) empObj.days_present++;
      } else if (attStat === 'PL') {
        empObj.days_leave++;
        empObj.pl_count++;
      } else if (attStat === 'UL') {
        empObj.days_leave++;
        empObj.ul_count++;
      } else if (attStat === 'HL') {
        empObj.days_leave += 0.5;
        empObj.hl_count++;
      } else if (attStat !== 'NOT_CLOCKED_IN') {
        empObj.days_leave++;
      }

      // STRICT ENFORCEMENT: If employee is on Full-Day Leave or if date is in the FUTURE, Worked Hours MUST be 0.00!
      var shiftHrs = 0;
      if (!isFullDayLeave && !isFutureDate) {
        shiftHrs = Number(log.total_shift_hours) || 0;
        if (shiftHrs <= 0 && log.start_time && (shiftStat === 'OPEN' || shiftStat === 'ON_BREAK') && log.date === todayStr) {
          try {
            var st = parseTimeStringToDate(log.start_time, log.date, timeZone);
            var nowD = new Date();
            if (st && nowD > st) {
              shiftHrs = Math.max(0, ((nowD.getTime() - st.getTime()) / (1000 * 60 * 60)) - ((Number(log.break_minutes) || 0) / 60));
            }
          } catch(err) {}
        } else if (shiftHrs <= 0 && log.start_time && log.end_time) {
          try {
            var st = parseTimeStringToDate(log.start_time, log.date, timeZone);
            var et = parseTimeStringToDate(log.end_time, log.date, timeZone);
            if (st && et && et > st) {
              shiftHrs = Math.max(0, ((et.getTime() - st.getTime()) / (1000 * 60 * 60)) - ((Number(log.break_minutes) || 0) / 60));
            }
          } catch(err) {}
        }
        if (isHalfDayLeave) {
          shiftHrs = Math.min(4.5, shiftHrs); // Cap half day worked hours at 4.5 hrs max
        }
        empObj.total_shift_hours = (empObj.total_shift_hours || 0) + shiftHrs;
        empObj.total_net_worked = (empObj.total_net_worked || 0) + shiftHrs;
        empObj.total_break_minutes += (Number(log.break_minutes) || 0);
      }
    }

    for (var tIdx = 0; tIdx < filteredTaskLogs.length; tIdx++) {
      var task = filteredTaskLogs[tIdx];
      if (task.date && task.date > todayStr) continue; // Skip future task logs if any
      var rawId = task.emp_id ? String(task.emp_id).trim() : '';
      var normId = rawId.toUpperCase();

      if (!empMonitorMap[normId]) {
        continue; // Only track tasks for active directory staff (skip management/inactive accounts)
      }
      var empObj = empMonitorMap[normId];
      var tActual = Number(task.actual_worked_hours) || 0;
      empObj.total_task_hours = (empObj.total_task_hours || 0) + tActual;
      empObj.total_actual_worked = (empObj.total_actual_worked || 0) + tActual;
      empObj.total_earned_hours += (Number(task.productive_hours) || 0);
      empObj.total_leakage_hours += (Number(task.leakage_hours) || 0);
      empObj.total_idle_hours += (Number(task.idle_hours) || 0);
      empObj.total_bench_hours += (Number(task.bench_hours) || 0);
      empObj.total_blocked_hours += (Number(task.blocked_hours) || 0);
      empObj.tasks_count++;
      empObj.tasks_list.push(task);
    }

    // Calculate non-unique days_not_clocked for each employee across period dates
    var monitorKeys = Object.keys(empMonitorMap);
    for (var mk = 0; mk < monitorKeys.length; mk++) {
      var e = empMonitorMap[monitorKeys[mk]];
      // Strict constraint: total logged days per employee cannot exceed the number of days in the period
      e.days_present = Math.min(periodDaysCount, Number(e.days_present) || 0);
      e.days_leave = Math.min(periodDaysCount - e.days_present, Number(e.days_leave) || 0);
      var loggedDays = (e.days_present || 0) + (e.days_leave || 0);
      e.days_not_clocked = Math.max(0, periodDaysCount - loggedDays);
      var effBasis = e.total_task_hours > 0 ? e.total_task_hours : (e.total_shift_hours || 0);
      e.speed_efficiency_pct = effBasis > 0 ? Math.round((e.total_earned_hours / effBasis) * 100) : 100;
    }

    var empMonitorList = [];
    for (var mk = 0; mk < monitorKeys.length; mk++) {
      empMonitorList.push(empMonitorMap[monitorKeys[mk]]);
    }

    // Helper function to check if present employee is Utilized
    function checkEmpUtilized(e) {
      var presDays = Number(e.days_present) || 0;
      if (presDays <= 0) return false;

      // Case 1: If employee HAS entered actual task logs, check actual logged project name
      if (e.tasks_list && e.tasks_list.length > 0) {
        var hasProdTask = false;
        for (var tk = 0; tk < e.tasks_list.length; tk++) {
          var t = e.tasks_list[tk];
          var pUpper = String(t.project || t.project_name || t.project_title || '').toUpperCase();
          var cUpper = String(t.client || '').toUpperCase();
          var isBenchP = (pUpper.indexOf('BENCH') !== -1 || pUpper.indexOf('FLOOR MANAGEMENT') !== -1 || pUpper.indexOf('UNALLOCATED') !== -1 || pUpper.indexOf('IDLE') !== -1);
          var isBenchC = (cUpper === 'BENCH' || cUpper === 'FLOOR MANAGEMENT' || cUpper === 'UNALLOCATED');
          if (!isBenchP && !isBenchC) {
            hasProdTask = true;
            break;
          }
        }
        return hasProdTask;
      }

      // Case 2: If NO task logs entered yet (e.g. Morning 10 AM), check User_Master default client & default project
      var clientUpper = String(e.client || '').toUpperCase();
      var isBenchClient = (clientUpper === 'BENCH' || clientUpper === 'FLOOR MANAGEMENT' || clientUpper === 'UNALLOCATED');
      if (isBenchClient) return false;

      var defProj = String(e.default_project || e.project || e.project_name || '').trim().toUpperCase();
      if (!defProj) return false; // Default project empty -> Bench until task log entry

      var isDefBenchP = (defProj.indexOf('BENCH') !== -1 || defProj.indexOf('FLOOR MANAGEMENT') !== -1 || defProj.indexOf('UNALLOCATED') !== -1 || defProj.indexOf('IDLE') !== -1);
      return !isDefBenchP;
    }

    // Helper function to determine effective client (prioritizing production/billable client tasks over bench tasks)
    function getEmpEffectiveClient(e) {
      var tList = e.tasks_list || e.tasks || [];
      if (tList && tList.length > 0) {
        var prodTask = null;
        for (var tk = 0; tk < tList.length; tk++) {
          var t = tList[tk];
          var pUpper = String(t.project || t.project_name || t.active_project || '').toUpperCase();
          var cUpper = String(t.client || t.active_client || '').toUpperCase();
          var isBenchP = (pUpper.indexOf('BENCH') !== -1 || pUpper.indexOf('FLOOR MANAGEMENT') !== -1 || pUpper.indexOf('UNALLOCATED') !== -1 || pUpper.indexOf('IDLE') !== -1);
          var isBenchC = (cUpper === 'BENCH' || cUpper === 'FLOOR MANAGEMENT' || cUpper === 'UNALLOCATED');
          if (!isBenchP && !isBenchC) {
            prodTask = t;
            break;
          }
        }
        if (prodTask) {
          return prodTask.client || prodTask.active_client || e.client || 'iMerit';
        }
        return tList[0].client || tList[0].active_client || e.client || 'iMerit';
      }
      return e.client || e.default_client || 'iMerit';
    }

    // Build Client Summary (Non-Unique Staff Instances)
    var clientSummaryMap = {};
    for (var em = 0; em < empMonitorList.length; em++) {
      var e = empMonitorList[em];
      var empClient = getEmpEffectiveClient(e);
      var c = empClient;
      if (!clientSummaryMap[c]) {
        clientSummaryMap[c] = { client: c, staff: 0, present: 0, not_clocked: 0, utilized: 0, bench: 0, leave: 0, net_worked: 0, earned: 0, leakage: 0, speed_pct: 100 };
      }
      var cObj = clientSummaryMap[c];
      var presDays = Number(e.days_present) || 0;
      var levDays = Number(e.days_leave) || 0;
      var notClkDays = Number(e.days_not_clocked) || 0;
      var isUtil = checkEmpUtilized(e);

      cObj.present += presDays;
      cObj.leave += levDays;
      cObj.not_clocked += notClkDays;
      cObj.staff += (presDays + levDays + notClkDays);
      cObj.utilized += (isUtil ? presDays : 0);
      cObj.bench += (!isUtil && presDays > 0 ? presDays : 0);
      cObj.net_worked += Number(e.total_net_worked) || 0;
      cObj.earned += Number(e.total_earned_hours) || 0;
      cObj.leakage += Number(e.total_leakage_hours) || 0;
    }

    var csKeys = Object.keys(clientSummaryMap);
    for (var ck = 0; ck < csKeys.length; ck++) {
      var cObj = clientSummaryMap[csKeys[ck]];
      cObj.speed_pct = cObj.net_worked > 0 ? Math.round((cObj.earned / cObj.net_worked) * 100) : 100;
    }

    // Build Team Summary (Non-Unique Staff Instances)
    var teamSummaryMap = {};
    for (var em = 0; em < empMonitorList.length; em++) {
      var e = empMonitorList[em];
      var tm = e.team || 'Annotation';
      if (!teamSummaryMap[tm]) {
        teamSummaryMap[tm] = { team: tm, staff: 0, present: 0, not_clocked: 0, utilized: 0, bench: 0, leave: 0, net_worked: 0, earned: 0, leakage: 0, speed_pct: 100 };
      }
      var tmObj = teamSummaryMap[tm];
      var presDays = Number(e.days_present) || 0;
      var levDays = Number(e.days_leave) || 0;
      var notClkDays = Number(e.days_not_clocked) || 0;
      var isUtil = checkEmpUtilized(e);

      tmObj.present += presDays;
      tmObj.leave += levDays;
      tmObj.not_clocked += notClkDays;
      tmObj.staff += (presDays + levDays + notClkDays);
      tmObj.utilized += (isUtil ? presDays : 0);
      tmObj.bench += (!isUtil && presDays > 0 ? presDays : 0);
      tmObj.net_worked += Number(e.total_net_worked) || 0;
      tmObj.earned += Number(e.total_earned_hours) || 0;
      tmObj.leakage += Number(e.total_leakage_hours) || 0;
    }

    var tsKeys = Object.keys(teamSummaryMap);
    for (var tk = 0; tk < tsKeys.length; tk++) {
      var tmObj = teamSummaryMap[tsKeys[tk]];
      tmObj.speed_pct = tmObj.net_worked > 0 ? Math.round((tmObj.earned / tmObj.net_worked) * 100) : 100;
    }

    // Build PM Summary (Non-Unique Staff Instances)
    var pmSummaryMap = {};
    for (var em = 0; em < empMonitorList.length; em++) {
      var e = empMonitorList[em];
      var pm = e.pm || 'Dinesh Kumar Rajagopal';
      if (!pmSummaryMap[pm]) {
        pmSummaryMap[pm] = { pm: pm, staff: 0, present: 0, not_clocked: 0, utilized: 0, bench: 0, leave: 0, net_worked: 0, earned: 0, leakage: 0, speed_pct: 100 };
      }
      var pmObj = pmSummaryMap[pm];
      var presDays = Number(e.days_present) || 0;
      var levDays = Number(e.days_leave) || 0;
      var notClkDays = Number(e.days_not_clocked) || 0;
      var isUtil = checkEmpUtilized(e);

      pmObj.present += presDays;
      pmObj.leave += levDays;
      pmObj.not_clocked += notClkDays;
      pmObj.staff += (presDays + levDays + notClkDays);
      pmObj.utilized += (isUtil ? presDays : 0);
      pmObj.bench += (!isUtil && presDays > 0 ? presDays : 0);
      pmObj.net_worked += Number(e.total_net_worked) || 0;
      pmObj.earned += Number(e.total_earned_hours) || 0;
      pmObj.leakage += Number(e.total_leakage_hours) || 0;
    }

    var psKeys = Object.keys(pmSummaryMap);
    for (var pk = 0; pk < psKeys.length; pk++) {
      var pmObj = pmSummaryMap[psKeys[pk]];
      pmObj.speed_pct = pmObj.net_worked > 0 ? Math.round((pmObj.earned / pmObj.net_worked) * 100) : 100;
    }

    var finalResult = {
      success: true,
      empMonitorList: empMonitorList,
      clientSummary: Object.keys(clientSummaryMap).map(function(k) { return clientSummaryMap[k]; }),
      teamSummary: Object.keys(teamSummaryMap).map(function(k) { return teamSummaryMap[k]; }),
      pmSummary: Object.keys(pmSummaryMap).map(function(k) { return pmSummaryMap[k]; }),
      availableMonths: availableMonths,
      monthWeeksMap: formattedMonthWeeksMap
    };

    return JSON.parse(JSON.stringify(finalResult));

  } catch (err) {
    return { success: false, message: 'Executive Analytics Backend Error: ' + err.toString() };
  }
}
