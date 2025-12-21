/**
 * ER Shift Management System - Web App Version
 * 
 * Setup:
 * 1. Create a Google Sheet.
 * 2. Config Sheet: "Year", "Month", "Calendar ID".
 * 3. Deploy as Web App: Include this script.
 */

const CONFIG_SHEET_NAME = 'Config';
const HOLIDAY_SHEET_NAME = 'Holidays';
const MASTER_LOG_SHEET_NAME = 'Master_Log';
const DOCTOR_SHEET_NAME = 'Doctors';
const DEBUG_LOG_SHEET_NAME = 'Debug_Log';

// Cache duration in seconds (10 minutes)
const CACHE_DURATION = 600;

// Shifts Sheet Name
const SHIFTS_SHEET_NAME = 'Shifts';

// ===== SHEET MANAGEMENT FUNCTIONS =====

/**
 * Get or create Shifts sheet with proper structure
 */
function getShiftsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHIFTS_SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHIFTS_SHEET_NAME);
    
    // Set up headers
    const headers = [
      '날짜', '근무명', '시작시간', '종료시간', 
      '상태', '의사이름', '의사이메일', '생성일시', '수정일시'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Format header row
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#4285f4')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    
    // Freeze header row
    sheet.setFrozenRows(1);
    
    // Set column widths
    sheet.setColumnWidth(1, 100); // 날짜
    sheet.setColumnWidth(2, 120); // 근무명
    sheet.setColumnWidth(3, 80);  // 시작시간
    sheet.setColumnWidth(4, 80);  // 종료시간
    sheet.setColumnWidth(5, 100); // 상태
    sheet.setColumnWidth(6, 100); // 의사이름
    sheet.setColumnWidth(7, 200); // 의사이메일
    sheet.setColumnWidth(8, 150); // 생성일시
    sheet.setColumnWidth(9, 150); // 수정일시
  }
  
  return sheet;
}

/**
 * Find shift row by date and shift name
 * Returns row number or -1 if not found
 */
function findShiftRow(sheet, dateStr, shiftName) {
  const data = sheet.getDataRange().getValues();
  
  // Skip header row (index 0)
  for (let i = 1; i < data.length; i++) {
    const rowDate = data[i][0]; // Column A: 날짜
    const rowShiftName = data[i][1]; // Column B: 근무명
    
    // Convert date to string for comparison
    const rowDateStr = typeof rowDate === 'string' 
      ? rowDate 
      : Utilities.formatDate(new Date(rowDate), sheet.getParent().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
    
    if (rowDateStr === dateStr && rowShiftName === shiftName) {
      return i + 1; // Return 1-indexed row number
    }
  }
  
  return -1; // Not found
}

/**
 * Safely format time value (handles both string "HH:mm" and Date objects)
 */
function formatTimeValue(timeValue) {
  if (!timeValue) return '00:00';
  
  // If it's already a string in HH:mm format, return it
  if (typeof timeValue === 'string') {
    return timeValue;
  }
  
  // If it's a Date object, extract hours and minutes
  if (timeValue instanceof Date) {
    const hours = String(timeValue.getHours()).padStart(2, '0');
    const minutes = String(timeValue.getMinutes()).padStart(2, '0');
    return `${hours}:${minutes}`;
  }
  
  // If it's a number (Excel serial date), convert to Date first
  if (typeof timeValue === 'number') {
    const date = new Date(timeValue);
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    return `${hours}:${minutes}`;
  }
  
  return '00:00';
}


/**
 * Convert sheet row to shift object
 */
function getShiftByRow(sheet, rowNum) {
  const row = sheet.getRange(rowNum, 1, 1, 9).getValues()[0];
  const tz = sheet.getParent().getSpreadsheetTimeZone();
  
  // Parse date
  const dateStr = typeof row[0] === 'string' 
    ? row[0] 
    : Utilities.formatDate(new Date(row[0]), tz, 'yyyy-MM-dd');
  
  const shiftName = row[1];
  const startTime = formatTimeValue(row[2]); // Convert to HH:mm string
  const endTime = formatTimeValue(row[3]);   // Convert to HH:mm string
  const status = row[4];    // OPEN or CONFIRMED
  const doctorName = row[5] || '';
  const doctorEmail = row[6] || '';
  
  // Create full datetime strings for start and end
  const [startH, startM] = startTime.split(':');
  const [endH, endM] = endTime.split(':');
  
  const startDate = new Date(dateStr);
  startDate.setHours(parseInt(startH), parseInt(startM), 0);
  
  let endDate = new Date(dateStr);
  // If end time is earlier than start time, it's next day
  if (parseInt(endH) < parseInt(startH)) {
    endDate.setDate(endDate.getDate() + 1);
  }
  endDate.setHours(parseInt(endH), parseInt(endM), 0);
  
  return {
    id: `${dateStr}_${shiftName}`, // Composite key
    title: status === 'OPEN' ? shiftName : `${doctorName} (${shiftName})`,
    shiftName: shiftName,
    start: Utilities.formatDate(startDate, tz, "yyyy-MM-dd'T'HH:mm:ss"),
    end: Utilities.formatDate(endDate, tz, "yyyy-MM-dd'T'HH:mm:ss"),
    status: status,
    doctorName: doctorName
  };
}

// Shift Configurations
// Default Shift Configurations (Used for initialization)
const DEFAULT_SHIFTS_WEEKDAY = [
  { name: '평일 오전', start: '08:30', end: '13:00' },
  { name: '평일 오후', start: '13:00', end: '17:30' },
  { name: '평일 야간', start: '17:30', end: '08:30' }
];

const DEFAULT_SHIFTS_HOLIDAY = [
  { name: '휴일 주간', start: '08:00', end: '20:00' },
  { name: '휴일 야간', start: '20:00', end: '08:00' }
];

// Helper to parse "HH:mm" or number 8 or Date object -> {h: 8, m: 0}
function parseTime(input) {
  // Handle Date object (from spreadsheet cells formatted as time)
  if (input instanceof Date) {
    return { h: input.getHours(), m: input.getMinutes() };
  }
  // Handle number
  if (typeof input === 'number') {
    return { h: input, m: 0 };
  }
  // Handle string "HH:mm"
  if (typeof input === 'string') {
    const parts = input.split(':');
    return { 
      h: parseInt(parts[0]), 
      m: parts.length > 1 ? parseInt(parts[1]) : 0 
    };
  }
  return { h: 0, m: 0 };
}

// --- Web App Entry Point ---

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('ER 근무 신청 시스템')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// --- API Functions (Called from Client) ---

/**
 * Fetches shift events for the configured month or requested month.
 * NOW READS FROM SHIFTS SHEET (5-10x faster than Calendar API)
 */
function getShiftData(reqYear, reqMonth) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig(ss);
  
  // Use requested date if valid, otherwise use config
  const year = reqYear ? parseInt(reqYear) : config.year;
  const month = reqMonth ? parseInt(reqMonth) : config.month;

  // Get Shifts sheet
  const shiftsSheet = getShiftsSheet();
  const data = shiftsSheet.getDataRange().getValues();
  
  const shifts = [];
  const tz = ss.getSpreadsheetTimeZone();
  
  // Skip header row (index 0)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    // Parse date
    const dateStr = typeof row[0] === 'string' 
      ? row[0] 
      : Utilities.formatDate(new Date(row[0]), tz, 'yyyy-MM-dd');
    
    // Check if this shift belongs to the requested month
    const shiftDate = new Date(dateStr);
    if (shiftDate.getFullYear() !== year || (shiftDate.getMonth() + 1) !== month) {
      continue; // Skip shifts from other months
    }
    
    const shiftName = row[1];
    const startTime = formatTimeValue(row[2]); // Convert to HH:mm string
    const endTime = formatTimeValue(row[3]);   // Convert to HH:mm string
    const status = row[4];    // OPEN or CONFIRMED
    const doctorName = row[5] || '';
    const doctorEmail = row[6] || '';
    
    // Create full datetime strings for start and end
    const [startH, startM] = startTime.split(':');
    const [endH, endM] = endTime.split(':');
    
    const startDate = new Date(dateStr);
    startDate.setHours(parseInt(startH), parseInt(startM), 0);
    
    let endDate = new Date(dateStr);
    // If end time is earlier than start time, it's next day
    if (parseInt(endH) < parseInt(startH)) {
      endDate.setDate(endDate.getDate() + 1);
    }
    endDate.setHours(parseInt(endH), parseInt(endM), 0);
    
    shifts.push({
      id: `${dateStr}_${shiftName}`, // Composite key
      title: status === 'OPEN' ? shiftName : `${doctorName} (${shiftName})`,
      shiftName: shiftName,
      start: Utilities.formatDate(startDate, tz, "yyyy-MM-dd'T'HH:mm:ss"),
      end: Utilities.formatDate(endDate, tz, "yyyy-MM-dd'T'HH:mm:ss"),
      status: status,
      doctorName: doctorName
    });
  }
  
  // Return sorted by date + Doctor Email Map
  return {
    year: year,
    month: month,
    // Custom Sort: Morning < Afternoon < Night
    shifts: shifts.sort((a, b) => {
      // 1. Sort by Date first
      const dateA = new Date(a.start).setHours(0,0,0,0);
      const dateB = new Date(b.start).setHours(0,0,0,0);
      if (dateA !== dateB) return dateA - dateB;
      
      // 2. Sort by Shift Name Priority
      const getPriority = (name) => {
        if (name.includes('오전') || name.includes('주간')) return 1;
        if (name.includes('오후')) return 2;
        if (name.includes('야간')) return 3;
        return 4;
      };
      
      return getPriority(a.shiftName) - getPriority(b.shiftName);
    })
    // doctorEmails: getDoctorEmails(ss) // REMOVED: Not used by frontend
    // holidays: getHolidays(ss, year, month) // REMOVED: Not used by frontend, causes 1+ second delay
  };
}

/**
 * Books a shift for a doctor.
 * NOW WRITES TO SHIFTS SHEET (5x faster than Calendar API)
 */
function bookShift(shiftId, doctorName, doctorEmail) {
  if (!doctorName || doctorName.trim() === '') {
    throw new Error('이름을 입력해주세요.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Parse composite ID: "YYYY-MM-DD_ShiftName"
  const [dateStr, shiftName] = shiftId.split('_');
  
  if (!dateStr || !shiftName) {
    throw new Error('잘못된 근무 ID입니다.');
  }
  
  // Find shift in Shifts sheet
  const shiftsSheet = getShiftsSheet();
  const rowNum = findShiftRow(shiftsSheet, dateStr, shiftName);
  
  if (rowNum === -1) {
    throw new Error('근무를 찾을 수 없습니다.');
  }
  
  // Update shift row
  const now = new Date();
  shiftsSheet.getRange(rowNum, 5).setValue('CONFIRMED'); // Status
  shiftsSheet.getRange(rowNum, 6).setValue(doctorName);   // Doctor Name
  shiftsSheet.getRange(rowNum, 7).setValue(doctorEmail || ''); // Doctor Email
  shiftsSheet.getRange(rowNum, 9).setValue(now);          // Updated At
  
  // Get shift details for email
  const startTime = formatTimeValue(shiftsSheet.getRange(rowNum, 3).getValue());
  const endTime = formatTimeValue(shiftsSheet.getRange(rowNum, 4).getValue());
  
  // Log to Master_Log
  logShift(ss, new Date(dateStr), `${doctorName} (${doctorEmail || 'No Email'})`, `${doctorName} (${shiftName})`);
  
  // Optional: Create Calendar event if email provided
  if (doctorEmail && doctorEmail.trim() !== '') {
    try {
      createCalendarEventForShift(ss, dateStr, shiftName, startTime, endTime, doctorName, doctorEmail);
    } catch (e) {
      console.error('Calendar event creation failed (non-blocking): ' + e);
      // Don't fail the booking if Calendar creation fails
    }
  }
  
  // Send Email Notification
  logToDebugSheet('[BOOK] Starting email process for: ' + doctorName);
  logToDebugSheet('[BOOK] Doctor Email: ' + doctorEmail);
  
  if (doctorEmail && doctorEmail.trim() !== '') {
    logToDebugSheet('[BOOK] Email present. Preparing to send...');
    try {
      const dateTimeStr = `${dateStr} ${startTime}`;
      const subject = `[응급실 당직 확정] ${doctorName} 선생님 - ${dateTimeStr}`;
      const body = `
        안녕하세요, ${doctorName} 선생님.
        
        응급실 당직 근무가 확정되었습니다.
        
        - 일시: ${dateTimeStr}
        - 근무: ${shiftName}
        - 상태: 확정 (Confirmed)
        
        이 메일은 시스템에서 자동으로 발송되었습니다.
        ${doctorEmail ? '캘린더 초대장이 함께 발송되었으니 일정을 확인해 주세요.' : ''}
        
        감사합니다.
      `;
      
      logToDebugSheet('[BOOK] Sending to: ' + doctorEmail);
      logToDebugSheet('[BOOK] Subject: ' + subject);
      
      MailApp.sendEmail({
        to: doctorEmail,
        subject: subject,
        body: body
      });
      
      logToDebugSheet('[BOOK] SUCCESS: Email sent.');
      
    } catch (e) {
      logToDebugSheet('[BOOK] ERROR: Failed to send email. ' + e.message);
      console.error('[ERROR] Failed to send email: ' + e);
      console.error('[ERROR] Stack trace: ' + e.stack);
      // Email failure should not fail the booking
    }
  } else {
    logToDebugSheet('[BOOK] SKIP: No email provided or empty.');
  }
  
  const successMsg = doctorEmail ? '신청 및 이메일 발송이 완료되었습니다.' : '신청이 완료되었습니다.';
  return { success: true, message: successMsg };
}

/**
 * Creates a Calendar event for a confirmed shift (optional, for notifications)
 */
function createCalendarEventForShift(ss, dateStr, shiftName, startTime, endTime, doctorName, doctorEmail) {
  const config = getConfig(ss);
  
  if (!config.calendarId) {
    console.log('Calendar ID not configured, skipping Calendar event creation');
    return;
  }
  
  const calendar = CalendarApp.getCalendarById(config.calendarId);
  if (!calendar) {
    console.log('Calendar not found, skipping Calendar event creation');
    return;
  }
  
  const tz = ss.getSpreadsheetTimeZone();
  
  // Parse times
  const [startH, startM] = startTime.split(':');
  const [endH, endM] = endTime.split(':');
  
  const startDate = new Date(dateStr);
  startDate.setHours(parseInt(startH), parseInt(startM), 0);
  
  let endDate = new Date(dateStr);
  // If end time is earlier than start time, it's next day
  if (parseInt(endH) < parseInt(startH)) {
    endDate.setDate(endDate.getDate() + 1);
  }
  endDate.setHours(parseInt(endH), parseInt(endM), 0);
  
  const title = `${doctorName} (${shiftName})`;
  const description = `Status: CONFIRMED\nDoctor: ${doctorName}\nShift: ${shiftName}\nCreated via Web App`;
  
  const event = calendar.createEvent(title, startDate, endDate, {
    description: description
  });
  
  // Add guest
  if (doctorEmail) {
    try {
      event.addGuest(doctorEmail);
    } catch (e) {
      console.error('Failed to add guest to Calendar event: ' + e);
    }
  }
  
  console.log(`Calendar event created for ${doctorName} on ${dateStr}`);
}



/**
 * Cancels a confirmed shift.
 * NOW WRITES TO SHIFTS SHEET (5x faster than Calendar API)
 */
function cancelShift(shiftId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Parse composite ID: "YYYY-MM-DD_ShiftName"
  const [dateStr, shiftName] = shiftId.split('_');
  
  if (!dateStr || !shiftName) {
    throw new Error('잘못된 근무 ID입니다.');
  }
  
  // Find shift in Shifts sheet
  const shiftsSheet = getShiftsSheet();
  const rowNum = findShiftRow(shiftsSheet, dateStr, shiftName);
  
  if (rowNum === -1) {
    throw new Error('근무를 찾을 수 없습니다.');
  }
  
  // Get current shift data before clearing
  const status = shiftsSheet.getRange(rowNum, 5).getValue();
  const doctorName = shiftsSheet.getRange(rowNum, 6).getValue();
  const doctorEmail = shiftsSheet.getRange(rowNum, 7).getValue();
  
  if (status !== 'CONFIRMED') {
    throw new Error('취소할 수 없는 상태입니다.');
  }
  
  // Revert to OPEN status
  const now = new Date();
  shiftsSheet.getRange(rowNum, 5).setValue('OPEN');      // Status
  shiftsSheet.getRange(rowNum, 6).setValue('');          // Clear Doctor Name
  shiftsSheet.getRange(rowNum, 7).setValue('');          // Clear Doctor Email
  shiftsSheet.getRange(rowNum, 9).setValue(now);         // Updated At
  
  // Log cancellation
  logShift(ss, new Date(dateStr), 'CANCELLATION', `Cancelled: ${doctorName} (${shiftName})`);
  
  // Send Cancellation Email
  if (doctorEmail) {
    try {
      const startTime = formatTimeValue(shiftsSheet.getRange(rowNum, 3).getValue());
      const dateTimeStr = `${dateStr} ${startTime}`;
      const subject = `[응급실 당직 취소] ${shiftName} - ${dateTimeStr}`;
      const body = `
        안녕하세요.
        
        다음 당직 근무 신청이 취소되었습니다.
        
        - 일시: ${dateTimeStr}
        - 근무: ${shiftName}
        - 상태: 취소됨 (Cancelled)
        
        이 메일은 시스템에서 자동으로 발송되었습니다.
        감사합니다.
      `;
      
      MailApp.sendEmail({
        to: doctorEmail,
        subject: subject,
        body: body
      });
      
      logToDebugSheet('[CANCEL] Cancel email sent to: ' + doctorEmail);
    } catch(e) {
      logToDebugSheet('[CANCEL] ERROR: ' + e.message);
      console.error('Cancellation email error: ' + e);
    }
  }
  
  // Optional: Delete Calendar event if it exists
  try {
    deleteCalendarEventForShift(ss, dateStr, shiftName, doctorName);
  } catch (e) {
    console.error('Calendar event deletion failed (non-blocking): ' + e);
  }

  return { success: true, message: '신청이 취소되었습니다.' };
}

/**
 * Deletes Calendar event for a cancelled shift (optional)
 */
function deleteCalendarEventForShift(ss, dateStr, shiftName, doctorName) {
  const config = getConfig(ss);
  
  if (!config.calendarId) {
    return;
  }
  
  const calendar = CalendarApp.getCalendarById(config.calendarId);
  if (!calendar) {
    return;
  }
  
  // Search for event on that date with matching title
  const searchDate = new Date(dateStr);
  const nextDay = new Date(searchDate);
  nextDay.setDate(nextDay.getDate() + 1);
  
  const events = calendar.getEvents(searchDate, nextDay);
  const titleToFind = `${doctorName} (${shiftName})`;
  
  for (let event of events) {
    if (event.getTitle() === titleToFind) {
      event.deleteEvent();
      console.log(`Calendar event deleted for ${doctorName} on ${dateStr}`);
      return;
    }
  }
}


// --- Admin / Helper Functions ---

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('ER 근무 관리')
    .addItem('근무표 슬롯 생성', 'generateMonthlySlots')
    .addItem('근무표 슬롯 삭제 (초기화)', 'deleteMonthlySlots')
    .addSeparator()
    .addItem('캐시 삭제 (데이터 갱신)', 'clearCache')
    .addToUi();
}

function generateMonthlySlots() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configParams = getConfig(ss);
  
  const year = configParams.year;
  const month = configParams.month - 1; 
  
  const startDate = new Date(year, month, 1);
  const endDate = new Date(year, month + 1, 0); 
  
  const holidays = getHolidays(ss, year, month + 1); // getHolidays expects 1-based month for API
  const shiftConfig = getShiftConfig(ss);
  const shiftsSheet = getShiftsSheet();
  const tz = ss.getSpreadsheetTimeZone();
  
  // Prepare bulk data for insertion
  const rowsToInsert = [];
  const now = new Date();

  for (let d = 1; d <= endDate.getDate(); d++) {
    const currentDate = new Date(year, month, d);
    const dateString = Utilities.formatDate(currentDate, tz, 'yyyy-MM-dd');
    
    // Check if holiday (formatting matches keys in getHolidays)
    const holidayName = holidays[dateString]; 
    const isHoliday = isWeekend(currentDate) || !!holidayName;
    
    // Check Next Day for Night Shift Logic
    const nextDate = new Date(currentDate);
    nextDate.setDate(currentDate.getDate() + 1);
    const nextDateString = Utilities.formatDate(nextDate, tz, 'yyyy-MM-dd');
    const isNextDayHoliday = isWeekend(nextDate) || !!holidays[nextDateString];

    // Choose shift types dynamically
    const shiftTypes = isHoliday ? shiftConfig.holiday : shiftConfig.weekday;
    
    shiftTypes.forEach(shift => {
      // Calculate end time based on next day logic
      let endTime = shift.end;
      
      if (shift.name === '평일 야간' && isNextDayHoliday) {
        endTime = '08:00';
      } else if (shift.name === '휴일 야간' && !isNextDayHoliday) {
        endTime = '08:30';
      }
      
      // Add row: [날짜, 근무명, 시작시간, 종료시간, 상태, 의사이름, 의사이메일, 생성일시, 수정일시]
      rowsToInsert.push([
        dateString,
        shift.name,
        shift.start,
        endTime,
        'OPEN',
        '',
        '',
        now,
        now
      ]);
    });
  }
  
  // Bulk insert all rows at once (much faster than individual inserts)
  if (rowsToInsert.length > 0) {
    const lastRow = shiftsSheet.getLastRow();
    shiftsSheet.getRange(lastRow + 1, 1, rowsToInsert.length, 9).setValues(rowsToInsert);
  }
  
  SpreadsheetApp.getUi().alert(`슬롯 생성이 완료되었습니다. (총 ${rowsToInsert.length}개)`);
}

function deleteMonthlySlots() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configParams = getConfig(ss);
  
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '슬롯 삭제 확인',
    `${configParams.year}년 ${configParams.month}월의 OPEN 상태 슬롯을 모두 삭제하시겠습니까?\n(주의: 확정된 근무는 삭제되지 않습니다.)`,
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const shiftsSheet = getShiftsSheet();
  const year = configParams.year;
  const month = configParams.month;
  const tz = ss.getSpreadsheetTimeZone();
  
  // Get all data
  const data = shiftsSheet.getDataRange().getValues();
  let deletedCount = 0;
  
  // Delete from bottom to top to avoid row number shifting
  for (let i = data.length - 1; i >= 1; i--) { // Skip header row
    const row = data[i];
    const dateStr = typeof row[0] === 'string' 
      ? row[0] 
      : Utilities.formatDate(new Date(row[0]), tz, 'yyyy-MM-dd');
    const status = row[4];
    
    // Check if this shift belongs to the target month
    const shiftDate = new Date(dateStr);
    if (shiftDate.getFullYear() === year && (shiftDate.getMonth() + 1) === month) {
      // Delete if OPEN status
      if (status === 'OPEN') {
        shiftsSheet.deleteRow(i + 1); // +1 because row numbers are 1-indexed
        deletedCount++;
      }
    }
  }

  ui.alert(`삭제 완료: 총 ${deletedCount}개의 슬롯이 삭제되었습니다.`);
}

// createSlotEvent is no longer needed - we insert directly to Shifts sheet in generateMonthlySlots()

function getConfig(ss) {
  let sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) {
      sheet = ss.insertSheet(CONFIG_SHEET_NAME);
      sheet.getRange("A1:A3").setValues([["Year"], ["Month"], ["Calendar ID"]]);
      sheet.getRange("B1:B3").setValues([[new Date().getFullYear()], [new Date().getMonth() + 1], ["YOUR_CALENDAR_ID"]]);
      
      // Initialize Shift Config Table at Row 5
      sheet.getRange("A5:D5").setValues([["Shift Type", "Shift Name", "Start Hour", "End Hour"]]);
      sheet.getRange("A5:D5").setBackground("#efefef").setFontWeight("bold");
      
      const defaults = [
        ["WEEKDAY", "평일 오전", "08:30", "13:00"],
        ["WEEKDAY", "평일 오후", "13:00", "17:30"],
        ["WEEKDAY", "평일 야간", "17:30", "08:30"],
        ["HOLIDAY", "휴일 주간", "08:00", "20:00"],
        ["HOLIDAY", "휴일 야간", "20:00", "08:00"]
      ];
      sheet.getRange(6, 1, defaults.length, 4).setValues(defaults);
  }
  return {
    year: sheet.getRange("B1").getValue(),
    month: sheet.getRange("B2").getValue(),
    calendarId: sheet.getRange("B3").getValue()
  };
}

function getShiftConfig(ss) {
  let sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  // Row 5 is header, Data starts from Row 6
  // Columns: A=Type, B=Name, C=Start, D=End
  
  if (!sheet || sheet.getLastRow() < 6) {
    // Return Defaults if not configured
    return {
      weekday: DEFAULT_SHIFTS_WEEKDAY,
      holiday: DEFAULT_SHIFTS_HOLIDAY
    };
  }
  
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(6, 1, lastRow - 5, 4).getValues();
  
  const weekdayShifts = [];
  const holidayShifts = [];
  
  data.forEach(row => {
    const type = row[0]; // WEEKDAY or HOLIDAY
    const name = row[1];
    const start = row[2];
    const end = row[3];
    
    if (type && name && start !== '' && end !== '') {
      const shift = { name: name, start: start, end: end };
      if (type === 'WEEKDAY') {
        weekdayShifts.push(shift);
      } else if (type === 'HOLIDAY') {
        holidayShifts.push(shift);
      }
    }
  });
  
  // Fallback if empty
  if (weekdayShifts.length === 0) return { weekday: DEFAULT_SHIFTS_WEEKDAY, holiday: DEFAULT_SHIFTS_HOLIDAY };
  
  return {
    weekday: weekdayShifts,
    holiday: holidayShifts
  };
}

function getHolidays(ss, year, month) {
  // Returns object: { 'YYYY-MM-DD': 'Holiday Name' }
  const cacheKey = `holidays_${year}_${month}`;
  const cache = CacheService.getScriptCache();
  
  // Try to get from cache first
  const cached = cache.get(cacheKey);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {
      console.error('Cache parse error: ' + e);
    }
  }
  
  let holidays = {};
  
  // 1. Fetch from Google Holiday Calendar
  try {
    const calendarId = 'ko.south_korea#holiday@group.v.calendar.google.com';
    const calendar = CalendarApp.getCalendarById(calendarId);
    if (calendar) {
      // Fetch for the whole month
      const start = new Date(year, month - 1, 1);
      const end = new Date(year, month, 0);
      const events = calendar.getEvents(start, end);
      events.forEach(e => {
        const dateStr = Utilities.formatDate(e.getStartTime(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
        holidays[dateStr] = e.getTitle();
      });
    }
  } catch (e) {
    console.error('Failed to fetch Google Holidays: ' + e);
  }

  // 2. Fetch from Manual Holiday Sheet (Merge, overwrites Google)
  const sheet = ss.getSheetByName(HOLIDAY_SHEET_NAME);
  if (sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      // Expecting Col A: Date, Col B: Name
      const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
      data.forEach(row => {
        const d = row[0];
        const name = row[1];
        if (d) {
           const dateStr = Utilities.formatDate(new Date(d), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
           // If name is '평일' or 'WEEKDAY', remove from holidays map (Override Google)
           if (name && (name.trim() === '평일' || name.trim().toUpperCase() === 'WEEKDAY')) {
             delete holidays[dateStr];
           } else {
             holidays[dateStr] = name || '휴일';
           }
        }
      });
    }
  }
  
  // Cache the result
  try {
    cache.put(cacheKey, JSON.stringify(holidays), CACHE_DURATION);
  } catch (e) {
    console.error('Failed to cache holidays: ' + e);
  }
  
  return holidays;
}

function isWeekend(date) {
  const day = date.getDay();
  return day === 0 || day === 6; // Sun or Sat
}


function getDoctorEmails(ss) {
  const cacheKey = 'doctor_emails';
  const cache = CacheService.getScriptCache();
  
  // Try to get from cache first
  const cached = cache.get(cacheKey);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {
      console.error('Cache parse error: ' + e);
    }
  }
  
  let sheet = ss.getSheetByName(DOCTOR_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(DOCTOR_SHEET_NAME);
    sheet.appendRow(['Name', 'Email']);
    return {};
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {}; // Header only
  
  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const map = {};
  data.forEach(row => {
    const name = row[0];
    const email = row[1];
    if (name && email) {
      map[name] = email;
    }
  });
  
  // Cache the result
  try {
    cache.put(cacheKey, JSON.stringify(map), CACHE_DURATION);
  } catch (e) {
    console.error('Failed to cache doctor emails: ' + e);
  }
  
  return map;
}

function logShift(ss, date, nameOrEmail, title) {
  let sheet = ss.getSheetByName(MASTER_LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(MASTER_LOG_SHEET_NAME);
    sheet.appendRow(['Shift Date', 'Name', 'Shift Title', 'Timestamp']);
  }
  sheet.appendRow([date, nameOrEmail, title, new Date()]);
}

function logToDebugSheet(message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(DEBUG_LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(DEBUG_LOG_SHEET_NAME);
    sheet.appendRow(['Timestamp', 'Message']);
  }
  sheet.appendRow([new Date(), message]);
}

// ===== CACHE MANAGEMENT =====
/**
 * Clears all cached data. Run this manually when you update:
 * - Doctors sheet (email addresses)
 * - Holidays sheet (holiday dates)
 */
function clearCache() {
  const cache = CacheService.getScriptCache();
  cache.removeAll(['doctor_emails']);
  
  // Clear holiday caches for current and nearby months
  const now = new Date();
  const year = now.getFullYear();
  for (let month = 1; month <= 12; month++) {
    cache.remove(`holidays_${year}_${month}`);
    cache.remove(`holidays_${year-1}_${month}`);
    cache.remove(`holidays_${year+1}_${month}`);
  }
  
  Logger.log('Cache cleared successfully!');
  SpreadsheetApp.getUi().alert('캐시가 성공적으로 삭제되었습니다.\n다음 로딩 시 최신 데이터가 반영됩니다.');
}

// ===== TEST FUNCTION - 이메일 발송 테스트용 =====
// 이 함수를 직접 실행해서 이메일 발송을 테스트할 수 있습니다
function testEmailSending() {
  const testEmail = 'YOUR_EMAIL@gmail.com'; // 여기에 본인 이메일 주소를 입력하세요
  
  console.log('[TEST] Starting email test...');
  console.log('[TEST] Target email: ' + testEmail);
  
  try {
    MailApp.sendEmail({
      to: testEmail,
      subject: '[테스트] 응급실 당직 시스템 이메일 발송 테스트',
      body: '이 메일이 도착했다면 이메일 발송 기능이 정상 작동하는 것입니다.\n\n테스트 시간: ' + new Date()
    });
    
    console.log('[TEST] Email sent successfully!');
    logToDebugSheet('[TEST] SUCCESS: Email sent to ' + testEmail);
    Logger.log('[TEST] 이메일 발송 성공! 받은편지함을 확인하세요.');
    return '성공: 이메일이 발송되었습니다. 받은편지함을 확인하세요.';
    
  } catch (e) {
    console.error('[TEST ERROR] Failed: ' + e);
    logToDebugSheet('[TEST] ERROR: ' + e.message);
    Logger.log('[TEST ERROR] 실패: ' + e.message);
    return '실패: ' + e.message;
  }
}
