/**
 * ER Shift Management System
 * 
 * Setup:
 * 1. Create a Google Sheet.
 * 2. Create a "Config" sheet with "Year", "Month", and "Calendar ID".
 * 3. Create a "Holidays" sheet with a list of update dates (YYYY-MM-DD).
 * 4. Create a "Doctors" sheet with "Name" (Col A) and "Email" (Col B).
 * 5. Tools > Script Editor > Paste this code.
 */

const CONFIG_SHEET_NAME = 'Config';
const HOLIDAY_SHEET_NAME = 'Holidays';
const MASTER_LOG_SHEET_NAME = 'Master_Log';
const DOCTOR_SHEET_NAME = 'Doctors';

// Shift Configurations
const SHIFTS_WEEKDAY = [
  { name: '평일 오전', start: 8, end: 13 },
  { name: '평일 오후', start: 13, end: 18 },
  { name: '평일 야간', start: 18, end: 32 }  // Ends next day 8am
];

const SHIFTS_HOLIDAY = [
  { name: '휴일 주간', start: 8, end: 18 },
  { name: '휴일 야간', start: 18, end: 32 }
];

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('ER 근무 관리')
    .addItem('근무표 슬롯 생성', 'generateMonthlySlots')
    .addItem('신청 내역 동기화', 'syncCalendarChanges')
    .addToUi();
}

/**
 * Generates empty shift slots for the specified month.
 */
function generateMonthlySlots() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configParams = getConfig(ss);
  
  if (!configParams.calendarId) {
    SpreadsheetApp.getUi().alert('캘린더 ID가 설정되지 않았습니다.');
    return;
  }

  const calendar = CalendarApp.getCalendarById(configParams.calendarId);
  const year = configParams.year;
  const month = configParams.month - 1; // JS Month is 0-indexed
  
  const startDate = new Date(year, month, 1);
  const endDate = new Date(year, month + 1, 0); // Last day of month
  
  const holidays = getHolidays(ss);

  for (let d = 1; d <= endDate.getDate(); d++) {
    const currentDate = new Date(year, month, d);
    const dateString = Utilities.formatDate(currentDate, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
    const isHoliday = isWeekend(currentDate) || holidays.includes(dateString);
    
    const shiftTypes = isHoliday ? SHIFTS_HOLIDAY : SHIFTS_WEEKDAY;
    
    shiftTypes.forEach(shift => {
      createSlotEvent(calendar, currentDate, shift);
    });
  }
  
  SpreadsheetApp.getUi().alert('슬롯 생성이 완료되었습니다.');
}

function createSlotEvent(calendar, date, shift) {
  let start = new Date(date);
  start.setHours(shift.start, 0, 0);
  
  let end = new Date(date);
  if (shift.end >= 24) {
    end.setDate(date.getDate() + 1);
    end.setHours(shift.end - 24, 0, 0);
  } else {
    end.setHours(shift.end, 0, 0);
  }
  
  calendar.createEvent(`[모집] ${shift.name}`, start, end, {
    description: `ShiftType: ${shift.name}\nStatus: OPEN`
    // guests option removed to avoid invalid argument error
  });
}

/**
 * Triggered manually or by time trigger to sync changes
 */
function syncCalendarChanges() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configParams = getConfig(ss);
  const calendar = CalendarApp.getCalendarById(configParams.calendarId);
  const doctorMap = getDoctorMap(ss); // Cache doctor names
  
  // Check the whole month
  const events = calendar.getEvents(
    new Date(configParams.year, configParams.month - 1, 1),
    new Date(configParams.year, configParams.month + 1, 0)
  );

  let updateCount = 0;

  events.forEach(event => {
    const title = event.getTitle();
    const guests = event.getGuestList();
    const realGuests = guests.filter(g => g.getEmail() !== calendar.getId());

    // Case A: New Application ([모집] + Guest exists)
    if (title.startsWith('[모집]') && realGuests.length > 0) {
        // Someone applied!
        const doctor = realGuests[0]; 
        const doctorEmail = doctor.getEmail();
        
        // Normalize for robust matching (trim & lowercase)
        const normalizedEmail = String(doctorEmail).trim().toLowerCase();
        
        Logger.log(`[Processing] Guest: ${doctorEmail} / Normalized: ${normalizedEmail}`);
        
        // Look up name
        const doctorName = doctorMap[normalizedEmail] || doctor.getName() || doctorEmail.split('@')[0];
        
        // 1. Confirm Event
        event.setTitle(`[확정] ${doctorName} (${title.replace('[모집] ', '')})`);
        event.setDescription(event.getDescription().replace('Status: OPEN', 'Status: CONFIRMED'));
        
        // 2. Log to Sheet
        logShift(ss, event.getStartTime(), doctorEmail, title);
        
        updateCount++;
    }
    
    // Case B: Cancellation ([확정] + No Guests)
    else if (title.startsWith('[확정]') && realGuests.length === 0) {
        // Revert to OPEN
        // Extract original shift name from "[확정] Name (ShiftName)"
        // Regex looks for content inside the last parenthesis
        const match = title.match(/\((.+)\)$/);
        const originalShiftName = match ? match[1] : title; // Fallback if regex fails
        
        event.setTitle(`[모집] ${originalShiftName}`);
        event.setDescription(event.getDescription().replace('Status: CONFIRMED', 'Status: OPEN'));
        
        // Optional: Log cancellation
        logShift(ss, event.getStartTime(), 'CANCELED', `Reverted: ${title}`);
        
        updateCount++;
    }
  });
  
  if(updateCount > 0) {
    SpreadsheetApp.getUi().alert(`${updateCount}건의 근무 신청이 확정되었습니다.\n로그를 확인하세요.`);
  } else {
    SpreadsheetApp.getUi().alert('새로운 신청 내역이 없습니다.');
  }
}

// --- Helpers ---

function getConfig(ss) {
  let sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) {
      sheet = ss.insertSheet(CONFIG_SHEET_NAME);
      sheet.getRange("A1:A3").setValues([["Year"], ["Month"], ["Calendar ID"]]);
      sheet.getRange("B1:B3").setValues([[new Date().getFullYear()], [new Date().getMonth() + 1], ["YOUR_CALENDAR_ID"]]);
  }
  return {
    year: sheet.getRange("B1").getValue(),
    month: sheet.getRange("B2").getValue(),
    calendarId: sheet.getRange("B3").getValue()
  };
}

function getDoctorMap(ss) {
  const sheet = ss.getSheetByName(DOCTOR_SHEET_NAME);
  const map = {};
  if (!sheet) return map;
  
  const data = sheet.getDataRange().getValues();
  // Assume Row 1 is header, start from Row 2
  for (let i = 1; i < data.length; i++) {
    const name = data[i][0]; // Col A
    const email = data[i][1]; // Col B
    if (email) {
      const key = String(email).trim().toLowerCase();
      map[key] = name;
    }
  }
  return map;
}

function getHolidays(ss) {
  const sheet = ss.getSheetByName(HOLIDAY_SHEET_NAME);
  if (!sheet) return [];
  const startRow = 2;
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return [];
  
  const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 1).getValues();
  return data.flat().map(d => Utilities.formatDate(new Date(d), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd'));
}

function isWeekend(date) {
  const day = date.getDay();
  return day === 0 || day === 6; // Sun or Sat
}

function logShift(ss, date, email, title) {
  let sheet = ss.getSheetByName(MASTER_LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(MASTER_LOG_SHEET_NAME);
    sheet.appendRow(['Shift Date', 'Doctor Email', 'Shift Title', 'Timestamp']);
  }
  sheet.appendRow([date, email, title, new Date()]);
}
