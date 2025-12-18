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

// --- Web App Entry Point ---

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('ER 근무 신청 시스템')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// --- API Functions (Called from Client) ---

/**
 * Fetches shift events for the configured month.
 */
function getShiftData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig(ss);
  
  if (!config.calendarId) {
    throw new Error('Calendar ID is not configured.');
  }

  const calendar = CalendarApp.getCalendarById(config.calendarId);
  if (!calendar) {
    throw new Error('Calendar not found. Check ID.');
  }

  const start = new Date(config.year, config.month - 1, 1);
  const end = new Date(config.year, config.month, 0); // Last day of month
  
  // Expand range slightly to cover overnight shifts ending on 1st of next month
  const fetchEnd = new Date(end);
  fetchEnd.setDate(fetchEnd.getDate() + 1);

  const events = calendar.getEvents(start, fetchEnd);
  
  const shifts = events.map(event => {
    const title = event.getTitle();
    const id = event.getId();
    const startTime = event.getStartTime();
    const endTime = event.getEndTime();
    
    let status = 'UNKNOWN';
    let doctorName = '';
    let shiftName = title;

    if (title.startsWith('[모집]')) {
      status = 'OPEN';
      shiftName = title.replace('[모집] ', '');
    } else if (title.startsWith('[확정]')) {
      status = 'CONFIRMED';
      // Format: [확정] DoctorName (ShiftName)
      const match = title.match(/\[확정\]\s+(.+)\s+\((.+)\)/);
      if (match) {
        doctorName = match[1];
        shiftName = match[2];
      } else {
        shiftName = title.replace('[확정] ', '');
      }
    }

    return {
      id: id,
      title: title,
      shiftName: shiftName,
      start: startTime.toISOString(),
      end: endTime.toISOString(),
      status: status,
      doctorName: doctorName
    };
  });
  
  // Return sorted by date
  return {
    year: config.year,
    month: config.month,
    shifts: shifts.sort((a, b) => new Date(a.start) - new Date(b.start))
  };
}

/**
 * Books a shift for a doctor.
 */
function bookShift(eventId, doctorName) {
  if (!doctorName || doctorName.trim() === '') {
    throw new Error('이름을 입력해주세요.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig(ss);
  const calendar = CalendarApp.getCalendarById(config.calendarId);
  const event = calendar.getEventById(eventId);
  
  if (!event) {
    throw new Error('이벤트를 찾을 수 없습니다.');
  }

  const title = event.getTitle();
  if (!title.startsWith('[모집]')) {
    throw new Error('이미 마감되었거나 신청할 수 없는 슬롯입니다.');
  }

  const shiftName = title.replace('[모집] ', '');
  
  // Update Event
  const newTitle = `[확정] ${doctorName} (${shiftName})`;
  event.setTitle(newTitle);
  event.setDescription(`Status: CONFIRMED\nDoctor: ${doctorName}\nUpdated via Web App`);
  
  // Log to Sheet
  logShift(ss, event.getStartTime(), doctorName, newTitle);
  
  return { success: true, message: '신청이 완료되었습니다.' };
}


// --- Admin / Helper Functions ---

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('ER 근무 관리')
    .addItem('근무표 슬롯 생성', 'generateMonthlySlots')
    .addToUi();
}

function generateMonthlySlots() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configParams = getConfig(ss);
  
  if (!configParams.calendarId) {
    SpreadsheetApp.getUi().alert('캘린더 ID가 설정되지 않았습니다.');
    return;
  }

  const calendar = CalendarApp.getCalendarById(configParams.calendarId);
  const year = configParams.year;
  const month = configParams.month - 1; 
  
  const startDate = new Date(year, month, 1);
  const endDate = new Date(year, month + 1, 0); 
  
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
  });
}

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

function logShift(ss, date, nameOrEmail, title) {
  let sheet = ss.getSheetByName(MASTER_LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(MASTER_LOG_SHEET_NAME);
    sheet.appendRow(['Shift Date', 'Name', 'Shift Title', 'Timestamp']);
  }
  sheet.appendRow([date, nameOrEmail, title, new Date()]);
}
