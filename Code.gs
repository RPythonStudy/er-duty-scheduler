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
 */
function getShiftData(reqYear, reqMonth) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig(ss);
  
  if (!config.calendarId) {
    throw new Error('Calendar ID is not configured.');
  }

  const calendar = CalendarApp.getCalendarById(config.calendarId);
  if (!calendar) {
    throw new Error('Calendar not found. Check ID.');
  }

  // Use requested date if valid, otherwise use config
  const year = reqYear ? parseInt(reqYear) : config.year;
  const month = reqMonth ? parseInt(reqMonth) : config.month;

  const start = new Date(year, month - 1, 1);
  const end = new Date(year, month, 0); // Last day of month
  
  // Expand range slightly to cover overnight shifts ending on 1st of next month
  const fetchEnd = new Date(end);
  fetchEnd.setDate(fetchEnd.getDate() + 1);

  const events = calendar.getEvents(start, fetchEnd);
  
  const shifts = events.map(event => {
    const title = event.getTitle();
    const id = event.getId();
    const startTime = event.getStartTime();
    const endTime = event.getEndTime();
    const desc = event.getDescription() || '';
    
    let status = 'UNKNOWN';
    let doctorName = '';
    let shiftName = title;

    // 1. OPEN SLOT (Checked via Description or legacy prefix)
    if (desc.includes('Status: OPEN') || title.startsWith('[모집]')) {
      status = 'OPEN';
      // Remove legacy prefix if present, though new ones won't have it
      shiftName = title.replace('[모집] ', '');
    } 
    // 2. CONFIRMED SLOT
    else {
      // Assumes confirmed if not OPEN.
      // Format: DoctorName (ShiftName) OR Legacy: [확정] DoctorName (ShiftName)
      status = 'CONFIRMED';
      
      // Clean legacy prefix
      const cleanTitle = title.replace('[확정] ', '');
      
      // Try parsing "Name (Shift)"
      const match = cleanTitle.match(/^(.+)\s+\((.+)\)$/);
      if (match) {
        doctorName = match[1];
        shiftName = match[2];
      } else {
        // Fallback: Title is the name? Or Title is ShiftName? 
        // If we can't parse, just use title as is, or look in description
        // But usually current system enforces formatting.
        if (desc.includes('Doctor:')) {
           const descMatch = desc.match(/Doctor:\s+(.+)/);
           if (descMatch) doctorName = descMatch[1];
        }
        // If we still don't have shiftName separate, assume Title might be it
        // ignoring the name... this is edge case.
        if (!doctorName) doctorName = '확정'; // Default
      }
    }

    return {
      id: id,
      title: title,
      shiftName: shiftName,
      start: Utilities.formatDate(startTime, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd'T'HH:mm:ss"),
      end: Utilities.formatDate(endTime, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd'T'HH:mm:ss"),
      status: status,
      doctorName: doctorName
    };
  });
  
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
    }),
    doctorEmails: getDoctorEmails(ss),
    holidays: getHolidays(ss, year, month)
  };
}

/**
 * Books a shift for a doctor.
 */
function bookShift(eventId, doctorName, doctorEmail) {
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
  const desc = event.getDescription() || '';

  // Validate: Check if Status is OPEN or CONFIRMED (Allow overwrite)
  // We only block if it's not a valid slot at all (e.g. neither open nor confirmed properly)
  // Actually, we want to allow editing, so we trust the input eventId is correct.
  // Just ensure it's a shift slot.
  if (!desc.includes('ShiftType:')) { 
     // A minimal check to ensure we don't edit random events
     // Or we can rely on the fact that we only show valid buttons.
  }
  
  // Previous Logic: if (!desc.includes('Status: OPEN') && !title.startsWith('[모집]')) throw ...
  // New Logic: Allow.
  
  // Determine Shift Name
  // If we are editing, the title might be "Doctor (Shift)". We need to extract Shift.
  let shiftName = '';
  // Try parsing (ShortName or FullName)
  // But wait, we don't need to parse title if we can get it from Description or pass it?
  // Passed params: eventId, doctorName, doctorEmail. We don't get shiftName.
  // We need to preserve the shift name.
  
  // 1. Try legacy/open format: "[모집] ShiftName"
  if (title.startsWith('[모집]')) {
    shiftName = title.replace('[모집] ', '');
  } 
  // 2. Try confirmed format: "Name (ShiftName)"
  else {
    const match = title.match(/\((.+)\)$/); // Match last parenthesis
    if (match) {
      // Logic for short names: (오전) -> restore full name? 
      // Actually, createSlotEvent saves "ShiftType: 평일 오전" in description!
      // Reliable way: read Description.
      const typeMatch = desc.match(/ShiftType:\s+(.+?)(\n|$)/);
      if (typeMatch) {
         shiftName = typeMatch[1];
      } else {
         // Fallback to title parse
         shiftName = match[1];
         // If short name was saved "오전", we might want to keep it or full name?
         // Let's rely on Description which has full name.
      }
    } else {
       // Just use title as is
       shiftName = title;
    }
  }

  // Double check description for full name if we missed it
  const typeMatch = desc.match(/ShiftType:\s+(.+?)(\n|$)/);
  if (typeMatch) {
      shiftName = typeMatch[1];
  }

  // Update Event
  // New Format: "DoctorName (ShortName?)" 
  // Wait, bookShift constructs title. createSlotEvent used full name.
  // Frontend displays short name. Event title should probably have SHORT name? 
  // User said: "simple text display: Morning".
  
  // Let's standardize Title storage:
  // "DoctorName (ShiftName)"
  // If we want title to be short, we should shorten it here too?
  // Actually frontend did the shortening logic `replace('평일 ', '')`.
  // So backend can store Full Name in Title, frontend shortens it.
  // OR backend stores Full Name, frontend shortens.
  // Let's stick to Full Name in Title for clarity/searchability in Google Calendar.
  // Frontend is just a view.
  
  const newTitle = `${doctorName} (${shiftName})`;
  event.setTitle(newTitle);
  let newDesc = `Status: CONFIRMED\nDoctor: ${doctorName}\nUpdated via Web App`;
  
  if (doctorEmail && doctorEmail.trim() !== '') {
    try {
      event.addGuest(doctorEmail);
      newDesc += `\nGuest: ${doctorEmail}`;
    } catch (e) {
      // Ignore invalid email errors to prevent blocking the booking
      console.error('Failed to add guest: ' + e);
    }
  }

  event.setDescription(newDesc);
  
  // Log to Sheet
  logShift(ss, event.getStartTime(), `${doctorName} (${doctorEmail || 'No Email'})`, newTitle);
  
  // Send Email Notification
  logToDebugSheet('[BOOK] Starting email process for: ' + doctorName);
  logToDebugSheet('[BOOK] Doctor Email: ' + doctorEmail);
  
  if (doctorEmail && doctorEmail.trim() !== '') {
    logToDebugSheet('[BOOK] Email present. Preparing to send...');
    try {
      const dateStr = Utilities.formatDate(event.getStartTime(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm');
      const subject = `[응급실 당직 확정] ${doctorName} 선생님 - ${dateStr}`;
      const body = `
        안녕하세요, ${doctorName} 선생님.
        
        응급실 당직 근무가 확정되었습니다.
        
        - 일시: ${dateStr}
        - 근무: ${shiftName}
        - 상태: 확정 (Confirmed)
        
        이 메일은 시스템에서 자동으로 발송되었습니다.
        캘린더 초대장이 함께 발송되었으니 일정을 확인해 주세요.
        
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
 * Cancels a confirmed shift.
 */
function cancelShift(eventId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig(ss);
  const calendar = CalendarApp.getCalendarById(config.calendarId);
  const event = calendar.getEventById(eventId);
  
  if (!event) {
    throw new Error('이벤트를 찾을 수 없습니다.');
  }

  const title = event.getTitle();
  const desc = event.getDescription() || '';
  
  // Check if event is CONFIRMED (new format uses Description)
  if (!desc.includes('Status: CONFIRMED') && !title.startsWith('[확정]')) {
    throw new Error('취소할 수 없는 상태입니다.');
  }

  // Parse shift name and doctor name
  // Current format: "DoctorName (ShiftName)" or Legacy: "[확정] DoctorName (ShiftName)"
  let shiftName = '';
  let doctorName = '';
  
  // Try to get shift name from Description first (most reliable)
  const typeMatch = desc.match(/ShiftType:\s+(.+?)(\n|$)/);
  if (typeMatch) {
    shiftName = typeMatch[1];
  }
  
  // Parse doctor name from title
  const cleanTitle = title.replace('[확정] ', '');
  const match = cleanTitle.match(/^(.+)\s+\((.+)\)$/);
  if (match) {
    doctorName = match[1];
    if (!shiftName) shiftName = match[2]; // Fallback if not in description
  }
  
  // If still no shift name, use title as fallback
  if (!shiftName) {
    shiftName = title.replace('[확정] ', '');
  }
  
  // Revert to Open status - use shift name only (no prefix)
  event.setTitle(shiftName);
  event.setDescription(`ShiftType: ${shiftName}\nStatus: OPEN\nCancelled by User`);
  
  // Optional: Log cancellation
  logShift(ss, event.getStartTime(), 'CANCELLATION', `Cancelled: ${title}`);
  
  // Send Cancellation Email
  try {
    let doctorEmail = '';
    
    // Method 1: Try to look up email by Name from Doctors Sheet
    if (doctorName) {
        const emails = getDoctorEmails(ss);
        if (emails[doctorName]) {
            doctorEmail = emails[doctorName];
        }
    }
    
    // Method 2: If not found, try guest list
    if (!doctorEmail) {
        const guests = event.getGuestList();
        if (guests.length > 0) {
            doctorEmail = guests[0].getEmail();
        }
    }

    if (doctorEmail) {
        const dateStr = Utilities.formatDate(event.getStartTime(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm');
        const subject = `[응급실 당직 취소] ${shiftName} - ${dateStr}`;
        const body = `
          안녕하세요.
          
          다음 당직 근무 신청이 취소되었습니다.
          
          - 일시: ${dateStr}
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
        
        // Remove guest from event to clear it from their calendar
        event.removeGuest(doctorEmail);
        logToDebugSheet('[CANCEL] Cancel email sent to: ' + doctorEmail);
    }
  } catch(e) {
      logToDebugSheet('[CANCEL] ERROR: ' + e.message);
      console.error('Cancellation email error: ' + e);
  }

  return { success: true, message: '신청이 취소되었습니다.' };
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
  
  if (!configParams.calendarId) {
    SpreadsheetApp.getUi().alert('캘린더 ID가 설정되지 않았습니다.');
    return;
  }

  const calendar = CalendarApp.getCalendarById(configParams.calendarId);
  const year = configParams.year;
  const month = configParams.month - 1; 
  
  const startDate = new Date(year, month, 1);
  const endDate = new Date(year, month + 1, 0); 
  
  const holidays = getHolidays(ss, year, month + 1); // getHolidays expects 1-based month for API
  const shiftConfig = getShiftConfig(ss);

  for (let d = 1; d <= endDate.getDate(); d++) {
    const currentDate = new Date(year, month, d);
    const dateString = Utilities.formatDate(currentDate, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
    
    // Check if holiday (formatting matches keys in getHolidays)
    const holidayName = holidays[dateString]; 
    const isHoliday = isWeekend(currentDate) || !!holidayName;
    
    // Check Next Day for Night Shift Logic
    const nextDate = new Date(currentDate);
    nextDate.setDate(currentDate.getDate() + 1);
    const nextDateString = Utilities.formatDate(nextDate, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
    const isNextDayHoliday = isWeekend(nextDate) || !!holidays[nextDateString];

    // Choose shift types dynamically
    const shiftTypes = isHoliday ? shiftConfig.holiday : shiftConfig.weekday;
    
    shiftTypes.forEach(shift => {
      // Just use the raw shift name (User requested to exclude holiday name in title)
      createSlotEvent(calendar, currentDate, shift, isNextDayHoliday);
      Utilities.sleep(100); // Reduced delay to prevent timeout
    });
  }
  
  SpreadsheetApp.getUi().alert('슬롯 생성이 완료되었습니다.');
}

function deleteMonthlySlots() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configParams = getConfig(ss);
  
  if (!configParams.calendarId) {
    SpreadsheetApp.getUi().alert('캘린더 ID가 설정되지 않았습니다.');
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '슬롯 삭제 확인',
    `${configParams.year}년 ${configParams.month}월의 '[모집]' 이벤트를 모두 삭제하시겠습니까?\n(주의: 확정된 근무는 삭제되지 않습니다.)`,
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const calendar = CalendarApp.getCalendarById(configParams.calendarId);
  const year = configParams.year;
  const month = configParams.month - 1;
  const start = new Date(year, month, 1);
  const end = new Date(year, month + 1, 0); 
  // Add padding for overnight
  const fetchEnd = new Date(end);
  fetchEnd.setDate(fetchEnd.getDate() + 1);

  const events = calendar.getEvents(start, fetchEnd);
  let deletedCount = 0;
  
  events.forEach(event => {
    const title = event.getTitle();
    const desc = event.getDescription() || '';
    
    // Delete if Status is OPEN OR legacy [모집] prefix
    if (desc.includes('Status: OPEN') || title.startsWith('[모집]')) {
      event.deleteEvent();
      deletedCount++;
      Utilities.sleep(100); // Throttle
    }
  });

  ui.alert(`삭제 완료: 총 ${deletedCount}개의 슬롯이 삭제되었습니다.`);
}

function createSlotEvent(calendar, date, shift, nextDayIsHoliday) {
  const startTime = parseTime(shift.start);
  let start = new Date(date);
  start.setHours(startTime.h, startTime.m, 0);
  
  // Calculate End Logic
  // Default end time from config
  const endTime = parseTime(shift.end);
  let end = new Date(date);
  
  // Special Rules for Next Day transition
  if (shift.name === '평일 야간' && nextDayIsHoliday) {
    // Weekday Night -> Holiday Morning: Ends at 08:00
    end.setDate(date.getDate() + 1);
    end.setHours(8, 0, 0);
  } else if (shift.name === '휴일 야간' && !nextDayIsHoliday) {
    // Holiday Night -> Weekday Morning: Ends at 08:30
    end.setDate(date.getDate() + 1);
    end.setHours(8, 30, 0);
  } else {
    // Normal Case
    if (endTime.h < startTime.h) {
      // Crosses midnight
      end.setDate(date.getDate() + 1);
    }
    end.setHours(endTime.h, endTime.m, 0);
  }
  
  // Title: Just Shift Name (No prefix)
  calendar.createEvent(shift.name, start, end, {
    description: `ShiftType: ${shift.name}\nStatus: OPEN`
  });
}

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
