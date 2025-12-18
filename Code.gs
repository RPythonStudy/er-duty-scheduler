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
  
  // Return sorted by date + Doctor Email Map
  return {
    year: config.year,
    month: config.month,
    shifts: shifts.sort((a, b) => new Date(a.start) - new Date(b.start)),
    doctorEmails: getDoctorEmails(ss)
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
  if (!title.startsWith('[모집]')) {
    throw new Error('이미 마감되었거나 신청할 수 없는 슬롯입니다.');
  }

  const shiftName = title.replace('[모집] ', '');
  
  // Update Event
  const newTitle = `[확정] ${doctorName} (${shiftName})`;
  event.setTitle(newTitle);
  let desc = `Status: CONFIRMED\nDoctor: ${doctorName}\nUpdated via Web App`;
  
  if (doctorEmail && doctorEmail.trim() !== '') {
    try {
      event.addGuest(doctorEmail);
      desc += `\nGuest: ${doctorEmail}`;
    } catch (e) {
      // Ignore invalid email errors to prevent blocking the booking
      console.error('Failed to add guest: ' + e);
    }
  }

  event.setDescription(desc);
  
  // Log to Sheet
  logShift(ss, event.getStartTime(), `${doctorName} (${doctorEmail || 'No Email'})`, newTitle);
  
  // Send Email Notification
  console.log('[DEBUG] Email sending section reached');
  console.log('[DEBUG] doctorEmail value: ' + doctorEmail);
  
  if (doctorEmail && doctorEmail.trim() !== '') {
    console.log('[DEBUG] Email validation passed, attempting to send...');
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
      
      console.log('[DEBUG] Sending email to: ' + doctorEmail);
      console.log('[DEBUG] Subject: ' + subject);
      
      MailApp.sendEmail({
        to: doctorEmail,
        subject: subject,
        body: body
      });
      
      console.log('[DEBUG] Email sent successfully!');
      
    } catch (e) {
      console.error('[ERROR] Failed to send email: ' + e);
      console.error('[ERROR] Stack trace: ' + e.stack);
      // Email failure should not fail the booking
    }
  } else {
    console.log('[DEBUG] Email NOT sent - doctorEmail is empty or invalid');
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
  if (!title.startsWith('[확정]')) {
    throw new Error('취소할 수 없는 상태입니다.');
  }

  // Parse original shift name from "[확정] DoctorName (ShiftName)"
  // or fallback if regex fails
  let shiftName = '';
  const match = title.match(/\[확정\]\s+(.+)\s+\((.+)\)/);
  if (match) {
    shiftName = match[2];
  } else {
    // Fallback logic if format is different
    shiftName = title.replace('[확정] ', '');
  }
  
  // Revert to Open status
  const newTitle = `[모집] ${shiftName}`;
  event.setTitle(newTitle);
  event.setDescription(`ShiftType: ${shiftName}\nStatus: OPEN\nCancelled by User`);
  
  // Optional: Log cancellation
  logShift(ss, event.getStartTime(), 'CANCELLATION', `Cancelled: ${title}`);
  
  // Send Cancellation Email
  try {
    let doctorEmail = '';
    // Method 1: Try to look up email by Name from Doctors Sheet (Most reliable if configured)
    const match = title.match(/\[확정\]\s+(.+)\s+\(/);
    if (match) {
        const docName = match[1];
        const emails = getDoctorEmails(ss);
        if (emails[docName]) {
            doctorEmail = emails[docName];
        }
    }
    
    // Method 2: If not found, try generic guest list (unreliable if multiple guests)
    if (!doctorEmail) {
        // Fallback: This might be hard if we don't store it explicitly, 
        // but often the user is the only guest added via this app.
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
    }
  } catch(e) {
      console.error('Cancellation email error: ' + e);
  }

  return { success: true, message: '신청이 취소되었습니다.' };
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


function getDoctorEmails(ss) {
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
    Logger.log('[TEST] 이메일 발송 성공! 받은편지함을 확인하세요.');
    return '성공: 이메일이 발송되었습니다. 받은편지함을 확인하세요.';
    
  } catch (e) {
    console.error('[TEST ERROR] Failed: ' + e);
    Logger.log('[TEST ERROR] 실패: ' + e.message);
    return '실패: ' + e.message;
  }
}
