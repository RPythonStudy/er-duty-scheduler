// ===== COMPREHENSIVE DEBUG FUNCTION =====
// Run this function to diagnose time display issues
function debugTimeIssue() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig(ss);
  
  Logger.log('=== CONFIG CHECK ===');
  Logger.log('Year: ' + config.year);
  Logger.log('Month: ' + config.month);
  Logger.log('Calendar ID: ' + config.calendarId);
  
  Logger.log('\n=== SHIFT CONFIG CHECK ===');
  const shiftConfig = getShiftConfig(ss);
  Logger.log('Weekday Shifts:');
  shiftConfig.weekday.forEach(shift => {
    Logger.log('  - ' + shift.name + ': ' + shift.start + ' to ' + shift.end);
  });
  Logger.log('Holiday Shifts:');
  shiftConfig.holiday.forEach(shift => {
    Logger.log('  - ' + shift.name + ': ' + shift.start + ' to ' + shift.end);
  });
  
  Logger.log('\n=== CALENDAR EVENTS CHECK ===');
  const calendar = CalendarApp.getCalendarById(config.calendarId);
  const start = new Date(config.year, config.month - 1, 1);
  const end = new Date(config.year, config.month, 0);
  const events = calendar.getEvents(start, end);
  
  Logger.log('Total events: ' + events.length);
  
  // Check first 5 events
  for (let i = 0; i < Math.min(5, events.length); i++) {
    const event = events[i];
    const title = event.getTitle();
    const startTime = event.getStartTime();
    const endTime = event.getEndTime();
    
    Logger.log('\nEvent ' + (i+1) + ':');
    Logger.log('  Title: ' + title);
    Logger.log('  Start (raw): ' + startTime);
    Logger.log('  End (raw): ' + endTime);
    Logger.log('  Start (formatted): ' + Utilities.formatDate(startTime, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss'));
    Logger.log('  End (formatted): ' + Utilities.formatDate(endTime, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss'));
    Logger.log('  Description: ' + event.getDescription());
  }
  
  Logger.log('\n=== GETSHIFTDATA OUTPUT CHECK ===');
  const shiftData = getShiftData(config.year, config.month);
  Logger.log('First shift from getShiftData:');
  if (shiftData.shifts.length > 0) {
    const firstShift = shiftData.shifts[0];
    Logger.log('  Shift Name: ' + firstShift.shiftName);
    Logger.log('  Start: ' + firstShift.start);
    Logger.log('  End: ' + firstShift.end);
    Logger.log('  Status: ' + firstShift.status);
  }
  
  Logger.log('\n=== DIAGNOSIS COMPLETE ===');
  Logger.log('Check the logs above to identify where the time is going wrong.');
  Logger.log('Expected: Times should match Config sheet settings (e.g., 08:30-13:00)');
  Logger.log('If calendar events show wrong times, you need to delete and recreate slots.');
  
  return 'Check Logs (View > Logs)';
}
