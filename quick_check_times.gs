// Quick check of actual calendar event times
function quickCheckCalendarTimes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig(ss);
  const calendar = CalendarApp.getCalendarById(config.calendarId);
  
  const start = new Date(config.year, config.month - 1, 1);
  const end = new Date(config.year, config.month, 0);
  const events = calendar.getEvents(start, end);
  
  Logger.log('=== CALENDAR EVENT TIMES ===');
  Logger.log('Checking first 3 events...\n');
  
  for (let i = 0; i < Math.min(3, events.length); i++) {
    const event = events[i];
    Logger.log('Event ' + (i+1) + ': ' + event.getTitle());
    Logger.log('  Start: ' + event.getStartTime());
    Logger.log('  End: ' + event.getEndTime());
    Logger.log('  Start Hour: ' + event.getStartTime().getHours());
    Logger.log('  Start Minute: ' + event.getStartTime().getMinutes());
    Logger.log('');
  }
  
  Logger.log('=== CONFIG SHIFT TIMES ===');
  const shiftConfig = getShiftConfig(ss);
  Logger.log('Expected times from Config:');
  shiftConfig.weekday.forEach(shift => {
    Logger.log('  ' + shift.name + ': ' + shift.start + ' - ' + shift.end);
  });
  
  Logger.log('\n=== CONCLUSION ===');
  Logger.log('If calendar events show 00:00:00, you need to:');
  Logger.log('1. Check Config sheet (row 6+) has correct times');
  Logger.log('2. Delete all slots: ER 근무 관리 → 근무표 슬롯 삭제');
  Logger.log('3. Recreate slots: ER 근무 관리 → 근무표 슬롯 생성');
  
  return 'Check logs (View → Logs)';
}
