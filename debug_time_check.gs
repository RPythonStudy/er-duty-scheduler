// ===== DEBUG FUNCTION - 시간 확인용 =====
// 이 함수를 직접 실행해서 실제 캘린더 이벤트의 시간을 확인할 수 있습니다
function debugCheckEventTimes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig(ss);
  const calendar = CalendarApp.getCalendarById(config.calendarId);
  
  const year = config.year;
  const month = config.month;
  const start = new Date(year, month - 1, 1);
  const end = new Date(year, month, 0);
  
  const events = calendar.getEvents(start, end);
  
  Logger.log('=== 캘린더 이벤트 시간 확인 ===');
  Logger.log('총 이벤트 수: ' + events.length);
  
  // 평일 오전 이벤트만 찾기
  events.forEach(event => {
    const title = event.getTitle();
    if (title.includes('평일 오전') || title.includes('오전')) {
      const startTime = event.getStartTime();
      const endTime = event.getEndTime();
      
      Logger.log('\n--- 이벤트 발견 ---');
      Logger.log('제목: ' + title);
      Logger.log('시작 시간 (원본): ' + startTime);
      Logger.log('종료 시간 (원본): ' + endTime);
      Logger.log('시작 시간 (포맷): ' + Utilities.formatDate(startTime, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss'));
      Logger.log('종료 시간 (포맷): ' + Utilities.formatDate(endTime, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss'));
      Logger.log('시작 시간 (ISO): ' + startTime.toISOString());
      Logger.log('종료 시간 (ISO): ' + endTime.toISOString());
      Logger.log('설명: ' + event.getDescription());
    }
  });
  
  Logger.log('\n=== Config 확인 ===');
  const shiftConfig = getShiftConfig(ss);
  Logger.log('평일 근무 설정:');
  shiftConfig.weekday.forEach(shift => {
    Logger.log('  - ' + shift.name + ': ' + shift.start + ' ~ ' + shift.end);
  });
  
  Logger.log('\n실행 완료! 위의 로그를 확인하세요.');
  Logger.log('Apps Script 에디터 상단 메뉴: 보기 > 로그 또는 실행 로그');
  
  return '로그를 확인하세요 (보기 > 로그)';
}
