function testHolidayFetch() {
  // South Korea Holidays Public Calendar ID
  const calendarId = 'ko.south_korea#holiday@group.v.calendar.google.com';
  const calendar = CalendarApp.getCalendarById(calendarId);
  
  if (!calendar) {
    console.log('Holiday Calendar NOT found');
    return;
  }
  
  const start = new Date('2025-01-01');
  const end = new Date('2025-12-31');
  const events = calendar.getEvents(start, end);
  
  console.log('Found ' + events.length + ' holidays for 2025');
  events.forEach(e => console.log(e.getTitle() + ': ' + e.getStartTime().toISOString().split('T')[0]));
}
