// API Route: /api/calendar-config
// Google Apps Script의 getCalendarConfig() 대체

import fs from 'fs';
import path from 'path';

export default function handler(req, res) {
  // Query parameters에서 year, month 가져오기
  const { year: reqYear, month: reqMonth } = req.query;

  // config.json 파일 읽기
  const configPath = path.join(process.cwd(), 'config.json');
  const configData = JSON.parse(fs.readFileSync(configPath, 'utf8'));

  // 요청된 year/month가 있으면 사용, 없으면 config 사용
  const year = reqYear ? parseInt(reqYear) : configData.year;
  const month = reqMonth ? parseInt(reqMonth) : configData.month;

  // 전체 설정 반환 (shifts, holidays, doctors 포함)
  res.status(200).json({
    year: year,
    month: month,
    shifts: configData.shifts,
    holidays: configData.holidays,
    doctors: configData.doctors
  });
}
