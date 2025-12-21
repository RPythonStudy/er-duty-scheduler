// API Route: /api/get-shifts
// 특정 월의 근무 데이터를 반환

import fs from 'fs';
import path from 'path';

export default function handler(req, res) {
    const { year: reqYear, month: reqMonth } = req.query;

    // config.json 읽기
    const configPath = path.join(process.cwd(), 'config.json');
    const configData = JSON.parse(fs.readFileSync(configPath, 'utf8'));

    // shifts.json 읽기
    const shiftsPath = path.join(process.cwd(), 'shifts.json');
    const shiftsData = JSON.parse(fs.readFileSync(shiftsPath, 'utf8'));

    const year = reqYear ? parseInt(reqYear) : configData.year;
    const month = reqMonth ? parseInt(reqMonth) : configData.month;

    // 해당 월의 모든 날짜에 대해 근무 슬롯 생성
    const shifts = [];
    const daysInMonth = new Date(year, month, 0).getDate();

    for (let day = 1; day <= daysInMonth; day++) {
        const date = new Date(year, month - 1, day);
        const dateStr = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        const dayOfWeek = date.getDay();

        // 주말 또는 공휴일 확인
        const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
        const isHoliday = configData.holidays[dateStr] !== undefined;
        const isHolidayOrWeekend = isWeekend || isHoliday;

        // 근무 타입 선택 (평일 vs 휴일)
        const shiftTypes = isHolidayOrWeekend ? configData.shifts.holiday : configData.shifts.weekday;

        // 각 근무 타입에 대해 슬롯 생성
        shiftTypes.forEach(shiftConfig => {
            const shiftId = `${dateStr}_${shiftConfig.name}`;
            const bookedShift = shiftsData[shiftId];

            // 시작/종료 시간 계산
            const [startH, startM] = shiftConfig.start.split(':');
            const [endH, endM] = shiftConfig.end.split(':');

            const startDate = new Date(year, month - 1, day, parseInt(startH), parseInt(startM));
            let endDate = new Date(year, month - 1, day, parseInt(endH), parseInt(endM));

            // 종료 시간이 시작 시간보다 이르면 다음날
            if (parseInt(endH) < parseInt(startH)) {
                endDate.setDate(endDate.getDate() + 1);
            }

            shifts.push({
                id: shiftId,
                date: dateStr,
                shiftName: shiftConfig.name,
                start: startDate.toISOString(),
                end: endDate.toISOString(),
                status: bookedShift ? bookedShift.status : 'OPEN',
                doctorName: bookedShift ? bookedShift.doctorName : '',
                doctorEmail: bookedShift ? bookedShift.doctorEmail : '',
                title: bookedShift && bookedShift.status === 'CONFIRMED'
                    ? `${bookedShift.doctorName} (${shiftConfig.name})`
                    : shiftConfig.name
            });
        });
    }

    res.status(200).json({
        year,
        month,
        shifts
    });
}
