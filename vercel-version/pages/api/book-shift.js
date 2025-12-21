// API Route: /api/book-shift
// 근무 예약

import fs from 'fs';
import path from 'path';

export default function handler(req, res) {
    if (req.method !== 'POST') {
        return res.status(405).json({ error: 'Method not allowed' });
    }

    const { shiftId, doctorName, doctorEmail } = req.body;

    if (!shiftId || !doctorName) {
        return res.status(400).json({ error: '근무 ID와 이름은 필수입니다.' });
    }

    // shifts.json 읽기
    const shiftsPath = path.join(process.cwd(), 'shifts.json');
    const shiftsData = JSON.parse(fs.readFileSync(shiftsPath, 'utf8'));

    // 이미 예약된 근무인지 확인
    if (shiftsData[shiftId] && shiftsData[shiftId].status === 'CONFIRMED') {
        return res.status(400).json({ error: '이미 예약된 근무입니다.' });
    }

    // 근무 예약
    shiftsData[shiftId] = {
        status: 'CONFIRMED',
        doctorName: doctorName,
        doctorEmail: doctorEmail || '',
        bookedAt: new Date().toISOString()
    };

    // shifts.json 저장
    fs.writeFileSync(shiftsPath, JSON.stringify(shiftsData, null, 2));

    res.status(200).json({
        success: true,
        message: '근무 예약이 완료되었습니다.'
    });
}
