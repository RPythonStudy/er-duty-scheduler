// API Route: /api/cancel-shift
// 근무 취소

import fs from 'fs';
import path from 'path';

export default function handler(req, res) {
    if (req.method !== 'POST') {
        return res.status(405).json({ error: 'Method not allowed' });
    }

    const { shiftId } = req.body;

    if (!shiftId) {
        return res.status(400).json({ error: '근무 ID는 필수입니다.' });
    }

    // shifts.json 읽기
    const shiftsPath = path.join(process.cwd(), 'shifts.json');
    const shiftsData = JSON.parse(fs.readFileSync(shiftsPath, 'utf8'));

    // 예약된 근무가 아니면 에러
    if (!shiftsData[shiftId] || shiftsData[shiftId].status !== 'CONFIRMED') {
        return res.status(400).json({ error: '취소할 수 있는 근무가 아닙니다.' });
    }

    // 근무 취소 (삭제)
    delete shiftsData[shiftId];

    // shifts.json 저장
    fs.writeFileSync(shiftsPath, JSON.stringify(shiftsData, null, 2));

    res.status(200).json({
        success: true,
        message: '근무 취소가 완료되었습니다.'
    });
}
