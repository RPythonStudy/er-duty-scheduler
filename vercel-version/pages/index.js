// 메인 페이지: 응급실 근무 신청 시스템

import { useState, useEffect } from 'react';
import styles from '../styles/Calendar.module.css';

export default function Home() {
    const [currentYear, setCurrentYear] = useState(null);
    const [currentMonth, setCurrentMonth] = useState(null);
    const [shifts, setShifts] = useState([]);
    const [holidays, setHolidays] = useState({});
    const [doctors, setDoctors] = useState([]);
    const [loading, setLoading] = useState(true);
    const [selectedShift, setSelectedShift] = useState(null);
    const [doctorName, setDoctorName] = useState('');
    const [doctorEmail, setDoctorEmail] = useState('');

    // 달력 데이터 로드
    const loadCalendar = async (year, month) => {
        setLoading(true);

        // 설정 로드
        const configParams = new URLSearchParams();
        if (year) configParams.append('year', year);
        if (month) configParams.append('month', month);

        const configResponse = await fetch(`/api/calendar-config?${configParams}`);
        const configData = await configResponse.json();

        setCurrentYear(configData.year);
        setCurrentMonth(configData.month);
        setHolidays(configData.holidays || {});
        setDoctors(configData.doctors || []);

        // 근무 데이터 로드
        const shiftsParams = new URLSearchParams();
        shiftsParams.append('year', configData.year);
        shiftsParams.append('month', configData.month);

        const shiftsResponse = await fetch(`/api/get-shifts?${shiftsParams}`);
        const shiftsData = await shiftsResponse.json();

        setShifts(shiftsData.shifts || []);
        setLoading(false);
    };

    // 초기 로드
    useEffect(() => {
        loadCalendar();
    }, []);

    // 월 변경
    const changeMonth = (offset) => {
        let newYear = currentYear;
        let newMonth = currentMonth + offset;

        if (newMonth > 12) {
            newYear++;
            newMonth = 1;
        } else if (newMonth < 1) {
            newYear--;
            newMonth = 12;
        }

        loadCalendar(newYear, newMonth);
    };

    // 근무 예약
    const bookShift = async () => {
        if (!doctorName.trim()) {
            alert('이름을 입력해주세요.');
            return;
        }

        try {
            const response = await fetch('/api/book-shift', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    shiftId: selectedShift.id,
                    doctorName: doctorName,
                    doctorEmail: doctorEmail
                })
            });

            const result = await response.json();

            if (response.ok) {
                alert(result.message);
                setSelectedShift(null);
                setDoctorName('');
                setDoctorEmail('');
                loadCalendar(currentYear, currentMonth);
            } else {
                alert(result.error || '예약에 실패했습니다.');
            }
        } catch (error) {
            alert('예약 중 오류가 발생했습니다.');
        }
    };

    // 근무 취소
    const cancelShift = async (shiftId) => {
        if (!confirm('정말 취소하시겠습니까?')) return;

        try {
            const response = await fetch('/api/cancel-shift', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ shiftId })
            });

            const result = await response.json();

            if (response.ok) {
                alert(result.message);
                loadCalendar(currentYear, currentMonth);
            } else {
                alert(result.error || '취소에 실패했습니다.');
            }
        } catch (error) {
            alert('취소 중 오류가 발생했습니다.');
        }
    };

    // 날짜별 근무 그룹화
    const getShiftsByDate = (dateStr) => {
        return shifts.filter(s => s.date === dateStr);
    };

    // 달력 렌더링
    const renderCalendar = () => {
        if (!currentYear || !currentMonth) return null;

        const firstDay = new Date(currentYear, currentMonth - 1, 1);
        const lastDay = new Date(currentYear, currentMonth, 0);
        const startDayOfWeek = firstDay.getDay();
        const daysInMonth = lastDay.getDate();

        const days = [];

        // 빈 칸
        for (let i = 0; i < startDayOfWeek; i++) {
            days.push(<div key={`empty-${i}`} className={styles.emptyCell}></div>);
        }

        // 날짜
        for (let d = 1; d <= daysInMonth; d++) {
            const currentDate = new Date(currentYear, currentMonth - 1, d);
            const dateStr = `${currentYear}-${String(currentMonth).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
            const isWeekend = currentDate.getDay() === 0 || currentDate.getDay() === 6;
            const isHoliday = holidays[dateStr] !== undefined;
            const dayShifts = getShiftsByDate(dateStr);

            days.push(
                <div key={d} className={`${styles.dayCell} ${isWeekend || isHoliday ? styles.holiday : ''}`}>
                    <div className={styles.dateHeader}>
                        <span className={styles.dateLabel}>{d}</span>
                        {isHoliday && <span className={styles.holidayLabel}>{holidays[dateStr]}</span>}
                    </div>
                    <div className={styles.shiftsContainer}>
                        {dayShifts.map(shift => (
                            <div
                                key={shift.id}
                                className={`${styles.shiftSlot} ${shift.status === 'CONFIRMED' ? styles.confirmed : styles.open}`}
                                onClick={() => {
                                    if (shift.status === 'OPEN') {
                                        setSelectedShift(shift);
                                    }
                                }}
                            >
                                <div className={styles.shiftName}>{shift.shiftName}</div>
                                {shift.status === 'CONFIRMED' ? (
                                    <div className={styles.shiftInfo}>
                                        <div className={styles.doctorName}>{shift.doctorName}</div>
                                        <button
                                            className={styles.cancelBtn}
                                            onClick={(e) => {
                                                e.stopPropagation();
                                                cancelShift(shift.id);
                                            }}
                                        >
                                            취소
                                        </button>
                                    </div>
                                ) : null}
                            </div>
                        ))}
                    </div>
                </div>
            );
        }

        return days;
    };

    if (loading) {
        return <div className={styles.container}>로딩 중...</div>;
    }

    return (
        <div className={styles.container}>
            <div className={styles.headerNav}>
                <button className={styles.navBtn} onClick={() => changeMonth(-1)}>&lt;</button>
                <h1>{currentYear}년 {currentMonth}월 응급실 근무표</h1>
                <button className={styles.navBtn} onClick={() => changeMonth(1)}>&gt;</button>
            </div>

            <div className={styles.calendarGrid}>
                <div className={`${styles.dayHeader} ${styles.sunday}`}>일</div>
                <div className={styles.dayHeader}>월</div>
                <div className={styles.dayHeader}>화</div>
                <div className={styles.dayHeader}>수</div>
                <div className={styles.dayHeader}>목</div>
                <div className={styles.dayHeader}>금</div>
                <div className={`${styles.dayHeader} ${styles.saturday}`}>토</div>
                {renderCalendar()}
            </div>

            {selectedShift && (
                <div className={styles.modal} onClick={() => setSelectedShift(null)}>
                    <div className={styles.modalContent} onClick={(e) => e.stopPropagation()}>
                        <h2>근무 신청</h2>
                        <p><strong>날짜:</strong> {selectedShift.date}</p>
                        <p><strong>근무:</strong> {selectedShift.shiftName}</p>
                        <input
                            type="text"
                            placeholder="이름"
                            value={doctorName}
                            onChange={(e) => setDoctorName(e.target.value)}
                            className={styles.input}
                        />
                        <input
                            type="email"
                            placeholder="이메일 (선택)"
                            value={doctorEmail}
                            onChange={(e) => setDoctorEmail(e.target.value)}
                            className={styles.input}
                            list="doctorEmails"
                        />
                        <datalist id="doctorEmails">
                            {doctors.map((doc, idx) => (
                                <option key={idx} value={doc.email}>{doc.name}</option>
                            ))}
                        </datalist>
                        <div className={styles.modalButtons}>
                            <button onClick={bookShift} className={styles.confirmBtn}>확인</button>
                            <button onClick={() => setSelectedShift(null)} className={styles.cancelBtnModal}>취소</button>
                        </div>
                    </div>
                </div>
            )}

            <div className={styles.version}>
                System Version: v5.0 (Vercel - Full Local Config)
            </div>
        </div>
    );
}
