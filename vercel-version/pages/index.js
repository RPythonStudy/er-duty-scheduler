// 메인 페이지: index_minimal.html을 React로 변환

import { useState, useEffect } from 'react';
import styles from '../styles/Calendar.module.css';

export default function Home() {
    const [currentYear, setCurrentYear] = useState(null);
    const [currentMonth, setCurrentMonth] = useState(null);
    const [loading, setLoading] = useState(true);

    // 달력 데이터 로드
    const loadCalendar = async (year, month) => {
        setLoading(true);

        const params = new URLSearchParams();
        if (year) params.append('year', year);
        if (month) params.append('month', month);

        const response = await fetch(`/api/calendar-config?${params}`);
        const data = await response.json();

        setCurrentYear(data.year);
        setCurrentMonth(data.month);
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
            const isWeekend = currentDate.getDay() === 0 || currentDate.getDay() === 6;

            days.push(
                <div key={d} className={`${styles.dayCell} ${isWeekend ? styles.weekend : ''}`}>
                    <span className={styles.dateLabel}>{d}</span>
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

            <div className={styles.version}>
                System Version: v4.0 (Vercel - Maximum Speed)
            </div>
        </div>
    );
}
