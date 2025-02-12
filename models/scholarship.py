from dataclasses import dataclass
from typing import Dict
import pandas as pd
from datetime import datetime, time
from config.settings import MORNING_CONFIG, AFTERNOON_CONFIG, BONUS_CONFIG


@dataclass
class SessionStats:
    """
       Data class for storing session attendance and scholarship statistics.

       Attributes:
           base (float): Base scholarship amount
           bonus (float): Additional bonus amount earned
           total_hours (float): Total hours attended
           attended_days (int): Number of days with attendance
           late_days (int): Number of days with late arrival
           absent_days (int): Number of days with no attendance
           perfect_days (int): Number of days with perfect attendance
           partial_bonus_days (int): Number of days eligible for partial bonus
           missed_hours (float): Total hours missed from expected attendance
           very_late_days (int): Number of days with significantly late arrival
           early_attendance_bonus (float): Bonus amount for consistent early attendance
       """
    base: float
    bonus: float
    total_hours: float
    attended_days: int
    late_days: int
    absent_days: int
    perfect_days: int
    partial_bonus_days: int
    missed_hours: float
    very_late_days: int
    early_attendance_bonus: float


class KollelScholarship:
    """
       Handles the calculation of Kollel student scholarships based on attendance and performance.

       This class processes attendance data and calculates scholarship amounts based on various
       criteria including punctuality, consistency, and study session participation.
    """

    def _parse_time(self, time_str: str) -> time:
        """
            Convert string time representation to time object.

            Args:
                time_str (str): Time string to parse

            Returns:
                time: Parsed time object
        """
        if isinstance(time_str, str):
            dt = pd.to_datetime(time_str).time()
            return dt
        return time_str

    def _calculate_session_hours(self, entry_time: time, exit_time: time, session_type: str) -> float:

        """
            Calculate actual hours attended within defined time windows.

                Args:
                    entry_time (time): Session entry time
                    exit_time (time): Session exit time
                    session_type (str): Type of session ('בוקר' for morning or 'צהריים' for afternoon)

                Returns:
                    float: Calculated hours attended, rounded to 2 decimal places
        """
        def to_minutes(t: time) -> int:
            return t.hour * 60 + t.minute

        if session_type == 'בוקר':
            valid_start = MORNING_CONFIG['START']
            valid_end = MORNING_CONFIG['END']
        else:
            valid_start = AFTERNOON_CONFIG['START']
            valid_end = AFTERNOON_CONFIG['END']

        entry_minutes = to_minutes(entry_time)
        exit_minutes = to_minutes(exit_time)
        start_minutes = to_minutes(valid_start)
        end_minutes = to_minutes(valid_end)

        actual_start = max(entry_minutes, start_minutes)
        actual_end = min(exit_minutes, end_minutes)

        if actual_end > actual_start:
            hours = (actual_end - actual_start) / 60
            return round(hours, 2)
        return 0

    def _get_empty_stats(self, working_days: int, session_type: str) -> SessionStats:
        """
        Create empty session statistics for periods with no attendance.

            Args:
                working_days (int): Number of expected working days
                session_type (str): Type of session ('בוקר' for morning or 'צהריים' for afternoon)

            Returns:
                SessionStats: Statistics object with default values
        """
        hours = 3.5 if session_type == 'בוקר' else 3.0
        return SessionStats(
            base=0,
            bonus=0,
            total_hours=0,
            attended_days=0,
            late_days=0,
            absent_days=working_days,
            perfect_days=0,
            partial_bonus_days=0,
            missed_hours=working_days * hours,
            very_late_days=0,
            early_attendance_bonus=0
        )

    def calculate_session_stats(self, session_data: pd.DataFrame, session_type: str, working_days: int,
                                has_afternoon: bool) -> SessionStats:
        """
            Calculate comprehensive statistics for a study session.

            Args:
                session_data (pd.DataFrame): Attendance data for the session
                session_type (str): Type of session ('בוקר' for morning or 'צהריים' for afternoon)
                working_days (int): Number of expected working days
                has_afternoon (bool): Whether the student attends afternoon sessions

            Returns:
                SessionStats: Calculated statistics for the session
        """
        if session_data.empty:
            return self._get_empty_stats(working_days, session_type)

        config = MORNING_CONFIG if session_type == 'בוקר' else AFTERNOON_CONFIG

        session_data = session_data.copy()
        session_data['שעת_כניסה'] = session_data['שעת_כניסה'].apply(self._parse_time)
        session_data['שעת_יציאה'] = session_data['שעת_יציאה'].apply(self._parse_time)

        daily_stats = session_data.groupby('תאריך').agg({
            'שעת_כניסה': 'min',
            'שעת_יציאה': 'max',
            'רצופות': lambda x: 'כן' in x.values
        })

        daily_stats['סך שעות'] = daily_stats.apply(
            lambda row: self._calculate_session_hours(
                row['שעת_כניסה'],
                row['שעת_יציאה'],
                session_type
            ),
            axis=1
        )

        attended_days = min(len(daily_stats), working_days)

        def compare_times(t1, t2):
            return (t1.hour * 60 + t1.minute) > (t2.hour * 60 + t2.minute)

        late_threshold = self._parse_time(config['LATE_THRESHOLD'])
        very_late_threshold = self._parse_time(config['VERY_LATE_THRESHOLD'])
        perfect_start = self._parse_time(config['PERFECT_START'])
        session_end = config['END']

        late_mask = daily_stats['שעת_כניסה'].apply(lambda x: compare_times(x, late_threshold)) & \
                    daily_stats['שעת_כניסה'].apply(lambda x: not compare_times(x, very_late_threshold))
        very_late_mask = daily_stats['שעת_כניסה'].apply(lambda x: compare_times(x, very_late_threshold))

        perfect_mask = (
                daily_stats['שעת_כניסה'].apply(lambda x: not compare_times(x, perfect_start)) &
                daily_stats['שעת_יציאה'].apply(lambda x: compare_times(x, config['END'])) &
                daily_stats['רצופות']
        )

        late_days = min(late_mask.sum(), working_days)
        very_late_days = min(very_late_mask.sum(), working_days)
        perfect_days = min(perfect_mask.sum(), working_days)

        partial_mask = daily_stats['רצופות'] & ~perfect_mask
        partial_bonus_days = min(partial_mask.sum(), working_days)

        expected_hours = working_days * (3.5 if session_type == 'בוקר' else 3.0)
        total_hours = min(daily_stats['סך שעות'].sum(), expected_hours)
        missed_hours = max(0, expected_hours - total_hours)

        if late_days > 2:
            base = attended_days * config['LATE_DAILY_RATE']
        else:
            base = config['BASE']

        daily_bonus = (perfect_days * BONUS_CONFIG['DAILY_PERFECT'] +
                       partial_bonus_days * BONUS_CONFIG['DAILY_PARTIAL'])

        early_attendance_bonus = 0
        if session_type == 'בוקר':
            nine_am = time(9, 0)
            late_after_nine = sum(compare_times(t, nine_am) for t in daily_stats['שעת_כניסה'])

            if late_after_nine <= 1:
                early_attendance_bonus = 200 if has_afternoon else 100
            else:
                weekly_bonus = self._calculate_weekly_early_attendance(session_data)
                early_attendance_bonus = min(weekly_bonus, 200 if has_afternoon else 100)

        return SessionStats(
            base=max(0, base),
            bonus=daily_bonus,
            total_hours=round(total_hours, 2),
            attended_days=attended_days,
            late_days=late_days,
            absent_days=working_days - attended_days,
            perfect_days=perfect_days,
            partial_bonus_days=partial_bonus_days,
            missed_hours=missed_hours,
            very_late_days=very_late_days,
            early_attendance_bonus=early_attendance_bonus
        )

    def _calculate_weekly_early_attendance(self, attendance_data: pd.DataFrame) -> float:
        """
            Calculate bonus for consistent early attendance on a weekly basis.

            Args:
                attendance_data (pd.DataFrame): Attendance data for analysis

            Returns:
                float: Calculated early attendance bonus
        """
        if attendance_data.empty:
            return 0

        attendance_data = attendance_data.copy()
        attendance_data['תאריך'] = pd.to_datetime(attendance_data['תאריך'])
        attendance_data['week'] = attendance_data['תאריך'].dt.isocalendar().week

        nine_am = time(9, 0)

        def is_late(t):
            if isinstance(t, str):
                t = self._parse_time(t)
            return (t.hour * 60 + t.minute) > (nine_am.hour * 60 + nine_am.minute)

        weekly_lates = (
            attendance_data
            .groupby('week')
            .agg({'שעת_כניסה': lambda x: sum(is_late(t) for t in x)})
        )

        good_weeks = sum(weekly_lates['שעת_כניסה'] <= 1)
        return good_weeks * 35

    def calculate_student_scholarship(self, student_data: pd.DataFrame, working_days: int) -> Dict:
        """
        Calculate total scholarship amount for a student based on attendance and performance.

        This method processes both morning and afternoon sessions, calculating base amounts,
        bonuses, and additional benefits based on attendance patterns and academic performance.

        Args:
            student_data (pd.DataFrame): Complete student attendance data
            working_days (int): Number of expected working days

        Returns:
            Dict: Comprehensive scholarship calculation results including:
                - Student identification details
                - Base scholarship amount
                - Various bonus amounts
                - Session-specific statistics
                - Total scholarship amount
        """
        morning_data = student_data[student_data['סדר'] == 'בוקר']
        afternoon_data = student_data[student_data['סדר'] == 'צהריים']

        has_afternoon = not afternoon_data.empty

        morning_stats = self.calculate_session_stats(morning_data, 'בוקר', working_days, has_afternoon)
        afternoon_stats = self.calculate_session_stats(afternoon_data, 'צהריים', working_days, has_afternoon)

        total_base = morning_stats.base + afternoon_stats.base
        total_bonus = morning_stats.bonus + afternoon_stats.bonus

        if not morning_data.empty and not afternoon_data.empty:
            if (morning_stats.missed_hours <= BONUS_CONFIG['MORNING_TIER1_MAX_MISSING'] and
                    afternoon_stats.missed_hours <= BONUS_CONFIG['AFTERNOON_TIER1_MAX_MISSING']):
                total_bonus += BONUS_CONFIG['MONTHLY_TIER1']

            if (morning_stats.missed_hours <= BONUS_CONFIG['MORNING_TIER2_MAX_MISSING'] and
                    afternoon_stats.missed_hours <= BONUS_CONFIG['AFTERNOON_TIER2_MAX_MISSING']):
                total_bonus += BONUS_CONFIG['MONTHLY_TIER2']

            if (morning_stats.absent_days == 0 and afternoon_stats.absent_days == 0 and
                    morning_stats.late_days == 0 and afternoon_stats.late_days == 0):
                total_bonus += BONUS_CONFIG['PERFECT_ATTENDANCE']

        return {
            'זהות': student_data['זהות'].iloc[0],
            'שם מלא': f"{student_data['שם פרטי'].iloc[0]} {student_data['שם משפחה'].iloc[0]}",
            'מלגת_בסיס': total_base,
            'בונוסים': total_bonus,
            'בונוס_נוכחות_מוקדמת': morning_stats.early_attendance_bonus,
            'סך_הכל': total_base + total_bonus + morning_stats.early_attendance_bonus,
            **{f'בוקר_{k}': v for k, v in vars(morning_stats).items()},
            **{f'צהריים_{k}': v for k, v in vars(afternoon_stats).items()}
        }
