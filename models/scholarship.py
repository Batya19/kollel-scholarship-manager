from dataclasses import dataclass
from datetime import datetime, time
from typing import Dict

import pandas as pd
from config.settings import (
    BONUS_CONFIG,
    MAX_ALLOWED_LATES_OR_ABSENCES, NINE_AM_TIME
)
from utils.time_utils import (
    compare_times, compare_times_gte,
    calculate_hours_between, get_session_config, get_expected_hours
)


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

    def _build_disqualification_warning(self, stats, session_name: str) -> str:
        """
        Build warning message for base scholarship disqualification.

        Args:
            stats: Session statistics object
            session_name (str): Name of session ('בוקר' or 'צהריים')

        Returns:
            str: Warning message or empty string
        """
        if stats.late_days <= MAX_ALLOWED_LATES_OR_ABSENCES and \
           stats.absent_days <= MAX_ALLOWED_LATES_OR_ABSENCES:
            return ""

        reasons = []
        if stats.late_days > MAX_ALLOWED_LATES_OR_ABSENCES:
            reasons.append(f"{stats.late_days} איחורים ב{session_name}")
        if stats.absent_days > MAX_ALLOWED_LATES_OR_ABSENCES:
            reasons.append(f"{stats.absent_days} היעדרויות ב{session_name}")

        return f"❌ מלגת בסיס {session_name} בוטלה ({', '.join(reasons)})"

    def _calculate_attendance_stats(
            self,
            daily_stats: pd.DataFrame,
            config: dict,
            working_days: int) -> tuple:
        """
        Calculate attendance statistics (late days, very late days, perfect days, etc.)

        Args:
            daily_stats (pd.DataFrame): Daily aggregated statistics
            config (dict): Session configuration (morning or afternoon)
            working_days (int): Number of expected working days

        Returns:
            tuple: (late_days, very_late_days, perfect_days, partial_bonus_days)
        """
        late_threshold = self._parse_time(config['LATE_THRESHOLD'])
        very_late_threshold = self._parse_time(config['VERY_LATE_THRESHOLD'])
        perfect_start = self._parse_time(config['PERFECT_START'])

        late_mask = daily_stats['שעת כניסה'].apply(
            lambda x: compare_times(
                x, late_threshold)) & daily_stats['שעת כניסה'].apply(
            lambda x: not compare_times(
                x, very_late_threshold))
        very_late_mask = daily_stats['שעת כניסה'].apply(
            lambda x: compare_times(x, very_late_threshold))

        perfect_mask = (
            daily_stats['שעת כניסה'].apply(
                lambda x: not compare_times(
                    x, perfect_start)) & daily_stats['שעת יציאה'].apply(
                lambda x: compare_times_gte(
                    x, config['END'])) & daily_stats['רצופות'])

        late_days = min(late_mask.sum(), working_days)
        very_late_days = min(very_late_mask.sum(), working_days)
        perfect_days = min(perfect_mask.sum(), working_days)

        partial_mask = daily_stats['רצופות'] & ~perfect_mask
        partial_bonus_days = min(partial_mask.sum(), working_days)

        return late_days, very_late_days, perfect_days, partial_bonus_days

    def _calculate_base_scholarship(
            self,
            late_days: int,
            absent_days: int,
            attended_days: int,
            very_late_days: int,
            config: dict) -> float:
        """
        Calculate base scholarship amount based on attendance.

        Args:
            late_days (int): Number of late days
            absent_days (int): Number of absent days
            attended_days (int): Number of attended days
            very_late_days (int): Number of very late days
            config (dict): Session configuration

        Returns:
            float: Base scholarship amount
        """
        if late_days > MAX_ALLOWED_LATES_OR_ABSENCES or absent_days > MAX_ALLOWED_LATES_OR_ABSENCES:
            regular_days = attended_days - very_late_days
            base = (regular_days * config['LATE_DAILY_RATE']) + \
                (very_late_days * BONUS_CONFIG['DAILY_PARTIAL'])
        else:
            base = config['BASE']

        return base

    def _calculate_early_attendance_bonus(
            self,
            session_data: pd.DataFrame,
            daily_stats: pd.DataFrame,
            absent_days: int,
            has_afternoon: bool) -> float:
        """
        Calculate early attendance bonus for morning sessions.

        Args:
            session_data (pd.DataFrame): Session attendance data
            daily_stats (pd.DataFrame): Daily aggregated statistics
            absent_days (int): Number of absent days
            has_afternoon (bool): Whether student attends afternoon sessions

        Returns:
            float: Early attendance bonus amount
        """
        late_after_nine = sum(compare_times(t, NINE_AM_TIME)
                              for t in daily_stats['שעת כניסה'])

        if absent_days <= MAX_ALLOWED_LATES_OR_ABSENCES:
            if late_after_nine <= 1:
                return 200 if has_afternoon else 100

            weekly_bonus = self._calculate_weekly_early_attendance(
                session_data)
            return min(weekly_bonus, 200 if has_afternoon else 100)

        return 0

    def _parse_time(self, time_str):
        """
        Convert string time representation to time object.
        Handles various input types including NaT values.
        """
        try:
            if pd.isna(time_str) or time_str is pd.NaT:
                return time(0, 0)

            if isinstance(time_str, str):
                try:
                    dt = pd.to_datetime(time_str).time()
                    return dt
                except (ValueError, TypeError, AttributeError) as e:
                    print(
                        f"Failed to parse time string: {time_str}, error: {e}")
                    return time(0, 0)
            elif isinstance(time_str, float):
                if pd.isna(time_str):
                    return time(0, 0)
                total_seconds = int(time_str * 86400)
                hours = total_seconds // 3600
                minutes = (total_seconds % 3600) // 60
                seconds = total_seconds % 60
                return time(hours % 24, minutes, seconds)
            elif isinstance(time_str, datetime):
                return time_str.time()
            elif hasattr(time_str, 'hour'):
                return time_str
            else:
                print(
                    f"Unknown time format: {time_str}, type: {type(time_str)}")
                return time(0, 0)
        except Exception as e:
            print(
                f"Error in _parse_time: {time_str}, type: {type(time_str)}, error: {e}")
            return time(0, 0)

    def _calculate_session_hours(
            self,
            entry_time: time,
            exit_time: time,
            session_type: str) -> float:
        """
            Calculate actual hours attended within defined time windows.

                Args:
                    entry_time (time): Session entry time
                    exit_time (time): Session exit time
                    session_type (str): Type of session ('בוקר' for morning or 'צהריים' for afternoon)

                Returns:
                    float: Calculated hours attended, rounded to 2 decimal places
        """
        config = get_session_config(session_type)
        return calculate_hours_between(
            entry_time, exit_time, config['START'], config['END'])

    def _get_empty_stats(
            self,
            working_days: int,
            session_type: str) -> SessionStats:
        """
        Create empty session statistics for periods with no attendance.

            Args:
                working_days (int): Number of expected working days
                session_type (str): Type of session ('בוקר' for morning or 'צהריים' for afternoon)

            Returns:
                SessionStats: Statistics object with default values
        """
        hours = get_expected_hours(session_type)
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

    def calculate_session_stats(
            self,
            session_data: pd.DataFrame,
            session_type: str,
            working_days: int,
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

        config = get_session_config(session_type)

        session_data = session_data.copy()
        session_data['שעת כניסה'] = session_data['שעת כניסה'].apply(
            self._parse_time)
        session_data['שעת יציאה'] = session_data['שעת יציאה'].apply(
            self._parse_time)

        # Filter days without valid exit time (00:00 = didn't sign out)
        # These days won't be counted at all - as if the student wasn't present
        session_data = session_data[session_data['שעת יציאה'] != time(0, 0)]

        # If after filtering there's no data - return empty statistics
        if session_data.empty:
            return self._get_empty_stats(working_days, session_type)

        daily_stats = session_data.groupby('תאריך').agg({
            'שעת כניסה': 'min',
            'שעת יציאה': 'max',
            'רצופות': lambda x: 'כן' in x.values
        })

        daily_stats['סך שעות'] = daily_stats.apply(
            lambda row: self._calculate_session_hours(
                row['שעת כניסה'],
                row['שעת יציאה'],
                session_type
            ),
            axis=1
        )

        attended_days = min(len(daily_stats), working_days)

        late_days, very_late_days, perfect_days, partial_bonus_days = \
            self._calculate_attendance_stats(daily_stats, config, working_days)

        expected_hours = working_days * get_expected_hours(session_type)
        total_hours = min(daily_stats['סך שעות'].sum(), expected_hours)
        missed_hours = max(0, expected_hours - total_hours)

        # Calculate absences
        absent_days = working_days - attended_days

        # Full base scholarship only if: up to 2 lates AND up to 2 absences
        base = self._calculate_base_scholarship(
            late_days, absent_days, attended_days, very_late_days, config)

        daily_bonus = (perfect_days * BONUS_CONFIG['DAILY_PERFECT'] +
                       partial_bonus_days * BONUS_CONFIG['DAILY_PARTIAL'])

        early_attendance_bonus = 0
        if session_type == 'בוקר':
            early_attendance_bonus = self._calculate_early_attendance_bonus(
                session_data, daily_stats, absent_days, has_afternoon
            )

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

    def _calculate_weekly_early_attendance(
            self, attendance_data: pd.DataFrame) -> float:
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

        def is_late(t):
            if isinstance(t, str):
                t = self._parse_time(t)
            return compare_times(t, NINE_AM_TIME)

        weekly_lates = (
            attendance_data
            .groupby('week')
            .agg({'שעת כניסה': lambda x: sum(is_late(t) for t in x)})
        )

        good_weeks = sum(weekly_lates['שעת כניסה'] <= 1)
        return good_weeks * 35

    def calculate_student_scholarship(
            self,
            student_data: pd.DataFrame,
            working_days: int) -> Dict:
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
        # Identify days without valid exit time
        warnings = []
        student_data_copy = student_data.copy()
        student_data_copy['שעת יציאה_parsed'] = student_data_copy['שעת יציאה'].apply(
            self._parse_time)
        missing_exit_days = len(
            student_data_copy[student_data_copy['שעת יציאה_parsed'] == time(0, 0)])

        if missing_exit_days > 0:
            warnings.append(
                f"⚠️ {missing_exit_days} ימים ללא זמן יציאה (לא נספרו)")

        morning_data = student_data[student_data['סדר'] == 'בוקר']
        afternoon_data = student_data[student_data['סדר'] == 'צהריים']

        has_afternoon = not afternoon_data.empty

        morning_stats = self.calculate_session_stats(
            morning_data, 'בוקר', working_days, has_afternoon)
        afternoon_stats = self.calculate_session_stats(
            afternoon_data, 'צהריים', working_days, has_afternoon)

        # Add warnings about base scholarship disqualification
        if not morning_data.empty:
            warning = self._build_disqualification_warning(
                morning_stats, 'בוקר')
            if warning:
                warnings.append(warning)

        if not afternoon_data.empty:
            warning = self._build_disqualification_warning(
                afternoon_stats, 'צהריים')
            if warning:
                warnings.append(warning)

        # Warning about early attendance bonus reduction
        if not morning_data.empty:
            max_bonus = 200 if has_afternoon else 100
            if morning_stats.early_attendance_bonus < max_bonus and \
               morning_stats.absent_days > MAX_ALLOWED_LATES_OR_ABSENCES:
                warnings.append(
                    f"⚠️ בונוס הגעה מוקדמת מופחת ({morning_stats.absent_days} היעדרויות)")

        total_base = morning_stats.base + afternoon_stats.base
        total_bonus = morning_stats.bonus + afternoon_stats.bonus

        # Monthly bonuses breakdown - not added to total_bonus as they will be
        # added in the "final total" formula
        monthly_tier1_bonus = 0
        monthly_tier2_bonus = 0
        perfect_attendance_bonus = 0

        if not morning_data.empty and not afternoon_data.empty:
            if (morning_stats.missed_hours <= BONUS_CONFIG['MORNING_TIER1_MAX_MISSING'] and
                    afternoon_stats.missed_hours <= BONUS_CONFIG['AFTERNOON_TIER1_MAX_MISSING']):
                monthly_tier1_bonus = BONUS_CONFIG['MONTHLY_TIER1']

            if (morning_stats.missed_hours <= BONUS_CONFIG['MORNING_TIER2_MAX_MISSING'] and
                    afternoon_stats.missed_hours <= BONUS_CONFIG['AFTERNOON_TIER2_MAX_MISSING']):
                monthly_tier2_bonus = BONUS_CONFIG['MONTHLY_TIER2']

            if (morning_stats.absent_days == 0 and afternoon_stats.absent_days ==
                    0 and                     morning_stats.late_days == 0 and afternoon_stats.late_days == 0):
                perfect_attendance_bonus = BONUS_CONFIG['PERFECT_ATTENDANCE']

        # Handle NaN in names
        last_name = str(
            student_data['שם משפחה'].iloc[0]) if pd.notna(
            student_data['שם משפחה'].iloc[0]) else ''
        first_name = str(
            student_data['שם פרטי'].iloc[0]) if pd.notna(
            student_data['שם פרטי'].iloc[0]) else ''

        return {
            'מספר זהות': student_data['זהות'].iloc[0],
            'שם מלא': f"{last_name} {first_name}".strip(),
            'מלגת בסיס': total_base,
            'תוספות': total_bonus,
            'בונוס יומי בוקר': morning_stats.bonus,
            'בונוס יומי צהריים': afternoon_stats.bonus,
            'תוספת דרגה 1': monthly_tier1_bonus,
            'תוספת דרגה 2': monthly_tier2_bonus,
            'נוכחות מושלמת': perfect_attendance_bonus,
            'הגעה מוקדמת': morning_stats.early_attendance_bonus,
            'סך הכל': total_base + total_bonus + morning_stats.early_attendance_bonus + monthly_tier1_bonus + monthly_tier2_bonus + perfect_attendance_bonus,
            'התראות': ' | '.join(warnings) if warnings else '',
            **{f'בוקר_{k}': v for k, v in vars(morning_stats).items()},
            **{f'צהריים_{k}': v for k, v in vars(afternoon_stats).items()}
        }
