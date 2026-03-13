from datetime import time, datetime

def time_to_minutes(t: time) -> int:
    """
    Convert time object to minutes since midnight.

    Args:
        t (time): Time object to convert

    Returns:
        int: Total minutes since midnight
    """
    return t.hour * 60 + t.minute


def compare_times(t1: time, t2: time) -> bool:
    """
    Compare if t1 is later than t2.

    Args:
        t1 (time): First time
        t2 (time): Second time

    Returns:
        bool: True if t1 > t2
    """
    return time_to_minutes(t1) > time_to_minutes(t2)


def compare_times_gte(t1: time, t2: time) -> bool:
    """
    Compare if t1 is later than or equal to t2.

    Args:
        t1 (time): First time
        t2 (time): Second time

    Returns:
        bool: True if t1 >= t2
    """
    return time_to_minutes(t1) >= time_to_minutes(t2)


def calculate_hours_between(entry_time: time, exit_time: time,
                            valid_start: time, valid_end: time) -> float:
    """
    Calculate actual hours between entry and exit within valid time window.

    Args:
        entry_time (time): Session entry time
        exit_time (time): Session exit time
        valid_start (time): Valid session start time
        valid_end (time): Valid session end time

    Returns:
        float: Calculated hours, rounded to 2 decimal places
    """
    entry_minutes = time_to_minutes(entry_time)
    exit_minutes = time_to_minutes(exit_time)
    start_minutes = time_to_minutes(valid_start)
    end_minutes = time_to_minutes(valid_end)

    actual_start = max(entry_minutes, start_minutes)
    actual_end = min(exit_minutes, end_minutes)

    if actual_end > actual_start:
        hours = (actual_end - actual_start) / 60
        return round(hours, 2)
    return 0


def get_session_config(session_type: str):
    """
    Get configuration for session type.

    Args:
        session_type (str): Type of session ('בוקר' or 'צהריים')

    Returns:
        dict: Session configuration
    """
    from config.settings import MORNING_CONFIG, AFTERNOON_CONFIG
    return MORNING_CONFIG if session_type == 'בוקר' else AFTERNOON_CONFIG


def get_expected_hours(session_type: str) -> float:
    """
    Get expected hours for session type.

    Args:
        session_type (str): Type of session ('בוקר' or 'צהריים')

    Returns:
        float: Expected hours for the session
    """
    return 3.5 if session_type == 'בוקר' else 3.0


def is_invalid_record(entry_time, exit_time, config) -> tuple:
    if not isinstance(entry_time, time) or not isinstance(exit_time, time):
        return True, "חסר זמן כניסה/יציאה"
    
    if entry_time == exit_time or entry_time > exit_time:
        return True, "זמן כניסה/יציאה זהים או הפוכים"
    if exit_time <= config['START']:
        return True, "יציאה לפני תחילת הסדר"
    if entry_time >= config['END']:
        return True, "כניסה לאחר סוף הסדר"
    
    hours = calculate_hours_between(entry_time, exit_time, config['START'], config['END'])
    if hours == 0:
        return True, "אפס שעות בפועל"
    
    return False, ""


def is_before_nine(entry_time: time) -> bool:
    """Returns True if entry time is before 09:00"""
    from config.settings import NINE_AM_MINUTES
    return time_to_minutes(entry_time) < NINE_AM_MINUTES


def is_perfect_day(entry_time: time, exit_time: time, is_continuous: bool, config: dict) -> bool:
    """Returns True if this is a perfect attendance day"""
    return (
        time_to_minutes(entry_time) <= time_to_minutes(config['PERFECT_START']) and
        time_to_minutes(exit_time) >= time_to_minutes(config['END']) and
        is_continuous
    )


def parse_time(time_str) -> time:
    """
    Convert various time representations to a time object.
    Handles strings, floats, datetime objects, and NaT values.
    Returns time(0, 0) for invalid/missing values.
    """
    import pandas as pd
    try:
        if pd.isna(time_str) or time_str is pd.NaT:
            return time(0, 0)
        if isinstance(time_str, str):
            try:
                return pd.to_datetime(time_str).time()
            except (ValueError, TypeError, AttributeError) as e:
                print(f"Failed to parse time string: {time_str}, error: {e}")
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
            print(f"Unknown time format: {time_str}, type: {type(time_str)}")
            return time(0, 0)
    except Exception as e:
        print(f"Error in parse_time: {time_str}, type: {type(time_str)}, error: {e}")
        return time(0, 0)