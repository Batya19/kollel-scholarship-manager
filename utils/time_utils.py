from datetime import time


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
