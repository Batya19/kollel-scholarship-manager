from datetime import time

# General rules
MAX_ALLOWED_LATES_OR_ABSENCES = 2
NINE_AM_MINUTES = 9 * 60

# Time to separate morning and afternoon sessions
# Any entry before this time is 'morning', any entry at or after is 'afternoon'
SESSION_SEPARATOR_TIME = time(13, 30)

MORNING_CONFIG = {
    'BASE': 400,
    'START': time(9, 30),
    'PERFECT_START': time(9, 30),
    'LATE_THRESHOLD': time(10, 0),
    'VERY_LATE_THRESHOLD': time(10, 30),
    'END': time(13, 0),
    'EARLY_LEAVE_THRESHOLD': time(12, 30),
    'LATE_DAILY_RATE': 15,
    'HOURS': 3.5
}

AFTERNOON_CONFIG = {
    'BASE': 250,
    'START': time(14, 0),
    'PERFECT_START': time(14, 0),
    'LATE_THRESHOLD': time(14, 15),
    'VERY_LATE_THRESHOLD': time(14, 30),
    'END': time(17, 0),
    'EARLY_LEAVE_THRESHOLD': time(16, 30),
    'LATE_DAILY_RATE': 10,
    'HOURS': 3.0
}

BONUS_CONFIG = {
    'DAILY_PERFECT': 20,
    'DAILY_PARTIAL': 5,
    'MONTHLY_TIER1': 190,
    'MONTHLY_TIER2': 200,
    'PERFECT_ATTENDANCE': 200,
    'MORNING_TIER1_MAX_MISSING': 7.0,
    'MORNING_TIER2_MAX_MISSING': 3.0,
    'AFTERNOON_TIER1_MAX_MISSING': 6.0,
    'AFTERNOON_TIER2_MAX_MISSING': 2.0
}

def get_max_early_bonus(has_afternoon: bool) -> int:
    """Maximum early attendance bonus based on session participation."""
    return 200 if has_afternoon else 100