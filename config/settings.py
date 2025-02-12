from datetime import time

MORNING_CONFIG = {
    'BASE': 400,
    'START': time(9, 30),
    'PERFECT_START': time(9, 30),
    'LATE_THRESHOLD': time(10, 0),
    'VERY_LATE_THRESHOLD': time(10, 30),
    'END': time(13, 0),
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
