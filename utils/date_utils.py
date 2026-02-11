def get_hebrew_month_name(month_number: int) -> str:
    """
    Get Hebrew month name from month number.

    Args:
        month_number (int): Month number (1-12)

    Returns:
        str: Hebrew month name
    """
    month_names = {
        1: "ינואר", 2: "פברואר", 3: "מרץ", 4: "אפריל",
        5: "מאי", 6: "יוני", 7: "יולי", 8: "אוגוסט",
        9: "ספטמבר", 10: "אוקטובר", 11: "נובמבר", 12: "דצמבר"
    }
    return month_names.get(month_number, str(month_number))


def get_hebrew_day_name(weekday: int) -> str:
    """
    Get Hebrew abbreviated day name from weekday number.

    Args:
        weekday (int): Weekday number (0=Monday, 6=Sunday)

    Returns:
        str: Hebrew abbreviated day name
    """
    day_names = {
        0: 'ב\'', 1: 'ג\'', 2: 'ד\'',
        3: 'ה\'', 4: 'ו\'', 5: 'ש\'', 6: 'א\''
    }
    return day_names.get(weekday, '')
