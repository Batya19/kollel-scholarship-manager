from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def format_summary_sheet(excel_file: str):
    """
    מעצב את גליון הסיכום בצורה יפה וקריאה.
    
    Args:
        excel_file (str): נתיב לקובץ Excel
    """
    wb = load_workbook(excel_file)
    ws = wb['סיכום'] if 'סיכום' in wb.sheetnames else wb.active
    
    # מצא את שורת הכותרות (שורה 1)
    header_row = 1
    
    # עצב את שורת הכותרות - עיצוב קלאסי ופשוט
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # עצב כותרות - פשוט וקלאסי
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=header_row, column=col)
        
        # עיצוב מינימליסטי
        cell.font = Font(bold=True, size=11)
        cell.fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")  # אפור בהיר
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border
    
    # גובה שורת כותרות
    ws.row_dimensions[header_row].height = 25
    
    # עצב את תאי הנתונים - פשוט וקלאסי
    for row in range(header_row + 1, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = thin_border
            
            # יישור מרכז למספרים
            if isinstance(cell.value, (int, float)):
                cell.alignment = Alignment(horizontal='center', vertical='center')
            else:
                cell.alignment = Alignment(horizontal='right', vertical='center')
    
    # התאם רוחב עמודות
    column_widths = {
        'מספר זהות': 12,
        'שם מלא': 20,
        'התראות': 35,
        'מלגת בסיס': 12,
        'תוספות': 10,
        'סך הכל': 12,
        'ציון חבורה': 10,
        'חבורה': 10,
        'ציון מוסר': 10,
        'מוסר': 10,
        'ציון סיכומים': 10,
        'סיכומים': 10,
        'ציון מבחן הלכה': 12,
        'מבחן הלכה': 12,
        'ציון מבחן שס': 12,
        'מבחן שס': 12,
        'תוספת דרגה 1': 12,
        'אנס דרגה 1': 12,
        'תוספת דרגה 2': 12,
        'אנס דרגה 2': 12,
        'נוכחות מושלמת': 12,
        'אנס נוכחות מושלמת': 15,
        'הגעה מוקדמת': 12,
        'אנס הגעה מוקדמת': 15,
        'סך סופי': 12,
    }
    
    for col in range(1, ws.max_column + 1):
        header = ws.cell(row=header_row, column=col).value
        col_letter = get_column_letter(col)
        
        if header in column_widths:
            ws.column_dimensions[col_letter].width = column_widths[header]
        else:
            # רוחב ברירת מחדל
            ws.column_dimensions[col_letter].width = 12
    
    # הקפא את שורת הכותרות
    ws.freeze_panes = ws[f'A{header_row + 1}']
    
    # שמור
    wb.save(excel_file)
