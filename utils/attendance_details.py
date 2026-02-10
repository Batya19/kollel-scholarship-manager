import pandas as pd
from datetime import time
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from config.settings import MORNING_CONFIG, AFTERNOON_CONFIG


def add_detailed_sheets(excel_file: str, input_data: pd.DataFrame, scholarship_results: list, working_days: int):
    """
    מוסיף גליונות מפורטים לכל אברך בקובץ Excel קיים.
    
    Args:
        excel_file (str): נתיב לקובץ Excel (שכבר מכיל את גליון הסיכום)
        input_data (pd.DataFrame): נתוני הנוכחות המקוריים
        scholarship_results (list): תוצאות חישוב המלגות
        working_days (int): מספר ימי עבודה בחודש
    """
    
    # טען את הקובץ הקיים
    wb = load_workbook(excel_file)
    
    # שנה את שם הגליון הראשון ל"סיכום"
    if 'Sheet' in wb.sheetnames or wb.sheetnames[0]:
        ws_summary = wb.active
        ws_summary.title = "סיכום"
        # הגדר כיוון מימין לשמאל (RTL)
        ws_summary.sheet_view.rightToLeft = True
    
    # עבור על כל אברך וצור לו גליון
    for result in scholarship_results:
        student_id = result['מספר זהות']
        student_name = result['שם מלא']
        
        # סנן את נתוני האברך
        student_data = input_data[input_data['זהות'] == student_id].copy()
        
        if student_data.empty:
            continue
        
        # צור גליון חדש לאברך (מגביל לـ31 תווים - מגבלת Excel)
        sheet_name = student_name[:31]
        ws = wb.create_sheet(title=sheet_name)
        
        # הגדר כיוון מימין לשמאל (RTL)
        ws.sheet_view.rightToLeft = True
        
        # הוסף כותרת
        ws.merge_cells('A1:M1')
        ws['A1'] = f"פירוט נוכחות - {student_name}"
        ws['A1'].font = Font(size=16, bold=True)
        ws['A1'].alignment = Alignment(horizontal='center')
        
        ws.merge_cells('A2:M2')
        ws['A2'] = f"מספר זהות: {student_id}"
        ws['A2'].font = Font(size=12)
        ws['A2'].alignment = Alignment(horizontal='center')
        
        # הוסף כותרות עמודות (שורה 4)
        headers = [
            'תאריך', 'יום', 'סדר', 'שעת כניסה', 'שעת יציאה', 
            'רצופות', 'סך שעות', 'סטטוס', 'זמן איחור (דק\')',
            'הגיע לפני 9:00?', 'בונוס יומי', 'שעות חסרות', 'הערות'
        ]
        
        for col, header in enumerate(headers, start=1):
            cell = ws.cell(row=4, column=col)
            cell.value = header
            cell.font = Font(bold=True, size=11)
            cell.fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
            cell.alignment = Alignment(horizontal='center')
            cell.border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
        
        # הוסף את נתוני הנוכחות
        student_data = student_data.sort_values(['תאריך', 'סדר'])
        
        row = 5
        for _, record in student_data.iterrows():
            # תאריך
            ws.cell(row=row, column=1).value = record['תאריך'].strftime('%d/%m/%Y') if hasattr(record['תאריך'], 'strftime') else str(record['תאריך'])
            
            # יום בשבוע
            day_names = {0: 'ב\'', 1: 'ג\'', 2: 'ד\'', 3: 'ה\'', 4: 'ו\'', 5: 'ש\'', 6: 'א\''}
            if hasattr(record['תאריך'], 'weekday'):
                ws.cell(row=row, column=2).value = day_names.get(record['תאריך'].weekday(), '')
            
            # סדר
            ws.cell(row=row, column=3).value = record['סדר']
            
            # שעת כניסה
            entry_time = record['שעת כניסה']
            if isinstance(entry_time, time):
                ws.cell(row=row, column=4).value = entry_time.strftime('%H:%M')
            elif hasattr(entry_time, 'time'):
                ws.cell(row=row, column=4).value = entry_time.time().strftime('%H:%M')
            else:
                ws.cell(row=row, column=4).value = str(entry_time)
            
            # שעת יציאה
            exit_time = record['שעת יציאה']
            if isinstance(exit_time, time):
                if exit_time == time(0, 0):
                    ws.cell(row=row, column=5).value = "חסר"
                    ws.cell(row=row, column=5).font = Font(color="FF0000")
                else:
                    ws.cell(row=row, column=5).value = exit_time.strftime('%H:%M')
            elif hasattr(exit_time, 'time'):
                ws.cell(row=row, column=5).value = exit_time.time().strftime('%H:%M')
            else:
                ws.cell(row=row, column=5).value = str(exit_time)
            
            # רצופות
            ws.cell(row=row, column=6).value = record.get('רצופות', '')
            
            # סך שעות - נחשב
            session_type = record['סדר']
            config = MORNING_CONFIG if session_type == 'בוקר' else AFTERNOON_CONFIG
            
            if isinstance(entry_time, time) and isinstance(exit_time, time) and exit_time != time(0, 0):
                entry_minutes = entry_time.hour * 60 + entry_time.minute
                exit_minutes = exit_time.hour * 60 + exit_time.minute
                start_minutes = config['START'].hour * 60 + config['START'].minute
                end_minutes = config['END'].hour * 60 + config['END'].minute
                
                actual_start = max(entry_minutes, start_minutes)
                actual_end = min(exit_minutes, end_minutes)
                
                if actual_end > actual_start:
                    hours = round((actual_end - actual_start) / 60, 2)
                    ws.cell(row=row, column=7).value = hours
                else:
                    ws.cell(row=row, column=7).value = 0
            else:
                ws.cell(row=row, column=7).value = 0
            
            # סטטוס (איחור/בזמן)
            if isinstance(entry_time, time):
                entry_minutes = entry_time.hour * 60 + entry_time.minute
                late_threshold_minutes = config['LATE_THRESHOLD'].hour * 60 + config['LATE_THRESHOLD'].minute
                very_late_threshold_minutes = config['VERY_LATE_THRESHOLD'].hour * 60 + config['VERY_LATE_THRESHOLD'].minute
                
                if entry_minutes >= very_late_threshold_minutes:
                    ws.cell(row=row, column=8).value = "❌ איחור משמעותי"
                    ws.cell(row=row, column=8).font = Font(color="FF0000")
                elif entry_minutes >= late_threshold_minutes:
                    ws.cell(row=row, column=8).value = "⚠️ איחור"
                    ws.cell(row=row, column=8).font = Font(color="FFA500")
                else:
                    ws.cell(row=row, column=8).value = "✅ בזמן"
                    ws.cell(row=row, column=8).font = Font(color="008000")
            
            # זמן איחור
            if isinstance(entry_time, time):
                entry_minutes_calc = entry_time.hour * 60 + entry_time.minute
                perfect_start_minutes = config['PERFECT_START'].hour * 60 + config['PERFECT_START'].minute
                
                if entry_minutes_calc > perfect_start_minutes:
                    late_minutes = entry_minutes_calc - perfect_start_minutes
                    ws.cell(row=row, column=9).value = late_minutes
                else:
                    ws.cell(row=row, column=9).value = 0
            
            # הגיע לפני 9:00? (בבוקר בלבד)
            if session_type == 'בוקר' and isinstance(entry_time, time):
                entry_minutes_calc = entry_time.hour * 60 + entry_time.minute
                nine_am_minutes = 9 * 60
                if entry_minutes_calc < nine_am_minutes:
                    ws.cell(row=row, column=10).value = "כן"
                    ws.cell(row=row, column=10).font = Font(color="008000")
                else:
                    ws.cell(row=row, column=10).value = "לא"
            else:
                ws.cell(row=row, column=10).value = "-"
            
            # בונוס יומי - חישוב מקורב
            bonus = 0
            if isinstance(entry_time, time) and isinstance(exit_time, time) and exit_time != time(0, 0):
                entry_minutes_calc = entry_time.hour * 60 + entry_time.minute
                perfect_start_minutes = config['PERFECT_START'].hour * 60 + config['PERFECT_START'].minute
                end_minutes = config['END'].hour * 60 + config['END'].minute
                
                is_perfect = (
                    entry_minutes_calc <= perfect_start_minutes and
                    (exit_time.hour * 60 + exit_time.minute) >= end_minutes and
                    record.get('רצופות', '') == 'כן'
                )
                
                if is_perfect:
                    bonus = 20
                elif record.get('רצופות', '') == 'כן':
                    bonus = 5
            
            ws.cell(row=row, column=11).value = f"{bonus} ₪" if bonus > 0 else "-"
            
            # שעות חסרות
            expected_hours = config['HOURS']
            actual_hours = ws.cell(row=row, column=7).value or 0
            missed_hours = round(expected_hours - actual_hours, 2)
            ws.cell(row=row, column=12).value = missed_hours if missed_hours > 0 else 0
            
            # הערות
            notes = []
            if exit_time == time(0, 0):
                notes.append("חסר זמן יציאה")
            if isinstance(entry_time, time):
                entry_minutes_calc = entry_time.hour * 60 + entry_time.minute
                very_late_threshold_minutes = config['VERY_LATE_THRESHOLD'].hour * 60 + config['VERY_LATE_THRESHOLD'].minute
                if entry_minutes_calc >= very_late_threshold_minutes:
                    notes.append("איחור חמור")
            
            ws.cell(row=row, column=13).value = ", ".join(notes) if notes else ""
            
            row += 1
        
        # הוסף סיכום בתחתית
        summary_row = row + 2
        ws.merge_cells(f'A{summary_row}:M{summary_row}')
        ws[f'A{summary_row}'] = "סיכום"
        ws[f'A{summary_row}'].font = Font(size=14, bold=True)
        ws[f'A{summary_row}'].alignment = Alignment(horizontal='center')
        ws[f'A{summary_row}'].fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        ws[f'A{summary_row}'].font = Font(size=14, bold=True, color="FFFFFF")
        
        summary_row += 2
        
        # נתוני סיכום מתוך result
        summary_data = [
            ("ימי נוכחות - בוקר:", result.get('בוקר_attended_days', 0)),
            ("ימי נוכחות - צהריים:", result.get('צהריים_attended_days', 0)),
            ("סה\"כ איחורים - בוקר:", result.get('בוקר_late_days', 0)),
            ("סה\"כ איחורים - צהריים:", result.get('צהריים_late_days', 0)),
            ("סה\"כ היעדרויות - בוקר:", result.get('בוקר_absent_days', 0)),
            ("סה\"כ היעדרויות - צהריים:", result.get('צהריים_absent_days', 0)),
            ("סה\"כ שעות - בוקר:", result.get('בוקר_total_hours', 0)),
            ("סה\"כ שעות - צהריים:", result.get('צהריים_total_hours', 0)),
            ("", ""),
            ("מלגת בסיס:", f"{result.get('מלגת בסיס', 0)} ₪"),
            ("תוספות:", f"{result.get('תוספות', 0)} ₪"),
            ("בונוס הגעה מוקדמת:", f"{result.get('תוספת נוכחות מוקדמת', 0)} ₪"),
            ("סך הכל:", f"{result.get('סך הכל', 0)} ₪"),
        ]
        
        for label, value in summary_data:
            ws.cell(row=summary_row, column=1).value = label
            ws.cell(row=summary_row, column=1).font = Font(bold=True)
            ws.cell(row=summary_row, column=2).value = value
            summary_row += 1
        
        # התאם רוחב עמודות
        ws.column_dimensions['A'].width = 12
        ws.column_dimensions['B'].width = 6
        ws.column_dimensions['C'].width = 10
        ws.column_dimensions['D'].width = 12
        ws.column_dimensions['E'].width = 12
        ws.column_dimensions['F'].width = 10
        ws.column_dimensions['G'].width = 10
        ws.column_dimensions['H'].width = 18
        ws.column_dimensions['I'].width = 15
        ws.column_dimensions['J'].width = 15
        ws.column_dimensions['K'].width = 12
        ws.column_dimensions['L'].width = 14
        ws.column_dimensions['M'].width = 20
    
    # שמור את הקובץ
    wb.save(excel_file)
