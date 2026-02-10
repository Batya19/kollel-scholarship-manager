from tkinter import messagebox
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from models.scholarship import KollelScholarship
from utils.attendance_details import add_detailed_sheets
from utils.summary_formatter import format_summary_sheet


def process_kollel_attendance(input_file: str, output_file: str, working_days: int) -> bool:
    """
    Process Kollel attendance data and generate scholarship calculations.

    This function reads attendance data from an Excel file, processes it to calculate
    scholarships, and generates a new Excel file with detailed calculations and
    additional manual input columns.

    Args:
        input_file (str): Path to input Excel file containing attendance data
        output_file (str): Path where the output Excel file will be saved
        working_days (int): Number of working days in the period

    Returns:
        bool: True if processing was successful, False otherwise

    Raises:
        ValueError: If required columns are missing in the input file
        Exception: For other processing errors
    """
    try:
        df = pd.read_excel(input_file, dtype={
            'זהות': str,
            'שם משפחה': str,
            'שם פרטי': str,
            'רצופות': str
        })
        df = df.dropna(how='all')

        for col in ['כניסה', 'יציאה']:
            df[col] = pd.to_datetime(df[col])

        first_date = df['כניסה'].iloc[0]
        month_number = first_date.month
        year = first_date.year

        month_names = {
            1: "ינואר", 2: "פברואר", 3: "מרץ", 4: "אפריל",
            5: "מאי", 6: "יוני", 7: "יולי", 8: "אוגוסט",
            9: "ספטמבר", 10: "אוקטובר", 11: "נובמבר", 12: "דצמבר"
        }

        month_name = month_names.get(month_number, str(month_number))
        output_file = f"מלגות חודשיות {month_name} {year}.xlsx"

        df['סך שעות'] = df['יציאה'] - df['כניסה']
        df['תאריך'] = df['כניסה'].dt.date
        df['שעת כניסה'] = df['כניסה'].dt.time
        df['שעת יציאה'] = df['יציאה'].dt.time
        df['סדר'] = df['שעת כניסה'].apply(lambda x: 'בוקר' if x.hour < 12 else 'צהריים')

        calculator = KollelScholarship()
        results = [
            calculator.calculate_student_scholarship(student_data, working_days)
            for _, student_data in df.groupby('זהות')
        ]

        results_df = pd.DataFrame(results)

        # חשב תוספות חודשיות לפני שמוחקים את העמודות הסטטיסטיות
        results_df['תוספת דרגה 1'] = results_df.apply(
            lambda row: 190 if (
                    row['בוקר_missed_hours'] <= 7.0 and
                    row['צהריים_missed_hours'] <= 6.0 and
                    row['צהריים_attended_days'] > 0
            ) else 0,
            axis=1
        )
        results_df['אנס דרגה 1'] = ''

        results_df['תוספת דרגה 2'] = results_df.apply(
            lambda row: 200 if (
                    row['בוקר_missed_hours'] <= 3.0 and
                    row['צהריים_missed_hours'] <= 2.0 and
                    row['צהריים_attended_days'] > 0
            ) else 0,
            axis=1
        )
        results_df['אנס דרגה 2'] = ''

        results_df['נוכחות מושלמת'] = results_df.apply(
            lambda row: 200 if (
                    row['בוקר_absent_days'] == 0 and
                    row['בוקר_late_days'] == 0 and
                    row['צהריים_absent_days'] == 0 and
                    row['צהריים_late_days'] == 0 and
                    row['צהריים_attended_days'] > 0
            ) else 0,
            axis=1
        )
        results_df['אנס נוכחות מושלמת'] = ''

        results_df['הגעה מוקדמת'] = results_df['תוספת נוכחות מוקדמת']
        results_df['אנס הגעה מוקדמת'] = ''

        # עכשיו הסר עמודות סטטיסטיקה מפורטות (בוקר_, צהריים_) - הן רק מבלבלות בסיכום
        # וגם הסר "תוספת נוכחות מוקדמת" כי יש לנו "הגעה מוקדמת" (זה אותו דבר!)
        stats_cols_to_remove = [col for col in results_df.columns 
                                if col.startswith('בוקר_') or col.startswith('צהריים_') 
                                or col == 'תוספת נוכחות מוקדמת']
        results_df = results_df.drop(columns=stats_cols_to_remove)

        # סדר מחדש את העמודות - התראות מיד אחרי השם
        base_cols = ['מספר זהות', 'שם מלא', 'התראות']
        other_cols = [col for col in results_df.columns if col not in base_cols]
        results_df = results_df[base_cols + other_cols]

        # הוסף עמודות ציונים ידניות
        manual_columns = [
            'ציון חבורה', 'חבורה',
            'ציון מוסר', 'מוסר',
            'ציון סיכומים', 'סיכומים',
            'ציון מבחן הלכה', 'מבחן הלכה',
            'ציון מבחן שס', 'מבחן שס'
        ]
        for col in manual_columns:
            results_df[col] = ''

        results_df['סך סופי'] = 0

        results_df.to_excel(output_file, index=False)

        wb = load_workbook(output_file)
        ws = wb.active
        
        # הגדר כיוון מימין לשמאל (RTL)
        ws.sheet_view.rightToLeft = True

        headers = [cell.value for cell in ws[1]]

        try:
            total_col = get_column_letter(headers.index('סך הכל') + 1)

            chabura_grade_col = get_column_letter(headers.index('ציון חבורה') + 1)
            musar_grade_col = get_column_letter(headers.index('ציון מוסר') + 1)
            sikumim_grade_col = get_column_letter(headers.index('ציון סיכומים') + 1)
            halacha_grade_col = get_column_letter(headers.index('ציון מבחן הלכה') + 1)
            shas_grade_col = get_column_letter(headers.index('ציון מבחן שס') + 1)

            chabura_col = get_column_letter(headers.index('חבורה') + 1)
            musar_col = get_column_letter(headers.index('מוסר') + 1)
            sikumim_col = get_column_letter(headers.index('סיכומים') + 1)
            halacha_col = get_column_letter(headers.index('מבחן הלכה') + 1)
            shas_col = get_column_letter(headers.index('מבחן שס') + 1)

            tier1_col = get_column_letter(headers.index('תוספת דרגה 1') + 1)
            tier1_reason_col = get_column_letter(headers.index('אנס דרגה 1') + 1)

            tier2_col = get_column_letter(headers.index('תוספת דרגה 2') + 1)
            tier2_reason_col = get_column_letter(headers.index('אנס דרגה 2') + 1)

            perfect_col = get_column_letter(headers.index('נוכחות מושלמת') + 1)
            perfect_reason_col = get_column_letter(headers.index('אנס נוכחות מושלמת') + 1)

            early_col = get_column_letter(headers.index('הגעה מוקדמת') + 1)
            early_reason_col = get_column_letter(headers.index('אנס הגעה מוקדמת') + 1)

            final_col = get_column_letter(headers.index('סך סופי') + 1)
        except ValueError as e:
            raise ValueError(f"לא נמצאה אחת העמודות הנדרשות: {str(e)}")

        for row in range(2, ws.max_row + 1):
            ws[
                f'{chabura_col}{row}'] = f'=IF(ISNUMBER({chabura_grade_col}{row}),IF({chabura_grade_col}{row}>=96,150,ROUND(150*{chabura_grade_col}{row}/100,0)),"")'

            ws[
                f'{musar_col}{row}'] = f'=IF(ISNUMBER({musar_grade_col}{row}),IF({musar_grade_col}{row}>=96,100,ROUND(100*{musar_grade_col}{row}/100,0)),"")'

            ws[
                f'{sikumim_col}{row}'] = f'=IF(ISNUMBER({sikumim_grade_col}{row}),IF({sikumim_grade_col}{row}>=96,150,ROUND(150*{sikumim_grade_col}{row}/100,0)),"")'

            ws[
                f'{halacha_col}{row}'] = f'=IF(ISNUMBER({halacha_grade_col}{row}),IF({halacha_grade_col}{row}>=96,200,ROUND(200*{halacha_grade_col}{row}/100,0)),"")'

            ws[
                f'{shas_col}{row}'] = f'=IF(ISNUMBER({shas_grade_col}{row}),IF({shas_grade_col}{row}>=96,390,ROUND(390*{shas_grade_col}{row}/100,0)),"")'

            formula = (
                f'={total_col}{row}+'
                f'IF(ISNUMBER({chabura_col}{row}),{chabura_col}{row},0)+'
                f'IF(ISNUMBER({musar_col}{row}),{musar_col}{row},0)+'
                f'IF(ISNUMBER({sikumim_col}{row}),{sikumim_col}{row},0)+'
                f'IF(ISNUMBER({halacha_col}{row}),{halacha_col}{row},0)+'
                f'IF(ISNUMBER({shas_col}{row}),{shas_col}{row},0)+'
                f'IF(AND({tier1_reason_col}{row}<>"",{tier1_reason_col}{row}<>0),IF(ISNUMBER({tier1_col}{row}),IF({tier1_col}{row}=0,190,{tier1_col}{row}),190),IF(ISNUMBER({tier1_col}{row}),{tier1_col}{row},0))+'
                f'IF(AND({tier2_reason_col}{row}<>"",{tier2_reason_col}{row}<>0),IF(ISNUMBER({tier2_col}{row}),IF({tier2_col}{row}=0,200,{tier2_col}{row}),200),IF(ISNUMBER({tier2_col}{row}),{tier2_col}{row},0))+'
                f'IF(AND({perfect_reason_col}{row}<>"",{perfect_reason_col}{row}<>0),IF(ISNUMBER({perfect_col}{row}),IF({perfect_col}{row}=0,200,{perfect_col}{row}),200),IF(ISNUMBER({perfect_col}{row}),{perfect_col}{row},0))+'
                f'IF(AND({early_reason_col}{row}<>"",{early_reason_col}{row}<>0),IF(ISNUMBER({early_col}{row}),IF({early_col}{row}=0,200,{early_col}{row}),200),IF(ISNUMBER({early_col}{row}),{early_col}{row},0))'
            )
            ws[f'{final_col}{row}'] = formula

        wb.save(output_file)
        
        # הוסף גליונות מפורטים לכל אברך
        try:
            add_detailed_sheets(output_file, df, results, working_days)
        except Exception as e:
            print(f"Warning: Could not add detailed sheets: {e}")
            # ממשיכים גם אם הוספת הגליונות נכשלה
        
        # עצב את גליון הסיכום
        try:
            format_summary_sheet(output_file)
        except Exception as e:
            print(f"Warning: Could not format summary sheet: {e}")
            # ממשיכים גם אם העיצוב נכשל
        
        return True

    except Exception as e:
        print(f"An error occurred: {e}")
        messagebox.showerror("שגיאה", f"אירעה שגיאה: {str(e)}")
        return False