from tkinter import messagebox
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from models.scholarship import KollelScholarship


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

        df['סך שעות'] = df['יציאה'] - df['כניסה']
        df['תאריך'] = df['כניסה'].dt.date
        df['שעת_כניסה'] = df['כניסה'].dt.time
        df['שעת_יציאה'] = df['יציאה'].dt.time
        df['סדר'] = df['שעת_כניסה'].apply(lambda x: 'בוקר' if x.hour < 12 else 'צהריים')

        calculator = KollelScholarship()
        results = [
            calculator.calculate_student_scholarship(student_data, working_days)
            for _, student_data in df.groupby('זהות')
        ]

        results_df = pd.DataFrame(results)

        manual_columns = ['חבורה', 'מוסר', 'סיכומים', 'מבחן הלכה', 'מבחן שס']
        for col in manual_columns:
            results_df[col] = ''

        results_df['תוספת_דרגה_1'] = results_df.apply(
            lambda row: 190 if (
                    row['בוקר_missed_hours'] <= 7.0 and
                    row['צהריים_missed_hours'] <= 6.0 and
                    row['צהריים_attended_days'] > 0
            ) else 0,
            axis=1
        )
        results_df['סיבה - תוספת דרגה 1'] = ''

        results_df['תוספת_דרגה_2'] = results_df.apply(
            lambda row: 200 if (
                    row['בוקר_missed_hours'] <= 3.0 and
                    row['צהריים_missed_hours'] <= 2.0 and
                    row['צהריים_attended_days'] > 0
            ) else 0,
            axis=1
        )
        results_df['סיבה - תוספת דרגה 2'] = ''

        results_df['נוכחות_מושלמת'] = results_df.apply(
            lambda row: 200 if (
                    row['בוקר_absent_days'] == 0 and
                    row['בוקר_late_days'] == 0 and
                    row['צהריים_absent_days'] == 0 and
                    row['צהריים_late_days'] == 0 and
                    row['צהריים_attended_days'] > 0
            ) else 0,
            axis=1
        )
        results_df['סיבה - נוכחות מושלמת'] = ''

        results_df['הגעה_מוקדמת'] = results_df['בונוס_נוכחות_מוקדמת']
        results_df['סיבה - הגעה מוקדמת'] = ''

        results_df['סך סופי'] = 0

        results_df.to_excel(output_file, index=False)

        wb = load_workbook(output_file)
        ws = wb.active

        headers = [cell.value for cell in ws[1]]

        try:
            total_col = get_column_letter(headers.index('סך_הכל') + 1)

            chabura_col = get_column_letter(headers.index('חבורה') + 1)
            musar_col = get_column_letter(headers.index('מוסר') + 1)
            sikumim_col = get_column_letter(headers.index('סיכומים') + 1)
            halacha_col = get_column_letter(headers.index('מבחן הלכה') + 1)
            shas_col = get_column_letter(headers.index('מבחן שס') + 1)

            tier1_col = get_column_letter(headers.index('תוספת_דרגה_1') + 1)
            tier1_reason_col = get_column_letter(headers.index('סיבה - תוספת דרגה 1') + 1)

            tier2_col = get_column_letter(headers.index('תוספת_דרגה_2') + 1)
            tier2_reason_col = get_column_letter(headers.index('סיבה - תוספת דרגה 2') + 1)

            perfect_col = get_column_letter(headers.index('נוכחות_מושלמת') + 1)
            perfect_reason_col = get_column_letter(headers.index('סיבה - נוכחות מושלמת') + 1)

            early_col = get_column_letter(headers.index('הגעה_מוקדמת') + 1)
            early_reason_col = get_column_letter(headers.index('סיבה - הגעה מוקדמת') + 1)

            final_col = get_column_letter(headers.index('סך סופי') + 1)
        except ValueError as e:
            raise ValueError(f"לא נמצאה אחת העמודות הנדרשות: {str(e)}")

        for row in range(2, ws.max_row + 1):
            formula = (
                f'={total_col}{row}+'
                f'IF({chabura_col}{row}="V",150,0)+'
                f'IF({musar_col}{row}="V",100,0)+'
                f'IF({sikumim_col}{row}="V",150,0)+'
                f'IF({halacha_col}{row}="V",200,0)+'
                f'IF({shas_col}{row}="V",390,0)+'
                f'IF(AND({tier1_reason_col}{row}="V",{tier1_col}{row}=0),190,0)+'
                f'IF(AND({tier2_reason_col}{row}="V",{tier2_col}{row}=0),200,0)+'
                f'IF(AND({perfect_reason_col}{row}="V",{perfect_col}{row}=0),200,0)+'
                f'IF(AND({early_reason_col}{row}="V",{early_col}{row}=0),200,0)'
            )
            ws[f'{final_col}{row}'] = formula

        wb.save(output_file)
        return True

    except Exception as e:
        print(f"An error occurred: {e}")
        messagebox.showerror("שגיאה", f"אירעה שגיאה: {str(e)}")
        return False