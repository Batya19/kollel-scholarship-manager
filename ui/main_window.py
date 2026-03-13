import customtkinter as ctk
from tkinter import filedialog
from .dialog import ModernDialog
from .rtl_messagebox import RTLMessageBox
from utils.excel_processor import process_kollel_attendance
from utils.date_utils import get_hebrew_month_name
import pandas as pd


class ModernKollelUI:
    """
    Main application window for the Kollel Scholarship Manager.
    """

    def __init__(self):
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("green")

        self.root = ctk.CTk()
        self.root.configure(fg_color="#ffffff")
        self.root.title("חישוב מלגות כולל")

        import os, sys
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.join(os.path.dirname(__file__), '..')
        icon_path = os.path.join(base_dir, 'icon.ico')
        if os.path.exists(icon_path):
            self.root.iconbitmap(icon_path)

        window_width = 500
        window_height = 420
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        center_x = int(screen_width / 2 - window_width / 2)
        center_y = int(screen_height / 2 - window_height / 2)
        self.root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

        # Icon
        import tkinter as tk
        canvas = tk.Canvas(self.root, width=72, height=72,
                           bg="#ffffff", highlightthickness=0)
        canvas.pack(pady=(40, 0))
        canvas.create_rectangle(4,  48, 20, 68, fill="#2ecc71", outline="")
        canvas.create_rectangle(26, 34, 42, 68, fill="#2ecc71", outline="")
        canvas.create_rectangle(48, 22, 64, 68, fill="#27ae60", outline="")

        self.title_label = ctk.CTkLabel(
            self.root,
            text="חישוב מלגות כולל",
            font=("Heebo", 26, "bold"),
            text_color="#1a1a1a"
        )
        self.title_label.pack(pady=(14, 32))

        self.frame = ctk.CTkFrame(self.root, fg_color="transparent")
        self.frame.pack(padx=60, fill="both", expand=True)


        self.btn_open = ctk.CTkButton(
            self.frame,
            text="בחר קובץ לחישוב",
            command=self.open_file,
            width=280,
            height=50,
            corner_radius=12,
            font=("Heebo", 15),
            fg_color="#2ecc71",
            hover_color="#27ae60",
            text_color="#ffffff"
        )
        self.btn_open.pack(pady=(0, 14))

        self.btn_exit = ctk.CTkButton(
            self.frame,
            text="יציאה",
            command=self.root.quit,
            width=280,
            height=50,
            corner_radius=12,
            font=("Heebo", 14),
            fg_color="transparent",
            border_width=1,
            border_color="#dddddd",
            text_color="#aaaaaa",
            hover_color="#f8f8f8"
        )
        self.btn_exit.pack()

        self.version_label = ctk.CTkLabel(
            self.root,
            text="v1.3",
            font=("Heebo", 10),
            text_color="#cccccc"
        )
        self.version_label.pack(pady=(12, 16))

    def _get_output_filename(self, input_file: str) -> str:
        try:
            df = pd.read_excel(input_file)
            df['כניסה'] = pd.to_datetime(df['כניסה'])
            first_date = df['כניסה'].iloc[0]
            month_number = first_date.month
            year = first_date.year
            month_name = get_hebrew_month_name(month_number)
            return f"מלגות חודשיות {month_name} {year}.xlsx"
        except Exception:
            return "מלגות חודשיות.xlsx"

    def open_file(self):
        filepath = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx")],
            title="בחר קובץ Excel"
        )

        if filepath:
            dialog = ModernDialog(self.root)
            self.root.wait_window(dialog)

            if dialog.result is not None:
                working_days = dialog.result
                output_filename = self._get_output_filename(filepath)

                try:
                    if process_kollel_attendance(filepath, output_filename, working_days):
                        msg = RTLMessageBox(
                            parent=self.root,
                            title="הצלחה",
                            message=f"החישוב הסתיים!\nהקובץ '{output_filename}' נשמר.",
                            icon="check"
                        )
                        self.root.wait_window(msg)
                except Exception as e:
                    msg = RTLMessageBox(
                        parent=self.root,
                        title="שגיאה",
                        message=f"אירעה שגיאה:\n{str(e)}",
                        icon="cancel"
                    )
                    self.root.wait_window(msg)

    def run(self):
        self.root.mainloop()