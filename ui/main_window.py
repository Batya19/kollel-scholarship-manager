import customtkinter as ctk
from tkinter import filedialog
from CTkMessagebox import CTkMessagebox
from .dialog import ModernDialog
from utils.excel_processor import process_kollel_attendance


class ModernKollelUI:
    """
    Main application window for the Kollel Scholarship Manager.

    Provides a modern user interface for selecting input files and processing
    scholarship calculations. Handles file selection, user input, and displays
    processing results.

    Methods:
        run(): Start the application
        open_file(): Handle file selection and processing
    """

    def __init__(self):
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("green")

        self.root = ctk.CTk()
        self.root.title("חישוב מלגות כולל")
        self.root.geometry("500x400")

        window_width = 500
        window_height = 400
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        center_x = int(screen_width / 2 - window_width / 2)
        center_y = int(screen_height / 2 - window_height / 2)
        self.root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

        self.title_label = ctk.CTkLabel(
            self.root,
            text="📊 חישוב מלגות כולל",
            font=("Heebo", 24, "bold")
        )
        self.title_label.pack(pady=30)

        self.frame = ctk.CTkFrame(self.root, fg_color="transparent")
        self.frame.pack(pady=20, padx=40, fill="both", expand=True)

        self.btn_open = ctk.CTkButton(
            self.frame,
            text="📁 בחר קובץ לחישוב",
            command=self.open_file,
            width=300,
            height=50,
            font=("Heebo", 14),
            fg_color="#2ecc71",
            hover_color="#27ae60"
        )
        self.btn_open.pack(pady=20)

        self.btn_exit = ctk.CTkButton(
            self.frame,
            text="❌ יציאה",
            command=self.root.quit,
            width=300,
            height=50,
            font=("Heebo", 14),
            fg_color="#e74c3c",
            hover_color="#c0392b"
        )
        self.btn_exit.pack(pady=20)

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
                try:
                    if process_kollel_attendance(filepath, "מלגות חודשיות.xlsx", working_days):
                        CTkMessagebox(
                            title="הצלחה",
                            message="החישוב הסתיים! הקובץ 'מלגות חודשיות.xlsx' נשמר.",
                            icon="check",
                            fade_in_duration=1
                        )
                except Exception as e:
                    CTkMessagebox(
                        title="שגיאה",
                        message=f"אירעה שגיאה: {str(e)}",
                        icon="cancel"
                    )

    def run(self):
        self.root.mainloop()
