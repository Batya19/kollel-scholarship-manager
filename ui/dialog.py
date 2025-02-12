import customtkinter as ctk
from CTkMessagebox import CTkMessagebox


class ModernDialog(ctk.CTkToplevel):
    """
    Modal dialog for entering working days count.

    A modern-styled dialog window that prompts the user to enter the number
    of working days in the current month. Includes input validation and
    error handling.

    Attributes:
        result (int): The entered number of working days (None if canceled)
    """
    def __init__(self, parent):
        super().__init__(parent)

        self.title("ימי עבודה")
        self.geometry("300x200")
        self.result = None

        window_width = 300
        window_height = 200
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        center_x = int(screen_width / 2 - window_width / 2)
        center_y = int(screen_height / 2 - window_height / 2)
        self.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

        self.label = ctk.CTkLabel(
            self,
            text="כמה ימי עבודה יש בחודש זה?",
            font=("Heebo", 14)
        )
        self.label.pack(pady=20)

        self.entry = ctk.CTkEntry(
            self,
            placeholder_text="הזן מספר ימים...",
            width=200,
            height=35,
            font=("Heebo", 12)
        )
        self.entry.pack(pady=10)

        self.button = ctk.CTkButton(
            self,
            text="אישור",
            command=self.ok_click,
            width=120,
            height=35,
            font=("Heebo", 12),
            fg_color="#2ecc71",
            hover_color="#27ae60"
        )
        self.button.pack(pady=20)

        self.entry.focus()

        self.bind('<Return>', lambda e: self.ok_click())

        self.transient(parent)
        self.grab_set()

    def ok_click(self):
        try:
            value = int(self.entry.get())
            if 1 <= value <= 31:
                self.result = value
                self.destroy()
            else:
                CTkMessagebox(
                    title="שגיאה",
                    message="נא להזין מספר בין 1 ל-31",
                    icon="warning"
                )
        except ValueError:
            CTkMessagebox(
                title="שגיאה",
                message="נא להזין מספר תקין",
                icon="warning"
            )
