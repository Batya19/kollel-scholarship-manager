import customtkinter as ctk
from .rtl_messagebox import RTLMessageBox


class ModernDialog(ctk.CTkToplevel):
    """
    Modal dialog for entering working days count.
    """
    def __init__(self, parent):
        super().__init__(parent)
        self.configure(fg_color="#ffffff")
        self.title("ימי עבודה")
        self.result = None

        window_width = 320
        window_height = 220
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        center_x = int(screen_width / 2 - window_width / 2)
        center_y = int(screen_height / 2 - window_height / 2)
        self.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

        self.label = ctk.CTkLabel(
            self,
            text="כמה ימי עבודה יש בחודש זה?",
            font=("Heebo", 14),
            text_color="#1a1a1a"
        )
        self.label.pack(pady=(28, 12))

        self.entry = ctk.CTkEntry(
            self,
            placeholder_text="הזן מספר ימים...",
            width=200,
            height=38,
            font=("Heebo", 13),
            corner_radius=8,
            border_color="#dddddd",
            fg_color="#ffffff",
            text_color="#1a1a1a"
        )
        self.entry.pack(pady=(0, 16))

        self.button = ctk.CTkButton(
            self,
            text="אישור",
            command=self.ok_click,
            width=120,
            height=38,
            corner_radius=8,
            font=("Heebo", 13),
            fg_color="#2ecc71",
            hover_color="#27ae60",
            text_color="#ffffff"
        )
        self.button.pack()

        self.bind('<Return>', lambda e: self.ok_click())
        self.transient(parent)
        self.grab_set()
        self.after(100, lambda: self.entry.focus_force())

    def ok_click(self):
        try:
            value = int(self.entry.get())
            if 1 <= value <= 31:
                self.result = value
                self.destroy()
            else:
                msg = RTLMessageBox(
                    parent=self,
                    title="שגיאה",
                    message="נא להזין מספר בין 1 ל-31",
                    icon="warning"
                )
                self.wait_window(msg)
        except ValueError:
            msg = RTLMessageBox(
                parent=self,
                title="שגיאה",
                message="נא להזין מספר תקין",
                icon="warning"
            )
            self.wait_window(msg)