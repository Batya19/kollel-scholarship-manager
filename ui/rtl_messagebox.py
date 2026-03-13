import customtkinter as ctk


class RTLMessageBox(ctk.CTkToplevel):
    """
    Custom message box with RTL support for Hebrew text.
    """

    def __init__(self, parent=None, title="", message="", icon="check"):
        if parent is None:
            parent = ctk.CTk()
            parent.withdraw()
            self._temp_parent = parent
        else:
            self._temp_parent = None

        super().__init__(parent)
        self.configure(fg_color="#ffffff")
        self.title(title)

        window_width = 480
        window_height = 300
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        center_x = int(screen_width / 2 - window_width / 2)
        center_y = int(screen_height / 2 - window_height / 2)
        self.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

        icon_colors = {
            "check":   "#2ecc71",
            "cancel":  "#e74c3c",
            "warning": "#f39c12"
        }
        icon_symbols = {
            "check":   "✓",
            "cancel":  "✗",
            "warning": "⚠"
        }

        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=36, pady=32)

        icon_color = icon_colors.get(icon, "#2ecc71")
        icon_symbol = icon_symbols.get(icon, "✓")

        icon_label = ctk.CTkLabel(
            main_frame,
            text=icon_symbol,
            font=("Heebo", 40, "bold"),
            text_color=icon_color
        )
        icon_label.pack(pady=(0, 8))

        message_label = ctk.CTkLabel(
            main_frame,
            text=message,
            font=("Heebo", 14),
            text_color="#333333",
            wraplength=380,
            justify="right"
        )
        message_label.pack(pady=(0, 24))

        ok_button = ctk.CTkButton(
            main_frame,
            text="אישור",
            command=self.close,
            width=130,
            height=42,
            corner_radius=8,
            font=("Heebo", 14),
            fg_color=icon_color,
            hover_color=self._darken_color(icon_color),
            text_color="#ffffff"
        )
        ok_button.pack()

        self.bind('<Return>', lambda e: self.close())
        self.bind('<Escape>', lambda e: self.close())
        self.transient(parent)
        self.grab_set()
        self.after(100, ok_button.focus_set)

    def _darken_color(self, hex_color):
        hex_color = hex_color.lstrip('#')
        r, g, b = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        return f'#{int(r*0.8):02x}{int(g*0.8):02x}{int(b*0.8):02x}'

    def close(self):
        self.grab_release()
        self.destroy()
        if self._temp_parent:
            self._temp_parent.destroy()


def show_message(title="", message="", icon="check"):
    dialog = RTLMessageBox(None, title=title, message=message, icon=icon)
    dialog.wait_window()