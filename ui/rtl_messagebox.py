import customtkinter as ctk


class RTLMessageBox(ctk.CTkToplevel):
    """
    Custom message box with RTL (Right-to-Left) support for Hebrew text.
    
    Provides a modern-styled message dialog with proper Hebrew text alignment
    and direction.
    
    Attributes:
        title (str): Dialog window title
        message (str): Message content to display
        icon (str): Icon type - "check" for success, "cancel" for error, "warning" for warning
    """
    
    def __init__(self, parent=None, title="", message="", icon="check"):
        # If no parent provided, create a temporary root window
        if parent is None:
            parent = ctk.CTk()
            parent.withdraw()
            self._temp_parent = parent
        else:
            self._temp_parent = None
            
        super().__init__(parent)
        
        self.title(title)
        self.geometry("400x200")
        
        # Center the window
        window_width = 400
        window_height = 200
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        center_x = int(screen_width / 2 - window_width / 2)
        center_y = int(screen_height / 2 - window_height / 2)
        self.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
        
        # Icon color mapping
        icon_colors = {
            "check": "#2ecc71",
            "cancel": "#e74c3c",
            "warning": "#f39c12"
        }
        icon_symbols = {
            "check": "✓",
            "cancel": "✗",
            "warning": "⚠"
        }
        
        # Create main frame
        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Icon label
        icon_color = icon_colors.get(icon, "#2ecc71")
        icon_symbol = icon_symbols.get(icon, "✓")
        
        icon_label = ctk.CTkLabel(
            main_frame,
            text=icon_symbol,
            font=("Heebo", 48, "bold"),
            text_color=icon_color
        )
        icon_label.pack(pady=(10, 0))
        
        # Message label with RTL support
        message_label = ctk.CTkLabel(
            main_frame,
            text=message,
            font=("Heebo", 13),
            wraplength=350,
            justify="right"  # Right-align for Hebrew RTL
        )
        message_label.pack(pady=(10, 20))
        
        # OK button
        ok_button = ctk.CTkButton(
            main_frame,
            text="אישור",
            command=self.close,
            width=120,
            height=35,
            font=("Heebo", 12),
            fg_color=icon_color,
            hover_color=self._darken_color(icon_color)
        )
        ok_button.pack(pady=(0, 10))
        
        # Bind Enter key to close
        self.bind('<Return>', lambda e: self.close())
        self.bind('<Escape>', lambda e: self.close())
        
        # Make it modal
        self.transient(parent)
        self.grab_set()
        
        # Focus on OK button
        self.after(100, ok_button.focus_set)
        
    def _darken_color(self, hex_color):
        """Darken a hex color for hover effect."""
        hex_color = hex_color.lstrip('#')
        r, g, b = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        r = int(r * 0.8)
        g = int(g * 0.8)
        b = int(b * 0.8)
        return f'#{r:02x}{g:02x}{b:02x}'
    
    def close(self):
        """Close the dialog and cleanup."""
        self.grab_release()
        self.destroy()
        if self._temp_parent:
            self._temp_parent.destroy()


def show_message(title="", message="", icon="check"):
    """
    Show a modal message box with RTL support.
    
    Args:
        title (str): Dialog title
        message (str): Message content
        icon (str): Icon type - "check", "cancel", or "warning"
    """
    dialog = RTLMessageBox(None, title=title, message=message, icon=icon)
    dialog.wait_window()
