import tkinter as tk
from tkinter import ttk
import threading

class SplashScreen:
    def __init__(self):
        self.root = tk.Tk()
        self.root.overrideredirect(True)
        self.root.attributes('-topmost', True)
        self.root.configure(bg="#ffffff")

        width, height = 400, 260
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        self.root.geometry(f"{width}x{height}+{(sw-width)//2}+{(sh-height)//2}")

        # מסגרת עם גבול עדין
        outer = tk.Frame(self.root, bg="#cccccc")
        outer.pack(fill="both", expand=True, padx=1, pady=1)
        inner = tk.Frame(outer, bg="#ffffff")
        inner.pack(fill="both", expand=True)

        # אייקון גרף — מלבנים ירוקים
        canvas = tk.Canvas(inner, width=56, height=56,
                           bg="#ffffff", highlightthickness=0)
        canvas.pack(pady=(28, 0))
        canvas.create_rectangle(4,  38, 16, 52, fill="#2ecc71", outline="")
        canvas.create_rectangle(20, 28, 32, 52, fill="#2ecc71", outline="")
        canvas.create_rectangle(36, 20, 48, 52, fill="#27ae60", outline="")

        tk.Label(inner, text="חישוב מלגות כולל",
                 font=("Segoe UI", 16, "bold"),
                 bg="#ffffff", fg="#1a1a1a").pack(pady=(10, 2))

        tk.Label(inner, text="מערכת ניהול מלגות כולל",
                 font=("Segoe UI", 10),
                 bg="#ffffff", fg="#aaaaaa").pack()

        # Progress bar
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("G.Horizontal.TProgressbar",
                        troughcolor='#f0f0f0',
                        background='#2ecc71',
                        thickness=6,
                        borderwidth=0)

        bar_frame = tk.Frame(inner, bg="#ffffff")
        bar_frame.pack(fill="x", padx=32, pady=(18, 6))
        self.progress = ttk.Progressbar(
            bar_frame, style="G.Horizontal.TProgressbar",
            length=336, mode='determinate')
        self.progress.pack()

        self.status_label = tk.Label(inner, text="מאתחל...",
                 font=("Segoe UI", 9),
                 bg="#ffffff", fg="#bbbbbb")
        self.status_label.pack()

    def _update(self, value, text):
        self.progress['value'] = value
        self.status_label.config(text=text)
        self.root.update_idletasks()

    def run_with_app(self, app_func):
        steps = [
            (0,   "מאתחל..."),
            (25,  "טוען ספריות..."),
            (55,  "מכין ממשק..."),
            (85,  "כמעט מוכן..."),
            (100, "מוכן!"),
        ]
        def load():
            import time
            for val, text in steps:
                self.root.after(0, self._update, val, text)
                time.sleep(0.45)
            time.sleep(0.25)
            self.root.after(0, self._finish, app_func)

        threading.Thread(target=load, daemon=True).start()
        self.root.mainloop()

    def _finish(self, app_func):
        self.root.destroy()
        app_func()