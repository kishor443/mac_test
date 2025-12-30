import tkinter as tk
from tkinter import ttk
from gui.theme import apply_theme, spacing

class StatusPopup:
    def __init__(self, session_manager):
        self.session_manager = session_manager

    def show(self):
        window = tk.Toplevel()
        window.title("Status")

        apply_theme(window)
        root = ttk.Frame(window, padding=(spacing["lg"], spacing["md"]))
        root.pack(fill=tk.BOTH, expand=True)

        summary = self.session_manager.get_daily_summary()
        ttk.Label(root, text="Productivity Status", style="Header.TLabel").pack(anchor=tk.W, pady=(0, spacing["sm"]))

        row = ttk.Frame(root)
        row.pack(fill=tk.X)
        ttk.Label(row, text="Active:").pack(side=tk.LEFT)
        ttk.Label(row, text=summary["total_active_time"], style="Subheader.TLabel").pack(side=tk.LEFT, padx=(spacing["sm"], spacing["lg"]))
        ttk.Label(row, text="Break:").pack(side=tk.LEFT)
        ttk.Label(row, text=summary["total_break_time"], style="Subheader.TLabel").pack(side=tk.LEFT, padx=(spacing["sm"], spacing["lg"]))
        ttk.Label(row, text="Idle:").pack(side=tk.LEFT)
        ttk.Label(row, text=summary["total_idle_time"], style="Subheader.TLabel").pack(side=tk.LEFT)

        window.after(8000, window.destroy)
