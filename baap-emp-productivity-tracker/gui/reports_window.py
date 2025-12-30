import tkinter as tk
from tkinter import ttk
from gui.theme import apply_theme, spacing
import threading

# Global tkinter root window for creating Toplevel windows
_tk_root = None
_tk_root_lock = threading.Lock()

def _get_tk_root():
    """Get or create the hidden tkinter root window"""
    global _tk_root
    with _tk_root_lock:
        if _tk_root is None:
            _tk_root = tk.Tk()
            _tk_root.withdraw()  # Hide the root window
            
            # Run tkinter event loop in a separate thread
            # This works on Windows where tkinter can run in non-main threads
            def run_tkinter():
                try:
                    _tk_root.mainloop()
                except Exception:
                    # If mainloop fails, use update loop instead
                    import time
                    while True:
                        try:
                            _tk_root.update()
                            time.sleep(0.01)
                        except Exception:
                            break
            
            thread = threading.Thread(target=run_tkinter, daemon=True)
            thread.start()
        return _tk_root

class ReportsWindow:
    def __init__(self, attendance_api, session_manager):
        self.attendance_api = attendance_api
        self.session_manager = session_manager

    def _render_summary_grid(self, parent, summary: dict):
        grid = ttk.Frame(parent)
        grid.pack(fill=tk.X)
        for i, (key, value) in enumerate(summary.items()):
            ttk.Label(grid, text=str(key)).grid(row=i, column=0, sticky=tk.W, padx=(0, spacing["lg"]), pady=(0, spacing["sm"]))
            ttk.Label(grid, text=str(value), style="Subheader.TLabel").grid(row=i, column=1, sticky=tk.W, pady=(0, spacing["sm"]))
        return grid

    def _create_window(self):
        """Create the window - must be called from tkinter thread"""
        tk_root = _get_tk_root()
        window = tk.Toplevel(tk_root)
        window.title("Attendance Reports")

        apply_theme(window)

        root = ttk.Frame(window, padding=(spacing["xl"], spacing["lg"]))
        root.pack(fill=tk.BOTH, expand=True)

        notebook = ttk.Notebook(root)
        notebook.pack(fill=tk.BOTH, expand=True)

        # Daily (uses local session aggregation for accuracy in current run)
        daily_tab = ttk.Frame(notebook)
        notebook.add(daily_tab, text="Daily")
        from datetime import date as _date
        _today_str = _date.today().strftime("%A, %Y-%m-%d")
        ttk.Label(daily_tab, text=f"Today ({_today_str})", style="Header.TLabel").pack(anchor=tk.W, pady=(0, spacing["md"]))
        # Table view for today's Active/Break/Idle
        daily_table = ttk.Treeview(daily_tab, columns=("active", "break", "idle", "activity"), show="headings", height=1)
        daily_table.heading("active", text="Active")
        daily_table.heading("break", text="Break")
        daily_table.heading("idle", text="Idle")
        daily_table.heading("activity", text="Activity % (last min)")
        daily_table.column("active", width=120, anchor=tk.CENTER)
        daily_table.column("break", width=120, anchor=tk.CENTER)
        daily_table.column("idle", width=120, anchor=tk.CENTER)
        daily_table.column("activity", width=170, anchor=tk.CENTER)
        daily_table.pack(fill=tk.X)
        daily_summary = self.session_manager.get_daily_summary()
        activity_pct = getattr(self.session_manager, "last_min_activity_percent", 0.0)
        daily_table.insert("", tk.END, values=(
            daily_summary.get("total_active_time"),
            daily_summary.get("total_break_time"),
            daily_summary.get("total_idle_time"),
            f"{activity_pct:.1f}%",
        ))

        # Weekly
        weekly_tab = ttk.Frame(notebook)
        notebook.add(weekly_tab, text="Weekly")
        ttk.Label(weekly_tab, text="This Week", style="Header.TLabel").pack(anchor=tk.W, pady=(0, spacing["md"]))
        weekly_summary = self.attendance_api.get_weekly_summary()
        self._render_summary_grid(weekly_tab, weekly_summary)

        # Current Month
        month_tab = ttk.Frame(notebook)
        notebook.add(month_tab, text="Current Month")
        ttk.Label(month_tab, text="Current Month", style="Header.TLabel").pack(anchor=tk.W, pady=(0, spacing["md"]))
        monthly_summary = self.attendance_api.get_monthly_summary()
        self._render_summary_grid(month_tab, monthly_summary)

        # All-time
        all_time_tab = ttk.Frame(notebook)
        notebook.add(all_time_tab, text="All-time")
        ttk.Label(all_time_tab, text="All-time", style="Header.TLabel").pack(anchor=tk.W, pady=(0, spacing["md"]))
        all_time_summary = self.attendance_api.get_all_time_summary()
        self._render_summary_grid(all_time_tab, all_time_summary)

        # Full History
        history_tab = ttk.Frame(notebook)
        notebook.add(history_tab, text="History")
        ttk.Label(history_tab, text="Full History", style="Header.TLabel").pack(anchor=tk.W, pady=(0, spacing["md"]))
        history = self.attendance_api.get_history()
        # Prepend today's in-progress summary
        try:
            from datetime import date
            today = date.today().isoformat()
            today_summary = self.session_manager.get_daily_summary()
            today_row = {
                "date": today,
                "active": today_summary.get("total_active_time"),
                "break": today_summary.get("total_break_time"),
                "idle": today_summary.get("total_idle_time"),
            }
            # remove existing record for today if present, then insert at start
            history = [r for r in history if r.get("date") != today]
            history.insert(0, today_row)
        except Exception:
            pass
        tree = ttk.Treeview(history_tab, columns=("day", "date", "active", "break", "idle", "laptop_sleep", "activity"), show="headings")
        tree.heading("day", text="Day")
        tree.heading("date", text="Date")
        tree.heading("active", text="Active")
        tree.heading("break", text="Break")
        tree.heading("idle", text="Idle")
        tree.heading("laptop_sleep", text="Laptop Sleep")
        tree.heading("activity", text="Activity %")
        tree.column("day", width=110)
        tree.column("date", width=120)
        tree.column("active", width=100)
        tree.column("break", width=100)
        tree.column("idle", width=100)
        tree.column("laptop_sleep", width=120)
        tree.column("activity", width=100)
        tree.pack(fill=tk.BOTH, expand=True)
        from datetime import datetime as _dt
        for row in history:
            _date_str = row.get("date")
            _day = ""
            try:
                _day = _dt.fromisoformat(_date_str).strftime("%A") if _date_str else ""
            except Exception:
                _day = ""
            tree.insert("", tk.END, values=(
                _day,
                row.get("date"),
                row.get("active"),
                row.get("break"),
                row.get("idle"),
                row.get("laptop_sleep", "00:00:00"),
                row.get("activity", "")
            ))

        ttk.Frame(root, height=spacing["lg"]).pack()
        ttk.Button(root, text="Close", command=window.destroy).pack(anchor=tk.E)

    def show(self):
        """Show the window - schedules creation on tkinter thread"""
        tk_root = _get_tk_root()
        # Schedule window creation on tkinter thread
        tk_root.after(0, self._create_window)
