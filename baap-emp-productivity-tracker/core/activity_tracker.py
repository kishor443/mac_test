import time
import threading
from pynput import mouse, keyboard
from datetime import datetime

from utils.excel_storage import append_activity_event


def log_activity_to_excel(user_id, activity_status, active_window, timestamp=None, extra_details=None, session_id=None, client_id=None, webcam_name=None, mouse_clicks=0, keys_count=0):
    """
    Record an activity entry inside the Excel workbook.
    """
    ts = timestamp or datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    metadata = extra_details.copy() if isinstance(extra_details, dict) else {}
    
    # Don't set end_time here - it will be set when next activity starts
    # Duration will be calculated automatically
    append_activity_event(
        {
            "client_id": client_id or session_id or "",
            "user_id": user_id or "unknown",
            "tool": metadata.pop("tool", active_window or ""),
            "action": metadata.pop("action", activity_status),
            "duration": "",  # Will be calculated automatically
            "start_time": ts,
            "end_time": "",  # Will be set when next activity starts
            "task_id": metadata.pop("task_id", ""),
            "project_id": metadata.pop("project_id", ""),
            "activity_type": activity_status,
            "title": metadata.pop("title", active_window or ""),
            "metadata_json": metadata,
            "screenshots": metadata.pop("screenshots", ""),
            "webcam_photo": metadata.pop("webcam_photo", ""),
            "webcam_name": webcam_name or metadata.pop("webcam_name", ""),
            "mouse_clicks": mouse_clicks,
            "keys_count": keys_count,
        }
    )

class ActivityTracker:
    def __init__(self, on_activity=None, user_id='unknown', session_id=None, client_id=None, webcam_name=None):
        self.last_activity_time = time.time()
        self.on_activity = on_activity
        self._mouse_listener = None
        self._keyboard_listener = None
        self.user_id = user_id
        self.session_id = session_id
        self.client_id = client_id
        self.webcam_name = webcam_name or ""
        self._last_excel_log_time = 0
        self._excel_log_interval = 60.0  # Log to Excel only every 60 seconds max (reduced for performance)
        
        # Track mouse clicks and key counts
        self.mouse_clicks_today = 0
        self.keys_pressed_today = 0
        self._last_reset_date = datetime.now().date()

    def update_context(self, user_id=None, client_id=None, session_id=None, webcam_name=None):
        if user_id:
            self.user_id = user_id
        if client_id:
            self.client_id = client_id
        if session_id:
            self.session_id = session_id
        if webcam_name:
            self.webcam_name = webcam_name

    def start(self):
        # Keep references to listeners to prevent GC from stopping them
        self._mouse_listener = mouse.Listener(on_move=self._on_activity, on_click=self._on_mouse_click)
        self._keyboard_listener = keyboard.Listener(on_press=self._on_key_press, on_release=self._on_activity)
        self._mouse_listener.start()
        self._keyboard_listener.start()

    def stop(self):
        """Stop listeners safely (used on sleep)."""
        try:
            if self._mouse_listener:
                self._mouse_listener.stop()
        except Exception:
            pass
        try:
            if self._keyboard_listener:
                self._keyboard_listener.stop()
        except Exception:
            pass
        self._mouse_listener = None
        self._keyboard_listener = None

    def _reset_daily_counts_if_needed(self):
        """Reset daily counts if it's a new day"""
        current_date = datetime.now().date()
        if current_date != self._last_reset_date:
            self.mouse_clicks_today = 0
            self.keys_pressed_today = 0
            self._last_reset_date = current_date

    def _on_mouse_click(self, x, y, button, pressed):
        """Handle mouse click events"""
        if pressed:  # Only count on press, not release
            self._reset_daily_counts_if_needed()
            self.mouse_clicks_today += 1
        self._on_activity(x, y, button, pressed)

    def _on_key_press(self, key):
        """Handle key press events"""
        self._reset_daily_counts_if_needed()
        self.keys_pressed_today += 1
        self._on_activity(key)

    def _on_activity(self, *args, **kwargs):
        self.last_activity_time = time.time()
        current_time = time.time()
        
        # Throttle Excel logging - only log every 30 seconds to prevent hang
        should_log_to_excel = (current_time - self._last_excel_log_time) >= self._excel_log_interval
        
        if should_log_to_excel:
            try:
                from utils.active_window import get_active_window_title
                window_title = get_active_window_title()
            except Exception:
                window_title = ''
            try:
                log_activity_to_excel(
                    user_id=self.user_id,
                    activity_status='active',
                    active_window=window_title,
                    extra_details={'session_id': self.session_id} if self.session_id else None,
                    session_id=self.session_id,
                    client_id=self.client_id,
                    webcam_name=self.webcam_name,
                    mouse_clicks=self.mouse_clicks_today,
                    keys_count=self.keys_pressed_today
                )
                self._last_excel_log_time = current_time
            except Exception:
                # Don't crash on Excel write errors - just continue
                pass
        
        # Always call on_activity callback (for UI updates)
        if self.on_activity:
            try:
                self.on_activity()
            except Exception:
                pass
    
    def get_activity_counts(self):
        """Get mouse clicks and key counts for today"""
        self._reset_daily_counts_if_needed()
        return {
            "mouse_clicks": self.mouse_clicks_today,
            "keys_count": self.keys_pressed_today
        }

    def get_idle_time(self):
        return time.time() - self.last_activity_time

    def get_status(self, idle_threshold_seconds: int) -> str:
        idle_seconds = self.get_idle_time()
        return "idle" if idle_seconds >= idle_threshold_seconds else "active"
