from datetime import datetime, timedelta, timezone
import threading
import time
from collections import deque
from enum import Enum
from pathlib import Path
from config import (
    IDLE_TIMEOUT_SECONDS,
    ACTIVITY_SCORE_WINDOW_SECONDS,
    SCREENSHOT_INTERVAL_SECONDS,
    SCREENSHOTS_DIR,
    WEBCAM_PHOTOS_DIR,
    WEBCAM_DEVICE_NAME,
    MAX_BROWSER_TABS_CAPTURED,
)
from utils.active_window import get_active_window_title
from utils.screen_capture import capture_screenshot
from utils.webcam_capture import capture_webcam_photo
from utils.browser_tabs import collect_browser_tabs
from utils.logger import logger
from utils.excel_storage import (
    append_activity_event,
    finalize_last_activity_row,
    make_hyperlink,
    read_local_storage,
    write_local_storage,
)
from utils.capture_types import CaptureArtifact

class SessionState(Enum):
    LOGGED_OUT = "logged_out"
    CLOCKED_IN = "clocked_in"
    ON_BREAK = "on_break"
    IDLE = "idle"

class SessionManager:
    def __init__(self, attendance_api):
        self.attendance_api = attendance_api
        self.state = SessionState.LOGGED_OUT
        self.session_start = None
        self.break_start = None
        self.last_update = None
        self.is_idle = False
        
        self.current_idle_seconds = 0
        self.active_seconds = 0
        self.break_seconds = 0
        self.idle_seconds = 0
        # Activity score (last N seconds of active samples)
        self._activity_samples = deque(maxlen=ACTIVITY_SCORE_WINDOW_SECONDS)
        self.last_min_activity_percent = 0.0
        # Background capture
        self._bg_running = False
        self._screenshot_thread = None
        self._window_log_thread = None
        self._recent_windows = []  # capped list of recent window titles
        self._captured_screenshots = []  # filenames of captures during session
        # App usage tracking
        self._app_usage = {}  # {app_name: seconds}
        self._current_app = None
        self._current_app_start_time = None
        # Power/suspend detection
        self._suspend_monitor_thread = None
        self._last_tick = None
        self._sleep_break_active = False  # Track if break was started due to sleep
        # Worklog snapshots
        self._latest_tabs_snapshot = None
        self._last_screenshot_path = None
        self._last_webcam_path = None
        self._last_capture_time = None
        self._worklog_lock = threading.Lock()
        self._clear_erp_overrides()
        self._sleep_gap_prehandled = False
        # Track whether ERP currently has an open break for this shift
        self._erp_break_open = False

    def _clear_erp_overrides(self):
        self._erp_in_time = None
        self._erp_out_time = None
        self._erp_break_seconds = None
        self._erp_active_seconds = None
        self._erp_state_hint = None
        self._erp_override_timestamp = 0.0
        self._erp_break_open = False

    def clock_in(self):
        from datetime import timezone
        self.state = SessionState.CLOCKED_IN
        self.session_start = datetime.now(timezone.utc)
        self.last_update = self.session_start
        self._clear_erp_overrides()
        self.break_start = None
        self.is_idle = False
        self.current_idle_seconds = 0
        self.active_seconds = 0
        self.break_seconds = 0
        self.idle_seconds = 0
        # Don't reset app usage tracking on clock in - keep daily totals
        # Only initialize if not exists
        if not hasattr(self, '_app_usage'):
            self._app_usage = {}
        if not hasattr(self, '_current_app'):
            self._current_app = None
        if not hasattr(self, '_current_app_start_time'):
            self._current_app_start_time = None
        try:
            date_iso = self.session_start.date().isoformat()
            in_time_iso = self.session_start.replace(microsecond=0).isoformat().replace("+00:00", "Z")
            self.attendance_api.punch_in(
                client_id=self.attendance_api.client_id,
                date_in_iso_format=date_iso,
                in_time_iso_utc=in_time_iso,
                shift_id=None,
                latitude=None,
                longitude=None,
                status="W",
            )
        except Exception:
            pass
        try:
            logger.info("SessionManager: clocked in")
        except Exception:
            pass
        # start background tasks
        self._start_background_tasks()
        # --- TEST PATCH (remove after confirming sleep shows up in GUI) ---
        try:
            from datetime import timedelta

            data = read_local_storage()
            sleep_events = data.get("sleep_events", [])
            if not sleep_events:
                now = datetime.now()
                sleep_event = {
                    "date": now.date().isoformat(),
                    "sleep_start": (now.replace(microsecond=0) - timedelta(minutes=5)).isoformat(),
                    "wake_time": now.replace(microsecond=0).isoformat(),
                    "duration_seconds": 300,
                }
                sleep_events.append(sleep_event)
                data["sleep_events"] = sleep_events
                write_local_storage(data)
        except Exception:
            pass
        # --- END PATCH ---

    def clock_in_local(self):
        # Same as clock_in but without sending a punch_in to the server.
        from datetime import timezone
        self.state = SessionState.CLOCKED_IN
        self.session_start = datetime.now(timezone.utc)
        self.last_update = self.session_start
        self._clear_erp_overrides()
        self.break_start = None
        self.is_idle = False
        self.current_idle_seconds = 0
        self.active_seconds = 0
        self.break_seconds = 0
        self.idle_seconds = 0
        try:
            logger.info("SessionManager: local clock in (server already punched)")
        except Exception:
            pass
        self._start_background_tasks()

    def clock_out(self, reason="manual"):
        self._accumulate_until_now()
        # Persist a history record for today
        try:
            summary = self.get_daily_summary()
            record = {
                "date": datetime.now().date().isoformat(),
                "active": summary.get("total_active_time", "00:00:00"),
                "break": summary.get("total_break_time", "00:00:00"),
                "idle": summary.get("total_idle_time", "00:00:00"),
                "activity_percent": round(self.last_min_activity_percent, 1),
                "top_windows": self._recent_windows[-10:],
                "screenshots": list(self._captured_screenshots[-10:]),
            }
            data = read_local_storage()
            history = data.get("history", [])
            history.append(record)
            data["history"] = history
            write_local_storage(data)
        except Exception:
            pass
        # Accumulate final app usage time before stopping
        import time
        if self._current_app and self._current_app_start_time:
            current_time = time.time()
            elapsed = current_time - self._current_app_start_time
            if self._current_app not in self._app_usage:
                self._app_usage[self._current_app] = 0
            self._app_usage[self._current_app] += elapsed
            self._current_app = None
            self._current_app_start_time = None
        # stop background tasks
        self._stop_background_tasks()
        try:
            finalize_last_activity_row()
        except Exception:
            pass
        self.state = SessionState.LOGGED_OUT
        try:
            from datetime import timezone as _tz
            now_utc = datetime.now(_tz.utc)
            date_iso = now_utc.date().isoformat()
            out_time_iso = now_utc.replace(microsecond=0).isoformat().replace("+00:00", "Z")
            return self.attendance_api.punch_out(date_in_iso_format=date_iso, out_time_iso_utc=out_time_iso, status="W")
        except Exception as e:
            return False, {"error": str(e)}, 0
        try:
            logger.info(f"SessionManager: clocked out (reason={reason})")
        except Exception:
            pass

    def get_daily_summary(self):
        self._accumulate_until_now()
        break_total = max(0, int(self.break_seconds))
        if self._erp_break_seconds is not None:
            break_total = max(0, int(self._erp_break_seconds))
        idle_total = max(0, int(self.idle_seconds))
        active_total = int(self.active_seconds)

        # Prefer ERP-provided active seconds if present
        if self._erp_active_seconds is not None:
            active_total = max(0, int(self._erp_active_seconds))
        else:
            # When clocked in, show total shift duration minus breaks rather than raw input time
            start_dt = self._erp_in_time or self.session_start
            if start_dt:
                now = datetime.now(timezone.utc)
                # If already clocked out, freeze the display at the last update moment
                if self._erp_out_time:
                    effective_end = self._erp_out_time
                elif self.state == SessionState.LOGGED_OUT and self.last_update:
                    effective_end = self.last_update
                else:
                    effective_end = now
                if effective_end and effective_end < start_dt:
                    effective_end = now
                total_elapsed = max(0, int((effective_end - start_dt).total_seconds()))
                active_total = max(0, total_elapsed - break_total)
        return {
            "total_active_time": self._format_hms(active_total),
            "total_break_time": self._format_hms(break_total),
            "total_idle_time": self._format_hms(idle_total),
        }

    def get_app_usage_stats(self):
        """Get app usage statistics with formatted time"""
        try:
            import time
            # Initialize if not exists
            if not hasattr(self, '_app_usage'):
                self._app_usage = {}
            if not hasattr(self, '_current_app'):
                self._current_app = None
            if not hasattr(self, '_current_app_start_time'):
                self._current_app_start_time = None
            
            # Accumulate time for current app (temporary for this call only)
            temp_app_usage = {}
            if hasattr(self, '_app_usage') and self._app_usage:
                temp_app_usage = self._app_usage.copy()
            
            if self._current_app and self._current_app_start_time:
                current_time = time.time()
                elapsed = current_time - self._current_app_start_time
                if elapsed > 0:  # Only accumulate positive time
                    if self._current_app not in temp_app_usage:
                        temp_app_usage[self._current_app] = 0
                    temp_app_usage[self._current_app] += elapsed
            
            # Format and sort by usage time
            apps = []
            app_process_map = getattr(self, '_app_process_map', {})
            for app_name, seconds in temp_app_usage.items():
                if seconds > 0:  # Only include apps with usage time (at least 1 second)
                    hours = int(seconds // 3600)
                    minutes = int((seconds % 3600) // 60)
                    secs = int(seconds % 60)
                    # Format as HH:MM:SS
                    time_str = f"{hours:02d}:{minutes:02d}:{secs:02d}"
                    
                    app_data = {
                        "name": app_name,
                        "time": time_str,
                        "duration": time_str,
                        "seconds": seconds,
                        "process_name": app_process_map.get(app_name)  # Include process name for icon extraction
                    }
                    apps.append(app_data)
            
            # Sort by seconds (descending)
            apps.sort(key=lambda x: x["seconds"], reverse=True)
            return apps
        except Exception as e:
            logger.error(f"Error in get_app_usage_stats: {e}", exc_info=True)
            return []

    def apply_server_totals(
        self,
        in_time_iso: str | None = None,
        out_time_iso: str | None = None,
        break_seconds: int | None = None,
        on_break: bool | None = None,
        active_seconds: int | None = None,
    ):
        changed = False
        in_dt = self._parse_iso_datetime(in_time_iso)
        if in_dt:
            self._erp_in_time = in_dt
            if not self.session_start or in_dt < self.session_start:
                self.session_start = in_dt
            changed = True
        out_dt = self._parse_iso_datetime(out_time_iso)
        if out_dt:
            self._erp_out_time = out_dt
            changed = True
        if break_seconds is not None:
            self._erp_break_seconds = max(0, int(break_seconds))
            changed = True
        if active_seconds is not None:
            self._erp_active_seconds = max(0, int(active_seconds))
            changed = True
        if on_break is not None:
            self._erp_state_hint = SessionState.ON_BREAK if on_break else SessionState.CLOCKED_IN
            # Sync ERP open-break flag with server view
            self._erp_break_open = bool(on_break)
            changed = True
        if changed:
            self._erp_override_timestamp = time.time()
            if self._erp_state_hint and self.state != SessionState.LOGGED_OUT:
                self.state = self._erp_state_hint
            self.last_update = datetime.now(timezone.utc)

    def start_break(self):
        # Prevent duplicate break start API calls until a resume occurs.
        if self.state != SessionState.CLOCKED_IN or self.break_start is not None:
            return False
        self._accumulate_until_now()
        self.state = SessionState.ON_BREAK
        try:
            from datetime import timezone as _tz
            self.break_start = datetime.now(_tz.utc)
            date_iso = self.break_start.date().isoformat()
            start_iso = self.break_start.replace(microsecond=0).isoformat().replace("+00:00", "Z")
            # Send break API call only if ERP doesn't already have an open break
            ok = True
            data = {}
            status = 0
            if not self._erp_break_open:
                ok, data, status = self.attendance_api.start_break(
                    date_in_iso_format=date_iso,
                    break_start_iso_utc=start_iso,
                )
                if ok:
                    self._erp_break_open = True
                else:
                    try:
                        logger.warning(f"SessionManager: break start API call failed: {data}")
                    except Exception:
                        pass
        except Exception as e:
            try:
                logger.error(f"SessionManager: error starting break: {e}")
            except Exception:
                pass
        try:
            logger.info("SessionManager: break started")
        except Exception:
            pass
        return True

    def end_break(self, force: bool = False):
        if self.state != SessionState.ON_BREAK:
            if not force or not self.break_start:
                return False
            # Pretend we are currently on break so resume API can be sent
            self.state = SessionState.ON_BREAK
        self._accumulate_until_now()
        self.state = SessionState.CLOCKED_IN
        try:
            from datetime import timezone as _tz
            now_utc = datetime.now(_tz.utc)
            date_iso = now_utc.date().isoformat()
            end_iso = now_utc.replace(microsecond=0).isoformat().replace("+00:00", "Z")
            # Send break end API call
            ok, data, status = self.attendance_api.end_break(
                date_in_iso_format=date_iso,
                break_end_iso_utc=end_iso,
            )
            if ok:
                self._erp_break_open = False
            else:
                try:
                    logger.warning(f"SessionManager: break end API call failed: {data}")
                except Exception:
                    pass
        except Exception as e:
            try:
                logger.error(f"SessionManager: error ending break: {e}")
            except Exception:
                pass
        self.break_start = None
        self._sleep_break_active = False
        try:
            logger.info("SessionManager: break ended")
        except Exception:
            pass
        return True

    def update_activity(self, is_active: bool, idle_seconds: float):
        # Called frequently by the GUI to reflect user activity
        self.is_idle = not is_active
        self.current_idle_seconds = idle_seconds
        # Record activity sample (1 for active, 0 for idle)
        try:
            self._activity_samples.append(1 if is_active else 0)
            if len(self._activity_samples) > 0:
                self.last_min_activity_percent = 100.0 * sum(self._activity_samples) / len(self._activity_samples)
        except Exception:
            pass
        try:
            logger.debug(f"SessionManager: update_activity is_active={is_active} idle_seconds={idle_seconds:.1f} state={self.state.value}")
        except Exception:
            pass
        self._accumulate_until_now()

    def _accumulate_until_now(self):
        now = datetime.now(timezone.utc)
        if self.last_update is None:
            self.last_update = now
            return
        delta = (now - self.last_update).total_seconds()
        if delta <= 0:
            return
        if self.state == SessionState.CLOCKED_IN:
            if self.is_idle:
                # Prolonged idle counts as break; short idle remains idle
                if self.current_idle_seconds >= IDLE_TIMEOUT_SECONDS:
                    self.break_seconds += delta
                else:
                    self.idle_seconds += delta
            else:
                self.active_seconds += delta
        elif self.state == SessionState.ON_BREAK:
            self.break_seconds += delta
        # if logged out, no accumulation
        self.last_update = now

    def _log_capture_event(
        self,
        capture_type: str,
        timestamp: datetime,
        extra_metadata: dict | None = None,
        screenshot: CaptureArtifact | None = None,
        webcam: CaptureArtifact | None = None,
        screenshot_key: str | None = None,
        webcam_key: str | None = None,
    ) -> None:
        if not screenshot and not webcam:
            return
        try:
            metadata = {}
            if screenshot:
                metadata["screenshot"] = screenshot.filename
            if screenshot_key:
                metadata["screenshot_key"] = screenshot_key
            if webcam:
                metadata["webcam"] = webcam.filename
                metadata["webcam_name"] = WEBCAM_DEVICE_NAME
            if webcam_key:
                metadata["webcam_key"] = webcam_key
            if extra_metadata:
                try:
                    metadata.update(extra_metadata)
                except Exception:
                    pass
            
            # Determine tool and title
            if screenshot and webcam:
                tool = "capture"
                title = f"Screenshot + Webcam ({timestamp.strftime('%H:%M:%S')})"
            elif screenshot:
                tool = "screenshot"
                title = screenshot.filename
            else:
                tool = "webcam"
                title = webcam.filename if webcam else ""
            
            event_data = {
                "client_id": self.attendance_api.client_id or "",
                "user_id": getattr(self.attendance_api.auth_api, "user_id", "") or "unknown",
                "tool": tool,
                "action": tool,
                "start_time": timestamp.strftime("%Y-%m-%d %H:%M:%S"),
                "activity_type": tool,
                "title": title,
                "metadata_json": metadata,
                "screenshots": screenshot_key or "",
                "webcam_photo": webcam_key or "",
                "webcam_name": WEBCAM_DEVICE_NAME if webcam else "",
            }
            # Pass image artifacts to embed in Excel
            append_activity_event(
                event_data,
                screenshot_artifact=screenshot,
                webcam_artifact=webcam,
            )
        except Exception:
            pass

    # Background helpers
    def _start_background_tasks(self):
        if self._bg_running:
            return
        self._bg_running = True
        # reset session collections
        self._recent_windows = []
        self._captured_screenshots = []
        # Screenshot thread
        def _shot_loop():
            while self._bg_running:
                if self.state != SessionState.CLOCKED_IN:
                    time.sleep(5)
                    continue
                capture_ts_utc = datetime.now(timezone.utc)
                capture_ts_local = capture_ts_utc.astimezone()
                excel_timestamp = capture_ts_local.replace(tzinfo=None)
                screenshot_artifact = capture_screenshot(SCREENSHOTS_DIR)
                if screenshot_artifact:
                    self._captured_screenshots.append(screenshot_artifact.filename)
                with self._worklog_lock:
                    self._last_capture_time = capture_ts_utc
                    self._last_screenshot_path = screenshot_artifact.filename if screenshot_artifact else None
                webcam_artifact = capture_webcam_photo(WEBCAM_PHOTOS_DIR, excel_timestamp)
                with self._worklog_lock:
                    self._last_webcam_path = webcam_artifact.filename if webcam_artifact else None
                tabs_snapshot = collect_browser_tabs(
                    user_id=getattr(self.attendance_api.auth_api, "user_id", None),
                    max_tabs=MAX_BROWSER_TABS_CAPTURED,
                )
                with self._worklog_lock:
                    self._latest_tabs_snapshot = tabs_snapshot
                # Upload to server first (needs files)
                payload = {
                    "timestamp": capture_ts_utc.isoformat(),
                    "user_id": getattr(self.attendance_api.auth_api, "user_id", ""),
                    "session_state": self.state.value,
                    "session_start": self.session_start.isoformat() if self.session_start else None,
                }
                attachments = {}
                if screenshot_artifact:
                    attachments["screenshot"] = screenshot_artifact.filename
                if webcam_artifact:
                    attachments["webcam"] = webcam_artifact.filename
                if attachments:
                    payload["attachments"] = attachments
                try:
                    self.attendance_api.upload_worklog_event(
                        payload=payload,
                        screenshot_artifact=screenshot_artifact,
                        webcam_artifact=webcam_artifact,
                        tabs_snapshot=tabs_snapshot,
                    )
                except Exception as exc:
                    logger.error("SessionManager: worklog upload failed: %s", exc, exc_info=True)

                screenshot_key = self._upload_capture_artifact(screenshot_artifact, "screenshot")
                webcam_key = self._upload_capture_artifact(webcam_artifact, "webcam")

                if screenshot_artifact or webcam_artifact:
                    self._log_capture_event(
                        "screenshot" if screenshot_artifact else "webcam",
                        excel_timestamp,
                        extra_metadata={"tabs": tabs_snapshot} if tabs_snapshot else None,
                        screenshot=screenshot_artifact,
                        webcam=webcam_artifact,
                        screenshot_key=screenshot_key,
                        webcam_key=webcam_key,
                    )
                time.sleep(SCREENSHOT_INTERVAL_SECONDS)
        # Active window thread - tracks apps regardless of state
        def _window_loop():
            try:
                import time
                def extract_app_name(window_title):
                    """Extract clean app name from window title"""
                    if not window_title:
                        return "Unknown"
                    
                    # Try to get process name for better accuracy
                    process_exe = None
                    try:
                        import win32gui
                        import win32process
                        hwnd = win32gui.GetForegroundWindow()
                        if hwnd:
                            _, pid = win32process.GetWindowThreadProcessId(hwnd)
                            try:
                                import psutil
                                process = psutil.Process(pid)
                                exe_name = process.name()
                                process_exe = exe_name  # Store original exe name for icon extraction
                                # Remove .exe extension and clean up
                                if exe_name.endswith('.exe'):
                                    exe_name = exe_name[:-4]
                                # Map common executable names to friendly names
                                app_map = {
                                    'chrome': 'Google Chrome',
                                    'msedge': 'Microsoft Edge',
                                    'edge': 'Microsoft Edge',
                                    'firefox': 'Firefox',
                                    'Code': 'VS Code',
                                    'devenv': 'Visual Studio',
                                    'notepad++': 'Notepad++',
                                    'notepad': 'Notepad',
                                    'winword': 'Microsoft Word',
                                    'word': 'Microsoft Word',
                                    'excel': 'Microsoft Excel',
                                    'powerpnt': 'Microsoft PowerPoint',
                                    'powerpoint': 'Microsoft PowerPoint',
                                    'Teams': 'Microsoft Teams',
                                    'outlook': 'Microsoft Outlook',
                                    'figma': 'Figma',
                                    'slack': 'Slack',
                                    'discord': 'Discord',
                                    'zoom': 'Zoom',
                                    'whatsapp': 'WhatsApp',
                                    'telegram': 'Telegram',
                                    'Cursor': 'Cursor',
                                    'cursor': 'Cursor',
                                    'python': 'Python',
                                    'pythonw': 'Python',
                                }
                                # Check if we have a friendly name
                                exe_lower = exe_name.lower()
                                friendly_name = None
                                for key, mapped_name in app_map.items():
                                    if key in exe_lower:
                                        friendly_name = mapped_name
                                        break
                                
                                # Store process exe for icon extraction
                                if not hasattr(self, '_app_process_map'):
                                    self._app_process_map = {}
                                if friendly_name:
                                    self._app_process_map[friendly_name] = process_exe
                                    return friendly_name
                                else:
                                    app_name = exe_name.replace('_', ' ').title()
                                    self._app_process_map[app_name] = process_exe
                                    return app_name
                            except ImportError:
                                # psutil not available, fall through to title parsing
                                pass
                            except Exception:
                                # Process access failed, fall through to title parsing
                                pass
                    except Exception:
                        pass
                    
                    # Fallback to window title parsing
                    # Common patterns: "file.py - VS Code", "Document - Chrome", "Meeting - Teams"
                    if " - " in window_title:
                        parts = window_title.split(" - ")
                        app_name = parts[-1].strip()
                        # Clean up common suffixes
                        if app_name.endswith(".exe"):
                            app_name = app_name[:-4]
                        # Remove common prefixes
                        if app_name.startswith("Microsoft "):
                            app_name = app_name.replace("Microsoft ", "")
                        # Map common app names from window titles
                        title_map = {
                            'Chrome': 'Google Chrome',
                            'Google Chrome': 'Google Chrome',
                            'Edge': 'Microsoft Edge',
                            'Microsoft Edge': 'Microsoft Edge',
                            'VS Code': 'VS Code',
                            'Visual Studio Code': 'VS Code',
                            'Teams': 'Microsoft Teams',
                            'Microsoft Teams': 'Microsoft Teams',
                            'Outlook': 'Microsoft Outlook',
                            'Microsoft Outlook': 'Microsoft Outlook',
                            'Word': 'Microsoft Word',
                            'Microsoft Word': 'Microsoft Word',
                            'Excel': 'Microsoft Excel',
                            'Microsoft Excel': 'Microsoft Excel',
                            'PowerPoint': 'Microsoft PowerPoint',
                            'Microsoft PowerPoint': 'Microsoft PowerPoint',
                            'Figma': 'Figma',
                            'Slack': 'Slack',
                            'Discord': 'Discord',
                            'Zoom': 'Zoom',
                            'WhatsApp': 'WhatsApp',
                            'Telegram': 'Telegram',
                        }
                        # Check both exact match and case-insensitive
                        app_lower = app_name.lower()
                        for key, friendly_name in title_map.items():
                            if app_name == key or app_lower == key.lower():
                                return friendly_name
                        return app_name
                    
                    # If no separator, try to extract meaningful part
                    # Check if it's a URL
                    title = window_title
                    if "://" in title or "www." in title.lower() or ".com" in title.lower():
                        # Extract domain from URL
                        try:
                            from urllib.parse import urlparse
                            # Add http:// if missing
                            if not title.startswith(('http://', 'https://')):
                                title = 'https://' + title
                            parsed = urlparse(title)
                            domain = parsed.netloc or parsed.path.split('/')[0]
                            if domain:
                                # Map common domains to app names
                                domain_map = {
                                    'meet.google.com': 'Google Meet',
                                    'web.whatsapp.com': 'WhatsApp Web',
                                    'whatsapp.com': 'WhatsApp',
                                    'instagram.com': 'Instagram',
                                    'www.instagram.com': 'Instagram',
                                    'figma.com': 'Figma',
                                    'bookmyshow.com': 'Book My Show',
                                    'youtube.com': 'YouTube',
                                    'gmail.com': 'Gmail',
                                    'drive.google.com': 'Google Drive',
                                    'facebook.com': 'Facebook',
                                    'twitter.com': 'Twitter',
                                    'x.com': 'Twitter',
                                    'linkedin.com': 'LinkedIn',
                                    'snapchat.com': 'Snapchat',
                                    'tiktok.com': 'TikTok',
                                    'pinterest.com': 'Pinterest',
                                    'reddit.com': 'Reddit',
                                }
                                domain_lower = domain.lower()
                                for key, app_name in domain_map.items():
                                    if key in domain_lower:
                                        return app_name
                                # Extract main domain name
                                parts = domain.split('.')
                                if len(parts) >= 2:
                                    main_domain = parts[-2] if parts[-2] not in ['www', 'web'] else parts[-3] if len(parts) >= 3 else parts[-2]
                                    return main_domain.title()
                                return domain
                        except Exception:
                            pass
                    
                    # Return first 30 chars as fallback
                    return window_title[:30] if len(window_title) > 30 else window_title
                
                while self._bg_running:
                    try:
                        title = get_active_window_title()
                        current_time = time.time()
                        
                        # Track app usage time - initialize if needed
                        if not hasattr(self, '_app_usage'):
                            self._app_usage = {}
                        if not hasattr(self, '_current_app'):
                            self._current_app = None
                        if not hasattr(self, '_current_app_start_time'):
                            self._current_app_start_time = None
                        
                        if title:
                            # keep only last 50 titles
                            self._recent_windows.append(title)
                            if len(self._recent_windows) > 50:
                                self._recent_windows = self._recent_windows[-50:]
                            
                            # Track app usage time
                            app_name = extract_app_name(title)
                            
                            # If app changed, accumulate time for previous app
                            if self._current_app and self._current_app != app_name:
                                if self._current_app_start_time:
                                    elapsed = current_time - self._current_app_start_time
                                    if elapsed > 0:  # Only accumulate positive time
                                        if self._current_app not in self._app_usage:
                                            self._app_usage[self._current_app] = 0
                                        self._app_usage[self._current_app] += elapsed
                                        # Log for debugging (only occasionally to avoid spam)
                                        if len(self._app_usage) % 5 == 0:
                                            try:
                                                logger.debug(f"App usage updated: {self._current_app} += {elapsed:.1f}s")
                                            except Exception:
                                                pass
                            
                            # Update current app
                            if self._current_app != app_name:
                                self._current_app = app_name
                                self._current_app_start_time = current_time
                            elif self._current_app_start_time is None:
                                self._current_app_start_time = current_time
                        else:
                            # No active window - if we had a current app, accumulate its time
                            if self._current_app and self._current_app_start_time:
                                elapsed = current_time - self._current_app_start_time
                                if elapsed > 0:
                                    if self._current_app not in self._app_usage:
                                        self._app_usage[self._current_app] = 0
                                    self._app_usage[self._current_app] += elapsed
                                self._current_app = None
                                self._current_app_start_time = None
                    except Exception as e:
                        logger.error(f"Error in window loop iteration: {e}", exc_info=True)
                    
                    # Increased sleep interval to reduce CPU usage
                    time.sleep(10)  # Check every 10 seconds instead of 5
            except Exception as e:
                logger.error(f"Error in window tracking loop: {e}", exc_info=True)
                # Continue running even if there's an error
                import time
                time.sleep(5)
        # Power/suspend monitor thread: detect long gaps as system sleep
        def _power_monitor_loop():
            try:
                import time
                monitor_interval_seconds = 10
                self._last_tick = datetime.now(timezone.utc)
                while self._bg_running:
                    time.sleep(monitor_interval_seconds)
                    now = datetime.now(timezone.utc)
                    elapsed = (now - self._last_tick).total_seconds()
                    # If elapsed is significantly greater than the interval, assume suspend
                    gap = elapsed - monitor_interval_seconds
                    if gap > 60:  # treat >60s extra as a sleep gap
                        sleep_duration_seconds = int(max(0, elapsed - monitor_interval_seconds))
                        sleep_start = now - timedelta(seconds=sleep_duration_seconds)
                        try:
                            # Count laptop sleep time as break if session is active or already on break
                            if self.state in (SessionState.CLOCKED_IN, SessionState.ON_BREAK):
                                # If we were clocked in, start a break for the sleep period
                                # Only if break wasn't already started by on_sleep() callback
                                if (
                                    not self._sleep_gap_prehandled
                                    and self.state == SessionState.CLOCKED_IN
                                    and not self.break_start
                                ):
                                    try:
                                        self._accumulate_until_now()
                                        self.state = SessionState.ON_BREAK
                                        self.break_start = sleep_start
                                        self._sleep_break_active = True
                                        date_iso = sleep_start.date().isoformat()
                                        start_iso = sleep_start.replace(microsecond=0).isoformat().replace("+00:00", "Z")
                                        ok, data, status = self.attendance_api.start_break(date_in_iso_format=date_iso, break_start_iso_utc=start_iso)
                                        if not ok:
                                            try:
                                                logger.warning(f"SessionManager: sleep break start API call failed: {data}")
                                            except Exception:
                                                pass
                                    except Exception as e:
                                        try:
                                            logger.error(f"SessionManager: error handling sleep break: {e}")
                                        except Exception:
                                            pass
                                # Add sleep time to break seconds
                                # If already ON_BREAK, we still need to manually add the sleep duration
                                # because _accumulate_until_now() won't run during sleep
                                self.break_seconds += sleep_duration_seconds
                                self._sleep_gap_prehandled = False
                            # Move last_update forward to avoid double counting
                            self.last_update = now
                        except Exception:
                            pass
                        try:
                            self._record_sleep_event(sleep_start, now, sleep_duration_seconds)
                        except Exception:
                            pass
                    self._last_tick = now
            except Exception:
                pass
        self._screenshot_thread = threading.Thread(target=_shot_loop, daemon=True)
        self._window_log_thread = threading.Thread(target=_window_loop, daemon=True)
        self._suspend_monitor_thread = threading.Thread(target=_power_monitor_loop, daemon=True)
        try:
            self._screenshot_thread.start()
            self._window_log_thread.start()
            self._suspend_monitor_thread.start()
        except Exception:
            pass

    def _stop_background_tasks(self):
        self._bg_running = False
        # Threads are daemons; they will exit naturally on flag

    def get_latest_worklog_info(self):
        with self._worklog_lock:
            return {
                "timestamp": self._last_capture_time,
                "screenshot_path": self._last_screenshot_path,
                "webcam_path": self._last_webcam_path,
                "tabs": self._latest_tabs_snapshot,
            }

    def _upload_capture_artifact(self, artifact, name: str) -> str | None:
        if not artifact:
            return None
        try:
            ok, key = self.attendance_api.upload_capture_asset(artifact, capture_type=name)
            if ok:
                try:
                    logger.info("SessionManager: %s uploaded to work-log API (key=%s)", name, key)
                except Exception:
                    pass
                return key
            try:
                logger.warning("SessionManager: %s upload failed (%s)", name, artifact.filename)
            except Exception:
                pass
        except Exception as exc:
            try:
                logger.error("SessionManager: %s upload crashed: %s", name, exc, exc_info=True)
            except Exception:
                pass
        return None

    @staticmethod
    def _format_hms(total_seconds: float) -> str:
        seconds = int(total_seconds)
        hrs = seconds // 3600
        mins = (seconds % 3600) // 60
        secs = seconds % 60
        return f"{hrs:02d}:{mins:02d}:{secs:02d}"

    @staticmethod
    def _parse_iso_datetime(value: str | None):
        if not value:
            return None
        text = str(value).strip()
        if not text:
            return None
        text = text.replace("Z", "+00:00")
        try:
            dt = datetime.fromisoformat(text)
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=timezone.utc)
            return dt.astimezone(timezone.utc)
        except Exception:
            return None

    # Persistence helpers
    def _record_sleep_event(self, sleep_start: datetime, wake_time: datetime, duration_seconds: int) -> None:
        try:
            data = read_local_storage()
            events = data.get("sleep_events", [])
            events.append({
                "date": wake_time.date().isoformat(),
                "sleep_start": sleep_start.isoformat(timespec="seconds"),
                "wake_time": wake_time.isoformat(timespec="seconds"),
                "duration_seconds": int(duration_seconds),
            })
            data["sleep_events"] = events
            write_local_storage(data)
            try:
                logger.info(f"Recorded sleep event: start={sleep_start.isoformat()} wake={wake_time.isoformat()} duration={duration_seconds}s")
            except Exception:
                pass
            try:
                # Optionally log via attendance API for server-side visibility
                self.attendance_api.log_event("system_resume", wake_time, {"slept_seconds": int(duration_seconds)})
            except Exception:
                pass
        except Exception:
            pass
