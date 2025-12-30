"""
Scheduler for uploading activity_log.xlsx file every 2 hours.
"""
import sys
import io
import os

# CRITICAL: Set UTF-8 environment BEFORE any imports that might use stdout/stderr
os.environ["PYTHONUTF8"] = "1"

# Force UTF-8 encoding with error replacement - MUST be done before any output
# Note: We skip wrapping stdout/stderr if they're closed or invalid
# Instead, we rely on safe_print() and safe_log_*() functions to handle Unicode
try:
    # Only try to wrap if streams are available and not already wrapped
    if hasattr(sys.stdout, 'buffer'):
        try:
            # Check if we can access the buffer without errors
            _ = sys.stdout.buffer
            if not isinstance(sys.stdout, io.TextIOWrapper):
                sys.stdout = io.TextIOWrapper(
                    sys.stdout.buffer,
                    encoding="utf-8",
                    errors="replace"
                )
        except (ValueError, OSError, AttributeError, TypeError):
            # stdout buffer is closed or invalid - skip wrapping
            # safe_print() will handle Unicode errors instead
            pass
    if hasattr(sys.stderr, 'buffer'):
        try:
            # Check if we can access the buffer without errors
            _ = sys.stderr.buffer
            if not isinstance(sys.stderr, io.TextIOWrapper):
                sys.stderr = io.TextIOWrapper(
                    sys.stderr.buffer,
                    encoding="utf-8",
                    errors="replace"
                )
        except (ValueError, OSError, AttributeError, TypeError):
            # stderr buffer is closed or invalid - skip wrapping
            # safe_log_*() functions will handle Unicode errors instead
            pass
except Exception:
    # Ultimate fallback - continue without wrapping
    # All output will go through safe_print() and safe_log_*() which handle Unicode
    pass

# Monkey-patch default encoding to prevent cp1252 from being used
try:
    import codecs
    # Override the default encoding for Windows
    if sys.platform == 'win32':
        # Force UTF-8 as default encoding
        import locale
        try:
            locale.setlocale(locale.LC_ALL, 'en_US.UTF-8')
        except:
            pass
except Exception:
    pass

import threading
import time

# Import logger AFTER encoding is set up
try:
    from utils.logger import logger
except Exception:
    # Fallback logger if import fails
    import logging
    logger = logging.getLogger("ExcelUploadScheduler")
    logger.addHandler(logging.NullHandler())


def safe_str(text):
    """
    Safely convert any text to ASCII-safe string, replacing problematic Unicode.
    """
    if text is None:
        return ""
    try:
        return str(text).encode("ascii", "replace").decode("ascii")
    except Exception:
        return "[UnicodeError]"


def safe_print(*args, **kwargs):
    """
    Safe print function that handles Unicode encoding errors.
    Replaces problematic characters instead of crashing.
    """
    try:
        # Convert all args to safe strings first
        safe_args = [safe_str(arg) for arg in args]
        print(*safe_args, **kwargs)
    except Exception:
        # Ultimate fallback - silent fail to prevent EXE crash
        pass


def safe_log_info(message):
    """Safely log info message, sanitizing Unicode."""
    try:
        logger.info(safe_str(message))
    except Exception:
        pass


def safe_log_warning(message):
    """Safely log warning message, sanitizing Unicode."""
    try:
        logger.warning(safe_str(message))
    except Exception:
        pass


def safe_log_error(message, exc_info=False):
    """Safely log error message, sanitizing Unicode."""
    try:
        logger.error(safe_str(message), exc_info=exc_info)
    except Exception:
        pass


class ExcelUploadScheduler:
    """
    Schedules periodic uploads of activity_log.xlsx file.
    """
    
    def __init__(self, attendance_api, upload_interval_hours: float = 2.0):
        """
        Initialize the scheduler.
        
        Args:
            attendance_api: AttendanceAPI instance with upload_activity_log_excel method
            upload_interval_hours: Interval between uploads in hours (default: 2.0)
        """
        self.attendance_api = attendance_api
        self.upload_interval_seconds = upload_interval_hours * 3600
        self._running = False
        self._thread = None
        self._stop_event = threading.Event()
    
    def start(self):
        """Start the scheduler in a background thread."""
        # Wrap in Unicode-safe handler - catch ALL encoding errors
        try:
            # Re-ensure encoding is set (in case EXE context reset it)
            # Don't check .closed - just try to wrap and catch errors
            try:
                if hasattr(sys.stdout, 'buffer'):
                    try:
                        # Test if buffer is accessible by trying to access it
                        _ = sys.stdout.buffer
                        if not isinstance(sys.stdout, io.TextIOWrapper):
                            sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
                    except (ValueError, OSError, AttributeError, TypeError):
                        pass  # stdout is closed or invalid - safe_print() will handle Unicode
                if hasattr(sys.stderr, 'buffer'):
                    try:
                        # Test if buffer is accessible by trying to access it
                        _ = sys.stderr.buffer
                        if not isinstance(sys.stderr, io.TextIOWrapper):
                            sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
                    except (ValueError, OSError, AttributeError, TypeError):
                        pass  # stderr is closed or invalid - safe_log_*() will handle Unicode
            except Exception:
                pass  # Ultimate fallback - safe functions will handle Unicode
            
            if self._running:
                try:
                    safe_log_warning("Excel upload scheduler is already running")
                    safe_print("  Excel upload scheduler is already running")
                except:
                    pass
                return
            
            self._running = True
            self._stop_event.clear()
            self._thread = threading.Thread(target=self._scheduler_loop, daemon=True)
            self._thread.start()
            
            # Sanitize ALL strings before any output
            try:
                interval_seconds = self.upload_interval_seconds
                if interval_seconds < 60:
                    interval_str = safe_str(f"{int(interval_seconds)} seconds")
                else:
                    interval_str = safe_str(f"{interval_seconds/60:.1f} minutes")
                
                safe_log_info(f"Excel upload scheduler started (interval: {interval_str})")
                safe_print(f"[SCHEDULER] Excel upload scheduler started (interval: {interval_str})")
            except:
                # If even sanitized output fails, just continue silently
                pass
        except UnicodeEncodeError:
            # Catch Unicode encoding errors specifically - fail silently
            try:
                safe_print("[SCHEDULER] Started (encoding error suppressed)")
            except:
                pass
        except Exception as e:
            try:
                safe_print(f"[SCHEDULER] Error: {safe_str(str(e))}")
                safe_log_error(f"Error starting Excel upload scheduler: {safe_str(str(e))}", exc_info=False)
            except:
                pass  # Ultimate fallback - silent fail to prevent EXE crash
    
    def stop(self):
        """Stop the scheduler."""
        if not self._running:
            return
        
        self._running = False
        self._stop_event.set()
        if self._thread and self._thread.is_alive():
            self._thread.join(timeout=5.0)
        safe_log_info("Excel upload scheduler stopped")
    
    def _scheduler_loop(self):
        """Main scheduler loop that runs uploads at specified intervals."""
        try:
            # Wait a bit before first upload to ensure app is fully initialized
            initial_wait = min(10, self.upload_interval_seconds)  # Wait 10 seconds or interval, whichever is shorter
            safe_log_info(f"Excel upload scheduler: Waiting {initial_wait} seconds before first upload...")
            safe_print(f"[SCHEDULER] Waiting {initial_wait} seconds before first upload...")
            time.sleep(initial_wait)
            
            upload_count = 0
            while self._running and not self._stop_event.is_set():
                try:
                    upload_count += 1
                    safe_log_info(f"[Excel Upload #{upload_count}] Starting upload attempt...")
                    safe_print(f"\n[UPLOAD #{upload_count}] Starting upload attempt...")
                    
                    # Perform the upload (client_id will be extracted from token if not provided)
                    success, message = self.attendance_api.upload_activity_log_excel()
                    # Sanitize message before using it
                    safe_message = safe_str(message) if message else ""
                    
                    if success:
                        safe_log_info(f"[Excel Upload #{upload_count}]  SUCCESS: {safe_message}")
                        safe_print(f"    SUCCESS: {safe_message}")
                    else:
                        safe_log_warning(f"[Excel Upload #{upload_count}]  FAILED: {safe_message}")
                        safe_print(f"    FAILED: {safe_message}")
                    
                    # Log next upload time
                    if self.upload_interval_seconds < 60:
                        next_msg = f"Next upload in {self.upload_interval_seconds} seconds"
                        safe_log_info(f"[Excel Upload #{upload_count}] {next_msg}")
                        safe_print(f"    {next_msg}")
                    else:
                        next_msg = f"Next upload in {self.upload_interval_seconds/60:.1f} minutes"
                        safe_log_info(f"[Excel Upload #{upload_count}] {next_msg}")
                        safe_print(f"    {next_msg}")
                    
                    # Wait for the next interval (or until stop event is set)
                    if self._stop_event.wait(timeout=self.upload_interval_seconds):
                        # Stop event was set, exit loop
                        break
                except Exception as exc:
                    safe_log_error(f"[Excel Upload #{upload_count}]  ERROR in scheduler loop: {exc}", exc_info=True)
                    safe_print(f"    ERROR: {safe_str(exc)}")
                    # Wait a shorter time before retrying on error
                    error_wait = min(30, self.upload_interval_seconds)  # Wait 30 seconds or interval, whichever is shorter
                    if self._stop_event.wait(timeout=error_wait):
                        break
            
            self._running = False
            safe_log_info(f"Excel upload scheduler loop exited (total uploads attempted: {upload_count})")
            safe_print(f"\n[SCHEDULER] Excel upload scheduler stopped (total uploads: {upload_count})")
        except Exception as e:
            self._running = False
            safe_print(f"[SCHEDULER] Fatal error in scheduler loop: {safe_str(e)}")
            safe_log_error(f"Fatal error in Excel upload scheduler loop: {e}", exc_info=True)
    
    def trigger_upload_now(self) -> tuple[bool, str]:
        """
        Manually trigger an upload immediately (for testing or manual requests).
        Returns (success, message).
        Client ID will be extracted from token automatically.
        """
        return self.attendance_api.upload_activity_log_excel()

