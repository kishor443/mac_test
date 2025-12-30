# =========================
# SAFE UTF-8 HANDLING (FIX)
# =========================
import sys
import os
import io

os.environ["PYTHONUTF8"] = "1"

def _safe_wrap_stream(stream):
    try:
        if stream and hasattr(stream, "buffer"):
            return io.TextIOWrapper(
                stream.buffer,
                encoding="utf-8",
                errors="replace"
            )
    except Exception:
        pass
    return stream

sys.stdout = _safe_wrap_stream(sys.stdout)
sys.stderr = _safe_wrap_stream(sys.stderr)

# =========================
# IMPORTS
# =========================
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import threading
import atexit

from api.auth_api import AuthAPI
from api.attendance_api import AttendanceAPI
from api.project_api import ProjectAPI
from api.task_api import TaskAPI
from api.appointment_api import AppointmentAPI

from config import ERP_DEVICE_ID, IDLE_TIMEOUT_SECONDS

from core.activity_tracker import ActivityTracker
from core.session_manager import SessionManager, SessionState

from gui.main_window import MainWindow
from gui.login_screen import LoginScreen
from gui.group_select import GroupSelectDialog
from gui.shift_select import ShiftSelectDialog
from gui.theme import apply_theme, spacing, fonts

from utils.auto_startup import enable_auto_startup
from utils.data_retention import enforce_data_retention_async
from utils.excel_upload_scheduler import ExcelUploadScheduler
from utils.excel_storage import set_default_client_id
from utils.logger import logger

from win_event_hook import WinEventHook


# =========================
# MAIN APPLICATION
# =========================
def main():
    enforce_data_retention_async()

    # Enable auto-startup (first run popup only once)
    if enable_auto_startup():
        marker = os.path.join(os.path.expanduser("~"), ".productivity_tracker_autostarted")
        if not os.path.exists(marker):
            try:
                root = tk.Tk()
                root.withdraw()
                messagebox.showinfo(
                    "Auto-Startup Enabled",
                    "Productivity Tracker will now launch automatically on Windows startup."
                )
                with open(marker, "w", encoding="utf-8") as f:
                    f.write("enabled")
            except Exception as e:
                # Surface errors in terminal/logger so they are visible during setup
                logger.exception(f"Auto-startup enable failed: {e}")

    # =========================
    # AUTHENTICATION
    # =========================
    auth = AuthAPI()
    login_data = None
    success = False
    message = ""

    if auth.refresh_token and auth.user_id:
        success, message, login_data = auth.login_with_refresh(
            refresh_token=auth.refresh_token,
            user_id=auth.user_id,
            device_id=ERP_DEVICE_ID
        )

    if not success:
        login = LoginScreen()
        creds = login.show()
        if not creds:
            logger.info("Login cancelled")
            return

        # Reload tokens from storage (login_screen might have saved them)
        auth._load_tokens()
        
        # Check if login_screen already succeeded (token exists)
        if auth.access_token:
            logger.info("Login already successful (token found in storage)")
            success = True
            message = "Login successful"
            login_data = {}  # Will be fetched later if needed
        else:
            # Extract and validate credentials
            phone = (creds.get("phone") or "").strip()
            password = (creds.get("password") or "").strip()
            otp = (creds.get("otp") or "").strip()
            device_id = creds.get("device_id") or ERP_DEVICE_ID
            
            # Debug logging
            logger.info(f"Login attempt - method: {creds.get('method')}, phone: {phone[:3] + '***' if phone else '(empty)'}, has_password: {bool(password)}, has_otp: {bool(otp)}")
            
            if creds.get("method") == "password":
                if not password:
                    logger.error("Password is empty in credentials")
                    # Create root window for error dialog
                    root = tk.Tk()
                    root.withdraw()
                    messagebox.showerror("Login Failed", "Password cannot be empty")
                    root.destroy()
                    return
                success, message, login_data = auth.login(
                    phone=phone,
                    password=password,
                    device_id=device_id
                )
            else:
                if not otp:
                    logger.error("OTP is empty in credentials")
                    # Create root window for error dialog
                    root = tk.Tk()
                    root.withdraw()
                    messagebox.showerror("Login Failed", "OTP cannot be empty")
                    root.destroy()
                    return
                success, message, login_data = auth.login_via_otp(
                    phone=phone,
                    otp=otp,
                    device_id=device_id
                )

    if not success:
        # Create root window for error dialog if it doesn't exist
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Login Failed", message or "Unknown error occurred")
            root.destroy()
        except Exception as e:
            logger.error(f"Failed to show error dialog: {e}")
            print(f"Login Failed: {message or 'Unknown error occurred'}")
        return

    if not auth.access_token:
        # Create root window for error dialog
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Login Failed", "No access token received")
        root.destroy()
        return

    # Ensure user_id is available (extract from token if not set)
    if not auth.user_id:
        logger.info("User ID not found in auth, attempting to extract from token")
        auth.user_id = auth.get_user_id_from_token()
        if auth.user_id:
            logger.info(f"User ID extracted from token: {auth.user_id}")
            auth._save_tokens()  # Save the user_id
        
    if not auth.user_id:
        logger.error("User ID is still not available after token extraction")
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Login Error", "Unable to get user ID from access token. Please try logging in again.")
        root.destroy()
        return

    # =========================
    # ACTIVITY TRACKER
    # =========================
    tracker = ActivityTracker(user_id=auth.user_id or "unknown")
    tracker.start()

    # =========================
    # CLIENT / GROUP SELECTION
    # =========================
    logger.info(f"Fetching clients - user_id: {auth.user_id}, device_id: {ERP_DEVICE_ID[:50]}...")
    ok, error_message, clients = auth.fetch_clients(
        user_id=auth.user_id,
        device_id=ERP_DEVICE_ID
    )
    if not ok:
        logger.error(f"Failed to fetch clients: {error_message}")
        # Create root window for error dialog
        root = tk.Tk()
        root.withdraw()
        detailed_error = f"Failed to fetch clients.\n\nError: {error_message or 'Unknown error'}\n\nPlease check your internet connection and try again."
        messagebox.showerror("Error", detailed_error)
        root.destroy()
        return

    selector = GroupSelectDialog(clients)
    client_id = selector.show()
    if not client_id:
        messagebox.showwarning("Selection Required", "Client selection is mandatory")
        return

    # =========================
    # API INITIALIZATION
    # =========================
    attendance = AttendanceAPI(auth)
    project_api = ProjectAPI(auth)
    task_api = TaskAPI(auth)
    appointment_api = AppointmentAPI(auth)

    for api in (attendance, project_api, task_api, appointment_api):
        try:
            api.set_client(client_id)
        except Exception as e:
            logger.exception(f"Failed to set client for API {api.__class__.__name__}: {e}")

    try:
        set_default_client_id(client_id)
    except Exception as e:
        logger.exception(f"Failed to persist default client id: {e}")

    tracker.update_context(user_id=auth.user_id, client_id=client_id)

    # =========================
    # EXCEL UPLOAD SCHEDULER
    # =========================
    excel_upload_scheduler = ExcelUploadScheduler(
        attendance,
        upload_interval_hours=4.0
    )
    excel_upload_scheduler.start()

    # =========================
    # SESSION MANAGER
    # =========================
    session = SessionManager(attendance)
    atexit.register(lambda: session.clock_out(reason="process_exit"))

    main_window = None

    def _has_pending_break():
        try:
            return (
                session.state == SessionState.ON_BREAK or
                session.break_start is not None or
                getattr(session, "_sleep_break_active", False)
            )
        except Exception:
            return False

    # =========================
    # SYSTEM EVENT HANDLERS
    # =========================
    def on_lock():
        if session.state == SessionState.CLOCKED_IN and not _has_pending_break():
            session.start_break()
        tracker.stop()

    def on_unlock():
        if _has_pending_break():
            session.end_break(force=True)
        tracker.start()

    def on_sleep():
        if session.state == SessionState.CLOCKED_IN:
            session.start_break()
        tracker.stop()

    def on_wake():
        if _has_pending_break():
            session.end_break(force=True)
        tracker.start()

    def on_shutdown():
        session.clock_out(reason="system_shutdown")
        os._exit(0)

    WinEventHook(
        on_lock=on_lock,
        on_unlock=on_unlock,
        on_sleep=on_sleep,
        on_wake=on_wake,
        on_shutdown=on_shutdown
    )

    # =========================
    # LOGOUT HANDLER
    # =========================
    def handle_logout():
        try:
            session.clock_out(reason="logout")
            auth.logout()
        except Exception:
            pass
        os.execl(sys.executable, sys.executable, *sys.argv)

    # =========================
    # MAIN WINDOW
    # =========================
    main_window = MainWindow(
        session_manager=session,
        user_info=login_data,
        on_clock_out=lambda: None,
        auto_clock_in=True,
        on_logout=handle_logout,
        project_api=project_api,
        task_api=task_api,
        auth_api=auth,
        appointment_api=appointment_api
    )
    main_window.show()


# =========================
# ENTRY POINT
# =========================
if __name__ == "__main__":
    main()
