import os
import json
import webview
import threading
import time
from datetime import datetime, timedelta, timezone
from gui.shift_select import ShiftSelectDialog
from core.session_manager import SessionState
from gui.reports_window import ReportsWindow
from config import IDLE_TIMEOUT_SECONDS, WEBCAM_DEVICE_NAME
from core.activity_tracker import ActivityTracker
from utils.logger import logger
from utils.resource_path import get_html_path
try:
    from tkinter import messagebox
except ImportError:
    # Fallback if tkinter not available
    class messagebox:
        @staticmethod
        def showinfo(title, message):
            # Use logger instead of print to avoid closed file errors
            try:
                from utils.logger import logger
                logger.info(f"{title}: {message}")
            except:
                pass
        
        @staticmethod
        def showerror(title, message):
            # Use logger instead of print to avoid closed file errors
            try:
                from utils.logger import logger
                logger.error(f"{title}: {message}")
            except:
                pass
        
        @staticmethod
        def showwarning(title, message):
            # Use logger instead of print to avoid closed file errors
            try:
                from utils.logger import logger
                logger.warning(f"{title}: {message}")
            except:
                pass

# Global reference to main window instance
_main_window_ref = None

class MainWindowAPI:
    """API class for pywebview - keep it simple to avoid introspection issues"""
    
    def initialize(self):
        """Initialize the window"""
        global _main_window_ref
        if _main_window_ref:
            try:
                _main_window_ref._sync_tracker_identity()
                _main_window_ref._load_shifts()
                
                # Check server state before trying to clock in
                # This will sync the state and update UI accordingly
                logger.info("initialize: Starting state sync from server")
                _main_window_ref._sync_state_from_server()
                logger.info(f"initialize: After state sync, state is: {_main_window_ref.session_manager.state}")
                
                # Refresh attendance info after state sync
                _main_window_ref._refresh_attendance_info()
                
                # Ensure controls are updated after all sync operations
                # This is critical - it updates button states based on synced state
                _main_window_ref._update_controls()
                
                # Refresh dashboard metrics
                _main_window_ref._refresh_dashboard_metrics()

                # Refresh announcements/notices
                _main_window_ref._refresh_notices()
                
                # Load user profile picture
                _main_window_ref._load_user_profile()
                
                # Final update to ensure button states are correct after all operations
                _main_window_ref._update_controls()
                
                # Log current state for debugging
                logger.info(f"After initialization, final state is: {_main_window_ref.session_manager.state}")
                
                # Force UI update after initialization with a small delay
                import threading
                def force_ui_update():
                    import time
                    time.sleep(1.0)  # Wait 1 second for everything to settle
                    try:
                        _main_window_ref._update_controls()
                        _main_window_ref._refresh_attendance_info()
                        logger.info("Force UI update completed after initialization")
                    except Exception as e:
                        logger.error(f"Error in force UI update: {e}")
                threading.Thread(target=force_ui_update, daemon=True).start()
                
                if getattr(_main_window_ref.session_manager.attendance_api, "machine_punch_required", False):
                    _main_window_ref._update_status("Status: MACHINE PUNCH REQUIRED — Use attendance machine")
                    _main_window_ref._update_button_states({
                        "clockIn": False,
                        "clockOut": False,
                        "startBreak": False,
                        "endBreak": False
                    })
                    _main_window_ref.auto_clock_in = False
                
                # Only auto clock in if not already clocked in
                # If state was synced to CLOCKED_IN, this won't execute
                if _main_window_ref.auto_clock_in and _main_window_ref.session_manager.state == SessionState.LOGGED_OUT:
                    threading.Timer(0.1, _main_window_ref.clock_in).start()
            except Exception as e:
                logger.error(f"Initialize error: {e}")
                import traceback
                traceback.print_exc()
    
    def get_user_profile(self):
        """Return current user's profile information for JS (first_name, last_name)."""
        global _main_window_ref
        if not _main_window_ref:
            return {"first_name": "", "last_name": ""}
        try:
            return _main_window_ref._get_user_profile_for_js()
        except Exception as e:
            logger.error(f"get_user_profile error: {e}", exc_info=True)
            return {"first_name": "", "last_name": ""}
    
    def start_break(self):
        """Start break"""
        global _main_window_ref
        logger.info("[BUTTON CLICK] start_break button clicked")
        if _main_window_ref:
            _main_window_ref.start_break()
            logger.info("[BUTTON CLICK] start_break executed successfully")
        else:
            logger.error("[BUTTON CLICK] ERROR: Main window reference not available for start_break")
    
    def end_break(self):
        """End break"""
        global _main_window_ref
        logger.info("[BUTTON CLICK] end_break button clicked")
        if _main_window_ref:
            _main_window_ref.end_break()
            logger.info("[BUTTON CLICK] end_break executed successfully")
        else:
            logger.error("[BUTTON CLICK] ERROR: Main window reference not available for end_break")
    
    def logout(self):
        """Logout"""
        global _main_window_ref
        logger.info("[BUTTON CLICK] logout button clicked")
        if _main_window_ref:
            _main_window_ref.logout()
            logger.info("[BUTTON CLICK] logout executed successfully")
        else:
            logger.error("[BUTTON CLICK] ERROR: Main window reference not available for logout")
    
    def clock_in(self):
        """Clock in"""
        global _main_window_ref
        logger.info("[BUTTON CLICK] clock_in button clicked")
        if _main_window_ref:
            _main_window_ref.clock_in()
            logger.info("[BUTTON CLICK] clock_in executed successfully")
        else:
            logger.error("[BUTTON CLICK] ERROR: Main window reference not available for clock_in")
    
    def clock_out(self):
        """Clock out"""
        global _main_window_ref
        logger.info("[BUTTON CLICK] clock_out button clicked")
        if _main_window_ref:
            _main_window_ref.clock_out()
            logger.info("[BUTTON CLICK] clock_out executed successfully")
        else:
            logger.error("[BUTTON CLICK] ERROR: Main window reference not available for clock_out")
    
    def refresh_shifts(self):
        """Refresh shifts"""
        global _main_window_ref
        logger.info("[BUTTON CLICK] refresh_shifts button clicked")
        if _main_window_ref:
            _main_window_ref._load_shifts()
            logger.info("[BUTTON CLICK] refresh_shifts executed successfully")
        else:
            logger.error("[BUTTON CLICK] ERROR: Main window reference not available for refresh_shifts")
    
    def show_summary(self):
        """Show summary"""
        global _main_window_ref
        logger.info("[BUTTON CLICK] show_summary button clicked")
        if _main_window_ref:
            _main_window_ref.show_summary()
            logger.info("[BUTTON CLICK] show_summary executed successfully")
        else:
            logger.error("[BUTTON CLICK] ERROR: Main window reference not available for show_summary")
    
    def refresh_attendance(self):
        """Refresh attendance"""
        global _main_window_ref
        logger.info("[BUTTON CLICK] refresh_attendance button clicked")
        if _main_window_ref:
            _main_window_ref._refresh_attendance_info()
            logger.info("[BUTTON CLICK] refresh_attendance executed successfully")
        else:
            logger.error("[BUTTON CLICK] ERROR: Main window reference not available for refresh_attendance")

    def refresh_notices(self):
        """Refresh announcements/notices"""
        global _main_window_ref
        logger.info("[BUTTON CLICK] refresh_notices button clicked")
        if _main_window_ref:
            _main_window_ref._refresh_notices()
            logger.info("[BUTTON CLICK] refresh_notices executed successfully")
        else:
            logger.error("[BUTTON CLICK] ERROR: Main window reference not available for refresh_notices")
    
    def get_user_profile(self):
        """Get user profile including profile picture"""
        global _main_window_ref
        logger.info("[MAIN WINDOW API] get_user_profile called")
        
        if not _main_window_ref:
            return {"success": False, "error": "Main window reference not available"}
        
        try:
            attendance = _main_window_ref.session_manager.attendance_api
            if not attendance:
                return {"success": False, "error": "Attendance API not available"}
            
            # Fetch user profile
            success, data, status_code = attendance.fetch_user_profile()
            
            if success and data:
                # Extract profile picture URL if available - check all possible field names
                profile_picture = None
                if isinstance(data, dict):
                    profile_picture = (data.get('profile_image') or 
                                     data.get('profile_picture') or 
                                     data.get('avatar') or 
                                     data.get('image_url') or 
                                     data.get('photo'))
                    
                    # Log for debugging
                    if profile_picture:
                        logger.info(f"[MAIN WINDOW API] Found profile image URL: {profile_picture}")
                    else:
                        logger.info(f"[MAIN WINDOW API] No profile image found in data. Available keys: {list(data.keys())}")
                
                return {
                    "success": True,
                    "data": {
                        "profile_image": profile_picture,  # Primary field name from API
                        "profile_picture": profile_picture,
                        "avatar": profile_picture,
                        "image_url": profile_picture,
                        "is_online": data.get('status') == 'active' if isinstance(data, dict) else True,
                        **data  # Include all other profile data
                    }
                }
            else:
                return {"success": False, "error": data.get("message", "Failed to fetch user profile") if isinstance(data, dict) else "Unknown error"}
                
        except Exception as e:
            logger.error(f"[MAIN WINDOW API] ERROR: Exception in get_user_profile - {str(e)}", exc_info=True)
            return {"success": False, "error": str(e)}
    
    def fetch_user_profile(self):
        """Alias for get_user_profile for consistency"""
        return self.get_user_profile()
    
    def get_app_icon(self, app_name, process_name=None):
        """Get application icon as base64 encoded image"""
        try:
            from utils.app_icon import get_app_icon_base64
            icon_data = get_app_icon_base64(app_name, process_name)
            if icon_data:
                return {"success": True, "icon": icon_data}
            else:
                return {"success": False, "icon": None}
        except Exception as e:
            logger.error(f"Error getting app icon: {e}", exc_info=True)
            return {"success": False, "icon": None}
    
    def refresh_appointments(self):
        """Refresh appointments - triggers UI to reload appointments"""
        global _main_window_ref
        # This method is called from JavaScript to trigger appointment refresh
        # The actual refresh is handled in JavaScript
        return {"success": True, "message": "Appointments refresh triggered"}
    
    def fetch_appointments(self, page=1, limit=10):
        """Fetch appointments for the current user"""
        global _main_window_ref
        logger.info(f"[MAIN WINDOW API] fetch_appointments called with page={page}, limit={limit}")
        
        if not _main_window_ref:
            error_msg = "Main window reference not available"
            logger.error(f"[MAIN WINDOW API] ERROR: {error_msg}")
            return {"success": False, "error": error_msg, "appointments": []}
        
        if not _main_window_ref.appointment_api:
            error_msg = "Appointment API not available"
            logger.error(f"[MAIN WINDOW API] ERROR: {error_msg}")
            return {"success": False, "error": error_msg, "appointments": []}
        
        try:
            # Get client_id and user_id
            attendance = _main_window_ref.session_manager.attendance_api
            client_id = getattr(attendance, "client_id", None)
            user_id = getattr(_main_window_ref.auth_api, "user_id", None)
            
            if not client_id:
                error_msg = "Client ID not set"
                logger.error(f"[MAIN WINDOW API] ERROR: {error_msg}")
                return {"success": False, "error": error_msg, "appointments": []}
            
            if not user_id:
                error_msg = "User ID not available"
                logger.error(f"[MAIN WINDOW API] ERROR: {error_msg}")
                return {"success": False, "error": error_msg, "appointments": []}
            
            logger.info(f"[MAIN WINDOW API] Calling appointment_api.fetch_appointments...")
            success, appointments, message = _main_window_ref.appointment_api.fetch_appointments(
                client_id=client_id,
                user_id=user_id,
                page=page,
                limit=limit
            )
            
            logger.info(f"[MAIN WINDOW API] API Response:")
            logger.info(f"[MAIN WINDOW API]   Success: {success}")
            logger.info(f"[MAIN WINDOW API]   Message: {message}")
            logger.info(f"[MAIN WINDOW API]   Appointments Count: {len(appointments) if appointments else 0}")
            
            if success:
                return {"success": True, "appointments": appointments, "message": message}
            else:
                return {"success": False, "error": message, "appointments": []}
                
        except Exception as e:
            error_msg = str(e)
            logger.error(f"[MAIN WINDOW API] ERROR: Exception - {error_msg}", exc_info=True)
            return {"success": False, "error": error_msg, "appointments": []}
    
    def set_shift(self, shift_id):
        """Set shift"""
        global _main_window_ref
        logger.info(f"[BUTTON CLICK] set_shift called with shift_id: {shift_id}")
        if _main_window_ref:
            try:
                _main_window_ref.session_manager.attendance_api.set_shift(shift_id)
                logger.info(f"[BUTTON CLICK] set_shift executed successfully - shift_id: {shift_id}")
            except Exception as e:
                logger.error(f"[BUTTON CLICK] ERROR: set_shift failed - {e}")
        else:
            logger.error("[BUTTON CLICK] ERROR: Main window reference not available for set_shift")
    
    def refresh_button_states(self):
        """Refresh button states - called for auto-refresh"""
        global _main_window_ref, _refresh_button_states_call_count
        
        # Initialize call count if not exists
        if '_refresh_button_states_call_count' not in globals():
            _refresh_button_states_call_count = 0
        
        if _main_window_ref:
            # Only sync state from server occasionally to avoid too many API calls
            # Sync every 5th call (every ~2.5 seconds) for faster updates
            _refresh_button_states_call_count += 1
            
            # Sync state from server every 5 calls (every ~2.5 seconds)
            if _refresh_button_states_call_count % 5 == 0:
                try:
                    logger.info("refresh_button_states: Syncing state from server")
                    _main_window_ref._sync_state_from_server()
                except Exception as e:
                    logger.error(f"Error syncing state in refresh_button_states: {e}")
            else:
                # Just update controls based on current state
                _main_window_ref._update_controls()
    
    def get_hours_spent_data(self):
        """Get weekly hours spent data for chart (without overtime, only work hours and break time)"""
        global _main_window_ref
        if _main_window_ref:
            return _main_window_ref._get_weekly_hours_data()
        return {"work_hours": [], "break_time": [], "days": []}
    
    def get_task_category_data(self):
        """Get task category data for chart"""
        global _main_window_ref
        if _main_window_ref:
            return _main_window_ref._get_task_category_data()
        return {"categories": [], "counts": []}
    
    def open_teams_chat(self, chat_identifier, chat_id=None):
        """Open Teams chat/conversation"""
        global _main_window_ref
        if _main_window_ref:
            return _main_window_ref._open_teams_chat(chat_identifier, chat_id)
        return False
    
    def fetch_projects(self):
        """Fetch all projects from API"""
        global _main_window_ref
        if _main_window_ref and _main_window_ref.project_api:
            try:
                result = _main_window_ref.project_api.fetch_projects()
                return result
            except Exception as e:
                logger.error(f"Error fetching projects: {e}")
                return {"success": False, "error": str(e)}
        return {"success": False, "error": "Project API not available"}
    
    def fetch_project(self, project_id):
        """
        Fetch a single project by ID
        GET /project/client/{client_id}/{project_id}
        
        Args:
            project_id: Project UUID to fetch
        """
        global _main_window_ref
        logger.info(f"[MAIN WINDOW API] fetch_project called with project_id: {project_id}")
        
        if not _main_window_ref:
            error_msg = "Main window reference not available"
            logger.error(f"[MAIN WINDOW API] ERROR: {error_msg}")
            return {"success": False, "error": error_msg, "data": None}
        
        if not _main_window_ref.project_api:
            error_msg = "Project API not available"
            logger.error(f"[MAIN WINDOW API] ERROR: {error_msg}")
            return {"success": False, "error": error_msg, "data": None}
        
        try:
            logger.info(f"[MAIN WINDOW API] Calling project_api.fetch_project...")
            result = _main_window_ref.project_api.fetch_project(project_id)
            logger.info(f"[MAIN WINDOW API] Result: {result}")
            return result
        except Exception as e:
            error_msg = str(e)
            logger.error(f"[MAIN WINDOW API] ERROR: Exception - {error_msg}", exc_info=True)
            return {"success": False, "error": error_msg, "data": None}
    
    def fetch_task(self, task_id):
        """
        Fetch a single task by ID
        GET /task/client/{client_id}/task/{task_id}
        
        Args:
            task_id: Task UUID to fetch
        """
        global _main_window_ref
        logger.info(f"[MAIN WINDOW API] fetch_task called with task_id: {task_id}")
        
        if not _main_window_ref:
            error_msg = "Main window reference not available"
            logger.error(f"[MAIN WINDOW API] ERROR: {error_msg}")
            return {"success": False, "error": error_msg, "task": None}
        
        if not _main_window_ref.task_api:
            error_msg = "Task API not available"
            logger.error(f"[MAIN WINDOW API] ERROR: {error_msg}")
            return {"success": False, "error": error_msg, "task": None}
        
        try:
            logger.info(f"[MAIN WINDOW API] Calling task_api.fetch_task...")
            success, data, message = _main_window_ref.task_api.fetch_task(task_id)
            
            logger.info(f"[MAIN WINDOW API] API Response - Success: {success}, Message: {message}")
            
            if success:
                # Extract task from response - handle different response formats
                task = None
                if isinstance(data, dict):
                    # Try common response keys in order
                    if "data" in data:
                        task_data = data["data"]
                        # If data is a dict with task inside
                        if isinstance(task_data, dict):
                            task = task_data.get("task") or task_data.get("data") or task_data
                        elif isinstance(task_data, list) and len(task_data) > 0:
                            task = task_data[0]
                        else:
                            task = task_data
                    elif "task" in data:
                        task = data["task"]
                    else:
                        # Data itself is the task
                        task = data
                elif isinstance(data, list) and len(data) > 0:
                    task = data[0]
                else:
                    task = data
                
                if task:
                    logger.info(f"[MAIN WINDOW API] SUCCESS: Task extracted successfully")
                    task_name = task.get('task_name') or task.get('name') or 'N/A'
                    logger.info(f"[MAIN WINDOW API] Task Name: {task_name}")
                else:
                    logger.warning(f"[MAIN WINDOW API] WARNING: Task data is None after extraction")
                
                return {"success": True, "task": task, "message": message}
            else:
                logger.error(f"[MAIN WINDOW API] ERROR: Task fetch failed - {message}")
                return {"success": False, "error": message, "task": None}
                
        except Exception as e:
            error_msg = str(e)
            logger.error(f"[MAIN WINDOW API] ERROR: Exception - {error_msg}", exc_info=True)
            return {"success": False, "error": error_msg, "task": None}
    
    def create_project(self, project_data):
        """Create a new project"""
        global _main_window_ref
        logger.info(f"[MAIN WINDOW API] create_project called")
        
        if not _main_window_ref:
            error_msg = "Main window reference not available"
            logger.error(f"[MAIN WINDOW API] ERROR: {error_msg}")
            return {"success": False, "error": error_msg}
        
        if not _main_window_ref.project_api:
            error_msg = "Project API not available"
            logger.error(f"[MAIN WINDOW API] ERROR: {error_msg}")
            return {"success": False, "error": error_msg}
        
        try:
            logger.info(f"[MAIN WINDOW API] Calling project_api.create_project...")
            result = _main_window_ref.project_api.create_project(project_data)
            
            if result.get("success"):
                logger.info(f"[MAIN WINDOW API] SUCCESS: Project created successfully")
            else:
                logger.error(f"[MAIN WINDOW API] ERROR: Project creation failed - {result.get('error', 'Unknown error')}")
            
            return result
        except Exception as e:
            error_msg = str(e)
            logger.error(f"[MAIN WINDOW API] ERROR: Exception creating project - {error_msg}", exc_info=True)
            return {"success": False, "error": error_msg}
    
    def fetch_tasks(self, project_id=None):
        """
        Fetch tasks assigned to the current user
        
        Args:
            project_id: Optional project ID to filter tasks by project
        """
        global _main_window_ref
        logger.info("="*80)
        logger.info("[MAIN WINDOW API] ========== FETCH TASKS START ==========")
        logger.info("="*80)
        logger.info(f"[MAIN WINDOW API] Project ID: {project_id if project_id else 'None (all tasks)'}")
        
        if not _main_window_ref:
            error_msg = "Main window reference not available"
            logger.error(f"[MAIN WINDOW API] ERROR: {error_msg}")
            return {"success": False, "error": error_msg, "tasks": [], "task_count": 0}
        
        if not _main_window_ref.task_api:
            error_msg = "Task API not available"
            logger.error(f"[MAIN WINDOW API] ERROR: {error_msg}")
            return {"success": False, "error": error_msg, "tasks": [], "task_count": 0}
        
        try:
            # Get client_id and user_id
            client_id = getattr(_main_window_ref.task_api, 'client_id', None)
            logger.info(f"[MAIN WINDOW API] Client ID: {client_id if client_id else 'None'}")
            
            user_id = None
            if _main_window_ref.auth_api:
                user_id = getattr(_main_window_ref.auth_api, 'user_id', None)
                if not user_id:
                    try:
                        user_id = _main_window_ref.auth_api.get_user_id_from_token()
                        logger.info(f"[MAIN WINDOW API] User ID extracted from token: {user_id if user_id else 'None'}")
                    except Exception as e:
                        logger.warning(f"[MAIN WINDOW API] Could not extract user_id from token: {e}")
                        pass
            else:
                logger.warning(f"[MAIN WINDOW API] auth_api not available")
            
            # Fetch tasks with optional project_id filter
            logger.info(f"[MAIN WINDOW API] Calling task_api.fetch_tasks...")
            logger.info(f"[MAIN WINDOW API] Parameters: client_id={client_id}, user_id={user_id}, project_id={project_id}")
            
            success, data, status_code = _main_window_ref.task_api.fetch_tasks(
                client_id=client_id,
                user_id=user_id,
                project_id=project_id
            )
            
            logger.info(f"[MAIN WINDOW API] API Response:")
            logger.info(f"[MAIN WINDOW API]   Success: {success}")
            logger.info(f"[MAIN WINDOW API]   Status Code: {status_code}")
            logger.info(f"[MAIN WINDOW API]   Data Type: {type(data)}")
            
            if success:
                # Parse tasks from response
                tasks = []
                if isinstance(data, dict):
                    logger.info(f"[MAIN WINDOW API] Response is dict, keys: {list(data.keys())}")
                    if 'data' in data:
                        if isinstance(data['data'], list):
                            tasks = data['data']
                            logger.info(f"[MAIN WINDOW API] Found tasks in data.data (list): {len(tasks)} tasks")
                        elif isinstance(data['data'], dict):
                            logger.info(f"[MAIN WINDOW API] data.data is dict, keys: {list(data['data'].keys())}")
                            tasks = data['data'].get('tasks', data['data'].get('items', []))
                            logger.info(f"[MAIN WINDOW API] Found tasks in data.data (dict): {len(tasks)} tasks")
                    elif 'tasks' in data:
                        tasks = data['tasks']
                        logger.info(f"[MAIN WINDOW API] Found tasks in data.tasks: {len(tasks)} tasks")
                    elif 'items' in data:
                        tasks = data['items']
                        logger.info(f"[MAIN WINDOW API] Found tasks in data.items: {len(tasks)} tasks")
                    else:
                        logger.warning(f"[MAIN WINDOW API] No tasks found in response dict")
                elif isinstance(data, list):
                    tasks = data
                    logger.info(f"[MAIN WINDOW API] Response is list: {len(tasks)} tasks")
                else:
                    logger.warning(f"[MAIN WINDOW API] Unexpected data type: {type(data)}")
                
                logger.info(f"[MAIN WINDOW API] Total tasks parsed: {len(tasks)}")
                if len(tasks) > 0:
                    import json
                    first_task = tasks[0]
                    task_sample = json.dumps({k: v for k, v in first_task.items() if k not in ['description', 'details']}, indent=2, ensure_ascii=False, default=str)[:500]
                    logger.info(f"[MAIN WINDOW API] First task sample: {task_sample}")
                
                result = {
                    "success": True,
                    "tasks": tasks,
                    "message": "Tasks fetched successfully",
                    "task_count": len(tasks)
                }
                logger.info(f"[MAIN WINDOW API] Returning success with {len(tasks)} tasks")
                logger.info("="*80)
                return result
            else:
                error_msg = data if isinstance(data, str) else "Unknown error"
                logger.error(f"[MAIN WINDOW API] API call failed: {error_msg}")
                result = {
                    "success": False,
                    "error": error_msg,
                    "tasks": [],
                    "task_count": 0
                }
                logger.info("="*80)
                return result
                    
        except Exception as e:
            error_msg = str(e)
            logger.error(f"[MAIN WINDOW API] EXCEPTION: {error_msg}", exc_info=True)
            return {"success": False, "error": error_msg, "tasks": [], "task_count": 0}
    
    def fetch_task(self, task_id):
        """
        Fetch a single task by ID
        GET /task/client/{client_id}/task/{task_id}
        
        Args:
            task_id: Task UUID to fetch
        """
        global _main_window_ref
        if _main_window_ref and _main_window_ref.task_api:
            try:
                success, data, message = _main_window_ref.task_api.fetch_task(
                    task_id=task_id
                )
                return {
                    "success": success,
                    "data": data,
                    "message": message,
                    "task": data if success else None
                }
            except Exception as e:
                logger.error(f"Error fetching task: {e}")
                return {"success": False, "error": str(e), "data": None, "task": None}
        return {"success": False, "error": "Task API not available", "data": None, "task": None}
    
    def fetch_task_statuses(self, page=1, limit=50):
        """
        Fetch all task statuses for the current client.
        Returns a dict shaped for JS consumption: {success, statuses, message}
        """
        global _main_window_ref
        logger.info(f"[MAIN WINDOW API] fetch_task_statuses called (page={page}, limit={limit})")
        
        if not _main_window_ref:
            error_msg = "Main window reference not available"
            logger.error(f"[MAIN WINDOW API] ERROR: {error_msg}")
            return {"success": False, "error": error_msg, "statuses": [], "message": error_msg}
        
        if not _main_window_ref.task_api:
            error_msg = "Task API not available"
            logger.error(f"[MAIN WINDOW API] ERROR: {error_msg}")
            return {"success": False, "error": error_msg, "statuses": [], "message": error_msg}
        
        try:
            success, statuses, message = _main_window_ref.task_api.fetch_task_statuses()
            logger.info(f"[MAIN WINDOW API] fetch_task_statuses result: success={success}, count={len(statuses) if statuses else 0}")
            return {
                "success": success,
                "statuses": statuses if success else [],
                "message": message
            }
        except Exception as e:
            error_msg = str(e)
            logger.error(f"[MAIN WINDOW API] ERROR: Exception fetching task statuses - {error_msg}", exc_info=True)
            return {"success": False, "error": error_msg, "statuses": [], "message": error_msg}
    
    def update_task(self, task_id, task_data):
        """
        Update a task
        
        Args:
            task_id: Task UUID to update
            task_data: Dictionary containing task fields to update
        """
        global _main_window_ref
        logger.info(f"[MAIN WINDOW API] update_task called with task_id: {task_id}")
        
        if not _main_window_ref:
            error_msg = "Main window reference not available"
            logger.error(f"[MAIN WINDOW API] ERROR: {error_msg}")
            return {"success": False, "error": error_msg}
        
        if not _main_window_ref.task_api:
            error_msg = "Task API not available"
            logger.error(f"[MAIN WINDOW API] ERROR: {error_msg}")
            return {"success": False, "error": error_msg}
        
        try:
            logger.info(f"[MAIN WINDOW API] Calling task_api.update_task...")
            success, data, message = _main_window_ref.task_api.update_task(
                task_id=task_id,
                task_data=task_data
            )
            
            if success:
                logger.info(f"[MAIN WINDOW API] SUCCESS: Task updated successfully - {task_id}")
            else:
                logger.error(f"[MAIN WINDOW API] ERROR: Task update failed - {message}")
            
            return {
                "success": success,
                "data": data,
                "message": message
            }
        except Exception as e:
            error_msg = str(e)
            logger.error(f"[MAIN WINDOW API] ERROR: Exception updating task - {error_msg}", exc_info=True)
            return {"success": False, "error": error_msg}
    
    def maximize_window(self):
        """Maximize the window"""
        global _main_window_ref
        if _main_window_ref and _main_window_ref.window:
            try:
                # Use win32gui to maximize window on Windows
                try:
                    import platform
                    if platform.system() == 'Windows':
                        import win32gui
                        import win32con
                        
                        # Find window by title
                        def enum_handler(hwnd, ctx):
                            if win32gui.IsWindowVisible(hwnd):
                                title = win32gui.GetWindowText(hwnd)
                                if 'Productivity Tracker' in title or 'Productivity' in title:
                                    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                                    return False  # Stop enumeration
                            return True
                        
                        win32gui.EnumWindows(enum_handler, None)
                        return {"success": True}
                    else:
                        # For non-Windows, try pywebview's maximize if available
                        if hasattr(_main_window_ref.window, 'maximize'):
                            _main_window_ref.window.maximize()
                            return {"success": True}
                except ImportError:
                    logger.warning("win32gui not available, cannot maximize window")
                except Exception as win_e:
                    logger.warning(f"Could not maximize via win32: {win_e}")
                
                return {"success": True}
            except Exception as e:
                logger.error(f"Error maximizing window: {e}")
                return {"success": False, "error": str(e)}
        return {"success": False, "error": "Window not available"}

class MainWindow:
    def __init__(self, session_manager, user_info, on_clock_out, auto_clock_in=False, on_logout=None, project_api=None, task_api=None, auth_api=None, appointment_api=None):
        global _main_window_ref
        self.session_manager = session_manager
        self.user_info = user_info
        self.on_clock_out = on_clock_out
        self.auto_clock_in = auto_clock_in
        self.on_logout = on_logout
        self.project_api = project_api
        self.task_api = task_api
        self.auth_api = auth_api
        self.appointment_api = appointment_api
        self._break_auto = False
        self._auto_break_triggered = False
        self.latest_tab_snapshot = None
        self._last_idle_warning_time = None
        self._shift_id_by_name = {}
        self._shift_display_by_id = {}
        self._user_assigned_shift_ids = set()
        self.window = None
        self.tracker = None
        self._is_force_break_mode = False
        self._update_thread = None
        self._running = True
        # Track last seen messages for detecting new ones
        self._last_seen_messages = set()  # Set of (sender, message, time) tuples
        # Track which messages have been notified to avoid duplicate notifications
        self._notified_messages = set()  # Set of (sender, message) tuples
        _main_window_ref = self
    
    def _get_user_profile_for_js(self):
        """
        Fetch user profile from ERP and return a dict for JS:
        {
            "first_name": "...",
            "last_name": "...",
            "raw": <full API payload>
        }
        """
        attendance = self.session_manager.attendance_api
        client_id = getattr(attendance, "client_id", None)
        ok, data, _ = attendance.fetch_user_profile(client_id=client_id)
        if not ok or data is None:
            return {"first_name": "", "last_name": "", "raw": data}

        source = data
        # If wrapped in {"data": ...} or {"user": ...}
        if isinstance(source, dict):
            if isinstance(source.get("data"), (dict, list)):
                source = source["data"]
            elif isinstance(source.get("user"), (dict, list)):
                source = source["user"]

        # If list, use first dict element
        if isinstance(source, list):
            for item in source:
                if isinstance(item, dict):
                    source = item
                    break

        if not isinstance(source, dict):
            return {"first_name": "", "last_name": "", "raw": data}

        profile = source.get("profile") or {}
        first_name = (
            profile.get("first_name")
            or profile.get("firstName")
            or source.get("first_name")
            or source.get("firstName")
            or ""
        )
        last_name = (
            profile.get("last_name")
            or profile.get("lastName")
            or source.get("last_name")
            or source.get("lastName")
            or ""
        )

        # Fallback to "name" split if needed
        if (not first_name or not last_name) and isinstance(source.get("name"), str):
            parts = source["name"].strip().split()
            if parts:
                first_name = first_name or parts[0]
            if len(parts) > 1:
                last_name = last_name or " ".join(parts[1:])

        # Trim extraneous whitespace
        first_name = (first_name or "").strip()
        last_name = (last_name or "").strip()

        return {
            "first_name": first_name or "",
            "last_name": last_name or "",
            "raw": data,
        }

    def show(self):
        # Get the HTML file path (works in both dev and PyInstaller)
        html_path = get_html_path('main_window.html')
        
        # Create API instance
        api = MainWindowAPI()
        
        # Create webview window
        self.window = webview.create_window(
            'Productivity Tracker',
            html_path,
            width=1400,
            height=900,
            min_size=(1200, 700),
            resizable=True,
            js_api=api
        )
        
        # Start activity tracker
        self.tracker = ActivityTracker(on_activity=self.update_activity)
        self.tracker.start()
        
        # Start update thread
        self._start_update_thread()
        
        # Start webview (blocking)
        self._running = True
        webview.start(debug=False)
        self._running = False
    
    def _start_update_thread(self):
        """Start background thread for UI updates - optimized for performance"""
        def update_loop():
            attendance_refresh_counter = 0
            button_refresh_counter = 0
            messages_refresh_counter = 0
            dashboard_refresh_counter = 0
            activity_tick_counter = 0
            while self._running:
                try:
                    # Activity tick every 5 seconds (reduced from every second)
                    activity_tick_counter += 1
                    if activity_tick_counter >= 5:
                        activity_tick_counter = 0
                        self._tick_activity()
                    
                    # Dashboard metrics every 30 seconds (reduced frequency to prevent hanging)
                    dashboard_refresh_counter += 1
                    if dashboard_refresh_counter >= 30:
                        dashboard_refresh_counter = 0
                        self._refresh_dashboard_metrics()
                        self._update_idle_warning_label()
                    
                    # Auto-refresh button states every 5 seconds (reduced from 2 seconds)
                    button_refresh_counter += 1
                    if button_refresh_counter >= 5:
                        button_refresh_counter = 0
                        self._update_controls()
                    
                    # Auto-refresh Teams messages every 5 seconds (reduced from 1 second)
                    messages_refresh_counter += 1
                    if messages_refresh_counter >= 5:
                        messages_refresh_counter = 0
                        self._update_teams_messages()
                    
                    # Auto-refresh attendance every 60 seconds (reduced from 30 seconds)
                    attendance_refresh_counter += 1
                    if attendance_refresh_counter >= 60:
                        attendance_refresh_counter = 0
                        self._refresh_attendance_info()
                    
                    # Sleep for 1 second but operations run less frequently
                    time.sleep(1)
                except Exception:
                    pass
        
        self._update_thread = threading.Thread(target=update_loop, daemon=True)
        self._update_thread.start()
    
    def _sync_tracker_identity(self):
        try:
            user_id = (getattr(self.session_manager.attendance_api.auth_api, "user_id", None)
                       or (self.user_info or {}).get("user_id"))
        except Exception:
            user_id = (self.user_info or {}).get("user_id")
        client_id = getattr(self.session_manager.attendance_api, "client_id", None)
        if self.tracker:
            self.tracker.update_context(
                user_id=user_id,
                client_id=client_id,
                webcam_name=WEBCAM_DEVICE_NAME,
            )
    
    def _update_status(self, status):
        """Update status in UI"""
        if self.window:
            try:
                self.window.evaluate_js(f'updateStatus("{status}")')
            except Exception:
                pass
    
    def _update_button_states(self, states):
        """Update button states in UI - optimized logging"""
        # Reduced logging verbosity for performance
        if self.window:
            try:
                import json
                js_code = f'updateButtonStates({json.dumps(states)})'
                self.window.evaluate_js(js_code)
            except Exception as e:
                logger.error(f"Error updating button states: {e}")
        else:
            logger.warning("Window not available, cannot update button states")
    
    def _update_break_time(self, time_str):
        """Update break time in UI"""
        if self.window:
            try:
                self.window.evaluate_js(f'updateBreakTime("{time_str}")')
            except Exception:
                pass
    
    def _update_active_apps(self, apps):
        """Update active apps in UI"""
        if self.window:
            try:
                self.window.evaluate_js(f'updateActiveApps("{apps}")')
            except Exception:
                pass
    
    def _update_idle_warning_label(self):
        """Update idle warning label"""
        if not self.window:
            return
        try:
            if self._last_idle_warning_time:
                display = self._last_idle_warning_time.strftime("%I:%M %p")
                self.window.evaluate_js(f'updateIdleWarning("triggered at {display}")')
            else:
                self.window.evaluate_js('updateIdleWarning("none")')
        except Exception:
            pass
    
    def _update_activity_counts(self):
        """Update mouse clicks and key counts in UI"""
        if not self.window or not self.tracker:
            return
        try:
            counts = self.tracker.get_activity_counts()
            mouse_clicks = counts.get("mouse_clicks", 0)
            keys_count = counts.get("keys_count", 0)
            
            import json
            js_code = f'updateActivityCounts({json.dumps(counts)})'
            self.window.evaluate_js(js_code)
        except Exception as e:
            logger.error(f"Error updating activity counts: {e}")

    def _update_app_usage_list(self):
        """Update application usage overview with real tracked data."""
        if not self.window:
            return
        try:
            apps_list = self.session_manager.get_app_usage_stats() or []
            
            # get_app_usage_stats returns a list of dicts with: name, time, seconds
            items = []
            for app in apps_list:
                if isinstance(app, dict):
                    app_name = app.get("name", "Unknown")
                    time_str = app.get("time", "0m")  # Already formatted as "Xh Ym" or "Ym"
                    # Convert to HH:MM:SS format for display
                    seconds = app.get("seconds", 0)
                    def format_hms(secs: float) -> str:
                        try:
                            total = int(secs)
                            hrs = total // 3600
                            mins = (total % 3600) // 60
                            secs = total % 60
                            return f"{hrs:02d}:{mins:02d}:{secs:02d}"
                        except Exception:
                            return "00:00:00"
                    
                    items.append({
                        "name": app_name,
                        "category": "Application",
                        "duration": format_hms(seconds),
                        "icon": "[PC]",
                    })
            
            import json
            self.window.evaluate_js(f'updateAppUsage({json.dumps(items)})')
        except Exception as e:
            logger.error(f"Error updating app usage list: {e}")

    def _update_notices(self, notices):
        """Send notices to the webview"""
        if not self.window:
            return
        try:
            import json
            js_code = f'updateNotices({json.dumps(notices)})'
            self.window.evaluate_js(js_code)
        except Exception as e:
            logger.error(f"Error updating notices: {e}")

    def _update_teams_messages(self):
        """Update Teams messages in the Messages section"""
        if not self.window:
            return
        try:
            from utils.teams_notifications import get_teams_messages
            from datetime import datetime
            
            # Get access token if available
            access_token = None
            try:
                auth_api = self.session_manager.attendance_api.auth_api
                access_token = getattr(auth_api, 'access_token', None)
            except Exception:
                pass
            
            # Get Teams messages
            messages = get_teams_messages(access_token=access_token)
            
            # Debug logging
            if messages:
                logger.info(f"Found {len(messages)} Teams messages")
                for msg in messages:
                    logger.info(f"  - {msg.get('sender', 'Unknown')}: {msg.get('message', '')[:50]}")
            else:
                logger.info("No Teams messages found")
            
            # Track current messages to detect new ones
            current_message_keys = set()
            
            # Format messages for UI - show ALL messages, but highlight new ones
            formatted_messages = []
            for msg in messages:
                sender = msg.get('sender', 'Unknown')
                message = msg.get('message', '')
                time_str = msg.get('time', datetime.now().strftime('%I:%M %p'))
                chat_id = msg.get('chat_id', '')
                chat_identifier = msg.get('chat_identifier', sender)
                
                # Skip if no sender or empty message
                if not sender or sender == 'Unknown':
                    continue
                
                # If message is empty or placeholder, use a default
                if not message or message in ['Personal chat', 'New message', 'Active chat']:
                    message = 'New message'
                
                # Create unique key for this message - use sender + message content (not time)
                # This allows same sender's new messages to be detected
                # Use full message (not truncated) for better uniqueness
                full_msg = msg.get('full_message', message) if 'full_message' in msg else message
                message_key = (sender, full_msg)
                current_message_keys.add(message_key)
                
                # Check if this is a new message (not seen before)
                is_new = message_key not in self._last_seen_messages
                
                # Show Windows notification ONLY for truly new messages
                # Must be: new message + has actual content + not a placeholder + not already notified
                if is_new and message and message not in ['Personal chat', 'New message', 'Active chat']:
                    # Check if we've already notified about this exact message
                    notification_key = (sender, full_msg)
                    if notification_key not in self._notified_messages:
                        try:
                            from utils.windows_notifications import show_teams_notification
                            # Show notification in background thread to avoid blocking
                            import threading
                            threading.Thread(
                                target=show_teams_notification,
                                args=(sender, message, 5),
                                daemon=True
                            ).start()
                            # Mark as notified immediately to prevent duplicates
                            self._notified_messages.add(notification_key)
                            logger.info(f"Shown notification for NEW message from {sender}: {message[:50]}")
                        except Exception as e:
                            logger.debug(f"Failed to show notification: {e}")
                
                # Truncate message if too long (but keep full message for tooltip)
                display_message = message
                if len(message) > 50:
                    display_message = message[:47] + '...'
                
                formatted_messages.append({
                    'sender': sender,
                    'message': display_message,
                    'full_message': message,  # Keep full message
                    'time': time_str,
                    'chat_id': chat_id,
                    'chat_identifier': chat_identifier,
                    'is_new': is_new  # Flag for new messages
                })
            
            # Sort by time (newest first) - put new messages at top
            formatted_messages.sort(key=lambda x: (
                not x.get('is_new', False),  # New messages first
                x.get('time', '')
            ), reverse=True)
            
            # Update last seen messages - merge with existing to allow name updates
            # Only track messages from last 2 minutes to allow updates
            # This ensures names can update when window titles change
            self._last_seen_messages = current_message_keys
            
            # Limit size to prevent memory bloat (keep last 100)
            if len(self._last_seen_messages) > 100:
                # Convert to list, take last 100, convert back to set
                self._last_seen_messages = set(list(self._last_seen_messages)[-100:])
            
            # Also limit notified messages to prevent memory bloat
            if len(self._notified_messages) > 200:
                # Keep last 200 notified messages
                self._notified_messages = set(list(self._notified_messages)[-200:])
            
            # Update UI
            import json
            js_code = f'updateMessages({json.dumps(formatted_messages)})'
            self.window.evaluate_js(js_code)
        except Exception as e:
            logger.error(f"Error updating Teams messages: {e}")
    
    def _open_teams_chat(self, chat_identifier, chat_id=None):
        """Open Teams chat/conversation"""
        try:
            import subprocess
            import win32gui
            import win32process
            
            # Method 1: Try to find and activate Teams window with this chat
            if win32gui:
                def find_teams_window(hwnd, results):
                    try:
                        window_title = win32gui.GetWindowText(hwnd)
                        if window_title and ('Teams' in window_title or 'Microsoft Teams' in window_title):
                            # Check if this window matches the chat
                            if chat_identifier.lower() in window_title.lower():
                                results.append(hwnd)
                    except Exception:
                        pass
                    return True
                
                teams_windows = []
                win32gui.EnumWindows(find_teams_window, teams_windows)
                
                if teams_windows:
                    # Activate the first matching window
                    hwnd = teams_windows[0]
                    try:
                        win32gui.ShowWindow(hwnd, 9)  # SW_RESTORE
                        win32gui.SetForegroundWindow(hwnd)
                        win32gui.BringWindowToTop(hwnd)
                        return True
                    except Exception:
                        pass
            
            # Method 2: Use Teams protocol handler (msteams://)
            try:
                # Teams deep link format: msteams://teams.microsoft.com/l/chat/0/0?tenantId=...&topicName=...
                # For now, just open Teams and let user navigate
                # In production, you'd construct proper deep link with chat ID
                url = 'msteams://'
                subprocess.Popen(['start', url], shell=True)
                
                # Also try to find and activate Teams window
                import time
                time.sleep(1)  # Wait for Teams to open
                
                # Try to find Teams window and bring to front
                if win32gui:
                    def find_teams_window(hwnd, results):
                        try:
                            window_title = win32gui.GetWindowText(hwnd)
                            if window_title and ('Teams' in window_title or 'Microsoft Teams' in window_title):
                                results.append(hwnd)
                        except Exception:
                            pass
                        return True
                    
                    teams_windows = []
                    win32gui.EnumWindows(find_teams_window, teams_windows)
                    
                    if teams_windows:
                        hwnd = teams_windows[0]
                        try:
                            win32gui.ShowWindow(hwnd, 9)  # SW_RESTORE
                            win32gui.SetForegroundWindow(hwnd)
                            win32gui.BringWindowToTop(hwnd)
                        except Exception:
                            pass
                
                return True
            except Exception as e:
                logger.error(f"Error opening Teams via protocol: {e}")
            
            # Method 3: Launch Teams app directly
            try:
                localappdata = os.getenv('LOCALAPPDATA', '')
                programfiles = os.getenv('PROGRAMFILES', '')
                programfiles_x86 = os.getenv('PROGRAMFILES(X86)', '')
                
                teams_paths = [
                    os.path.join(localappdata, 'Microsoft', 'Teams', 'Update.exe'),
                    os.path.join(programfiles_x86, 'Microsoft', 'Teams', 'current', 'Teams.exe'),
                    os.path.join(programfiles, 'Microsoft', 'Teams', 'current', 'Teams.exe'),
                ]
                
                for path in teams_paths:
                    if os.path.exists(path):
                        if path.endswith('Update.exe'):
                            # Update.exe needs --processStart Teams.exe argument
                            subprocess.Popen([path, '--processStart', 'Teams.exe'])
                        else:
                            subprocess.Popen([path])
                        return True
            except Exception as e:
                logger.error(f"Error launching Teams: {e}")
            
            return False
        except Exception as e:
            logger.error(f"Error opening Teams chat: {e}")
            return False
    
    def _refresh_notices(self):
        """Fetch notices for the current client and push to UI"""
        attendance = self.session_manager.attendance_api
        client_id = getattr(attendance, "client_id", None)
        if not client_id:
            return
        try:
            ok, data, status = attendance.fetch_notices(client_id, page=1, limit=10)
            if not ok or not isinstance(data, dict):
                logger.warning(f"Failed to fetch notices (status {status}): {data}")
                self._update_notices([])
                return
            payload = data.get("data") or {}
            items = payload.get("notices") or payload.get("data") or payload.get("items") or []
            notices = []
            for it in items:
                if not isinstance(it, dict):
                    continue
                notices.append({
                    "title": it.get("title") or "Untitled",
                    "message": it.get("message") or "",
                    "valid_till": it.get("valid_till") or "",
                    "status": it.get("status") or "",
                })
            self._update_notices(notices)
        except Exception as e:
            logger.error(f"Error refreshing notices: {e}")
            self._update_notices([])
    
    def _load_shifts(self):
        """Load shifts and update UI"""
        try:
            attendance = self.session_manager.attendance_api
            items = []
            user_assigned_shift_ids = set()
            
            user_id = attendance.auth_api.user_id
            if user_id:
                ok, data, _ = attendance.fetch_user_shifts(attendance.client_id, user_id)
                if ok and isinstance(data, dict):
                    user_shifts = data.get("shifts") or []
                    items = user_shifts
                    for shift in user_shifts:
                        shift_id = shift.get("id") or shift.get("shift_id")
                        if shift_id:
                            user_assigned_shift_ids.add(shift_id)
                    if user_shifts:
                        for shift in user_shifts:
                            if shift.get("is_current", False):
                                shift_id = shift.get("id") or shift.get("shift_id")
                                if shift_id:
                                    attendance.set_shift(shift_id)
                                    break
                        if not attendance.selected_shift_id and user_shifts:
                            shift_id = user_shifts[0].get("id") or user_shifts[0].get("shift_id")
                            if shift_id:
                                attendance.set_shift(shift_id)
            
            if not items:
                ok, data, _ = attendance.fetch_shifts(attendance.client_id)
                if ok and isinstance(data, dict):
                    items = data.get("data") or data.get("items") or data.get("shifts") or []
            
            displays = []
            self._shift_id_by_name = {}
            self._shift_display_by_id = {}
            self._user_assigned_shift_ids = user_assigned_shift_ids
            
            for it in items:
                if not isinstance(it, dict):
                    continue
                name = (
                    it.get("shift_type")
                    or it.get("name")
                    or it.get("shift_name")
                    or it.get("title")
                    or (isinstance(it.get("location_data"), dict) and it.get("location_data", {}).get("name"))
                    or it.get("id")
                )
                sid = it.get("id") or it.get("shift_id")
                if name and sid:
                    in_t = it.get("in_time")
                    out_t = it.get("out_time")
                    if in_t or out_t:
                        in_ist = self._format_shift_time_ist(in_t)
                        out_ist = self._format_shift_time_ist(out_t)
                        formatted_range = f"{in_ist or '--'} - {out_ist or '--'}"
                        display = f"{name} ({formatted_range})"
                    else:
                        display = f"{name}"
                    displays.append({"id": str(sid), "name": display})
                    self._shift_id_by_name[display] = sid
                    self._shift_display_by_id[str(sid)] = display
            
            if displays and self.window:
                selected_id = str(attendance.selected_shift_id) if attendance.selected_shift_id else None
                import json
                self.window.evaluate_js(f'updateShifts({json.dumps(displays)}, {json.dumps(selected_id)})')
        except Exception as e:
            logger.error(f"Load shifts error: {e}")
    
    def _format_shift_time_ist(self, shift_time):
        """Convert server-provided HH:MM[:SS] times to IST-friendly display."""
        if not shift_time:
            return ""
        value = str(shift_time).strip()
        if not value:
            return ""
        lowered = value.lower()
        if "am" in lowered or "pm" in lowered or "ist" in lowered or "gmt" in lowered:
            return value
        try:
            fmt = "%H:%M:%S" if value.count(":") >= 2 else "%H:%M"
            base = datetime.strptime(value, fmt)
            dummy = datetime(2000, 1, 1, base.hour, base.minute, base.second)
            ist_time = dummy + timedelta(hours=5, minutes=30)
            return ist_time.strftime("%I:%M %p IST")
        except Exception:
            return value
    
    def clock_in(self):
        """Clock in - preserves all original logic"""
        # Validate: Can only clock in if logged out
        if self.session_manager.state != SessionState.LOGGED_OUT:
            try:
                if self.session_manager.state == SessionState.CLOCKED_IN:
                    messagebox.showinfo("Clock In", "You are already clocked in.")
                elif self.session_manager.state == SessionState.ON_BREAK:
                    messagebox.showinfo("Clock In", "You are currently on break. Please end break first.")
            except Exception:
                pass
            return
        
        try:
            if getattr(self.session_manager.attendance_api, "machine_punch_required", False):
                try:
                    messagebox.showinfo("Clock In Not Allowed", "Your profile requires Machine Punch. Please use the attendance machine.")
                except Exception:
                    pass
                return
        except Exception:
            pass
        
        try:
            a = self.session_manager.attendance_api
            user_id = a.auth_api.user_id
            user_assigned_shift_ids = getattr(self, "_user_assigned_shift_ids", set())
            selected_shift_id = None
            
            if user_id:
                ok, user_shifts_data, _ = a.fetch_user_shifts(a.client_id, user_id)
                if ok and isinstance(user_shifts_data, dict):
                    user_shifts = user_shifts_data.get("shifts") or []
                    if user_shifts:
                        user_assigned_shift_ids = {shift.get("id") or shift.get("shift_id") for shift in user_shifts if shift.get("id") or shift.get("shift_id")}
                        self._user_assigned_shift_ids = user_assigned_shift_ids
                        current_shift_id = a.selected_shift_id
                        if current_shift_id and current_shift_id not in user_assigned_shift_ids:
                            current_shift_id = None
                        if not current_shift_id:
                            for shift in user_shifts:
                                if shift.get("is_current", False):
                                    selected_shift_id = shift.get("id") or shift.get("shift_id")
                                    if selected_shift_id:
                                        break
                            if not selected_shift_id and user_shifts:
                                selected_shift_id = user_shifts[0].get("id") or user_shifts[0].get("shift_id")
                        else:
                            selected_shift_id = current_shift_id
            
            if selected_shift_id:
                a.set_shift(selected_shift_id)
            elif not a.selected_shift_id:
                a.auto_select_shift(a.client_id, use_user_shifts=True)
        except Exception:
            try:
                a = self.session_manager.attendance_api
                if not a.selected_shift_id:
                    a.auto_select_shift(a.client_id, use_user_shifts=True)
            except Exception:
                pass
        
        try:
            from datetime import datetime, timezone, timedelta
            from dateutil import parser
            
            # Check last punch-out time from server before attempting punch-in
            last_punch_out_time = None
            try:
                from datetime import date
                attendance = self.session_manager.attendance_api
                client_id = attendance.client_id
                today = date.today()
                ok_att, payload, _ = attendance.fetch_attendance(client_id=client_id, month=today.month, year=today.year)
                
                if ok_att and isinstance(payload, (dict, list)):
                    records = []
                    base = payload
                    if isinstance(base, dict):
                        for key in ("attendances", "data", "items", "rows", "attendance", "result", "records"):
                            val = base.get(key)
                            if isinstance(val, list):
                                records = val
                                break
                        if not records and isinstance(base.get("data"), dict):
                            records = [base.get("data")]
                        if not records:
                            records = [base]
                    else:
                        records = base
                    
                    # Find the most recent punch-out time
                    for r in records:
                        if not isinstance(r, dict):
                            continue
                        out_time = r.get("out_time") or r.get("punch_out") or r.get("punchOutTime") or r.get("clock_out_time")
                        if out_time:
                            try:
                                out_dt = parser.parse(out_time)
                                if out_dt.tzinfo is None:
                                    out_dt = out_dt.replace(tzinfo=timezone.utc)
                                if last_punch_out_time is None or out_dt > last_punch_out_time:
                                    last_punch_out_time = out_dt
                            except Exception:
                                pass
            except Exception:
                pass
            
            # Get current time
            now_utc = datetime.now(timezone.utc)
            
            # If last punch-out exists and current time is before it, add a small buffer
            if last_punch_out_time and now_utc <= last_punch_out_time:
                # Add 1 minute buffer after last punch-out to ensure punch-in is later
                now_utc = last_punch_out_time + timedelta(minutes=1)
                try:
                    messagebox.showwarning(
                        "Clock In Time Adjusted",
                        f"Your punch-in time has been adjusted to be after your last punch-out time ({last_punch_out_time.strftime('%I:%M %p')})."
                    )
                except Exception:
                    pass
            
            date_iso = now_utc.date().isoformat()
            in_time_iso_utc = now_utc.replace(microsecond=0).isoformat().replace("+00:00", "Z")
            ok, data, status = self.session_manager.attendance_api.punch_in(
                client_id=self.session_manager.attendance_api.client_id,
                date_in_iso_format=date_iso,
                in_time_iso_utc=in_time_iso_utc,
                shift_id=self.session_manager.attendance_api.selected_shift_id,
            )
            if not ok:
                msg = data.get("message") if isinstance(data, dict) else str(data)
                if isinstance(msg, str) and "machine punch" in msg.lower():
                    try:
                        self.session_manager.attendance_api.machine_punch_required = True
                        self._update_status("Status: MACHINE PUNCH REQUIRED — Use attendance machine")
                        self._update_button_states({
                            "clockIn": False,
                            "clockOut": False,
                            "startBreak": False,
                            "endBreak": False
                        })
                        messagebox.showinfo("Clock In Not Allowed", "Your profile requires Machine Punch. Please use the attendance machine.")
                    except Exception:
                        pass
                    return
                
                # Handle "already punched in" error - sync state from server
                if isinstance(msg, str) and ("already punched in" in msg.lower() or "already punched" in msg.lower() or "punch out before" in msg.lower()):
                    try:
                        # User is already clocked in on server, sync state
                        logger.info("Server indicates user is already punched in, syncing state from server")
                        # Sync state from server first - this will update state to CLOCKED_IN
                        self._sync_state_from_server()
                        # Force update controls to reflect the synced state
                        self._update_controls()
                        # Refresh attendance info
                        self._refresh_attendance_info()
                        # Double-check state and force update if needed
                        if self.session_manager.state == SessionState.CLOCKED_IN:
                            logger.info("State successfully synced to CLOCKED_IN, Clock Out button should be visible")
                        else:
                            logger.warning(f"State sync may have failed, current state: {self.session_manager.state}")
                            # Force state update
                            self.session_manager.state = SessionState.CLOCKED_IN
                            self._update_status("Status: CLOCKED IN")
                            self._update_controls()
                        # Show info message instead of error
                        try:
                            messagebox.showinfo(
                                "Already Clocked In",
                                "You are already clocked in. Your status has been updated. Please use Clock Out to end your session."
                            )
                        except Exception:
                            pass
                    except Exception as e:
                        logger.error(f"Error syncing state after 'already punched in' error: {e}")
                        import traceback
                        traceback.print_exc()
                        # Even if sync fails, try to update state manually
                        try:
                            # If server says we're punched in, update local state
                            self.session_manager.state = SessionState.CLOCKED_IN
                            self._update_status("Status: CLOCKED IN")
                            self._update_controls()
                            logger.info("Manually set state to CLOCKED_IN after sync failure")
                            try:
                                messagebox.showinfo(
                                    "Already Clocked In",
                                    "You are already clocked in. Please use Clock Out to end your session."
                                )
                            except Exception:
                                pass
                        except Exception as e2:
                            logger.error(f"Error manually updating state: {e2}")
                            try:
                                messagebox.showerror("Clock In Failed", f"Status {status}: {msg}")
                            except Exception:
                                pass
                    return
                
                # Handle "punch in time must be later than last punch out" error specifically
                if isinstance(msg, str) and ("punch in time" in msg.lower() and "punch out time" in msg.lower()):
                    try:
                        # Extract the last punch out time from error message if possible
                        import re
                        match = re.search(r'(\d{1,2}:\d{2}\s*(?:am|pm))', msg, re.IGNORECASE)
                        if match:
                            last_out_str = match.group(1)
                            messagebox.showerror(
                                "Clock In Failed",
                                f"Your punch-in time must be after your last punch-out time ({last_out_str}).\n\nPlease wait a moment and try again, or contact your administrator if this issue persists."
                            )
                        else:
                            messagebox.showerror("Clock In Failed", f"Status {status}: {msg}")
                    except Exception:
                        messagebox.showerror("Clock In Failed", f"Status {status}: {msg}")
                else:
                    try:
                        messagebox.showerror("Clock In Failed", f"Status {status}: {msg}")
                    except Exception:
                        pass
                return
        except Exception as e:
            logger.error(f"Clock in error: {e}")
        
        self.session_manager.clock_in_local()
        self._update_status("Status: CLOCKED IN")
        # Immediately update button states to disable Clock In and enable Clock Out
        self._update_controls()
        try:
            messagebox.showinfo("Clock In", "You are clocked in.")
        except Exception:
            pass
        self._refresh_attendance_info()
    
    def clock_out(self):
        """Clock out - only works if punched in"""
        # Validate: Can only clock out if already clocked in
        if self.session_manager.state == SessionState.LOGGED_OUT:
            try:
                messagebox.showwarning("Clock Out", "Please Clock In first before Clocking Out.")
            except Exception:
                pass
            return
        
        result = self.session_manager.clock_out()
        # Update state to LOGGED_OUT
        self.session_manager.state = SessionState.LOGGED_OUT
        self._update_status("Status: LOGGED OUT")
        # Immediately update button states to enable Clock In and disable Clock Out
        self._update_controls()
        logger.info("After clock out, state is: LOGGED_OUT, Clock In button should be visible")
        if isinstance(result, tuple) and len(result) == 3:
            ok, data, status = result
            if ok:
                try:
                    messagebox.showinfo("Clock Out", "ERP attendance punch-out successful.")
                except Exception:
                    pass
            else:
                try:
                    messagebox.showerror("Clock Out Failed", f"Status {status}: {data}")
                except Exception:
                    pass
        else:
            try:
                messagebox.showinfo("Clock Out", "You are clocked out.")
            except Exception:
                pass
        self._refresh_attendance_info()
        if self.on_clock_out:
            self.on_clock_out()
    
    def logout(self):
        """Logout"""
        try:
            if self.session_manager.state != SessionState.LOGGED_OUT:
                self.session_manager.clock_out()
        except Exception:
            pass
        try:
            if callable(self.on_logout):
                self.on_logout()
        finally:
            try:
                if self.window:
                    self.window.destroy()
            except Exception:
                pass
    
    def start_break(self):
        """Start break - only works if punched in"""
        # Validate: Can only start break if clocked in
        if self.session_manager.state == SessionState.LOGGED_OUT:
            try:
                messagebox.showwarning("Start Break", "Please Clock In first before starting a break.")
            except Exception:
                pass
            return
        
        if self.session_manager.state == SessionState.ON_BREAK:
            try:
                messagebox.showinfo("Start Break", "You are already on break.")
            except Exception:
                pass
            return
        
        self._break_auto = False
        self._auto_break_triggered = False
        if self.session_manager.start_break():
            self._update_status("Status: ON BREAK")
            self.force_break_popup()
            self._refresh_attendance_info()
            self._update_controls()
        else:
            try:
                messagebox.showwarning("Start Break", "Unable to start break. Please ensure you are clocked in.")
            except Exception:
                pass
    
    def end_break(self):
        """End break - only works if break is active"""
        # Validate: Can only end break if currently on break
        if self.session_manager.state != SessionState.ON_BREAK:
            try:
                messagebox.showwarning("End Break", "You are not currently on a break.")
            except Exception:
                pass
            return
        
        if self.session_manager.end_break():
            self._update_status("Status: CLOCKED IN")
            self._break_auto = False
            self._auto_break_triggered = False
            self._is_force_break_mode = False
            self._refresh_attendance_info()
            self._update_controls()
        else:
            try:
                messagebox.showwarning("End Break", "Unable to end break.")
            except Exception:
                pass
    
    def show_summary(self):
        """Show summary"""
        ReportsWindow(self.session_manager.attendance_api, self.session_manager).show()
    
    def update_activity(self):
        """Update activity"""
        pass
    
    def _tick_activity(self):
        """Tick activity"""
        try:
            idle = self.tracker.get_idle_time()
            status = "idle" if idle >= IDLE_TIMEOUT_SECONDS else "active"
            self.session_manager.update_activity(is_active=(status == "active"), idle_seconds=idle)
            
            if status == "idle" and self.session_manager.state == SessionState.CLOCKED_IN:
                if not hasattr(self, '_auto_break_triggered') or not self._auto_break_triggered:
                    self._auto_break_triggered = True
                    self._auto_start_break()
                if not self._last_idle_warning_time:
                    self._last_idle_warning_time = datetime.now()
                    self._update_idle_warning_label()
            elif status == "active" and self.session_manager.state == SessionState.ON_BREAK and self._break_auto:
                self._auto_break_triggered = False
                self.end_break()
            elif status == "active":
                self._auto_break_triggered = False
                if self._last_idle_warning_time:
                    self._last_idle_warning_time = None
                    self._update_idle_warning_label()
        except Exception:
            pass
    
    def _auto_start_break(self):
        """Auto start break"""
        self._break_auto = True
        if self.session_manager.start_break():
            self._update_status("Status: ON BREAK — Auto (5 min idle)")
            self.force_break_popup()
            self._refresh_attendance_info()
    
    def force_break_popup(self):
        """Force break popup"""
        self._is_force_break_mode = True
        if self.window:
            try:
                self.window.show()
                self.window.focus()
            except Exception:
                pass
        if self.session_manager.state == SessionState.ON_BREAK:
            if self._break_auto:
                self._update_status("Status: ON BREAK — Auto (5 min idle) — Click Clock In to Resume")
            else:
                self._update_status("Status: ON BREAK — Click Clock In to Resume")
    
    def _sync_state_from_server(self):
        """Sync session state from server attendance data"""
        try:
            from datetime import date, datetime, timezone
            from dateutil import parser
            
            attendance = self.session_manager.attendance_api
            client_id = attendance.client_id
            today = date.today()
            logger.info(f"_sync_state_from_server: Fetching attendance for client_id={client_id}, month={today.month}, year={today.year}")
            ok, payload, _ = attendance.fetch_attendance(client_id=client_id, month=today.month, year=today.year)
            
            if not ok or not isinstance(payload, (dict, list)):
                # If fetch fails, ensure controls are still updated based on current state
                logger.warning(f"Failed to fetch attendance data (ok={ok}, payload type={type(payload)}), using current state")
                self._update_controls()
                return
            
            logger.info(f"_sync_state_from_server: Received payload, type={type(payload)}")
            
            # Parse attendance records
            records = []
            base = payload
            if isinstance(base, dict):
                for key in ("attendances", "data", "items", "rows", "attendance", "result", "records"):
                    val = base.get(key)
                    if isinstance(val, list):
                        records = val
                        logger.info(f"_sync_state_from_server: Found records in key '{key}', count={len(records)}")
                        break
                if not records and isinstance(base.get("data"), dict):
                    records = [base.get("data")]
                    logger.info("_sync_state_from_server: Using single record from 'data' key")
                if not records:
                    records = [base]
                    logger.info("_sync_state_from_server: Using base dict as single record")
            else:
                records = base
                logger.info(f"_sync_state_from_server: Using base as records, count={len(records) if isinstance(records, list) else 'N/A'}")
            
            today_str = today.isoformat()
            logger.info(f"_sync_state_from_server: Looking for record with date containing '{today_str}'")
            rec = None
            for r in records:
                if not isinstance(r, dict):
                    continue
                d = r.get("date") or r.get("date_in_iso_format") or r.get("day") or r.get("attendance_date")
                logger.debug(f"_sync_state_from_server: Checking record with date field: {d}")
                if isinstance(d, str) and today_str in d:
                    rec = r
                    logger.info(f"_sync_state_from_server: Found today's record!")
                    break
            
            if rec is None:
                # No record found - ensure controls are updated based on current state
                logger.warning(f"_sync_state_from_server: No record found for today ({today_str})")
                self._update_controls()
                return
            
            # Check punch in/out status
            def pick(*keys):
                for k in keys:
                    v = rec.get(k)
                    if v:
                        return v
                return None
            
            in_time = pick("in_time", "punch_in", "punchInTime", "clock_in_time", "check_in_time", "start")
            out_time = pick("out_time", "punch_out", "punchOutTime", "clock_out_time", "check_out_time", "end")
            
            logger.info(f"_sync_state_from_server: in_time={in_time}, out_time={out_time}")
            
            # Check for open work period in work_periods array (for multiple punch in/out cycles)
            work_periods = rec.get("work_periods") or rec.get("workPeriods") or []
            open_work_period = None
            if isinstance(work_periods, list):
                for wp in work_periods:
                    if isinstance(wp, dict):
                        punch_in = wp.get("punch_in_time") or wp.get("punchInTime") or wp.get("in_time")
                        punch_out = wp.get("punch_out_time") or wp.get("punchOutTime") or wp.get("out_time")
                        is_open = wp.get("is_open") or wp.get("isOpen") or False
                        # Check if this is an open work period (punch_in exists, punch_out is None, and is_open is True)
                        if punch_in and not punch_out and is_open:
                            open_work_period = wp
                            logger.info(f"_sync_state_from_server: Found open work period: punch_in={punch_in}, is_open={is_open}")
                            # Use the open work period's punch_in_time as the current in_time
                            in_time = punch_in
                            out_time = None
                            break
            
            # Check for open break
            breaks = rec.get("breaks") or rec.get("break_list") or []
            on_break = False
            if isinstance(breaks, list):
                for b in breaks:
                    if isinstance(b, dict):
                        bs = b.get("break_time") or b.get("start") or b.get("break_start") or b.get("start_time")
                        be = b.get("resume_time") or b.get("end") or b.get("break_end") or b.get("end_time")
                        is_active = b.get("is_active") or b.get("isActive") or False
                        # Check if this is an active/open break
                        if bs and not be and is_active:
                            on_break = True
                            logger.info(f"_sync_state_from_server: Found open break (is_active=True)")
                            break
            
            # Sync state: if punched in and not punched out, sync to CLOCKED_IN or ON_BREAK
            if in_time and not out_time:
                logger.info(f"_sync_state_from_server: User is clocked in (in_time exists, out_time is None)")
                try:
                    in_dt = parser.parse(in_time)
                    if in_dt.tzinfo is None:
                        in_dt = in_dt.replace(tzinfo=timezone.utc)
                    
                    # Update session state
                    if on_break:
                        self.session_manager.state = SessionState.ON_BREAK
                        self._update_status("Status: ON BREAK")
                    else:
                        self.session_manager.state = SessionState.CLOCKED_IN
                        self._update_status("Status: CLOCKED IN")
                    
                    # Set session start time
                    self.session_manager.session_start = in_dt
                    self.session_manager.last_update = datetime.now(timezone.utc)
                    
                    # Apply server totals
                    self.session_manager.apply_server_totals(
                        in_time_iso=in_time,
                        break_seconds=None,
                        on_break=on_break
                    )
                    
                    # Immediately update button states after syncing state
                    self._update_controls()
                    logger.info(f"State synced from server: {self.session_manager.state}")
                    # Force a second update after a small delay to ensure UI is updated
                    import threading
                    def delayed_update():
                        import time
                        time.sleep(0.5)
                        self._update_controls()
                    threading.Thread(target=delayed_update, daemon=True).start()
                except Exception as e:
                    logger.error(f"Error syncing state from server: {e}")
                    # Even on error, try to update controls
                    self._update_controls()
            elif in_time and out_time and not open_work_period:
                # Already punched out (only if there's no open work period)
                self.session_manager.state = SessionState.LOGGED_OUT
                self._update_status("Status: LOGGED OUT")
                # Update button states - Clock In should be visible
                self._update_controls()
                logger.info("State synced: LOGGED_OUT (already punched out, no open work period)")
            else:
                # No in_time found - user is logged out
                self.session_manager.state = SessionState.LOGGED_OUT
                self._update_status("Status: LOGGED OUT")
                self._update_controls()
        except Exception as e:
            logger.error(f"Error in _sync_state_from_server: {e}")
            # On error, ensure controls are updated
            self._update_controls()
    
    def _refresh_dashboard_metrics(self):
        """Refresh dashboard metrics"""
        try:
            summary = self.session_manager.get_daily_summary()
            break_time = summary.get("total_break_time", "00:00:00")
            self._update_break_time(break_time)
            
            recent_windows = getattr(self.session_manager, "_recent_windows", [])
            if recent_windows:
                apps = ", ".join(dict.fromkeys([w.split(" - ")[-1][:20] for w in recent_windows[-3:]]))
            else:
                apps = "No recent apps"
            self._update_active_apps(apps)
            
            # Update hours spent chart
            self._update_hours_spent_chart()
            
            # Update task category chart
            self._update_task_category_chart()

            # Update app usage list
            self._update_app_usage_list()

            # Update shift/session/worked/break overview cards
            self._update_shift_overview_cards(summary)
            
            # Update activity counts (mouse clicks and keys count)
            self._update_activity_counts()
            
            # Update Teams messages
            self._update_teams_messages()
        except Exception:
            pass

    def _update_shift_overview_cards(self, summary):
        """Push compact shift / session / worked / break info to the UI cards."""
        if not self.window:
            return
        try:
            from datetime import timezone, timedelta

            attendance = self.session_manager.attendance_api

            # Shift label + timing (reuse text from dropdown if available)
            shift_label = "General Shift"
            shift_time = ""
            sid = getattr(attendance, "selected_shift_id", None)
            if sid is not None:
                key = str(sid)
                display = self._shift_display_by_id.get(key)
                if display:
                    # Display is like "General Shift (10:00 AM – 07:00 PM IST)"
                    if "(" in display and display.endswith(")"):
                        name, rng = display.split("(", 1)
                        shift_label = name.strip()
                        shift_time = rng.rstrip(")").strip()
                    else:
                        shift_label = display

            # Session start (IST friendly)
            session_start_str = "-"
            if self.session_manager.session_start:
                try:
                    dt = self.session_manager.session_start
                    if dt.tzinfo is None:
                        dt = dt.replace(tzinfo=timezone.utc)
                    ist = timezone(timedelta(hours=5, minutes=30))
                    local = dt.astimezone(ist)
                    session_start_str = local.strftime("%I:%M %p")
                except Exception:
                    session_start_str = "-"

            # Worked today & break time in "HH:MM hr" format
            def _fmt_hhmm(label):
                if not isinstance(label, str):
                    label = str(label or "")
                parts = label.split(":")
                if len(parts) >= 2:
                    hh = parts[0].zfill(2)
                    mm = parts[1].zfill(2)
                    return f"{hh}:{mm} hr"
                return f"{label} hr" if label else "00:00 hr"

            total_active = summary.get("total_active_time", "00:00:00")
            total_break = summary.get("total_break_time", "00:00:00")
            worked_today_hr = _fmt_hhmm(total_active)
            break_time_hr = _fmt_hhmm(total_break)

            def _time_to_hours(label):
                if not isinstance(label, str):
                    label = str(label or "")
                parts = label.split(":")
                if len(parts) >= 2:
                    try:
                        return int(parts[0]) + int(parts[1]) / 60.0
                    except Exception:
                        return 0.0
                try:
                    return float(label)
                except Exception:
                    return 0.0

            # Calculate simple percentage vs 9 hour reference
            STANDARD_SHIFT_HOURS = 9.0
            worked_hours = _time_to_hours(total_active)
            break_hours = _time_to_hours(total_break)
            worked_pct = ""
            break_pct = ""
            if worked_hours > 0 and STANDARD_SHIFT_HOURS > 0:
                worked_pct = f"{(worked_hours / STANDARD_SHIFT_HOURS) * 100:.1f}%"
            if break_hours > 0 and STANDARD_SHIFT_HOURS > 0:
                break_pct = f"{(break_hours / STANDARD_SHIFT_HOURS) * 100:.1f}%"

            import json
            payload = {
                "shift_label": shift_label,
                "shift_time": shift_time,
                "session_start": session_start_str,
                "worked_today": worked_today_hr,
                "break_time_hr": break_time_hr,
                "worked_pct": worked_pct,
                "break_pct": break_pct,
            }
            self.window.evaluate_js(f'updateShiftOverview({json.dumps(payload)})')
        except Exception as e:
            logger.error(f"Error updating shift overview cards: {e}")
    
    def _get_weekly_hours_data(self):
        """Get weekly hours data for stacked bar chart from ERP attendance API (actual data)"""
        try:
            from datetime import datetime, timedelta, date
            from dateutil import parser
            
            attendance = self.session_manager.attendance_api
            client_id = attendance.client_id
            if not client_id:
                logger.warning("No client_id available for fetching attendance data")
                return {"work_hours": [0]*7, "break_time": [0]*7, "overtime": [0]*7,
                        "days": ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"], 
                        "total_work_hours": 0, "total_break_time": 0, "total_overtime": 0}
            
            # Fetch attendance data for current month
            today = date.today()
            ok, payload, _ = attendance.fetch_attendance(client_id=client_id, month=today.month, year=today.year)
            
            if not ok or not isinstance(payload, (dict, list)):
                logger.warning("Failed to fetch attendance data from ERP, using session data")
                return self._get_weekly_hours_from_session()
            
            # Parse attendance records
            records = []
            base = payload
            if isinstance(base, dict):
                for key in ("attendances", "data", "items", "rows", "attendance", "result", "records"):
                    val = base.get(key)
                    if isinstance(val, list):
                        records = val
                        break
                if not records and isinstance(base.get("data"), dict):
                    records = [base.get("data")]
                if not records:
                    records = [base]
            else:
                records = base
            
            # Helper function to parse time duration to hours
            def parse_duration_to_hours(value):
                """Parse duration from ERP (could be seconds, HH:MM:SS string, or hours)"""
                if value is None:
                    return 0.0
                if isinstance(value, (int, float)):
                    # If it's seconds, convert to hours
                    if value > 1000:  # Likely seconds
                        return value / 3600
                    else:  # Likely already in hours
                        return float(value)
                if isinstance(value, str):
                    # Try to parse HH:MM:SS format
                    try:
                        parts = value.split(":")
                        if len(parts) == 3:
                            return int(parts[0]) + int(parts[1])/60 + int(parts[2])/3600
                        elif len(parts) == 2:
                            return int(parts[0]) + int(parts[1])/60
                        else:
                            return float(value)
                    except:
                        return 0.0
                return 0.0
            
            # Helper function to calculate work hours from in_time and out_time
            def calculate_work_hours_from_times(in_time, out_time):
                """Calculate work hours from punch in/out times"""
                if not in_time or not out_time:
                    return 0.0
                try:
                    in_dt = parser.parse(in_time)
                    out_dt = parser.parse(out_time)
                    if in_dt.tzinfo is None:
                        in_dt = in_dt.replace(tzinfo=parser.parse("UTC").tzinfo or None)
                    if out_dt.tzinfo is None:
                        out_dt = out_dt.replace(tzinfo=parser.parse("UTC").tzinfo or None)
                    duration = (out_dt - in_dt).total_seconds()
                    return max(0, duration / 3600)  # Convert to hours
                except:
                    return 0.0
            
            # Create a map of date -> record for easy lookup
            date_to_record = {}
            for r in records:
                if not isinstance(r, dict):
                    continue
                d = r.get("date") or r.get("date_in_iso_format") or r.get("day") or r.get("attendance_date")
                if d:
                    # Extract date part
                    try:
                        if isinstance(d, str):
                            date_part = d.split("T")[0] if "T" in d else d.split(" ")[0]
                            date_to_record[date_part] = r
                    except:
                        pass
            
            # Calculate data for last 7 days (with overtime calculation)
            # Standard shift hours: 10am to 7pm = 9 hours
            STANDARD_SHIFT_HOURS = 9.0
            
            days_data = []
            total_work_hours = 0
            total_break_time = 0
            total_overtime = 0
            
            for i in range(6, -1, -1):  # Last 7 days
                day_date = today - timedelta(days=i)
                day_str = day_date.isoformat()
                day_name = day_date.strftime("%a")
                
                work_hours = 0.0
                break_time = 0.0
                overtime = 0.0
                
                # For TODAY, always use real-time session manager data (actual live tracking)
                if i == 0:  # Today - use real-time data from session manager
                    summary = self.session_manager.get_daily_summary()
                    total_active = summary.get("total_active_time", "00:00:00")
                    parts = total_active.split(":")
                    if len(parts) == 3:
                        work_hours = int(parts[0]) + int(parts[1])/60 + int(parts[2])/3600
                    total_break = summary.get("total_break_time", "00:00:00")
                    parts = total_break.split(":")
                    if len(parts) == 3:
                        break_time = int(parts[0]) + int(parts[1])/60 + int(parts[2])/3600
                    
                    # Calculate overtime: work_hours > 9 means overtime (10am to 7pm = 9 hours)
                    if work_hours > STANDARD_SHIFT_HOURS:
                        overtime = work_hours - STANDARD_SHIFT_HOURS
                        # Work hours should only show standard hours, overtime is separate
                        work_hours = STANDARD_SHIFT_HOURS
                else:
                    # For previous days, use actual ERP data only (no fake/estimated data)
                    rec = date_to_record.get(day_str, {})
                    
                    # Try multiple keys for work hours from ERP
                    for key in ("total_work_hours", "work_hours", "total_work_time", "work_time", 
                               "hours_worked", "duration", "total_duration", "active_time", "total_active_time"):
                        val = rec.get(key)
                        if val is not None:
                            work_hours = parse_duration_to_hours(val)
                            break
                    
                    # If work hours not found in direct fields, calculate from in_time/out_time
                    if work_hours == 0:
                        in_time = rec.get("in_time") or rec.get("punch_in") or rec.get("punchInTime") or rec.get("clock_in_time")
                        out_time = rec.get("out_time") or rec.get("punch_out") or rec.get("punchOutTime") or rec.get("clock_out_time")
                        if in_time and out_time:
                            work_hours = calculate_work_hours_from_times(in_time, out_time)
                    
                    # Extract break time from ERP
                    for key in ("total_break_seconds", "break_seconds", "break_secs", "total_break",
                               "total_break_time", "break_time_total", "break_duration", "break_time"):
                        val = rec.get(key)
                        if val is not None:
                            break_time = parse_duration_to_hours(val)
                            break
                    
                    # Calculate overtime from ERP data
                    for key in ("overtime", "overtime_hours", "total_overtime", "overtime_seconds", 
                               "overtime_secs", "extra_hours"):
                        val = rec.get(key)
                        if val is not None:
                            overtime = parse_duration_to_hours(val)
                            break
                    
                    # If overtime not found in ERP but work_hours > 9, calculate it
                    if overtime == 0 and work_hours > STANDARD_SHIFT_HOURS:
                        overtime = work_hours - STANDARD_SHIFT_HOURS
                        work_hours = STANDARD_SHIFT_HOURS
                
                days_data.append({
                    "day": day_name,
                    "work_hours": work_hours,
                    "break_time": break_time,
                    "overtime": overtime
                })
                
                total_work_hours += work_hours
                total_break_time += break_time
                total_overtime += overtime
            
            return {
                "work_hours": [d["work_hours"] for d in days_data],
                "break_time": [d["break_time"] for d in days_data],
                "overtime": [d["overtime"] for d in days_data],
                "days": [d["day"] for d in days_data],
                "total_work_hours": total_work_hours,
                "total_break_time": total_break_time,
                "total_overtime": total_overtime
            }
        except Exception as e:
            logger.error(f"Error getting weekly hours data from ERP: {e}")
            # Fallback to session data
            return self._get_weekly_hours_from_session()
    
    def _get_weekly_hours_from_session(self):
        """Fallback: Get weekly hours data from session manager (local tracking)"""
        try:
            from datetime import datetime, timedelta
            
            summary = self.session_manager.get_daily_summary()
            total_active = summary.get("total_active_time", "00:00:00")
            total_break = summary.get("total_break_time", "00:00:00")
            
            def parse_time_to_hours(time_str):
                parts = time_str.split(":")
                if len(parts) == 3:
                    return int(parts[0]) + int(parts[1])/60 + int(parts[2])/3600
                elif len(parts) == 2:
                    return int(parts[0]) + int(parts[1])/60
                return 0
            
            # Standard shift hours: 10am to 7pm = 9 hours
            STANDARD_SHIFT_HOURS = 9.0
            
            work_hours_today = parse_time_to_hours(total_active)
            break_time_today = parse_time_to_hours(total_break)
            
            # Calculate overtime for today
            overtime_today = 0.0
            if work_hours_today > STANDARD_SHIFT_HOURS:
                overtime_today = work_hours_today - STANDARD_SHIFT_HOURS
                work_hours_today = STANDARD_SHIFT_HOURS
            
            today = datetime.now()
            days_data = []
            
            for i in range(6, -1, -1):
                day = today - timedelta(days=i)
                if i == 0:  # Today - use actual session data
                    days_data.append({
                        "day": day.strftime("%a"),
                        "work_hours": work_hours_today,
                        "break_time": break_time_today,
                        "overtime": overtime_today
                    })
                else:
                    # Previous days - zero (only real data from ERP, no fake estimates)
                    days_data.append({
                        "day": day.strftime("%a"),
                        "work_hours": 0.0,
                        "break_time": 0.0,
                        "overtime": 0.0
                    })
            
            return {
                "work_hours": [d["work_hours"] for d in days_data],
                "break_time": [d["break_time"] for d in days_data],
                "overtime": [d["overtime"] for d in days_data],
                "days": [d["day"] for d in days_data],
                "total_work_hours": work_hours_today,
                "total_break_time": break_time_today,
                "total_overtime": overtime_today
            }
        except Exception as e:
            logger.error(f"Error getting weekly hours data from session: {e}")
            return {"work_hours": [0]*7, "break_time": [0]*7, "overtime": [0]*7,
                    "days": ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"], 
                    "total_work_hours": 0, "total_break_time": 0, "total_overtime": 0}
    
    def _get_task_category_data(self):
        """Get task category data from app usage"""
        try:
            app_stats = self.session_manager.get_app_usage_stats()
            
            # Check if app_stats is None or empty
            if not app_stats:
                return {"categories": ["Mobile App", "Website", "Dashboard"], "counts": [0, 0, 0]}
            
            categories = {
                "Mobile App": 0,
                "Website": 0,
                "Dashboard": 0
            }
            
            # Categorize apps - handle both list and dict formats
            if isinstance(app_stats, list):
                # New format: list of dicts with 'name' and 'seconds'
                for app_data in app_stats:
                    if isinstance(app_data, dict):
                        app_name = app_data.get('name', '')
                        seconds = app_data.get('seconds', 0)
                        
                        app_lower = app_name.lower()
                        if "mobile" in app_lower or "app" in app_lower or "android" in app_lower or "ios" in app_lower:
                            categories["Mobile App"] += seconds / 3600
                        elif "browser" in app_lower or "chrome" in app_lower or "firefox" in app_lower or "edge" in app_lower or "safari" in app_lower or "website" in app_lower:
                            categories["Website"] += seconds / 3600
                        elif "dashboard" in app_lower or "admin" in app_lower or "panel" in app_lower:
                            categories["Dashboard"] += seconds / 3600
            elif isinstance(app_stats, dict):
                # Old format: dict with app_name as key and seconds as value
                for app_name, seconds in app_stats.items():
                    app_lower = app_name.lower()
                    if "mobile" in app_lower or "app" in app_lower or "android" in app_lower or "ios" in app_lower:
                        categories["Mobile App"] += seconds / 3600
                    elif "browser" in app_lower or "chrome" in app_lower or "firefox" in app_lower or "edge" in app_lower or "safari" in app_lower or "website" in app_lower:
                        categories["Website"] += seconds / 3600
                    elif "dashboard" in app_lower or "admin" in app_lower or "panel" in app_lower:
                        categories["Dashboard"] += seconds / 3600
            
            # Convert to counts (approximate - 1 hour = 10 tasks)
            result = {
                "categories": list(categories.keys()),
                "counts": [int(categories[cat] * 10) for cat in categories.keys()]
            }
            
            return result
        except Exception as e:
            logger.error(f"Error getting task category data: {e}", exc_info=True)
            return {"categories": ["Mobile App", "Website", "Dashboard"], "counts": [0, 0, 0]}
    
    def _update_hours_spent_chart(self):
        """Update hours spent chart with real weekly data (without overtime)"""
        try:
            data = self._get_weekly_hours_data()
            if self.window:
                import json
                self.window.evaluate_js(f'updateHoursSpentChart({json.dumps(data)})')
        except Exception as e:
            logger.error(f"Error updating hours spent chart: {e}")
    
    def _update_task_category_chart(self):
        """Update task category chart with real data"""
        try:
            data = self._get_task_category_data()
            if self.window:
                import json
                self.window.evaluate_js(f'updateTaskCategoryChart({json.dumps(data)})')
        except Exception as e:
            logger.error(f"Error updating task category chart: {e}")
    
    def _update_app_usage_list(self):
        """Update application usage list with actual app icons"""
        try:
            app_stats = self.session_manager.get_app_usage_stats()
            if not app_stats:
                app_stats = []
            
            # Format app data for frontend
            app_list = []
            for app_data in app_stats:
                if isinstance(app_data, dict):
                    app_list.append({
                        "name": app_data.get("name", "Unknown"),
                        "duration": app_data.get("duration") or app_data.get("time", "00:00:00"),
                        "category": "Application",
                        "process_name": app_data.get("process_name")
                    })
            
            if self.window:
                import json
                self.window.evaluate_js(f'window.updateAppUsage({json.dumps(app_list)})')
        except Exception as e:
            logger.error(f"Error updating app usage list: {e}", exc_info=True)
    
    def _load_user_profile(self):
        """Load and update user profile picture"""
        try:
            if not self.window:
                return
            
            attendance = self.session_manager.attendance_api
            if not attendance:
                return
            
            # Fetch user profile in background
            def fetch_profile():
                try:
                    success, data, status_code = attendance.fetch_user_profile()
                    if success and data and isinstance(data, dict):
                        # Extract profile picture URL - check all possible field names
                        profile_picture = (data.get('profile_image') or 
                                         data.get('profile_picture') or 
                                         data.get('avatar') or 
                                         data.get('image_url') or 
                                         data.get('photo'))
                        
                        if profile_picture:
                            # Update UI with profile picture
                            import json
                            profile_data = {
                                "profile_image": profile_picture,  # Primary field name from API
                                "profile_picture": profile_picture,
                                "avatar": profile_picture,
                                "image_url": profile_picture,
                                "is_online": data.get('status') == 'active',
                                **data
                            }
                            self.window.evaluate_js(f'window.updateUserProfile({json.dumps(profile_data)})')
                            logger.info(f"User profile picture updated: {profile_picture}")
                        else:
                            logger.warning(f"No profile image found in user data. Available keys: {list(data.keys())}")
                except Exception as e:
                    logger.error(f"Error loading user profile: {e}", exc_info=True)
            
            # Run in background thread
            import threading
            threading.Thread(target=fetch_profile, daemon=True).start()
            
        except Exception as e:
            logger.error(f"Error in _load_user_profile: {e}", exc_info=True)
    
    def _refresh_attendance_info(self):
        """Refresh attendance info - preserves all original logic"""
        try:
            from datetime import date
            attendance = self.session_manager.attendance_api
            client_id = attendance.client_id
            today = date.today()
            ok, payload, _ = attendance.fetch_attendance(client_id=client_id, month=today.month, year=today.year)
            if not ok or not isinstance(payload, (dict, list)):
                return
            
            records = []
            events = []
            base = payload
            if isinstance(base, dict):
                for key in ("attendances", "data", "items", "rows", "attendance", "result", "records"):
                    val = base.get(key)
                    if isinstance(val, list):
                        records = val
                        break
                if not records and isinstance(base.get("data"), dict):
                    records = [base.get("data")]
                for evk in ("events", "logs", "history"):
                    ev = base.get(evk)
                    if isinstance(ev, list):
                        events = ev
                        break
                if not records:
                    records = [base]
            else:
                records = base
            
            today_str = today.isoformat()
            rec = None
            for r in records:
                if not isinstance(r, dict):
                    continue
                d = r.get("date") or r.get("date_in_iso_format") or r.get("day") or r.get("attendance_date")
                if isinstance(d, str) and today_str in d:
                    rec = r
                    break
            if rec is None and records and isinstance(records[0], dict):
                rec = records[0]
            rec = rec or {}
            
            def pick(*keys):
                for k in keys:
                    v = rec.get(k)
                    if v:
                        return v
                return None
            
            in_time = pick("in_time", "punch_in", "punchInTime", "clock_in_time", "check_in_time", "start")
            out_time = pick("out_time", "punch_out", "punchOutTime", "clock_out_time", "check_out_time", "end")
            
            work_periods = []
            if isinstance(rec.get("work_periods"), list):
                for wp in rec["work_periods"]:
                    if isinstance(wp, dict):
                        pi = wp.get("punch_in_time") or wp.get("start")
                        po = wp.get("punch_out_time") or wp.get("end")
                        if pi or po:
                            work_periods.append({"in": self._fmt_ist(pi) if pi else "—", "out": self._fmt_ist(po) if po else "—"})
            if in_time or out_time:
                work_periods.insert(0, {"in": self._fmt_ist(in_time) if in_time else "—", "out": self._fmt_ist(out_time) if out_time else "—"})
            
            br_list = rec.get("breaks") or rec.get("break_list") or rec.get("break")
            br_intervals = []
            break_tuples = []  # Store ISO timestamps for calculation
            
            if isinstance(br_list, list):
                for b in br_list:
                    if not isinstance(b, dict):
                        continue
                    bs = b.get("break_time") or b.get("start") or b.get("break_start") or b.get("start_time") or b.get("pause_time")
                    be = b.get("resume_time") or b.get("end") or b.get("break_end") or b.get("end_time")
                    if bs or be:
                        br_intervals.append({"start": self._fmt_ist(bs) if bs else "—", "end": self._fmt_ist(be) if be else "—"})
                        break_tuples.append((bs, be))  # Store original ISO timestamps
            
            # Extract break time from server (mobile app data) - PRIORITY
            server_break_val = None
            for k in (
                "total_break_seconds",
                "totalBreakSeconds",
                "break_seconds",
                "break_secs",
                "breakDurationInSeconds",
                "totalBreakDuration",
                "breakDuration",
                "total_break",
                "total_break_time",
                "break_time_total",
                "break_duration",
                "break",
            ):
                v = rec.get(k)
                if v is not None:
                    server_break_val = v
                    break
            
            # Parse break duration to seconds
            def _parse_duration_to_seconds(value):
                if value is None:
                    return None
                if isinstance(value, (int, float)):
                    try:
                        return int(value)
                    except Exception:
                        return None
                text = str(value).strip()
                if not text:
                    return None
                try:
                    # Try HH:MM:SS format
                    parts = [int(p) for p in text.split(":", 2)]
                    while len(parts) < 3:
                        parts.insert(0, 0)
                    return parts[0] * 3600 + parts[1] * 60 + parts[2]
                except Exception:
                    try:
                        return int(float(text))
                    except Exception:
                        return None
            
            break_seconds = _parse_duration_to_seconds(server_break_val)
            
            # If no server break time, calculate from break intervals
            if break_seconds is None and break_tuples:
                break_seconds = self._compute_break_seconds_from_iso(break_tuples)
            elif break_seconds is None:
                break_seconds = 0
            
            # Extract work/active time from server if provided
            def _parse_duration_to_seconds(value):
                if value is None:
                    return None
                if isinstance(value, (int, float)):
                    try:
                        return int(value)
                    except Exception:
                        return None
                text = str(value).strip()
                if not text:
                    return None
                try:
                    parts = [int(p) for p in text.split(":", 2)]
                    while len(parts) < 3:
                        parts.insert(0, 0)
                    return parts[0] * 3600 + parts[1] * 60 + parts[2]
                except Exception:
                    try:
                        return int(float(text))
                    except Exception:
                        return None

            work_seconds = None
            for k in (
                "total_work_seconds", "work_seconds", "work_secs",
                "total_work_time", "work_time", "total_work_hours",
                "hours_worked", "duration", "total_duration", "active_time",
            ):
                val = rec.get(k)
                work_seconds = _parse_duration_to_seconds(val)
                if work_seconds is not None:
                    break
            
            # Update session manager with server break time (mobile app data)
            if break_seconds is not None:
                try:
                    # Check if currently on break from server
                    erp_on_break = False
                    if break_tuples:
                        # Check if last break has start but no end (open break)
                        for br_tuple in break_tuples:
                            bs, be = br_tuple
                            if bs and not be and not out_time:
                                erp_on_break = True
                                break
                    
                    # If punched in and on break, sync state to ON_BREAK
                    if in_time and not out_time and erp_on_break:
                        if self.session_manager.state != SessionState.ON_BREAK:
                            self.session_manager.state = SessionState.ON_BREAK
                            self._update_status("Status: ON BREAK")
                    # If punched in but not on break, sync state to CLOCKED_IN
                    elif in_time and not out_time:
                        if self.session_manager.state == SessionState.LOGGED_OUT:
                            self.session_manager.state = SessionState.CLOCKED_IN
                            self._update_status("Status: CLOCKED IN")
                    
                    # Apply server totals to sync break time from mobile app
                    self.session_manager.apply_server_totals(
                        in_time_iso=in_time,
                        out_time_iso=out_time,
                        break_seconds=break_seconds,
                        active_seconds=work_seconds,
                        on_break=erp_on_break if in_time and not out_time else None
                    )
                    
                    # Immediately update break time display with server data (mobile app)
                    summary = self.session_manager.get_daily_summary()
                    server_break_time = summary.get("total_break_time", "00:00:00")
                    self._update_break_time(server_break_time)
                    
                    # Update button states after applying server totals (in case state changed)
                    self._update_controls()
                except Exception as e:
                    logger.error(f"Error applying server totals: {e}")
            
            # Always update controls after checking state, even if no break data
            # This ensures buttons are correct if user is already clocked in
            if in_time and not out_time:
                # User is clocked in - ensure Clock Out button is enabled
                if self.session_manager.state == SessionState.CLOCKED_IN or self.session_manager.state == SessionState.ON_BREAK:
                    self._update_controls()
            elif in_time and out_time:
                # User is clocked out - ensure Clock In button is enabled
                if self.session_manager.state == SessionState.LOGGED_OUT:
                    self._update_controls()

            # Push server break time to UI immediately after fetch
            try:
                summary_after = self.session_manager.get_daily_summary()
                server_break = summary_after.get("total_break_time", "00:00:00")
                self._update_break_time(server_break)
                # Update the compact cards to reflect server totals
                self._update_shift_overview_cards(summary_after)
            except Exception as e:
                logger.error(f"Error updating break time from server data: {e}")
            
            if self.window:
                import json
                self.window.evaluate_js(f'updateWorkPeriods({json.dumps(work_periods)})')
                self.window.evaluate_js(f'updateBreaks({json.dumps(br_intervals)})')
        except Exception as e:
            logger.error(f"Refresh attendance error: {e}")
    
    @staticmethod
    def _fmt_ist(iso_ts):
        """Format ISO timestamp to IST"""
        try:
            from datetime import datetime, timezone, timedelta
            if not iso_ts:
                return "—"
            ts = iso_ts.replace("Z", "+00:00")
            dt = datetime.fromisoformat(ts)
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=timezone.utc)
            ist = timezone(timedelta(hours=5, minutes=30))
            local = dt.astimezone(ist)
            return local.strftime("%Y-%m-%d %H:%M:%S IST")
        except Exception:
            return iso_ts or "—"
    
    def _compute_break_seconds_from_iso(self, break_tuples):
        """Compute total break seconds from ISO timestamp tuples"""
        if not break_tuples:
            return 0
        try:
            from datetime import datetime, timezone
            total = 0
            now_utc = datetime.now(timezone.utc)
            for start_iso, end_iso in break_tuples:
                start_dt = None
                end_dt = None
                
                if start_iso:
                    try:
                        ts = str(start_iso).replace("Z", "+00:00")
                        start_dt = datetime.fromisoformat(ts)
                        if start_dt.tzinfo is None:
                            start_dt = start_dt.replace(tzinfo=timezone.utc)
                    except Exception:
                        pass
                
                if end_iso:
                    try:
                        ts = str(end_iso).replace("Z", "+00:00")
                        end_dt = datetime.fromisoformat(ts)
                        if end_dt.tzinfo is None:
                            end_dt = end_dt.replace(tzinfo=timezone.utc)
                    except Exception:
                        pass
                
                if start_dt:
                    end_dt = end_dt or now_utc
                    if end_dt >= start_dt:
                        total += int((end_dt - start_dt).total_seconds())
        except Exception as e:
            logger.error(f"Error computing break seconds: {e}")
        return total
    
    def _update_controls(self):
        """Update control states"""
        try:
            if getattr(self.session_manager.attendance_api, "machine_punch_required", False):
                self._update_button_states({
                    "clockIn": False,
                    "clockOut": False,
                    "startBreak": False,
                    "endBreak": False
                })
                return
        except Exception:
            pass
        
        state = self.session_manager.state
        logger.info(f"_update_controls called with state: {state}")
        
        if state == SessionState.LOGGED_OUT:
            logger.info("Setting button states for LOGGED_OUT: Clock In enabled, Clock Out disabled")
            self._update_button_states({
                "clockIn": True,
                "clockOut": False,
                "startBreak": False,
                "endBreak": False
            })
        elif state == SessionState.CLOCKED_IN:
            logger.info("Setting button states for CLOCKED_IN: Clock In disabled, Clock Out enabled, Start Break enabled")
            self._update_button_states({
                "clockIn": False,
                "clockOut": True,
                "startBreak": True,
                "endBreak": False
            })
        elif state == SessionState.ON_BREAK:
            logger.info("Setting button states for ON_BREAK: Clock In disabled, Clock Out enabled, End Break enabled")
            self._update_button_states({
                "clockIn": False,
                "clockOut": True,
                "startBreak": False,
                "endBreak": True
            })
        else:
            logger.warning(f"Unknown state: {state}, defaulting to LOGGED_OUT")
            self._update_button_states({
                "clockIn": True,
                "clockOut": False,
                "startBreak": False,
                "endBreak": False
            })
