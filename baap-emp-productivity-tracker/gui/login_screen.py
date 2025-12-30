import os
import sys
import io
import webview
from api.auth_api import AuthAPI
from config import ERP_DEVICE_ID
from utils.resource_path import get_html_path

# Safe print functions to handle closed stdout/stderr
def safe_str(text):
    """Safely convert any text to ASCII-safe string."""
    if text is None:
        return ""
    try:
        return str(text).encode("ascii", "replace").decode("ascii")
    except Exception:
        return "[UnicodeError]"

def safe_print(*args, **kwargs):
    """Safe print function that handles Unicode encoding errors and closed streams."""
    try:
        safe_args = [safe_str(arg) for arg in args]
        print(*safe_args, **kwargs)
    except (ValueError, OSError, AttributeError, TypeError):
        # stdout/stderr is closed or invalid - silent fail
        pass
    except Exception:
        # Any other error - silent fail
        pass

# Store login screen instance globally to avoid introspection issues
_login_screen_ref = None
_login_in_progress = False  # Flag to prevent multiple simultaneous login attempts
_auth_api_instance = None  # Reuse AuthAPI instance to preserve tokens

class LoginScreenAPI:
    """API class for pywebview - keep it simple to avoid introspection issues"""
    
    def request_otp(self, phone):
        """Handle OTP request"""
        if not phone:
            return {"success": False, "message": "Please enter your phone number"}
        
        # Normalize phone number (remove spaces, dashes, etc.)
        phone_normalized = phone.strip().replace(" ", "").replace("-", "").replace("(", "").replace(")", "")
        
        # Validate phone number format (basic validation)
        if len(phone_normalized) < 10:
            return {"success": False, "message": "Please enter a valid phone number"}
        
        try:
            # Create AuthAPI only when needed
            auth = AuthAPI()
            ok, msg = auth.request_otp(phone_normalized)
            if ok:
                return {"success": True, "message": msg or "OTP sent successfully to your phone"}
            else:
                return {"success": False, "message": msg or "Failed to send OTP. Please try again."}
        except Exception as e:
            return {"success": False, "message": f"An error occurred: {str(e)}"}
    
    def submit_login(self, credentials):
        """Handle login submission"""
        global _login_screen_ref, _login_in_progress, _auth_api_instance
        
        import time
        from utils.logger import logger
        call_id = int(time.time() * 1000) % 10000  # Unique call ID for debugging
        logger.info(f"[BUTTON CLICK] submit_login button clicked (call_id: {call_id})")
        logger.info(f"[BUTTON CLICK] Login method: {credentials.get('method', 'unknown')}")
        safe_print(f"[DEBUG] submit_login called (call_id: {call_id}), _login_in_progress: {_login_in_progress}")
        # Prevent multiple simultaneous login attempts
        if _login_in_progress:
            safe_print(f"[DEBUG] Call {call_id}: Login already in progress, rejecting duplicate call")
            return {"success": False, "message": "Login already in progress. Please wait..."}
        
        method = credentials.get("method")
        phone = credentials.get("phone", "").strip()
        
        if not phone:
            return {"success": False, "message": "Please enter your phone number"}
        
        # Validate phone number format (basic validation)
        if len(phone) < 10:
            return {"success": False, "message": "Please enter a valid phone number"}
        
        # Set flag to prevent concurrent calls IMMEDIATELY
        _login_in_progress = True
        safe_print(f"[DEBUG] Call {call_id}: Set _login_in_progress = True")
        
        try:
            # Reuse AuthAPI instance to preserve tokens across calls
            if _auth_api_instance is None:
                _auth_api_instance = AuthAPI()
                safe_print(f"[DEBUG] Call {call_id}: Created new AuthAPI instance")
            else:
                safe_print(f"[DEBUG] Call {call_id}: Reusing existing AuthAPI instance")
            auth = _auth_api_instance
            
            # Check if we already have access token BEFORE making any API calls
            if auth.access_token:
                safe_print(f"[DEBUG] Call {call_id}: Access token already exists (length: {len(auth.access_token)}), login already succeeded")
                _login_in_progress = False
                if _login_screen_ref:
                    _login_screen_ref.result = {
                        "method": method,
                        "phone": phone,
                        "device_id": ERP_DEVICE_ID
                    }
                try:
                    windows = webview.windows
                    if windows:
                        windows[0].destroy()
                except Exception:
                    pass
                return {"success": True, "message": "Login already successful"}
            
            # For OTP method, check token again after a brief delay (in case another call just succeeded)
            if method == "otp":
                import time
                time.sleep(0.1)  # Brief delay to allow any concurrent call to complete
                if auth.access_token:
                    safe_print(f"[DEBUG] Call {call_id}: Access token found after delay, login already succeeded")
                    _login_in_progress = False
                    if _login_screen_ref:
                        _login_screen_ref.result = {
                            "method": "otp",
                            "phone": phone,
                            "device_id": ERP_DEVICE_ID
                        }
                    try:
                        windows = webview.windows
                        if windows:
                            windows[0].destroy()
                    except Exception:
                        pass
                    return {"success": True, "message": "Login already successful"}
            
            if method == "otp":
                otp = credentials.get("otp", "").strip()
                if not otp:
                    _login_in_progress = False  # Reset flag on validation failure
                    return {"success": False, "message": "Please enter OTP"}
                
                # Validate OTP format (should be numeric, typically 4-6 digits)
                if not otp.isdigit():
                    _login_in_progress = False  # Reset flag on validation failure
                    return {"success": False, "message": "OTP should contain only numbers"}
                
                if len(otp) < 4 or len(otp) > 6:
                    _login_in_progress = False  # Reset flag on validation failure
                    return {"success": False, "message": "Please enter a valid OTP (4-6 digits)"}
                
                # Normalize phone number (remove spaces, dashes, etc.) before sending
                phone_normalized = phone.replace(" ", "").replace("-", "").replace("(", "").replace(")", "")
                
                # Check access token again right before calling (in case another call just succeeded)
                if auth.access_token:
                    safe_print(f"[DEBUG] Call {call_id}: Access token found before OTP verification, login already succeeded")
                    _login_in_progress = False
                    if _login_screen_ref:
                        _login_screen_ref.result = {
                            "method": "otp",
                            "phone": phone,
                            "device_id": ERP_DEVICE_ID
                        }
                    try:
                        windows = webview.windows
                        if windows:
                            windows[0].destroy()
                    except Exception:
                        pass
                    return {"success": True, "message": "Login already successful"}
                
                # Login via OTP (access token check already done above)
                safe_print(f"[DEBUG] Call {call_id}: Calling login_via_otp")
                success, message, login_data = auth.login_via_otp(
                    phone=phone_normalized,
                    otp=otp,
                    device_id=ERP_DEVICE_ID
                )
                safe_print(f"[DEBUG] Call {call_id}: login_via_otp returned success={success}, message={message}")
                
                # ABSOLUTE CRITICAL: Token existence is THE ONLY TRUTH - check immediately and repeatedly
                # IGNORE success/message from API - only token matters
                token_found = False
                
                # Immediate check
                if auth.access_token:
                    token_found = True
                    safe_print(f"[DEBUG] Call {call_id}: Token found IMMEDIATELY (length: {len(auth.access_token)})")
                
                # Multiple retries with increasing delays
                if not token_found:
                    for retry in range(8):  # More retries
                        time.sleep(0.15)  # Fixed delay
                        if auth.access_token:
                            token_found = True
                            safe_print(f"[DEBUG] Call {call_id}: Token found on retry {retry+1} (length: {len(auth.access_token)})")
                            break
                        # Reload from storage
                        try:
                            auth._load_tokens()
                            if auth.access_token:
                                token_found = True
                                safe_print(f"[DEBUG] Call {call_id}: Token found in storage on retry {retry+1}")
                                break
                        except:
                            pass
                
                # FINAL DECISION: If token exists, SUCCESS - period. Ignore everything else.
                if token_found or auth.access_token:
                    safe_print(f"[DEBUG] Call {call_id}: *** TOKEN EXISTS - FORCING SUCCESS ***")
                    success = True
                    message = "Login successful"
                    # Don't even look at the original success/message - token is truth
                else:
                    safe_print(f"[DEBUG] Call {call_id}: No token found - using API response")
                    # Only use API response if no token exists
                
                if success:
                    if _login_screen_ref:
                        _login_screen_ref.result = {
                            "method": "otp",
                            "phone": phone,
                            "otp": otp,
                            "device_id": ERP_DEVICE_ID
                        }
                    # Close window
                    try:
                        windows = webview.windows
                        if windows:
                            windows[0].destroy()
                    except Exception:
                        pass
                    _login_in_progress = False  # Reset flag on success
                    logger.info("[BUTTON CLICK] Login SUCCESS - OTP method")
                    return {"success": True, "message": "Login successful"}
                else:
                    # ONE MORE FINAL CHECK - maybe token was set just now
                    time.sleep(0.2)
                    auth._load_tokens()  # Reload from storage
                    if auth.access_token:
                        safe_print(f"[DEBUG] Call {call_id}: *** LAST SECOND TOKEN CHECK - TOKEN FOUND! FORCING SUCCESS ***")
                        _login_in_progress = False
                        if _login_screen_ref:
                            _login_screen_ref.result = {
                                "method": "otp",
                                "phone": phone,
                                "otp": otp,
                                "device_id": ERP_DEVICE_ID
                            }
                        try:
                            windows = webview.windows
                            if windows:
                                windows[0].destroy()
                        except Exception:
                            pass
                        return {"success": True, "message": "Login successful"}
                    
                    error_msg = message or "Login failed. Please check your OTP."
                    safe_print(f"[DEBUG] Call {call_id}: Returning error: {error_msg}")
                    logger.error(f"[BUTTON CLICK] Login FAILED - OTP method: {error_msg}")
                    _login_in_progress = False  # Reset flag on failure
                    return {"success": False, "message": error_msg}
            else:
                # Password login
                password = str(credentials.get("password", "") or "").strip()
                remember = credentials.get("remember", False)
                
                safe_print(f"[DEBUG] Password login - password type: {type(credentials.get('password'))}, length: {len(password)}")
                
                if not password:
                    _login_in_progress = False  # Reset flag on validation failure
                    return {"success": False, "message": "Please enter your password"}
                
                if len(password) < 4:
                    _login_in_progress = False  # Reset flag on validation failure
                    return {"success": False, "message": "Password should be at least 4 characters"}
                
                # Normalize phone number (remove spaces, dashes, etc.) before sending
                phone_normalized = phone.replace(" ", "").replace("-", "").replace("(", "").replace(")", "")
                
                # Login via password
                try:
                    safe_print(f"[INFO] Attempting login with phone: {phone} (normalized: {phone_normalized}), password length: {len(password)}")
                    success, message, login_data = auth.login(
                        phone=phone_normalized,
                        password=password,
                        device_id=ERP_DEVICE_ID
                    )
                    safe_print(f"[INFO] Login result: success={success}, message={message}")
                    
                    if success:
                        if _login_screen_ref:
                            _login_screen_ref.result = {
                                "method": "password",
                                "phone": phone,
                                "password": password,  # Store the trimmed password
                                "device_id": ERP_DEVICE_ID
                            }
                        safe_print(f"[INFO] Login successful, token saved, result stored with password length: {len(password)}")
                        # Close window
                        try:
                            windows = webview.windows
                            if windows:
                                windows[0].destroy()
                        except Exception:
                            pass
                        _login_in_progress = False  # Reset flag on success
                        logger.info("[BUTTON CLICK] Login SUCCESS - Password method")
                        return {"success": True, "message": "Login successful"}
                    else:
                        # Provide more detailed error message
                        error_msg = message or "Login failed. Please check your phone number and password."
                        
                        # Check for common error patterns and provide helpful messages
                        error_msg_lower = str(message).lower()
                        if "401" in str(message) or "unauthorized" in error_msg_lower or "invalid" in error_msg_lower:
                            error_msg = "Invalid phone number or password. Please check your credentials and try again."
                        elif "404" in str(message) or "not found" in error_msg_lower:
                            error_msg = "Login endpoint not found. Please contact support."
                        elif "network" in error_msg_lower or "connection" in error_msg_lower:
                            error_msg = "Network error. Please check your internet connection and try again."
                        elif "timeout" in error_msg_lower:
                            error_msg = "Request timed out. Please try again."
                        elif "email or password" in error_msg_lower:
                            error_msg = "Invalid phone number or password. Please verify your credentials."
                        
                        safe_print(f"[ERROR] Login failed: {error_msg}")
                        logger.error(f"[BUTTON CLICK] Login FAILED - Password method: {error_msg}")
                        _login_in_progress = False  # Reset flag on failure
                        return {"success": False, "message": error_msg}
                except Exception as e:
                    # Catch any unexpected errors
                    error_msg = str(e)
                    safe_print(f"[ERROR] Login exception: {error_msg}")
                    _login_in_progress = False  # Reset flag on exception
                    
                    # Provide user-friendly error messages based on error type
                    error_msg_lower = error_msg.lower()
                    if "i/o operation on closed file" in error_msg_lower:
                        # This error is now fixed, but handle gracefully if it occurs
                        return {"success": False, "message": "Login failed. Please try again."}
                    elif "network" in error_msg_lower or "connection" in error_msg_lower:
                        return {"success": False, "message": "Network error. Please check your internet connection and try again."}
                    elif "timeout" in error_msg_lower:
                        return {"success": False, "message": "Request timed out. Please try again."}
                    else:
                        return {"success": False, "message": f"Login failed: {error_msg[:200]}"}
        except Exception as e:
            error_msg = str(e)
            _login_in_progress = False  # Reset flag on exception
            safe_print(f"[ERROR] Outer exception in submit_login: {error_msg}")
            
            # Provide more user-friendly error messages
            error_msg_lower = error_msg.lower()
            if "i/o operation on closed file" in error_msg_lower:
                # This error is now fixed, but handle gracefully if it occurs
                return {"success": False, "message": "An error occurred. Please try again."}
            elif "network" in error_msg_lower or "connection" in error_msg_lower:
                return {"success": False, "message": "Network error. Please check your internet connection and try again."}
            elif "timeout" in error_msg_lower:
                return {"success": False, "message": "Request timed out. Please try again."}
            else:
                return {"success": False, "message": f"An error occurred: {error_msg[:200]}"}

class LoginScreen:
    def __init__(self):
        global _login_screen_ref
        self.result = None  # dict with method and credentials
        self.window = None
        _login_screen_ref = self

    def show(self):
        # Get the HTML file path (works in both dev and PyInstaller)
        html_path = get_html_path('login_screen.html')
        
        # Create API instance - simple class with no complex attributes
        api = LoginScreenAPI()
        
        # Create webview window
        self.window = webview.create_window(
            'Sign In - Your Central Workspace',
            html_path,
            width=1200,
            height=700,
            min_size=(900, 600),
            resizable=True,
            js_api=api
        )
        
        # Start webview (blocking)
        webview.start(debug=False)
        
        return self.result
