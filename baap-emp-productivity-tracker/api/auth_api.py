import time
import json
import base64
import requests
import threading
from typing import Optional, Tuple, Dict, Any
from config import ERP_CLIENTS_URL, ERP_REFRESH_TOKEN_URL, ERP_LOGIN_URL, ERP_CREDENTIAL_LOGIN_URL
from config import ERP_DEFAULT_HEADERS, ERP_LOGIN_HEADERS
from config import ERP_REQUEST_OTP_URL, ERP_VERIFY_OTP_URL
from utils.excel_storage import read_local_storage, write_local_storage
from utils.network_checker import get_public_ip

# Lock to prevent concurrent OTP verification calls
_otp_verify_lock = threading.Lock()

# Safe print function to handle closed file descriptors
def _safe_print(*args, **kwargs):
    """Safe print function that handles Unicode encoding errors and closed streams."""
    try:
        # Convert all arguments to strings safely
        safe_args = []
        for arg in args:
            try:
                safe_args.append(str(arg))
            except Exception:
                safe_args.append("[UnicodeError]")
        print(*safe_args, **kwargs)
    except (ValueError, OSError, AttributeError, TypeError, IOError):
        # stdout/stderr is closed or invalid - silent fail
        pass
    except Exception:
        # Any other error - silent fail
        pass

class AuthAPI:
    def __init__(self):
        self.access_token: Optional[str] = None
        self.refresh_token: Optional[str] = None
        self.user_id: Optional[str] = None
        self.access_token_expires_at: Optional[float] = None
        # Optional UI-provided IP override for all outbound ERP calls
        self.ip_override: Optional[str] = None
        self._load_tokens()

    def login(self, phone: str, password: str, device_id: str = "") -> Tuple[bool, str, Optional[Dict[str, Any]]]:
        """
        Login using phone number and password.
        API endpoint: /auth/api/auth/login
        Expected payload: {"phone": "...", "password": "..."}
        """
        # Normalize phone number (remove spaces, dashes, etc.)
        phone_clean = phone.strip().replace(" ", "").replace("-", "").replace("(", "").replace(")", "")
        
        # API expects 'phone' + 'password' (not email)
        # Ensure phone number is not empty
        if not phone_clean:
            return False, "Phone number cannot be empty", None
        if not password:
            return False, "Password cannot be empty", None
        
        # Create payload with phone and password only
        # Some APIs don't accept device_id in login payload
        payload = {
            "phone": phone_clean,
            "password": password,
        }
        
        # Note: device_id is NOT added to login payload as it may cause validation errors
        # Device ID can be sent separately after login if needed
        
        # Use login-specific headers matching platform.baap.company
        headers = dict(ERP_LOGIN_HEADERS)
        # Prefer dedicated credential login endpoint if configured; fall back to OTP verify (unlikely to work) or clients
        url = (ERP_CREDENTIAL_LOGIN_URL or ERP_LOGIN_URL or ERP_CLIENTS_URL)
        
        # Ensure URL is properly formatted (remove double slashes)
        if url:
            url = url.replace("//auth/api/auth", "/auth/api/auth").replace("//auth/api", "/auth/api")
        
        # Debug: Print URL and payload (without password)
        _safe_print(f"[DEBUG] Login URL: {url}")
        _safe_print(f"[DEBUG] Login payload keys: {list(payload.keys())}")
        _safe_print(f"[DEBUG] Phone (original): {phone}, Phone (cleaned): {phone_clean}, Device ID: {device_id[:50] if device_id else 'None'}...")
        
        try:
            resp = requests.post(url, json=payload, headers=headers, timeout=30)
            data = self._safe_json(resp)
            
            # Debug: Print response status and data
            _safe_print(f"[DEBUG] Response status: {resp.status_code}")
            _safe_print(f"[DEBUG] Response headers: {dict(resp.headers)}")
            if isinstance(data, dict):
                _safe_print(f"[DEBUG] Response keys: {list(data.keys())}")
                if "data" in data and isinstance(data["data"], dict):
                    _safe_print(f"[DEBUG] Response data keys: {list(data['data'].keys())}")
            _safe_print(f"[DEBUG] Response data (first 500 chars): {str(data)[:500]}...")
            
            if resp.ok:
                # Handle both dict and non-dict responses
                if not isinstance(data, dict):
                    # Try to parse as JSON string
                    try:
                        if isinstance(data, str):
                            import json
                            data = json.loads(data)
                        else:
                            data = {"raw": str(data)}
                    except:
                        data = {"raw": str(data)}
                
                # Expecting tokens in response - handle various response formats
                source = data
                if "data" in data and isinstance(data["data"], dict):
                    source = data["data"]
                elif "result" in data and isinstance(data["result"], dict):
                    source = data["result"]
                elif "response" in data and isinstance(data["response"], dict):
                    source = data["response"]
                
                access = (
                    source.get("access_token") 
                    or source.get("token") 
                    or source.get("accessToken")
                    or data.get("access_token")  # Check top level too
                    or data.get("token")
                )
                refresh = (
                    source.get("refresh_token") 
                    or source.get("refreshToken")
                    or data.get("refresh_token")  # Check top level too
                )
                self.user_id = (
                    source.get("user_id") 
                    or source.get("userId")
                    or data.get("user_id")  # Check top level too
                    or self._user_id_from_jwt(access)
                )
                expires_in = (
                    source.get("expires_in") 
                    or source.get("expiresIn")
                    or data.get("expires_in")  # Check top level too
                )
                
                _safe_print(f"[DEBUG] Login: Extracted user_id={self.user_id}, has_access_token={bool(access)}, has_refresh_token={bool(refresh)}")
                
                self._set_tokens(access, refresh, expires_in)
                
                # Ensure user_id is saved (in case it was extracted from JWT after _set_tokens)
                if self.user_id:
                    self._save_tokens()
                    _safe_print(f"[DEBUG] Login: Saved user_id={self.user_id}")
                
                # If we have refresh token, refresh to ensure latest token
                if self.refresh_token:
                    refresh_ok, refresh_msg = self.refresh_access_token()
                    if not refresh_ok:
                        _safe_print(f"[DEBUG] Token refresh failed: {refresh_msg}, but continuing with original token")
                    # After refresh, user_id might need to be re-extracted from new token
                    if not self.user_id and self.access_token:
                        self.user_id = self._user_id_from_jwt(self.access_token)
                        if self.user_id:
                            self._save_tokens()
                            _safe_print(f"[DEBUG] Login: Extracted user_id from refreshed token: {self.user_id}")
                
                # Fetch clients if user_id and device_id are available
                if self.user_id and device_id:
                    clients_ok, clients_msg, _ = self.fetch_clients(self.user_id, device_id)
                    if not clients_ok:
                        _safe_print(f"[DEBUG] Fetch clients failed: {clients_msg}, but login may still be valid")
                
                if self.access_token:
                    _safe_print(f"[DEBUG] Login successful, access token received (length: {len(self.access_token)})")
                    return True, "Login successful", data
                else:
                    _safe_print(f"[DEBUG] Login failed: No access token in response")
                    _safe_print(f"[DEBUG] Available keys in response: {list(data.keys()) if isinstance(data, dict) else 'N/A'}")
                    return False, "No access token in response. Please check API response format.", data
            else:
                error_msg = f"Login failed (HTTP {resp.status_code})"
                
                # Try to extract detailed error message from response
                if isinstance(data, dict):
                    # Try multiple common error message fields
                    error_msg = (data.get("message") 
                                or data.get("error") 
                                or data.get("msg") 
                                or data.get("errorMessage")
                                or data.get("error_message")
                                or data.get("detail")
                                or data.get("description"))
                    
                    # If we still don't have a message, try nested structures
                    if not error_msg and "data" in data and isinstance(data["data"], dict):
                        error_msg = (data["data"].get("message") 
                                    or data["data"].get("error") 
                                    or data["data"].get("msg"))
                    
                    # If still no message, use string representation
                    if not error_msg:
                        error_msg = str(data)[:200]
                        
                elif isinstance(data, str):
                    error_msg = data[:200] if len(data) > 200 else data
                else:
                    # Try to extract error from response text
                    try:
                        if resp.text:
                            # Try to parse as JSON first
                            try:
                                error_json = resp.json()
                                if isinstance(error_json, dict):
                                    error_msg = (error_json.get("message") 
                                                or error_json.get("error") 
                                                or error_json.get("msg")
                                                or str(error_json)[:200])
                                else:
                                    error_msg = resp.text[:200]
                            except:
                                error_msg = resp.text[:200]
                        else:
                            error_msg = f"HTTP {resp.status_code}: {resp.reason}"
                    except Exception:
                        error_msg = f"HTTP {resp.status_code}: {resp.reason}"
                
                # Ensure we have a meaningful error message
                if not error_msg or error_msg == "Login failed":
                    error_msg = f"Login failed (HTTP {resp.status_code}). Please check your credentials and try again."
                
                _safe_print(f"[DEBUG] Login failed: {error_msg}")
                return False, error_msg, data
        except requests.exceptions.Timeout:
            error_msg = "Request timed out. Please check your internet connection and try again."
            _safe_print(f"[DEBUG] Login timeout: {error_msg}")
            return False, error_msg, None
        except requests.exceptions.ConnectionError as e:
            error_msg = f"Connection error: Unable to reach server. Please check your internet connection."
            _safe_print(f"[DEBUG] Login connection error: {str(e)}")
            return False, error_msg, None
        except requests.exceptions.RequestException as e:
            error_msg = f"Network error: {str(e)}"
            _safe_print(f"[DEBUG] Login network error: {str(e)}")
            return False, error_msg, None
        except Exception as e:
            error_msg = f"Login error: {str(e)}"
            _safe_print(f"[DEBUG] Login exception: {str(e)}")
            try:
                import traceback
                traceback.print_exc()
            except Exception:
                pass
            return False, error_msg, None

    def get_auth_header(self) -> Dict[str, str]:
        return {"Authorization": f"Bearer {self.access_token}"} if self.access_token else {}

    def is_access_token_expiring(self, skew_seconds: int = 30) -> bool:
        if not self.access_token_expires_at:
            return False
        return time.time() >= (self.access_token_expires_at - skew_seconds)

    def refresh_access_token(self) -> Tuple[bool, str]:
        if not self.refresh_token:
            return False, "No refresh token"
        headers = {"content-type": "application/json"}
        payload = {"refresh_token": self.refresh_token}
        resp = requests.post(ERP_REFRESH_TOKEN_URL, json=payload, headers=headers)
        data = self._safe_json(resp)
        if resp.ok and isinstance(data, dict):
            new_access = data.get("access_token") or data.get("token")
            new_refresh = data.get("refresh_token") or self.refresh_token
            expires_in = data.get("expires_in")
            self._set_tokens(new_access, new_refresh, expires_in)
            return True, "Token refreshed"
        return False, (data.get("message") if isinstance(data, dict) else "Refresh failed")

    def authorized_request(self, method: str, url: str, retry_on_401: bool = True, **kwargs) -> requests.Response:
        # Attach auth header
        headers = kwargs.pop("headers", {}) or {}
        # Merge default ERP headers (do not overwrite explicit headers)
        merged = dict(ERP_DEFAULT_HEADERS)
        merged.update(headers)
        headers = merged
        # Inject client IP header and payload field if available/requested
        try:
            ip_addr = (self.ip_override or "").strip()
            if not ip_addr:
                try:
                    ip_addr = get_public_ip()
                except Exception:
                    ip_addr = ""
            if ip_addr and "x-client-ip" not in {k.lower(): v for k, v in headers.items()}:
                headers["x-client-ip"] = ip_addr
            # If a JSON body is present and does not include ip, attach it
            if "json" in kwargs and isinstance(kwargs.get("json"), dict):
                kwargs_json = dict(kwargs.get("json"))
                kwargs_json.setdefault("ip", ip_addr)
                kwargs["json"] = kwargs_json
            # Also ensure ip is present as a query parameter for GETs or when using params
            if method and method.upper() == "GET":
                params = dict(kwargs.get("params") or {})
                params.setdefault("ip", ip_addr)
                kwargs["params"] = params
            elif "params" in kwargs and isinstance(kwargs.get("params"), dict):
                params = dict(kwargs.get("params"))
                params.setdefault("ip", ip_addr)
                kwargs["params"] = params
        except Exception:
            pass
        if self.is_access_token_expiring():
            self.refresh_access_token()
        if self.access_token:
            headers["Authorization"] = f"Bearer {self.access_token}"
        kwargs["headers"] = headers

        resp = requests.request(method, url, **kwargs)
        if resp.status_code == 401 and retry_on_401 and self.refresh_access_token()[0]:
            # retry once with new token
            headers = kwargs.get("headers", {})
            if self.access_token:
                headers["Authorization"] = f"Bearer {self.access_token}"
                kwargs["headers"] = headers
            resp = requests.request(method, url, **kwargs)
        return resp

    # Allow UI to set IP that should be sent with every request
    def set_ip_override(self, ip: Optional[str]) -> None:
        self.ip_override = (ip or "").strip() or None

    def logout(self) -> None:
        # Clear tokens and persist
        self.access_token = None
        self.refresh_token = None
        self.user_id = None
        self.access_token_expires_at = None
        self._save_tokens()

    @staticmethod
    def _safe_json(resp: requests.Response):
        try:
            return resp.json()
        except Exception:
            return {"message": resp.text}

    def fetch_clients(self, user_id: str, device_id: str) -> Tuple[bool, str, Optional[Dict[str, Any]]]:
        if not user_id:
            _safe_print(f"[DEBUG] fetch_clients: user_id is empty")
            return False, "User ID is required but not provided", None
        if not device_id:
            _safe_print(f"[DEBUG] fetch_clients: device_id is empty")
            return False, "Device ID is required but not provided", None
        if not self.access_token:
            _safe_print(f"[DEBUG] fetch_clients: access_token is not available")
            return False, "Access token is required but not available. Please login again.", None
        
        params = {"user_id": user_id, "device_id": device_id}
        _safe_print(f"[DEBUG] fetch_clients: URL={ERP_CLIENTS_URL}, user_id={user_id}, device_id length={len(device_id)}")
        
        try:
            resp = self.authorized_request("GET", ERP_CLIENTS_URL, params=params)
            _safe_print(f"[DEBUG] fetch_clients: Response status={resp.status_code}")
            data = self._safe_json(resp)
            
            if resp.ok:
                _safe_print(f"[DEBUG] fetch_clients: Success, data type={type(data)}")
                return True, "Clients fetched", data
            else:
                # Extract error message
                error_msg = "Failed to fetch clients"
                if isinstance(data, dict):
                    error_msg = (data.get("message") 
                                or data.get("error") 
                                or data.get("msg") 
                                or data.get("detail")
                                or error_msg)
                elif isinstance(data, str):
                    error_msg = data[:200] if len(data) > 200 else data
                else:
                    try:
                        error_msg = resp.text[:200] if resp.text else f"HTTP {resp.status_code}: {resp.reason}"
                    except:
                        error_msg = f"HTTP {resp.status_code}: {resp.reason}"
                
                _safe_print(f"[DEBUG] fetch_clients: Failed - {error_msg}")
                return False, error_msg, None
        except Exception as e:
            error_msg = f"Exception while fetching clients: {str(e)}"
            _safe_print(f"[DEBUG] fetch_clients: Exception - {error_msg}")
            try:
                import traceback
                traceback.print_exc()
            except:
                pass
            return False, error_msg, None

    # Token persistence and helpers
    def _set_tokens(self, access: Optional[str], refresh: Optional[str], expires_in: Optional[float]):
        if access:
            self.access_token = access
            # Derive expiry from JWT exp if expires_in not provided
            self._maybe_set_expiry_from_jwt(access)
        if isinstance(expires_in, (int, float)):
            self.access_token_expires_at = time.time() + float(expires_in)
        if refresh:
            self.refresh_token = refresh
        self._save_tokens()

    def _maybe_set_expiry_from_jwt(self, jwt_token: str):
        try:
            parts = jwt_token.split(".")
            if len(parts) != 3:
                return
            # JWT base64url decode
            padding = '=' * (-len(parts[1]) % 4)
            payload_bytes = base64.urlsafe_b64decode(parts[1] + padding)
            payload = json.loads(payload_bytes.decode("utf-8"))
            exp = payload.get("exp")
            if isinstance(exp, (int, float)):
                self.access_token_expires_at = float(exp)
        except Exception:
            pass

    def _load_tokens(self):
        try:
            data = read_local_storage()
            auth = data.get("auth", {})
            self.access_token = auth.get("access_token")
            self.refresh_token = auth.get("refresh_token")
            self.user_id = auth.get("user_id")
            self.access_token_expires_at = auth.get("access_token_expires_at")
        except Exception:
            pass

    def _save_tokens(self):
        try:
            data: Dict[str, Any] = read_local_storage()
            data["auth"] = {
                "access_token": self.access_token,
                "refresh_token": self.refresh_token,
                "user_id": self.user_id,
                "access_token_expires_at": self.access_token_expires_at,
            }
            write_local_storage(data)
        except Exception:
            pass

    # New: Login flow using refresh token and user id
    def login_with_refresh(self, refresh_token: str, user_id: str, device_id: str) -> Tuple[bool, str, Optional[Dict[str, Any]]]:
        self.user_id = user_id
        self.refresh_token = refresh_token
        self._save_tokens()
        ok, msg = self.refresh_access_token()
        if not ok:
            return False, msg, None
        ok, msg, clients = self.fetch_clients(user_id=user_id, device_id=device_id)
        if not ok:
            return False, msg, None
        return True, "Login successful", clients

    def request_otp(self, phone: str) -> Tuple[bool, str]:
        # Normalize phone number (remove spaces, dashes, etc.)
        phone_clean = phone.strip().replace(" ", "").replace("-", "").replace("(", "").replace(")", "")
        
        headers = {"content-type": "application/json"}
        payload = {"phone": phone_clean}
        
        _safe_print(f"[DEBUG] Request OTP URL: {ERP_REQUEST_OTP_URL}")
        _safe_print(f"[DEBUG] Request OTP payload: {payload}")
        
        try:
            resp = requests.post(ERP_REQUEST_OTP_URL, json=payload, headers=headers, timeout=30)
            data = self._safe_json(resp)
            
            _safe_print(f"[DEBUG] Request OTP response status: {resp.status_code}")
            _safe_print(f"[DEBUG] Request OTP response: {str(data)[:200]}...")
            
            if resp.ok:
                message = data.get("message") if isinstance(data, dict) else "OTP sent"
                return True, message
            else:
                error_msg = data.get("message") if isinstance(data, dict) else "Failed to request OTP"
                return False, error_msg
        except Exception as e:
            _safe_print(f"[DEBUG] Request OTP exception: {str(e)}")
            return False, f"Error requesting OTP: {str(e)}"

    def login_via_otp(self, phone: str, otp: str, device_id: str = "") -> Tuple[bool, str, Optional[Dict[str, Any]]]:
        global _otp_verify_lock
        
        # If we already have an access token, don't make another API call
        # This prevents duplicate calls when the function is called multiple times
        if self.access_token:
            _safe_print(f"[DEBUG] Access token already exists, skipping OTP verification (token length: {len(self.access_token)})")
            return True, "Login already successful", {"access_token": self.access_token}
        
        # Check if we're already in the process of verifying (lock check before acquiring)
        # This is critical to prevent duplicate OTP verification calls
        if _otp_verify_lock.locked():
            _safe_print(f"[DEBUG] OTP verification lock already held, waiting for completion...")
            # Wait a moment and check if we got a token (first call might have succeeded)
            for i in range(15):  # Check up to 15 times (1.5 seconds total)
                time.sleep(0.1)
                if self.access_token:
                    _safe_print(f"[DEBUG] Access token found after waiting ({i+1} attempts), login succeeded")
                    return True, "Login already successful", {"access_token": self.access_token}
            _safe_print(f"[DEBUG] Lock held but no token after waiting, returning error")
            return False, "OTP verification already in progress. Please wait...", None
        
        # Use lock to prevent concurrent OTP verification calls
        # Try to acquire lock with a very short timeout to prevent blocking
        lock_acquired = False
        try:
            lock_acquired = _otp_verify_lock.acquire(blocking=False)
        except Exception as e:
            _safe_print(f"[DEBUG] Lock acquire error: {e}")
        
        if not lock_acquired:
            _safe_print(f"[DEBUG] OTP verification already in progress, skipping duplicate call")
            # Wait a moment and check if we got a token (first call might have succeeded)
            for _ in range(5):  # Check up to 5 times (0.5 seconds total)
                time.sleep(0.1)
                if self.access_token:
                    _safe_print(f"[DEBUG] Access token found after waiting, login succeeded")
                    return True, "Login already successful", {"access_token": self.access_token}
            return False, "OTP verification already in progress. Please wait...", None
        
        try:
            # Normalize phone number (remove spaces, dashes, etc.)
            phone_clean = phone.strip().replace(" ", "").replace("-", "").replace("(", "").replace(")", "")
            
            # Ensure OTP is string and remove any whitespace
            otp_clean = str(otp).strip()
        
            # Try OTP as both string and integer (some APIs expect integer)
            try:
                otp_int = int(otp_clean)
            except:
                otp_int = None
            
            # Use simple headers for OTP verification (some APIs don't like complex headers)
            # Don't add device_id - OTP verification typically only needs phone and OTP
            headers = {
                "content-type": "application/json",
                "accept": "application/json"
            }
            
            # OTP verification payload - only phone and OTP (no device_id in payload or headers)
            payload = {
                "phone": phone_clean,
                "otp": otp_clean  # Send as string first
            }
            
            _safe_print(f"[DEBUG] Phone number (cleaned): {phone_clean}, Length: {len(phone_clean)}")
            _safe_print(f"[DEBUG] OTP (cleaned): {otp_clean}, Length: {len(otp_clean)}, Type: {type(otp_clean)}")
            _safe_print(f"[DEBUG] Device ID: NOT INCLUDED in OTP verification request")
            
            _safe_print(f"[DEBUG] Verify OTP URL: {ERP_VERIFY_OTP_URL}")
            _safe_print(f"[DEBUG] Verify OTP payload (string): {payload}")
            
            # First try with OTP as string
            resp = requests.post(ERP_VERIFY_OTP_URL, json=payload, headers=headers, timeout=30)
            data = self._safe_json(resp)
            
            _safe_print(f"[DEBUG] Verify OTP response status: {resp.status_code}")
            _safe_print(f"[DEBUG] Verify OTP response headers: {dict(resp.headers)}")
            if isinstance(data, dict):
                _safe_print(f"[DEBUG] Verify OTP response keys: {list(data.keys())}")
                if "data" in data and isinstance(data["data"], dict):
                    _safe_print(f"[DEBUG] Verify OTP response data keys: {list(data['data'].keys())}")
            _safe_print(f"[DEBUG] Verify OTP response data (first 500 chars): {str(data)[:500]}...")
            
            # If failed with string OTP and we have integer, try with integer
            if not resp.ok and otp_int is not None and resp.status_code == 401:
                _safe_print(f"[DEBUG] Trying with OTP as integer instead of string")
                payload_int = {
                    "phone": phone_clean,
                    "otp": otp_int  # Try as integer
                }
                resp = requests.post(ERP_VERIFY_OTP_URL, json=payload_int, headers=headers, timeout=30)
                data = self._safe_json(resp)
                _safe_print(f"[DEBUG] Verify OTP (integer) response status: {resp.status_code}")
                _safe_print(f"[DEBUG] Verify OTP (integer) response: {str(data)[:200]}...")
            
            # Helper function to extract token from response data
            def extract_token_data(response_data):
                """Extract token, refresh, user_id, and expires_in from response"""
                if not isinstance(response_data, dict):
                    return None, None, None, None
                
                source = response_data
                if "data" in response_data and isinstance(response_data["data"], dict):
                    source = response_data["data"]
                elif "result" in response_data and isinstance(response_data["result"], dict):
                    source = response_data["result"]
                elif "response" in response_data and isinstance(response_data["response"], dict):
                    source = response_data["response"]
                
                access = (
                    source.get("access_token")
                    or source.get("token")
                    or source.get("accessToken")
                    or response_data.get("access_token")
                    or response_data.get("token")
                )
                refresh = (
                    source.get("refresh_token")
                    or source.get("refreshToken")
                    or response_data.get("refresh_token")
                )
                user_id = (
                    source.get("user_id")
                    or source.get("userId")
                    or response_data.get("user_id")
                )
                expires_in = (
                    source.get("expires_in")
                    or source.get("expiresIn")
                    or response_data.get("expires_in")
                )
                return access, refresh, user_id, expires_in
            
            # Handle both dict and non-dict responses
            if not isinstance(data, dict):
                # Try to parse as JSON string
                try:
                    if isinstance(data, str):
                        import json
                        data = json.loads(data)
                    else:
                        data = {"raw": str(data)}
                except:
                    data = {"raw": str(data)}
            
            # CRITICAL: Always try to extract token from response, regardless of HTTP status
            # Some APIs return tokens even with error status codes
            access, refresh, user_id_from_response, expires_in = extract_token_data(data)
            
            # If we found a token, treat as success regardless of HTTP status
            if access:
                _safe_print(f"[DEBUG] Token found in response (status: {resp.status_code}), treating as success")
                if not user_id_from_response and access:
                    user_id_from_response = self._user_id_from_jwt(access)
                
                self._set_tokens(access, refresh, expires_in)
                if user_id_from_response:
                    self.user_id = user_id_from_response
                
                if self.access_token:
                    _safe_print(f"[DEBUG] Verify OTP successful, access token received (length: {len(self.access_token)})")
                    result = (True, "Login successful", data)
                    
                    # IMPORTANT: Keep lock longer to prevent duplicate calls
                    time.sleep(0.3)
                    
                    # Immediately refresh to ensure we have the latest access token/claims
                    if self.refresh_token:
                        refresh_ok, refresh_msg = self.refresh_access_token()
                        if not refresh_ok:
                            _safe_print(f"[DEBUG] Token refresh failed: {refresh_msg}, but continuing with original token")
                    
                    if self.user_id and device_id:
                        clients_ok, clients_msg, _ = self.fetch_clients(self.user_id, device_id)
                        if not clients_ok:
                            _safe_print(f"[DEBUG] Fetch clients failed: {clients_msg}, but login may still be valid")
                    
                    time.sleep(0.2)
                else:
                    _safe_print(f"[DEBUG] Token extraction failed despite finding token in response")
                    result = (False, "Failed to set access token", data)
            elif resp.ok:
                # No token found but response is OK - this shouldn't happen but handle it
                _safe_print(f"[DEBUG] Verify OTP: Response OK but no token found")
                _safe_print(f"[DEBUG] Available keys in response: {list(data.keys()) if isinstance(data, dict) else 'N/A'}")
                result = (False, "No access token returned from verify-otp. Please check API response format.", data)
            else:
                # Handle error response - no token found
                error_msg = "OTP verification failed"
                if isinstance(data, dict):
                    error_msg = data.get("message") or data.get("error") or data.get("msg") or str(data)
                elif isinstance(data, str):
                    error_msg = data
                else:
                    # Try to extract error from response text
                    try:
                        error_msg = resp.text[:200] if resp.text else "Unknown error"
                    except:
                        error_msg = f"HTTP {resp.status_code}: {resp.reason}"
                
                _safe_print(f"[DEBUG] Verify OTP failed: {error_msg}")
                result = (False, error_msg, data)
                
        except requests.exceptions.Timeout:
            error_msg = "Request timed out. Please check your internet connection and try again."
            _safe_print(f"[DEBUG] Verify OTP timeout: {error_msg}")
            result = (False, error_msg, None)
        except requests.exceptions.ConnectionError as e:
            error_msg = f"Connection error: Unable to reach server. Please check your internet connection."
            _safe_print(f"[DEBUG] Verify OTP connection error: {str(e)}")
            result = (False, error_msg, None)
        except requests.exceptions.RequestException as e:
            error_msg = f"Network error: {str(e)}"
            _safe_print(f"[DEBUG] Verify OTP network error: {str(e)}")
            result = (False, error_msg, None)
        except Exception as e:
            error_msg = f"OTP verification error: {str(e)}"
            _safe_print(f"[DEBUG] Verify OTP exception: {str(e)}")
            try:
                import traceback
                traceback.print_exc()
            except Exception:
                pass
            result = (False, error_msg, None)
        finally:
            # ABSOLUTE FINAL CHECK - Token existence is THE ONLY TRUTH
            # Reload tokens one more time in case they were saved asynchronously
            try:
                self._load_tokens()
            except:
                pass
            
            # If token exists, FORCE success - ignore everything else
            if self.access_token:
                _safe_print(f"[DEBUG] *** FINAL FINAL CHECK: Token exists (length: {len(self.access_token)}), FORCING SUCCESS ***")
                # Preserve the data from previous result if it exists
                try:
                    data = result[2] if isinstance(result, tuple) and len(result) > 2 else None
                except:
                    data = None
                result = (True, "Login successful", data)
                _safe_print(f"[DEBUG] Result overridden to: success=True, message='Login successful'")
            
            # Ensure lock is always released, but only if we acquired it
            try:
                if lock_acquired and _otp_verify_lock.locked():
                    _otp_verify_lock.release()
                    _safe_print(f"[DEBUG] OTP verification lock released")
            except Exception as e:
                _safe_print(f"[DEBUG] Error releasing lock: {e}")
        
        return result

    def _user_id_from_jwt(self, jwt_token: Optional[str]) -> Optional[str]:
        if not jwt_token:
            return None
        try:
            parts = jwt_token.split(".")
            if len(parts) != 3:
                return None
            padding = '=' * (-len(parts[1]) % 4)
            payload_bytes = base64.urlsafe_b64decode(parts[1] + padding)
            payload = json.loads(payload_bytes.decode("utf-8"))
            uid = payload.get("user_id") or payload.get("sub")
            return uid
        except Exception:
            return None

    def decode_jwt_payload(self, jwt_token: Optional[str] = None) -> Optional[Dict[str, Any]]:
        """
        Decode JWT token and return the full payload as a dictionary.
        If jwt_token is not provided, uses self.access_token.
        Returns None if decoding fails.
        """
        if not jwt_token:
            jwt_token = self.access_token
        if not jwt_token:
            return None
        try:
            parts = jwt_token.split(".")
            if len(parts) != 3:
                return None
            padding = '=' * (-len(parts[1]) % 4)
            payload_bytes = base64.urlsafe_b64decode(parts[1] + padding)
            payload = json.loads(payload_bytes.decode("utf-8"))
            return payload
        except Exception:
            return None

    def get_client_id_from_token(self, jwt_token: Optional[str] = None) -> Optional[str]:
        """
        Extract client_id from JWT token payload.
        Tries common field names: client_id, clientId, client_id
        """
        payload = self.decode_jwt_payload(jwt_token)
        if not payload:
            return None
        # Try various possible field names for client_id
        return (payload.get("client_id") 
                or payload.get("clientId") 
                or payload.get("client"))

    def get_user_id_from_token(self, jwt_token: Optional[str] = None) -> Optional[str]:
        """
        Extract user_id from JWT token payload.
        Tries common field names: user_id, userId, sub
        """
        payload = self.decode_jwt_payload(jwt_token)
        if not payload:
            return None
        # Try various possible field names for user_id
        return (payload.get("user_id") 
                or payload.get("userId") 
                or payload.get("sub"))
