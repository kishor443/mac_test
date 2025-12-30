import requests
from typing import Optional, Tuple, List, Dict, Any
from config import BASE_URL
from utils.logger import logger

# Get appointment base URL if configured, otherwise use BASE_URL
try:
    from config import ERP_APPOINTMENT_BASE_URL
except ImportError:
    ERP_APPOINTMENT_BASE_URL = None


class AppointmentAPI:
    def __init__(self, auth_api):
        self.auth_api = auth_api
        self.client_id = None

    def set_client(self, client_id: str):
        """Set the client ID for appointment requests"""
        self.client_id = client_id

    def fetch_appointments(self, client_id: str = None, user_id: str = None, page: int = 1, limit: int = 10) -> Tuple[bool, List[Dict[str, Any]], str]:
        """
        Fetch appointments for a user.
        
        Uses: GET /appointment/client/{client_id}/user?user_id={user_id}&page={page}&limit={limit}
        
        Args:
            client_id: Client UUID (defaults to self.client_id)
            user_id: User UUID (defaults to auth_api.user_id)
            page: Page number for pagination (default: 1)
            limit: Number of items per page (default: 10)
        
        Returns:
            Tuple of (success, appointments_list, message)
        """
        # Use provided values or fall back to instance values
        client_id = client_id or self.client_id
        user_id = user_id or getattr(self.auth_api, "user_id", None)
        
        if not client_id:
            logger.error(f"[APPOINTMENT API] ERROR: Client ID is required")
            return False, [], "Client ID is required"
        if not user_id:
            logger.error(f"[APPOINTMENT API] ERROR: User ID is required")
            return False, [], "User ID is required"
        
        # Use ERP_APPOINTMENT_BASE_URL if configured, otherwise use BASE_URL
        base_url = ERP_APPOINTMENT_BASE_URL if ERP_APPOINTMENT_BASE_URL else BASE_URL
        url = f"{base_url}/appointment/client/{client_id}/user"
        
        params = {
            "user_id": user_id,
            "page": page,
            "limit": limit
        }
        
        try:
            logger.info(f"[AppointmentAPI] Fetching appointments: client_id={client_id}, user_id={user_id}, page={page}, limit={limit}")
            resp = self.auth_api.authorized_request("GET", url, params=params)
            
            try:
                data = resp.json()
            except Exception:
                data = {"message": resp.text}
            
            if resp.ok:
                # Extract appointments from response
                # Handle various response formats
                appointments = []
                if isinstance(data, dict):
                    # Try common response keys
                    for key in ("data", "appointments", "items", "results", "appointment"):
                        if key in data:
                            value = data[key]
                            if isinstance(value, list):
                                appointments = value
                                break
                            elif isinstance(value, dict) and "items" in value:
                                appointments = value["items"]
                                break
                            elif isinstance(value, dict) and "data" in value:
                                if isinstance(value["data"], list):
                                    appointments = value["data"]
                                    break
                    # If still no appointments, check if data itself contains appointment fields
                    if not appointments:
                        # Maybe the response is a single appointment object
                        if any(key in data for key in ("appointment_title", "title", "meeting_title", "appointment_date", "date", "start_time")):
                            appointments = [data]
                    # If data is a list directly in a dict wrapper
                    if not appointments and isinstance(data.get("data"), list):
                        appointments = data["data"]
                elif isinstance(data, list):
                    appointments = data
                
                logger.info(f"[AppointmentAPI] Successfully fetched {len(appointments)} appointments")
                if appointments:
                    logger.info(f"[AppointmentAPI] First appointment sample: {appointments[0]}")
                return True, appointments, "Success"
            else:
                error_msg = data.get("message") if isinstance(data, dict) else resp.text
                logger.error(f"[AppointmentAPI] Failed to fetch appointments: HTTP {resp.status_code} - {error_msg}")
                return False, [], error_msg or f"HTTP {resp.status_code}"
                
        except Exception as exc:
            logger.error(f"[AppointmentAPI] Exception while fetching appointments: {exc}", exc_info=True)
            return False, [], str(exc)

