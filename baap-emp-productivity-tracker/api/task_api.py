"""
Task API for fetching and managing tasks
"""
import requests
import json
from typing import Dict, Tuple, Any, List, Optional
from config import BASE_URL, ERP_DEFAULT_HEADERS
from utils.logger import logger


class TaskAPI:
    """
    API handler for task-related operations
    """
    
    def __init__(self, auth_api):
        """
        Initialize TaskAPI with auth_api reference
        
        Args:
            auth_api: Reference to AuthAPI instance for accessing tokens
        """
        self.auth_api = auth_api
        self.client_id = None
    
    def set_client(self, client_id: str):
        """
        Set the current client context
        
        Args:
            client_id: The client UUID to set as active
        """
        self.client_id = client_id
        logger.info(f"TaskAPI client_id set to: {client_id}")
    
    def _get_headers(self) -> Dict[str, str]:
        """
        Build request headers with authorization
        Based on curl command: uses qa.d3kq8oy4csoq2n.amplifyapp.com origin/referer
        
        Returns:
            Dict containing headers with Bearer token
        """
        headers = {
            "accept": "application/json, text/plain, */*",
            "accept-language": "en-US,en;q=0.9",
            # Align headers with the latest curl the user provided (QA environment)
            "origin": "https://qa.d3kq8oy4csoq2n.amplifyapp.com",
            "referer": "https://qa.d3kq8oy4csoq2n.amplifyapp.com/",
            "sec-ch-ua": '"Chromium";v="142", "Google Chrome";v="142", "Not_A Brand";v="99"',
            "sec-ch-ua-mobile": "?0",
            "sec-ch-ua-platform": '"Windows"',
            "sec-fetch-dest": "empty",
            "sec-fetch-mode": "cors",
            "sec-fetch-site": "cross-site",
            "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/142.0.0.0 Safari/537.36"
        }
        if self.auth_api and self.auth_api.access_token:
            headers["authorization"] = f"Bearer {self.auth_api.access_token}"
        return headers
    
    def fetch_tasks(
        self, 
        client_id: Optional[str] = None,
        user_id: Optional[str] = None,
        project_id: Optional[str] = None
    ) -> Tuple[bool, Any, int]:
        """
        Fetch tasks for a client/user/project
        
        Args:
            client_id: Client UUID (uses self.client_id if not provided)
            user_id: User UUID (optional, if provided, fetches tasks for that user)
            project_id: Project UUID (optional, if provided, filters tasks by project)
        
        Returns:
            Tuple of (success: bool, data: dict/str, status_code: int)
        """
        cid = client_id or self.client_id
        if not cid:
            logger.error("No client_id set for TaskAPI")
            return False, "Client ID not set", 0
        
        # Get user_id from auth_api if not provided
        if not user_id and self.auth_api:
            # Try multiple methods to get user_id
            user_id = getattr(self.auth_api, 'user_id', None)
            if not user_id:
                # Try to get from token using auth_api method
                try:
                    user_id = self.auth_api.get_user_id_from_token()
                except Exception:
                    pass
            if not user_id:
                # Try to extract from token directly
                try:
                    import jwt
                    if self.auth_api.access_token:
                        decoded = jwt.decode(self.auth_api.access_token, options={"verify_signature": False})
                        user_id = decoded.get('user_id')
                except Exception:
                    pass
        
        headers = self._get_headers()
        
        # Based on curl: /task/client/{client_id}/task/{task_id} (for specific task)
        # For fetching all tasks, try various patterns
        
        # Try endpoints in order of likelihood
        # Based on curl pattern: /task/client/{client_id}/task/{task_id}
        # For all tasks, try various patterns
        endpoints_to_try = [
            # Pattern 1: List endpoint (common REST pattern)
            f"{BASE_URL}/task/client/{cid}/list",
            f"{BASE_URL}/tasks/client/{cid}/list",
            
            # Pattern 2: Direct client endpoint
            f"{BASE_URL}/task/client/{cid}",
            f"{BASE_URL}/tasks/client/{cid}",
            
            # Pattern 3: With /tasks suffix
            f"{BASE_URL}/task/client/{cid}/tasks",
            f"{BASE_URL}/tasks/client/{cid}/tasks",
            
            # Pattern 4: With /all suffix
            f"{BASE_URL}/task/client/{cid}/all",
            f"{BASE_URL}/tasks/client/{cid}/all",
            
            # Pattern 5: With /all-tasks
            f"{BASE_URL}/task/client/{cid}/all-tasks",
            f"{BASE_URL}/tasks/client/{cid}/all-tasks",
        ]
        
        # Add user-specific patterns if user_id available
        if user_id:
            endpoints_to_try.extend([
                f"{BASE_URL}/task/client/{cid}/user/{user_id}",
                f"{BASE_URL}/tasks/client/{cid}/user/{user_id}",
                f"{BASE_URL}/task/client/{cid}/user/{user_id}/tasks",
                f"{BASE_URL}/tasks/client/{cid}/user/{user_id}/tasks",
                f"{BASE_URL}/task/client/{cid}/user/{user_id}/list",
                f"{BASE_URL}/tasks/client/{cid}/user/{user_id}/list",
            ])
        
        # Build query parameters
        query_params = {}
        if user_id:
            query_params['user_id'] = user_id
        if project_id:
            query_params['project_id'] = project_id
        
        # Try with query parameters - Priority endpoints based on user's curl
        query_endpoints = []
        
        # Pattern from user's curl: /task/client/{client_id}/user?project_id={project_id}&user_id={user_id}
        if user_id or project_id:
            query_endpoints.extend([
                (f"{BASE_URL}/task/client/{cid}/user", query_params.copy()),
                (f"{BASE_URL}/tasks/client/{cid}/user", query_params.copy()),
            ])
        
        # Add other query parameter patterns
        if query_params:
            query_endpoints.extend([
                (f"{BASE_URL}/task/client/{cid}/list", query_params.copy()),
                (f"{BASE_URL}/tasks/client/{cid}/list", query_params.copy()),
                (f"{BASE_URL}/task/client/{cid}", query_params.copy()),
                (f"{BASE_URL}/tasks/client/{cid}", query_params.copy()),
            ])
        
        # Detailed logging
        logger.info("="*80)
        logger.info("[TASK API] ========== FETCH TASKS START ==========")
        logger.info("="*80)
        logger.info(f"[TASK API] Client ID: {cid}")
        logger.info(f"[TASK API] User ID: {user_id if user_id else 'None'}")
        logger.info(f"[TASK API] Project ID: {project_id if project_id else 'None'}")
        logger.debug(f"[TASK API] Query Params: {json.dumps(query_params, indent=2)}")
        logger.info(f"Fetching tasks - Client: {cid}, User: {user_id}, Project: {project_id}")
        
        last_error = None
        tried_count = 0
        max_tries_before_summary = 3  # Only show detailed logs for first few attempts
        
        # Try endpoints WITH query parameters FIRST (matches user's curl pattern)
        for url, params in query_endpoints:
            if not params:  # Skip if no params
                continue
            tried_count += 1
            try:
                logger.info(f"[TASK API] Attempt {tried_count}: GET {url}")
                logger.debug(f"[TASK API] Params: {json.dumps(params, indent=2)}")
                logger.debug(f"[TASK API] Headers: {json.dumps({k: v if k != 'authorization' else f'{v[:20]}...' for k, v in headers.items()}, indent=2)}")
                
                response = requests.get(url, headers=headers, params=params, timeout=10)
                
                status_code = response.status_code
                logger.info(f"[TASK API] Response Status: {status_code}")
                logger.debug(f"[TASK API] Response Headers: {dict(response.headers)}")
                
                if status_code == 200:
                    data = response.json()
                    logger.info(f"[TASK API] Response Data Type: {type(data)}")
                    logger.debug(f"[TASK API] Response Data Keys: {list(data.keys()) if isinstance(data, dict) else 'N/A (not a dict)'}")
                    
                    # Parse task count
                    task_count = 0
                    if isinstance(data, dict):
                        if 'data' in data:
                            if isinstance(data['data'], list):
                                task_count = len(data['data'])
                            elif isinstance(data['data'], dict):
                                task_count = len(data['data'].get('tasks', data['data'].get('items', [])))
                        elif 'tasks' in data:
                            task_count = len(data['tasks']) if isinstance(data['tasks'], list) else 0
                        elif 'items' in data:
                            task_count = len(data['items']) if isinstance(data['items'], list) else 0
                    elif isinstance(data, list):
                        task_count = len(data)
                    
                    logger.info(f"[TASK API] SUCCESS: Fetched {task_count} tasks from {url}")
                    logger.debug(f"[TASK API] Response Sample (first 500 chars): {json.dumps(data, indent=2, ensure_ascii=False)[:500]}...")
                    logger.info("="*80)
                    return True, data, status_code
                elif status_code == 404:
                    last_error = "404 Not Found"
                    continue
                else:
                    try:
                        error_data = response.json()
                        last_error = error_data.get("message", f"HTTP {status_code}")
                    except Exception:
                        last_error = f"HTTP {status_code}"
                    continue
                    
            except (requests.exceptions.Timeout, requests.exceptions.ConnectionError, Exception) as e:
                last_error = str(e) if not isinstance(e, (requests.exceptions.Timeout, requests.exceptions.ConnectionError)) else type(e).__name__
                continue
        
        # Then try endpoints without query parameters
        for url in endpoints_to_try:
            tried_count += 1
            try:
                logger.info(f"[TASK API] Attempt {tried_count}: GET {url} (no params)")
                response = requests.get(url, headers=headers, timeout=10)
                status_code = response.status_code
                logger.info(f"[TASK API] Response Status: {status_code}")
                
                if status_code == 200:
                    data = response.json()
                    logger.info(f"[TASK API] Response Data Type: {type(data)}")
                    
                    # Parse task count
                    task_count = 0
                    if isinstance(data, dict):
                        if 'data' in data:
                            if isinstance(data['data'], list):
                                task_count = len(data['data'])
                            elif isinstance(data['data'], dict):
                                task_count = len(data['data'].get('tasks', data['data'].get('items', [])))
                        elif 'tasks' in data:
                            task_count = len(data['tasks']) if isinstance(data['tasks'], list) else 0
                        elif 'items' in data:
                            task_count = len(data['items']) if isinstance(data['items'], list) else 0
                    elif isinstance(data, list):
                        task_count = len(data)
                    
                    logger.info(f"[TASK API] SUCCESS: Fetched {task_count} tasks from {url}")
                    logger.debug(f"[TASK API] Response Sample (first 500 chars): {json.dumps(data, indent=2, ensure_ascii=False)[:500]}...")
                    logger.info("="*80)
                    return True, data, status_code
                elif status_code == 404:
                    last_error = "404 Not Found"
                    continue
                else:
                    try:
                        error_data = response.json()
                        last_error = error_data.get("message", f"HTTP {status_code}")
                    except Exception:
                        last_error = f"HTTP {status_code}"
                    continue
                    
            except (requests.exceptions.Timeout, requests.exceptions.ConnectionError, Exception) as e:
                last_error = str(e) if not isinstance(e, (requests.exceptions.Timeout, requests.exceptions.ConnectionError)) else type(e).__name__
                continue
        
        # If all endpoints failed, return empty result gracefully
        logger.warning(f"[TASK API] WARNING: All {tried_count} endpoints failed")
        logger.warning(f"[TASK API] Last Error: {last_error}")
        logger.warning(f"[TASK API] Returning empty task list")
        logger.info("="*80)
        return True, {"data": [], "tasks": [], "items": []}, 200
    
    def get_assigned_tasks(
        self, 
        client_id: Optional[str] = None
    ) -> Tuple[bool, List[Dict], str]:
        """
        Get tasks assigned to the current user
        
        Args:
            client_id: Client UUID (uses self.client_id if not provided)
        
        Returns:
            Tuple of (success: bool, tasks: list, message: str)
        """
        # Fetch tasks - API will filter by authenticated user from token
        # Don't pass user_id - let API use the Bearer token to determine user
        success, data, status_code = self.fetch_tasks(
            client_id=client_id,
            user_id=None  # API uses token to identify user
        )
        
        if not success:
            return False, [], f"Failed to fetch tasks: {data}"
        
        # Parse the response to extract tasks
        tasks = []
        try:
            if isinstance(data, dict):
                # Common response formats
                if "data" in data:
                    if isinstance(data["data"], list):
                        tasks = data["data"]
                    elif isinstance(data["data"], dict) and "tasks" in data["data"]:
                        tasks = data["data"]["tasks"]
                    elif isinstance(data["data"], dict) and "items" in data["data"]:
                        tasks = data["data"]["items"]
                elif "tasks" in data:
                    tasks = data["tasks"]
                elif "items" in data:
                    tasks = data["items"]
                elif "results" in data:
                    tasks = data["results"]
            elif isinstance(data, list):
                tasks = data
            
            return True, tasks, "Tasks fetched successfully"
            
        except Exception as e:
            logger.error(f"Error parsing task data: {str(e)}")
            return False, [], f"Error parsing tasks: {str(e)}"
    
    def update_task(
        self,
        task_id: str,
        task_data: Dict[str, Any],
        client_id: Optional[str] = None
    ) -> Tuple[bool, Any, str]:
        """
        Update a task
        PUT /task/client/{client_id}/task/{task_id}
        
        Args:
            task_id: Task UUID to update
            task_data: Dictionary containing task fields to update
            client_id: Client UUID (uses self.client_id if not provided)
        
        Returns:
            Tuple of (success: bool, data: dict, message: str)
        """
        cid = client_id or self.client_id
        if not cid:
            logger.error("No client_id set for TaskAPI")
            return False, None, "Client ID not set"
        
        if not task_id:
            return False, None, "Task ID is required"
        
        try:
            url = f"{BASE_URL}/task/client/{cid}/task/{task_id}"
            headers = self._get_headers()
            
            logger.info(f"[TASK API] Updating task {task_id[:8]}... -> status_id: {task_data.get('taskstatus_id', 'N/A')[:8]}...")
            response = requests.put(url, headers=headers, json=task_data, timeout=10)
            status_code = response.status_code
            
            if status_code in [200, 201]:
                data = response.json()
                logger.info(f"[TASK API] SUCCESS: Task updated successfully - Status: {status_code}")
                return True, data, "Task updated successfully"
            else:
                error_text = response.text[:200] if response.text else "No error message"
                logger.error(f"[TASK API] ERROR: Task update failed - Status: {status_code}, Error: {error_text}")
                logger.debug(f"[TASK API] Payload sent: {json.dumps(task_data, indent=2, ensure_ascii=False)}")
                return False, None, f"HTTP {status_code}: {error_text}"
                
        except Exception as e:
            logger.error(f"[TASK API] ERROR: Exception: {str(e)}", exc_info=True)
            return False, None, str(e)
    
    def fetch_task(
        self,
        task_id: str,
        client_id: Optional[str] = None
    ) -> Tuple[bool, Any, str]:
        """
        Fetch a single task by ID
        GET /task/client/{client_id}/task/{task_id}
        
        Args:
            task_id: Task UUID to fetch
            client_id: Client UUID (uses self.client_id if not provided)
        
        Returns:
            Tuple of (success: bool, data: dict, message: str)
        """
        cid = client_id or self.client_id
        if not cid:
            logger.error("No client_id set for TaskAPI")
            return False, None, "Client ID not set"
        
        if not task_id:
            return False, None, "Task ID is required"
        
        try:
            url = f"{BASE_URL}/task/client/{cid}/task/{task_id}"
            headers = self._get_headers()
            
            logger.info(f"[TASK API] Fetching task {task_id[:8]}... for client {cid[:8]}...")
            
            response = requests.get(url, headers=headers, timeout=10)
            status_code = response.status_code
            
            if status_code == 200:
                data = response.json()
                logger.info(f"[TASK API] SUCCESS: Task fetched successfully - Status: {status_code}")
                if isinstance(data, dict):
                    logger.info(f"[TASK API] Response keys: {list(data.keys())}")
                return True, data, "Task fetched successfully"
            else:
                try:
                    error_data = response.json()
                    error_message = error_data.get("message", f"HTTP {status_code}")
                except Exception:
                    error_message = f"HTTP {status_code}: {response.text[:200]}"
                
                logger.error(f"[TASK API] ERROR: Failed to fetch task ({status_code}) - {error_message}")
                return False, None, error_message
                
        except requests.exceptions.Timeout:
            error_msg = "Request timeout"
            logger.error(f"[TASK API] ERROR: {error_msg}")
            return False, None, error_msg
        except requests.exceptions.ConnectionError:
            error_msg = "Connection error"
            logger.error(f"[TASK API] ERROR: {error_msg}")
            return False, None, error_msg
        except Exception as e:
            error_msg = str(e)
            logger.error(f"[TASK API] ERROR: Exception - {error_msg}", exc_info=True)
            return False, None, error_msg
    
    def fetch_task_statuses(self, client_id: Optional[str] = None) -> Tuple[bool, List[Dict], str]:
        """
        Fetch all available task statuses for a client
        
        Args:
            client_id: Client UUID (uses self.client_id if not provided)
        
        Returns:
            Tuple of (success: bool, statuses: list, message: str)
        """
        cid = client_id or self.client_id
        if not cid:
            logger.error("No client_id set for TaskAPI")
            return False, [], "Client ID not set"
        
        headers = self._get_headers()
        
        # Try common endpoints for fetching task statuses
        endpoints_to_try = [
            f"{BASE_URL}/taskstatus/client/{cid}",
            f"{BASE_URL}/task-status/client/{cid}",
            f"{BASE_URL}/task/status/client/{cid}",
            f"{BASE_URL}/task/statuses/client/{cid}",
            f"{BASE_URL}/taskstatus",
            f"{BASE_URL}/task-status",
        ]
        
        print(f"[TASK API] Fetching task statuses for client {cid[:8]}...")
        
        for url in endpoints_to_try:
            try:
                response = requests.get(url, headers=headers, timeout=10)
                status_code = response.status_code
                
                if status_code == 200:
                    data = response.json()
                    
                    # Parse response
                    statuses = []
                    if isinstance(data, dict):
                        statuses = data.get('data', data.get('statuses', data.get('items', [])))
                    elif isinstance(data, list):
                        statuses = data
                    
                    if statuses:
                        logger.info(f"[TASK API] SUCCESS: Fetched {len(statuses)} task statuses")
                        return True, statuses, "Task statuses fetched successfully"
                    
            except Exception as e:
                continue
        
        logger.warning(f"[TASK API] WARNING: Could not fetch task statuses (API endpoint not available)")
        return False, [], "Task status endpoint not available"

