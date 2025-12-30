import json
import requests
from config import BASE_URL
from utils.logger import logger

class ProjectAPI:
    def __init__(self, auth_api):
        self.auth_api = auth_api
        self.client_id = None
        
    def set_client(self, client_id: str):
        self.client_id = client_id
    
    def _get_headers(self):
        """Get headers with authorization token"""
        token = self.auth_api.access_token
        if not token:
            logger.error("[PROJECT API] ERROR: No auth token available")
            raise Exception("No auth token available")
        
        headers = {
            'accept': 'application/json, text/plain, */*',
            'accept-language': 'en-US,en;q=0.9',
            'authorization': f'Bearer {token}',
            'content-type': 'application/json',
            'origin': 'https://development.d3kq8oy4csoq2n.amplifyapp.com',
            'referer': 'https://development.d3kq8oy4csoq2n.amplifyapp.com/',
            'priority': 'u=1, i',
            'sec-ch-ua': '"Google Chrome";v="143", "Chromium";v="143", "Not A(Brand";v="24"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
            'sec-fetch-dest': 'empty',
            'sec-fetch-mode': 'cors',
            'sec-fetch-site': 'cross-site',
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/143.0.0.0 Safari/537.36'
        }
        
        logger.debug(f"[PROJECT API] Headers prepared with token: {token[:20]}...")
        return headers
    
    def fetch_projects(self, client_id: str = None):
        """
        Fetch all projects for a client
        Try multiple endpoint patterns
        """
        client_id = client_id or self.client_id
        if not client_id:
            logger.error("No client_id available for fetching projects")
            return {"success": False, "error": "No client_id"}
        
        try:
            # Get user_id from auth_api
            user_id = getattr(self.auth_api, 'user_id', None)
            if not user_id:
                try:
                    user_id = self.auth_api.get_user_id_from_token()
                except Exception as e:
                    logger.warning(f"Could not get user_id: {e}")
                    user_id = None
            
            headers = self._get_headers()
            
            logger.info("="*80)
            logger.info("[PROJECT API] Fetching projects")
            logger.info("="*80)
            logger.info(f"[PROJECT API] Client ID: {client_id}")
            logger.info(f"[PROJECT API] User ID: {user_id}")
            
            # Log headers (mask token for security)
            headers_print = headers.copy()
            if "authorization" in headers_print:
                token = headers_print["authorization"]
                if len(token) > 20:
                    headers_print["authorization"] = token[:20] + "..." + token[-10:]
            logger.debug(f"[PROJECT API] Request Headers: {headers_print}")
            logger.info("="*80)
            
            # Try different endpoint patterns
            endpoints_to_try = [
                # Pattern 1: Just client_id
                f"{BASE_URL}/project/client/{client_id}",
                f"{BASE_URL}/project/client/{client_id}/",
                
                # Pattern 2: With /list suffix
                f"{BASE_URL}/project/client/{client_id}/list",
                f"{BASE_URL}/project/client/{client_id}/projects",
                
                # Pattern 3: With user_id if available
            ]
            
            if user_id:
                endpoints_to_try.extend([
                    f"{BASE_URL}/project/client/{client_id}/{user_id}",
                    f"{BASE_URL}/project/client/{client_id}/user/{user_id}",
                    f"{BASE_URL}/project/client/{client_id}/user/{user_id}/list",
                ])
            
            last_error = None
            last_response_data = None
            
            for url in endpoints_to_try:
                try:
                    logger.info(f"[PROJECT API] Trying URL: {url}")
                    response = requests.get(url, headers=headers, timeout=10)
                    
                    logger.info(f"[PROJECT API] Response Status: {response.status_code}")
                    
                    if response.status_code == 200:
                        data = response.json()
                        last_response_data = data
                        
                        logger.debug(f"[PROJECT API] Response Payload: {json.dumps(data, indent=2, ensure_ascii=False)}")
                        
                        # Check if response contains actual projects (not nested error)
                        has_projects = False
                        project_count = 0
                        
                        if isinstance(data, dict):
                            # Check for nested error response
                            if 'project' in data and isinstance(data['project'], dict):
                                if data['project'].get('success') == False:
                                    logger.warning(f"[PROJECT API] WARNING: Nested error: {data['project'].get('message')}")
                                    continue  # Try next endpoint
                            
                            # Try to find projects in response
                            if 'data' in data and isinstance(data['data'], dict) and 'projects' in data['data']:
                                has_projects = True
                                project_count = len(data['data']['projects'])
                            elif 'data' in data and isinstance(data['data'], list):
                                has_projects = True
                                project_count = len(data['data'])
                            elif 'projects' in data and isinstance(data['projects'], list):
                                has_projects = True
                                project_count = len(data['projects'])
                            elif 'project' in data and isinstance(data['project'], list):
                                has_projects = True
                                project_count = len(data['project'])
                        elif isinstance(data, list):
                            has_projects = True
                            project_count = len(data)
                        
                        if has_projects or project_count > 0:
                            logger.info(f"[PROJECT API] SUCCESS: Successfully fetched {project_count} projects from {url}")
                            return {"success": True, "data": data}
                        else:
                            logger.warning(f"[PROJECT API] WARNING: No projects found in response structure")
                            continue  # Try next endpoint
                            
                    elif response.status_code == 404:
                        logger.warning(f"[PROJECT API] ERROR: 404 Not Found")
                        last_error = "404 Not Found"
                        continue
                    else:
                        logger.error(f"[PROJECT API] ERROR: HTTP {response.status_code}")
                        logger.error(f"[PROJECT API] Error Response: {response.text[:200]}")
                        last_error = f"HTTP {response.status_code}"
                        continue
                        
                except requests.exceptions.Timeout:
                    logger.error(f"[PROJECT API] ERROR: Timeout")
                    last_error = "Timeout"
                    continue
                except Exception as e:
                    logger.error(f"[PROJECT API] ERROR: Exception: {str(e)}")
                    last_error = str(e)
                    continue
            
            # If we got here, all endpoints failed
            logger.warning(f"[PROJECT API] WARNING: All endpoints failed. Returning empty project list.")
            logger.info("="*80)
            
            # Return last response if available, otherwise empty
            if last_response_data:
                return {"success": True, "data": {"projects": []}, "message": "No projects found"}
            else:
                return {"success": True, "data": {"projects": []}, "message": last_error or "No projects available"}
                
        except Exception as e:
            logger.error(f"[PROJECT API] ERROR: Exception: {str(e)}", exc_info=True)
            return {"success": False, "error": str(e)}
    
    def create_project(self, project_data: dict):
        """
        Create a new project
        POST /auth/api/project/client/{client_id}/
        
        project_data should include:
        - project_name: str
        - project_details: str (optional)
        - project_type: str (e.g., "software")
        - visibility: str (e.g., "private")
        - assignees: list (optional)
        """
        client_id = project_data.get('client_id') or self.client_id
        if not client_id:
            logger.error("No client_id available for creating project")
            return {"success": False, "error": "No client_id"}
        
        try:
            url = f"{BASE_URL}/project/client/{client_id}/"
            headers = self._get_headers()
            
            # Get user_id from auth_api
            user_id = getattr(self.auth_api, 'user_id', None) or self.auth_api.get_user_id_from_token()
            
            # Prepare payload
            payload = {
                "project_name": project_data.get('project_name', ''),
                "project_details": project_data.get('project_details', ''),
                "project_type": project_data.get('project_type', 'software'),
                "visibility": project_data.get('visibility', 'private'),
                "assignees": project_data.get('assignees', []),
                "client_id": client_id,
                "created_by": user_id,
                "updated_by": user_id
            }
            
            logger.info(f"[PROJECT API] Creating project: {payload['project_name']}")
            response = requests.post(url, headers=headers, json=payload, timeout=10)
            
            if response.status_code in [200, 201]:
                data = response.json()
                logger.info(f"[PROJECT API] SUCCESS: Project created successfully - {payload['project_name']}")
                return {"success": True, "data": data}
            else:
                error_msg = response.text[:200] if response.text else "No error message"
                logger.error(f"[PROJECT API] ERROR: Failed to create project - Status: {response.status_code}, Error: {error_msg}")
                return {"success": False, "error": f"HTTP {response.status_code}", "message": response.text}
                
        except Exception as e:
            logger.error(f"Error creating project: {e}")
            return {"success": False, "error": str(e)}
    
    def update_project(self, project_id: str, project_data: dict):
        """
        Update an existing project
        PUT /auth/api/project/client/{client_id}/project/{project_id}
        """
        client_id = self.client_id
        if not client_id:
            logger.error("No client_id available for updating project")
            return {"success": False, "error": "No client_id"}
        
        try:
            url = f"{BASE_URL}/project/client/{client_id}/project/{project_id}"
            headers = self._get_headers()
            
            # Get user_id from auth_api
            user_id = getattr(self.auth_api, 'user_id', None) or self.auth_api.get_user_id_from_token()
            project_data['updated_by'] = user_id
            
            logger.info(f"[PROJECT API] Updating project: {project_id}")
            response = requests.put(url, headers=headers, json=project_data, timeout=10)
            
            if response.status_code == 200:
                data = response.json()
                logger.info(f"[PROJECT API] SUCCESS: Project updated successfully - {project_id}")
                return {"success": True, "data": data}
            else:
                error_msg = response.text[:200] if response.text else "No error message"
                logger.error(f"[PROJECT API] ERROR: Failed to update project - Status: {response.status_code}, Error: {error_msg}")
                return {"success": False, "error": f"HTTP {response.status_code}"}
                
        except Exception as e:
            logger.error(f"Error updating project: {e}")
            return {"success": False, "error": str(e)}
    
    def fetch_project(self, project_id: str, client_id: str = None):
        """
        Fetch a single project by ID
        GET /project/client/{client_id}/{project_id}
        
        Args:
            project_id: Project UUID to fetch
            client_id: Client UUID (uses self.client_id if not provided)
        
        Returns:
            Dictionary with success flag and project data
        """
        logger.info("="*80)
        logger.info("[PROJECT API] ========== FETCH PROJECT START ==========")
        logger.info("="*80)
        
        cid = client_id or self.client_id
        logger.info(f"[PROJECT API] Client ID: {cid}")
        logger.info(f"[PROJECT API] Project ID: {project_id}")
        
        if not cid:
            error_msg = "No client_id available for fetching project"
            logger.error(f"[PROJECT API] ERROR: {error_msg}")
            return {"success": False, "error": error_msg, "data": None}
        
        if not project_id:
            error_msg = "Project ID is required"
            logger.error(f"[PROJECT API] ERROR: {error_msg}")
            return {"success": False, "error": error_msg, "data": None}
        
        try:
            url = f"{BASE_URL}/project/client/{cid}/{project_id}"
            headers = self._get_headers()
            
            logger.info(f"[PROJECT API] URL: {url}")
            logger.info(f"[PROJECT API] Method: GET")
            logger.debug(f"[PROJECT API] Headers: {json.dumps({k: v if k != 'authorization' else f'{v[:20]}...' for k, v in headers.items()}, indent=2)}")
            
            logger.info(f"Fetching project: {project_id} for client: {cid}")
            logger.info(f"[PROJECT API] Making request...")
            
            response = requests.get(url, headers=headers, timeout=10)
            status_code = response.status_code
            
            logger.info(f"[PROJECT API] Response Status Code: {status_code}")
            logger.debug(f"[PROJECT API] Response Headers: {dict(response.headers)}")
            
            if status_code == 200:
                try:
                    data = response.json()
                    logger.debug(f"[PROJECT API] Response Data: {json.dumps(data, indent=2, ensure_ascii=False)}")
                    logger.info(f"[PROJECT API] SUCCESS: Project fetched successfully")
                    logger.info("="*80)
                    return {"success": True, "data": data, "project": data}
                except Exception as json_error:
                    logger.error(f"[PROJECT API] ERROR: Failed to parse JSON response: {json_error}")
                    logger.error(f"[PROJECT API] Response Text: {response.text[:500]}")
                    return {"success": False, "error": f"Invalid JSON response: {str(json_error)}", "data": None}
            else:
                try:
                    error_data = response.json()
                    error_message = error_data.get("message", f"HTTP {status_code}")
                    logger.error(f"[PROJECT API] Error Response JSON: {json.dumps(error_data, indent=2)}")
                except Exception:
                    error_message = f"HTTP {status_code}: {response.text[:200]}"
                    logger.error(f"[PROJECT API] Error Response Text: {response.text[:500]}")
                
                logger.error(f"[PROJECT API] ERROR: Failed to fetch project ({status_code}) - {error_message}")
                logger.info("="*80)
                return {"success": False, "error": error_message, "data": None}
                
        except requests.exceptions.Timeout:
            error_msg = "Request timeout"
            logger.error(f"[PROJECT API] ERROR: {error_msg}")
            logger.info("="*80)
            return {"success": False, "error": error_msg, "data": None}
        except requests.exceptions.ConnectionError as conn_error:
            error_msg = f"Connection error: {str(conn_error)}"
            logger.error(f"[PROJECT API] ERROR: {error_msg}")
            logger.info("="*80)
            return {"success": False, "error": error_msg, "data": None}
        except Exception as e:
            error_msg = str(e)
            logger.error(f"[PROJECT API] ERROR: Exception - {error_msg}", exc_info=True)
            logger.info("="*80)
            return {"success": False, "error": error_msg, "data": None}
    
    def delete_project(self, project_id: str):
        """
        Delete a project
        DELETE /auth/api/project/client/{client_id}/project/{project_id}
        """
        client_id = self.client_id
        if not client_id:
            logger.error("No client_id available for deleting project")
            return {"success": False, "error": "No client_id"}
        
        try:
            url = f"{BASE_URL}/project/client/{client_id}/project/{project_id}"
            headers = self._get_headers()
            
            logger.info(f"[PROJECT API] Deleting project: {project_id}")
            response = requests.delete(url, headers=headers, timeout=10)
            
            if response.status_code == 200:
                logger.info(f"[PROJECT API] SUCCESS: Project deleted successfully - {project_id}")
                return {"success": True}
            else:
                error_msg = response.text[:200] if response.text else "No error message"
                logger.error(f"[PROJECT API] ERROR: Failed to delete project - Status: {response.status_code}, Error: {error_msg}")
                return {"success": False, "error": f"HTTP {response.status_code}"}
                
        except Exception as e:
            logger.error(f"Error deleting project: {e}")
            return {"success": False, "error": str(e)}

