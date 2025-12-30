"""
Test script to find the correct task endpoint
"""
import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from api.auth_api import AuthAPI
from config import ERP_DEVICE_ID, BASE_URL
import requests
import json


def main():
    print("="*60)
    print("TESTING TASK ENDPOINTS")
    print("="*60)
    
    # Login
    auth = AuthAPI()
    if not (auth.refresh_token and auth.user_id):
        print("Please login first")
        return
        
    success, _, _ = auth.login_with_refresh(
        refresh_token=auth.refresh_token,
        user_id=auth.user_id,
        device_id=ERP_DEVICE_ID
    )
    
    if not success:
        print("Login failed")
        return
    
    # Get client
    ok, _, clients_response = auth.fetch_clients(user_id=auth.user_id, device_id=ERP_DEVICE_ID)
    clients = clients_response.get("clients", []) if isinstance(clients_response, dict) else clients_response
    
    if not clients:
        print("No clients found")
        return
        
    client_id = clients[0].get("id")
    user_id = auth.user_id
    print(f"Client ID: {client_id}")
    print(f"User ID: {user_id}")
    print(f"Base URL: {BASE_URL}\n")
    
    # Try different endpoints
    headers = {
        "authorization": f"Bearer {auth.access_token}",
        "accept": "application/json, text/plain, */*",
        "accept-language": "en-US,en;q=0.9",
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
    
    endpoints_to_try = [
        f"/task/client/{client_id}",
        f"/task/client/{client_id}/user/{user_id}",
        f"/task/client/{client_id}/tasks",
        f"/tasks/client/{client_id}",
        f"/tasks/client/{client_id}/user/{user_id}",
        f"/task/list/client/{client_id}",
        f"/task/client/{client_id}/assigned",
        f"/task/client/{client_id}/assigned/{user_id}",
    ]
    
    print("="*60)
    print("TRYING DIFFERENT ENDPOINTS...")
    print("="*60)
    
    for endpoint in endpoints_to_try:
        url = BASE_URL + endpoint
        print(f"\nTrying: {url}")
        
        try:
            response = requests.get(url, headers=headers, timeout=10)
            print(f"Status: {response.status_code}")
            
            if response.status_code == 200:
                data = response.json()
                print(f"SUCCESS! Found endpoint:")
                print(json.dumps(data, indent=2, default=str)[:500])
                print("\n... (truncated)")
                print(f"\nSUCCESS: Working endpoint: {url}")
                return
            elif response.status_code == 404:
                print("ERROR: Not found")
            else:
                print(f"WARNING: Error: {response.status_code}")
                try:
                    print(response.json())
                except:
                    pass
        except Exception as e:
            print(f"ERROR: Error: {str(e)}")
    
    print("\n" + "="*60)
    print("No working endpoint found.")
    print("Please check your API documentation for the correct endpoint")
    print("="*60)


if __name__ == "__main__":
    main()

