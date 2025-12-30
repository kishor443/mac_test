import socket
import json
import urllib.request

def is_online(host="8.8.8.8", port=53, timeout=3):
    try:
        socket.setdefaulttimeout(timeout)
        socket.socket(socket.AF_INET, socket.SOCK_STREAM).connect((host, port))
        return True
    except Exception:
        return False

def get_local_ip() -> str:
    try:
        hostname = socket.gethostname()
        return socket.gethostbyname(hostname)
    except Exception:
        return ""

def get_public_ip(timeout: int = 3) -> str:
    """
    Returns the public IPv4 address, falling back to local IP if external lookup fails.
    """
    try:
        with urllib.request.urlopen("https://api.ipify.org?format=json", timeout=timeout) as r:
            data = json.loads(r.read().decode("utf-8"))
            ip = data.get("ip")
            if isinstance(ip, str) and ip:
                return ip
    except Exception:
        pass
    return get_local_ip()
