"""
GUARANTEED terminal logging - writes to BOTH terminal AND file.
This ensures logs are visible even when webview blocks stdout.
"""
import sys
import os
from datetime import datetime

# Log file path
_log_file_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "terminal_output.log")

def _write_both(message):
    """Write to BOTH terminal AND file - guaranteed to work."""
    # Write to file (always works)
    try:
        with open(_log_file_path, 'a', encoding='utf-8', errors='replace') as f:
            timestamp = datetime.now().strftime("%H:%M:%S")
            f.write(f"[{timestamp}] {message}")
            f.flush()
    except:
        pass
    
    # Write to terminal (may be blocked by webview)
    try:
        # Try direct file descriptor write
        fd = sys.stdout.fileno()
        os.write(fd, message.encode('utf-8', errors='replace'))
    except:
        try:
            # Fallback to stdout
            sys.stdout.write(message)
            sys.stdout.flush()
        except:
            pass

def log(message):
    """Log message to both terminal and file."""
    _write_both(message)

def log_api(api_name, method, url, payload=None, response=None, error=None):
    """Log API call with full details."""
    if error:
        _write_both(f"\n[ERROR] {api_name} → {error}\n")
    elif response:
        status = response.get('status_code', '?')
        _write_both(f"\n[API] {api_name} ({method}) → {status}\n")
        _write_both(f"  URL: {url}\n")
        if payload:
            import json
            _write_both(f"  Payload: {json.dumps(payload, indent=2, default=str)}\n")
        _write_both(f"  Response: {json.dumps(response, indent=2, default=str)[:500]}\n")
    else:
        _write_both(f"\n[API] {api_name} ({method})\n")
        _write_both(f"  URL: {url}\n")
        if payload:
            import json
            _write_both(f"  Payload: {json.dumps(payload, indent=2, default=str)}\n")

def log_token(token_type, preview, length, user_id=None):
    """Log token generation."""
    _write_both(f"\n[TOKEN] {token_type} Generated\n")
    _write_both(f"  Preview: {preview}\n")
    _write_both(f"  Length: {length}\n")
    if user_id:
        _write_both(f"  User ID: {user_id}\n")

# Initialize log file
try:
    with open(_log_file_path, 'w', encoding='utf-8') as f:
        f.write(f"\n{'='*80}\n")
        f.write(f"TERMINAL OUTPUT LOG - Started at {datetime.now()}\n")
        f.write(f"{'='*80}\n\n")
except:
    pass


