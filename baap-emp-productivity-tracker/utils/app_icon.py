"""
Utility to extract application icons from Windows executables
"""
import os
import base64
import tempfile
import time
from typing import Optional
from functools import lru_cache
import threading

# Icon cache to avoid repeated heavy operations
_icon_cache = {}
_cache_lock = threading.Lock()

try:
    import win32gui
    import win32process
    import win32api
    import win32con
    from PIL import Image
    import io
except ImportError:
    win32gui = None
    win32process = None
    win32api = None
    win32con = None
    Image = None

# Cache for process name to exe path mapping (updated less frequently)
_process_path_cache = {}
_process_cache_lock = threading.Lock()
_last_process_scan = 0
PROCESS_CACHE_TTL = 300  # Cache process paths for 5 minutes

def _get_process_path_cached(process_name_exe: str) -> Optional[str]:
    """Get process executable path with caching to avoid scanning all processes frequently"""
    global _process_path_cache, _last_process_scan
    
    current_time = time.time() if 'time' in globals() else 0
    
    with _process_cache_lock:
        # Check cache first
        if process_name_exe.lower() in _process_path_cache:
            cached_path, cached_time = _process_path_cache[process_name_exe.lower()]
            if os.path.exists(cached_path) and (current_time - cached_time) < PROCESS_CACHE_TTL:
                return cached_path
        
        # Only scan processes if cache is stale (every 5 minutes)
        if (current_time - _last_process_scan) > PROCESS_CACHE_TTL:
            try:
                import psutil
                import time as time_module
                _last_process_scan = time_module.time()
                
                # Limit scan to first 100 processes to avoid hanging
                count = 0
                for proc in psutil.process_iter(['name', 'exe']):
                    if count >= 100:  # Limit to prevent hanging
                        break
                    count += 1
                    try:
                        if proc.info['name'] and proc.info['exe'] and os.path.exists(proc.info['exe']):
                            name_lower = proc.info['name'].lower()
                            _process_path_cache[name_lower] = (proc.info['exe'], _last_process_scan)
                    except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                        continue
            except ImportError:
                pass
        
        # Check cache again after potential update
        if process_name_exe.lower() in _process_path_cache:
            cached_path, _ = _process_path_cache[process_name_exe.lower()]
            if os.path.exists(cached_path):
                return cached_path
    
    return None

def get_app_icon_base64(app_name: str, process_name: str = None) -> Optional[str]:
    """
    Get application icon as base64 encoded image (with caching for performance)
    
    Args:
        app_name: Application name
        process_name: Process executable name (e.g., 'chrome.exe', 'python.exe')
    
    Returns:
        Base64 encoded PNG image string or None
    """
    # Check cache first
    cache_key = f"{app_name}_{process_name or ''}"
    with _cache_lock:
        if cache_key in _icon_cache:
            return _icon_cache[cache_key]
    
    try:
        if not win32gui or not Image:
            return None
        
        # Try to find executable path
        exe_path = None
        
        # If process_name is provided, try to find it using cached process paths
        if process_name:
            if not process_name.endswith('.exe'):
                process_name_exe = process_name + '.exe'
            else:
                process_name_exe = process_name
            
            # Use cached process path lookup (much faster)
            exe_path = _get_process_path_cached(process_name_exe)
        
        # If we found an executable, extract icon
        if exe_path and os.path.exists(exe_path):
            try:
                # Extract icon using win32api
                large, small = win32gui.ExtractIconEx(exe_path, 0)
                if large and len(large) > 0:
                    # Get icon handle
                    icon_handle = large[0]
                    
                    # Get icon info
                    icon_info = win32gui.GetIconInfo(icon_handle)
                    hbm = icon_info[3]  # Bitmap handle
                    
                    # Convert to PIL Image
                    bmp = win32gui.GetObject(hbm)
                    width = bmp.bmWidth
                    height = bmp.bmHeight
                    
                    # Get bitmap bits
                    bitmap_data = win32gui.GetBitmapBits(hbm, width * height * 4)
                    
                    # Convert to PIL Image
                    img = Image.frombytes('RGBA', (width, height), bitmap_data, 'raw', 'BGRA', 0, 1)
                    
                    # Resize to 32x32
                    img = img.resize((32, 32), Image.Resampling.LANCZOS)
                    
                    # Convert to base64
                    buffer = io.BytesIO()
                    img.save(buffer, format='PNG')
                    img_str = base64.b64encode(buffer.getvalue()).decode()
                    result = f"data:image/png;base64,{img_str}"
                    
                    # Cache the result
                    with _cache_lock:
                        _icon_cache[cache_key] = result
                        # Limit cache size to prevent memory issues
                        if len(_icon_cache) > 100:
                            # Remove oldest entries (simple FIFO)
                            oldest_key = next(iter(_icon_cache))
                            del _icon_cache[oldest_key]
                    
                    return result
            except Exception as e:
                # Cache None result to avoid repeated failed attempts
                with _cache_lock:
                    _icon_cache[cache_key] = None
                pass
        
        # Cache None result
        with _cache_lock:
            _icon_cache[cache_key] = None
        return None
    except Exception as e:
        # Cache None result on error
        with _cache_lock:
            _icon_cache[cache_key] = None
        return None

