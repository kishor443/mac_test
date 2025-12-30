"""
Resource path helper for PyInstaller compatibility.
This module provides a function to get resource paths that work both in development
and when packaged as an EXE with PyInstaller.
"""
import os
import sys
import shutil
import tempfile

# Cache for temporary HTML file paths
_temp_html_cache = {}


def resource_path(relative_path):
    """
    Get absolute path to resource, works for dev and PyInstaller.
    
    Args:
        relative_path: Path relative to the application root
        
    Returns:
        Absolute path to the resource
    """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except AttributeError:
        # Running in development mode
        base_path = os.path.abspath(os.path.dirname(os.path.dirname(__file__)))
    
    full_path = os.path.join(base_path, relative_path)
    
    # Normalize the path
    full_path = os.path.normpath(full_path)
    
    return full_path


def get_html_path(filename):
    """
    Get the file:// URL path for an HTML file in the gui directory.
    Works in both development and PyInstaller environments.
    
    For PyInstaller builds, copies HTML files to a temporary directory
    to avoid file:// URL access issues with webview.
    
    Args:
        filename: Name of the HTML file (e.g., 'login_screen.html')
        
    Returns:
        file:// URL path suitable for webview
    """
    # Check cache first
    if filename in _temp_html_cache:
        cached_path = _temp_html_cache[filename]
        if os.path.exists(cached_path):
            return _path_to_file_url(cached_path)
    
    # Get the original HTML file path
    html_path = resource_path(os.path.join('gui', filename))
    
    # Verify file exists
    if not os.path.exists(html_path):
        error_msg = f"HTML file not found: {html_path}"
        try:
            from utils.logger import logger
            logger.error(error_msg)
            logger.error(f"Base path: {os.path.dirname(html_path)}")
            logger.error(f"MEIPASS: {getattr(sys, '_MEIPASS', 'Not in PyInstaller')}")
            if os.path.exists(os.path.dirname(html_path)):
                logger.error(f"Files in directory: {os.listdir(os.path.dirname(html_path))}")
            else:
                logger.error("Directory does not exist")
                # Try to find where the file might be
                base_path = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(os.path.dirname(__file__))))
                logger.error(f"Searching in base path: {base_path}")
                if os.path.exists(base_path):
                    logger.error(f"Files in base path: {os.listdir(base_path)}")
        except Exception as e:
            try:
                from utils.logger import logger
                logger.error(f"Error during debugging: {e}")
            except:
                pass
        # Return the path anyway - webview will show the error
        return _path_to_file_url(html_path)
    
    # Check if we're running in PyInstaller
    is_pyinstaller = hasattr(sys, '_MEIPASS')
    
    if is_pyinstaller:
        # In PyInstaller, copy HTML file to a temporary directory
        # This avoids file:// URL access issues with webview
        try:
            temp_dir = tempfile.gettempdir()
            app_temp_dir = os.path.join(temp_dir, 'ProductivityTracker')
            os.makedirs(app_temp_dir, exist_ok=True)
            
            temp_html_path = os.path.join(app_temp_dir, filename)
            shutil.copy2(html_path, temp_html_path)
            
            # Cache the path
            _temp_html_cache[filename] = temp_html_path
            
            return _path_to_file_url(temp_html_path)
        except Exception as e:
            # If copying fails, fall back to original path
            try:
                from utils.logger import logger
                logger.warning(f"Failed to copy HTML to temp directory: {e}, using original path")
            except:
                pass
    
    # In development mode or if copy failed, use original path
    return _path_to_file_url(html_path)


def _path_to_file_url(file_path):
    """
    Convert a file path to a file:// URL.
    
    Args:
        file_path: Absolute file path
        
    Returns:
        file:// URL string
    """
    # Convert to absolute path
    file_path = os.path.abspath(file_path)
    
    # Convert to file:// URL format
    # Use pathname2url for proper encoding
    from urllib.request import pathname2url
    
    # On Windows, we need to handle drive letters specially
    if sys.platform == 'win32':
        # pathname2url handles Windows paths correctly
        url_path = pathname2url(file_path)
        # Ensure it starts with file://
        if not url_path.startswith('file://'):
            # Windows needs file:/// (three slashes)
            url_path = 'file:///' + url_path.lstrip('/')
        return url_path
    else:
        # Unix-like systems
        url_path = pathname2url(file_path)
        if not url_path.startswith('file://'):
            url_path = 'file://' + url_path
        return url_path

