"""
Windows Toast Notifications utility
Shows native Windows 10/11 toast notifications
"""
import logging

logger = logging.getLogger(__name__)

# Try to use win10toast library (simpler)
TOAST_AVAILABLE = False
USE_WIN10TOAST = False
USE_WIN32 = False

try:
    from win10toast import ToastNotifier
    TOAST_AVAILABLE = True
    USE_WIN10TOAST = True
except ImportError:
    logger.debug("win10toast not available, trying alternative method")
    # Fallback to Windows Toast Notification API
    try:
        import win32api
        import win32con
        import win32gui
        TOAST_AVAILABLE = True
        USE_WIN32 = True
    except ImportError:
        logger.warning("Neither win10toast nor win32api available for notifications")

def show_notification(title: str, message: str, duration: int = 5, icon_path: str = None):
    """
    Show a Windows toast notification
    
    Args:
        title: Notification title
        message: Notification message
        duration: Duration in seconds (default 5)
        icon_path: Path to icon file (optional)
    """
    try:
        if not TOAST_AVAILABLE:
            logger.warning("Toast notifications not available")
            return False
        
        # Method 1: Use win10toast (simpler and better - native toast)
        if USE_WIN10TOAST:
            try:
                from win10toast import ToastNotifier
                toaster = ToastNotifier()
                toaster.show_toast(
                    title=title,
                    msg=message,
                    duration=duration,
                    icon_path=icon_path,
                    threaded=True
                )
                logger.debug(f"Shown toast notification: {title} - {message}")
                return True
            except Exception as e:
                logger.debug(f"win10toast failed: {e}, trying alternative")
        
        # Method 2: Use Windows API directly (fallback - shows MessageBox)
        if USE_WIN32:
            try:
                import win32api
                import win32con
                
                # Create a simple notification using MessageBox (fallback)
                # This is not a toast but a simple notification
                win32api.MessageBox(
                    0,
                    message,
                    title,
                    win32con.MB_OK | win32con.MB_ICONINFORMATION
                )
                logger.debug(f"Shown notification (MessageBox): {title} - {message}")
                return True
            except Exception as e:
                logger.error(f"Windows API notification failed: {e}")
                return False
        
        return False
            
    except Exception as e:
        logger.error(f"Error showing notification: {e}")
        return False

def show_teams_notification(sender: str, message: str, duration: int = 5):
    """
    Show a Teams-specific notification (custom Teams-style window)
    
    NOTE: Teams notifications are disabled - this function does nothing
    
    Args:
        sender: Sender name
        message: Message content
        duration: Duration in seconds (default 5)
    """
    # Teams notifications disabled - return immediately without showing anything
    return False

