try:
    import win32gui
    import win32process
except Exception:
    win32gui = None
    win32process = None

def get_active_window_title() -> str:
    try:
        if not win32gui:
            return ""
        hwnd = win32gui.GetForegroundWindow()
        if not hwnd:
            return ""
        title = win32gui.GetWindowText(hwnd)
        return title or ""
    except Exception:
        return ""


