import threading
import win32con
import win32gui
import win32api
import win32event
import win32ts
import ctypes

# Import logger with safe fallback
try:
    from utils.logger import logger
except Exception:
    import logging
    logger = logging.getLogger("WinEventHook")
    logger.addHandler(logging.NullHandler())

WM_WTSSESSION_CHANGE = 0x02B1
WM_POWERBROADCAST = 0x218
WM_QUERYENDSESSION = 0x11
WTS_SESSION_LOCK = 0x7
WTS_SESSION_UNLOCK = 0x8
PBT_APMSUSPEND = 0x4
PBT_APMRESUMEAUTOMATIC = 0x12


class WinEventHook:
    def __init__(self, on_lock, on_unlock, on_sleep, on_wake, on_shutdown):
        self.on_lock = on_lock
        self.on_unlock = on_unlock
        self.on_sleep = on_sleep
        self.on_wake = on_wake
        self.on_shutdown = on_shutdown
        self.thread = threading.Thread(target=self._msg_loop, daemon=True)
        self.thread.start()

    def _msg_loop(self):
        wc = win32gui.WNDCLASS()
        hinst = wc.hInstance = win32api.GetModuleHandle(None)
        wc.lpszClassName = "WinEventHookListener"
        wc.lpfnWndProc = self._wnd_proc
        class_atom = win32gui.RegisterClass(wc)
        hwnd = win32gui.CreateWindow(wc.lpszClassName, "", 0, 0, 0, 0, 0, 0, 0, hinst, None)

        # Register for session notifications
        win32ts.WTSRegisterSessionNotification(hwnd, win32ts.NOTIFY_FOR_THIS_SESSION)

        # Message loop
        while True:
            bRet, msg = win32gui.GetMessage(hwnd, 0, 0)
            if bRet == 0:
                break  # WM_QUIT
            elif bRet == -1:
                continue  # error
            win32gui.TranslateMessage(msg)
            win32gui.DispatchMessage(msg)

        win32ts.WTSUnRegisterSessionNotification(hwnd)

    def _wnd_proc(self, hwnd, msg, wparam, lparam):
        if msg == WM_WTSSESSION_CHANGE:
            if wparam == WTS_SESSION_LOCK:
                try:
                    logger.info("System locked — switched to Break Mode")
                except Exception:
                    pass  # Silently fail if logger is not available
                if self.on_lock:
                    self.on_lock()
            elif wparam == WTS_SESSION_UNLOCK:
                try:
                    logger.info("System unlocked/resumed — showing break status UI")
                except Exception:
                    pass  # Silently fail if logger is not available
                if self.on_unlock:
                    self.on_unlock()
        elif msg == WM_POWERBROADCAST:
            if wparam == PBT_APMSUSPEND:
                try:
                    logger.info("System sleep/hibernate — switched to Break Mode")
                except Exception:
                    pass  # Silently fail if logger is not available
                if self.on_sleep:
                    self.on_sleep()
            elif wparam == PBT_APMRESUMEAUTOMATIC:
                try:
                    logger.info("System resumed from sleep/hibernate — showing break status UI")
                except Exception:
                    pass  # Silently fail if logger is not available
                if self.on_wake:
                    self.on_wake()
        elif msg == WM_QUERYENDSESSION:
            try:
                logger.info("System shutting down — auto Clock Out completed")
            except Exception:
                pass  # Silently fail if logger is not available
            if self.on_shutdown:
                self.on_shutdown()
            return 1  # allow shutdown
        return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)
