import sys, os, platform
from typing import Optional

def _startup_folder() -> Optional[str]:
    try:
        return os.path.join(os.environ.get("APPDATA", ""), "Microsoft", "Windows", "Start Menu", "Programs", "Startup")
    except Exception:
        return None

def _executable_path() -> str:
    # If bundled (PyInstaller), use the frozen executable; otherwise use current Python executable with script
    if getattr(sys, 'frozen', False):
        return sys.executable
    # Fallback to launching this project via pythonw if available
    pythonw = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
    if os.path.exists(pythonw):
        # Try to launch the repo's main module
        repo_root = os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir))
        main_py = os.path.join(repo_root, "main.py")
        if os.path.exists(main_py):
            return f'"{pythonw}" "{main_py}"'
    # Last resort: current interpreter and argv
    return f'"{sys.executable}" "{sys.argv[0]}"'

def _create_startup_shortcut(app_name: str) -> bool:
    startup = _startup_folder()
    if not startup or not os.path.isdir(startup):
        return False
    try:
        try:
            from win32com.client import Dispatch  # type: ignore
        except Exception:
            return False
        target = _executable_path()
        shortcut_path = os.path.join(startup, f"{app_name}.lnk")
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortcut(shortcut_path)
        # If target is a command line with quotes, split target and args
        if target.startswith('"') and '" ' in target:
            exe, args = target.split('" ', 1)
            exe = exe.strip('"')
            shortcut.TargetPath = exe
            shortcut.Arguments = args
        else:
            shortcut.TargetPath = target.strip('"')
        shortcut.WorkingDirectory = os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir)))
        shortcut.IconLocation = sys.executable if getattr(sys, 'frozen', False) else sys.executable
        shortcut.save()
        return True
    except Exception:
        return False

def _create_registry_run(app_name: str) -> bool:
    try:
        import winreg  # type: ignore
    except Exception:
        return False
    try:
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as key:
            cmd = _executable_path()
            winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, cmd)
        return True
    except Exception:
        return False

def enable_auto_startup(app_name="ProductivityTracker") -> bool:
    system = platform.system()
    if system != "Windows":
        return False
    # Prefer Startup shortcut; fallback to HKCU Run registry value
    if _create_startup_shortcut(app_name):
        return True
    return _create_registry_run(app_name)

def check_admin_privileges():
    try:
        import ctypes  # type: ignore
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False
