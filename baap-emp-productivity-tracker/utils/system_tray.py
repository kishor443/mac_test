import tkinter as tk
try:
    import pystray
    from PIL import Image, ImageDraw
    STRAY_AVAILABLE = True
except ImportError:
    STRAY_AVAILABLE = False

def create_tray_icon(main_window, on_quit=None):
    """Create a system tray icon for the application"""
    if not STRAY_AVAILABLE:
        return None
    
    try:
        # Create a simple icon
        image = Image.new('RGB', (64, 64), color='white')
        draw = ImageDraw.Draw(image)
        # Draw a simple "P" for Productivity Tracker
        draw.text((20, 15), "P", fill='black')
        draw.text((20, 35), "T", fill='black')
        
        menu = pystray.Menu(
            pystray.MenuItem("Show Window", lambda: _show_window(main_window)),
            pystray.MenuItem("Quit", lambda: _quit_app(main_window, on_quit))
        )
        
        icon = pystray.Icon("ProductivityTracker", image, "Productivity Tracker", menu)
        return icon
    except Exception:
        return None

def _show_window(main_window):
    """Show the main window"""
    if main_window and main_window.window:
        try:
            main_window.window.deiconify()
            main_window.window.lift()
            main_window.window.attributes('-topmost', True)
            main_window.window.after(100, lambda: main_window.window.attributes('-topmost', False))
        except Exception:
            pass

def _quit_app(main_window, on_quit):
    """Quit the application"""
    if on_quit:
        try:
            on_quit()
        except Exception:
            pass
    if main_window and main_window.window:
        try:
            main_window.window.destroy()
        except Exception:
            pass
