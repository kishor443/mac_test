"""
Custom Teams-style notification window
Creates a notification that looks exactly like Microsoft Teams notifications
"""
import tkinter as tk
from tkinter import ttk
import threading
import time
import logging
from PIL import Image, ImageDraw, ImageFont
import io
import base64

logger = logging.getLogger(__name__)

class TeamsNotificationWindow:
    """Custom Teams-style notification window"""
    
    def __init__(self, sender: str, message: str, duration: int = 5):
        self.sender = sender
        self.message = message
        self.duration = duration
        self.window = None
        self.closed = False
        
    def _create_avatar_image(self, name: str, size: int = 48):
        """Create a circular avatar with initials"""
        try:
            # Get first letter of first name and first letter of last name
            parts = name.split()
            if len(parts) >= 2:
                initials = (parts[0][0] + parts[-1][0]).upper()
            else:
                initials = name[0:2].upper() if len(name) >= 2 else name[0].upper()
            
            # Create image with Teams purple background
            img = Image.new('RGB', (size, size), color='#6264A7')
            draw = ImageDraw.Draw(img)
            
            # Draw circle
            draw.ellipse([0, 0, size-1, size-1], fill='#6264A7')
            
            # Draw text (initials)
            try:
                # Try to use a nice font
                font_size = size // 2
                font = ImageFont.truetype("arial.ttf", font_size)
            except:
                font = ImageFont.load_default()
            
            # Calculate text position (center)
            bbox = draw.textbbox((0, 0), initials, font=font)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]
            x = (size - text_width) // 2
            y = (size - text_height) // 2
            
            draw.text((x, y), initials, fill='white', font=font)
            
            # Convert to PhotoImage
            from PIL import ImageTk
            return ImageTk.PhotoImage(img)
        except Exception as e:
            logger.debug(f"Error creating avatar: {e}")
            # Return a simple colored circle
            img = Image.new('RGB', (size, size), color='#6264A7')
            from PIL import ImageTk
            return ImageTk.PhotoImage(img)
    
    def show(self, root=None):
        """Show the notification window"""
        try:
            if root is None:
                root = _get_notification_root()
                if root is None:
                    logger.error("No root window available for notification")
                    return
            
            # Create root window
            self.window = tk.Toplevel(root)
            self.window.overrideredirect(True)  # Remove window decorations
            self.window.attributes('-topmost', True)  # Always on top
            self.window.attributes('-alpha', 0.95)  # Slight transparency
            
            # Teams purple background color
            teams_purple = '#6264A7'
            teams_dark = '#464EB8'
            
            # Set window size
            width = 360
            height = 120
            self.window.geometry(f'{width}x{height}')
            
            # Position in top-right corner, stack multiple notifications
            screen_width = self.window.winfo_screenwidth()
            screen_height = self.window.winfo_screenheight()
            x = screen_width - width - 20
            
            # Stack notifications vertically (each new one appears below previous)
            global _active_notifications
            active_count = len([n for n in _active_notifications if not n.closed and n.window])
            y = 20 + (active_count * (height + 10))  # Stack with 10px gap
            
            self.window.geometry(f'{width}x{height}+{x}+{y}')
            
            # Main container
            container = tk.Frame(self.window, bg=teams_purple, padx=0, pady=0)
            container.pack(fill=tk.BOTH, expand=True)
            
            # Header
            header = tk.Frame(container, bg=teams_purple, height=32)
            header.pack(fill=tk.X, padx=12, pady=(8, 0))
            
            # Microsoft Teams logo and text
            logo_frame = tk.Frame(header, bg=teams_purple)
            logo_frame.pack(side=tk.LEFT)
            
            # Teams logo (simple "TT" text for now)
            logo_label = tk.Label(
                logo_frame,
                text="TT",
                bg=teams_purple,
                fg='white',
                font=('Segoe UI', 10, 'bold')
            )
            logo_label.pack(side=tk.LEFT, padx=(0, 6))
            
            teams_label = tk.Label(
                logo_frame,
                text="Microsoft Teams",
                bg=teams_purple,
                fg='white',
                font=('Segoe UI', 10)
            )
            teams_label.pack(side=tk.LEFT)
            
            # Header buttons (options and close)
            buttons_frame = tk.Frame(header, bg=teams_purple)
            buttons_frame.pack(side=tk.RIGHT)
            
            # Options button (three dots)
            options_btn = tk.Label(
                buttons_frame,
                text="⋯",
                bg=teams_purple,
                fg='white',
                font=('Segoe UI', 14),
                cursor='hand2',
                padx=4
            )
            options_btn.pack(side=tk.LEFT)
            
            # Close button
            close_btn = tk.Label(
                buttons_frame,
                text="✕",
                bg=teams_purple,
                fg='white',
                font=('Segoe UI', 12, 'bold'),
                cursor='hand2',
                padx=4
            )
            close_btn.pack(side=tk.LEFT)
            close_btn.bind('<Button-1>', lambda e: self.close())
            
            # Content area
            content = tk.Frame(container, bg=teams_purple)
            content.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 12))
            
            # Avatar
            avatar_img = self._create_avatar_image(self.sender)
            avatar_label = tk.Label(
                content,
                image=avatar_img,
                bg=teams_purple
            )
            avatar_label.image = avatar_img  # Keep a reference
            avatar_label.pack(side=tk.LEFT, padx=(0, 12))
            
            # Message content
            message_frame = tk.Frame(content, bg=teams_purple)
            message_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            # Sender name
            sender_label = tk.Label(
                message_frame,
                text=self.sender,
                bg=teams_purple,
                fg='white',
                font=('Segoe UI', 12, 'bold'),
                anchor='w'
            )
            sender_label.pack(fill=tk.X, pady=(0, 4))
            
            # Message preview
            # Truncate message if too long
            display_message = self.message
            if len(display_message) > 60:
                display_message = display_message[:57] + "..."
            
            message_label = tk.Label(
                message_frame,
                text=display_message,
                bg=teams_purple,
                fg='#EFEFEF',  # Light gray/white color (tkinter doesn't support rgba)
                font=('Segoe UI', 11),
                anchor='w',
                wraplength=240,
                justify='left'
            )
            message_label.pack(fill=tk.X)
            
            # Make window clickable to open Teams
            def on_click(event):
                try:
                    from utils.teams_notifications import open_teams_chat_by_identifier
                    open_teams_chat_by_identifier(self.sender)
                except Exception as e:
                    logger.debug(f"Error opening Teams chat: {e}")
                self.close()
            
            container.bind('<Button-1>', on_click)
            for widget in [content, message_frame, sender_label, message_label, avatar_label]:
                widget.bind('<Button-1>', on_click)
            
            # Auto-close after duration
            self.window.after(self.duration * 1000, self.close)
            
            # Fade in animation
            self._fade_in()
            
        except Exception as e:
            logger.error(f"Error showing Teams notification: {e}")
    
    def _fade_in(self):
        """Fade in animation"""
        try:
            alpha = 0.0
            step = 0.05
            def animate():
                nonlocal alpha
                if alpha < 0.95 and not self.closed:
                    alpha += step
                    self.window.attributes('-alpha', alpha)
                    self.window.after(20, animate)
            animate()
        except Exception:
            pass
    
    def close(self):
        """Close the notification window"""
        if self.closed:
            return
        self.closed = True
        try:
            if self.window:
                # Remove from active notifications
                global _active_notifications
                if self in _active_notifications:
                    _active_notifications.remove(self)
                
                # Fade out animation
                def fade_out():
                    try:
                        alpha = self.window.attributes('-alpha')
                        if alpha > 0:
                            alpha -= 0.1
                            self.window.attributes('-alpha', alpha)
                            self.window.after(30, fade_out)
                        else:
                            self.window.destroy()
                    except Exception:
                        if self.window:
                            self.window.destroy()
                fade_out()
        except Exception as e:
            logger.debug(f"Error closing notification: {e}")

# Global tkinter root for notifications
_notification_root = None
_active_notifications = []  # Track active notifications for stacking

def _get_notification_root():
    """Get or create a tkinter root for notifications"""
    global _notification_root
    if _notification_root is None:
        try:
            _notification_root = tk.Tk()
            _notification_root.withdraw()  # Hide the root window
            _notification_root.attributes('-topmost', False)
            # Keep root alive - schedule update only if main loop is running
            try:
                _notification_root.after(100, _update_notifications)
            except Exception as e:
                logger.debug(f"Could not schedule initial notification update: {e}")
        except Exception as e:
            logger.error(f"Error creating notification root: {e}")
            return None
    else:
        # Check if existing root is still valid
        try:
            if not _notification_root.winfo_exists():
                _notification_root = None
                return _get_notification_root()  # Recursively create new root
        except (tk.TclError, RuntimeError):
            _notification_root = None
            return _get_notification_root()  # Recursively create new root
    return _notification_root

def _update_notifications():
    """Update all active notifications"""
    global _notification_root, _active_notifications
    if _notification_root:
        try:
            # Check if root window is still valid
            if not _notification_root.winfo_exists():
                # Root window was destroyed, clear it
                _notification_root = None
                return
            
            # Remove closed notifications
            _active_notifications = [n for n in _active_notifications if not n.closed and n.window]
            
            # Update root
            try:
                _notification_root.update()
            except Exception as e:
                # Root window destroyed or main loop not running
                logger.debug(f"Notification root update failed: {e}")
                _notification_root = None
                return
            
            # Schedule next update only if root is still valid
            try:
                _notification_root.after(100, _update_notifications)
            except Exception as e:
                # Root window destroyed or main loop not running
                logger.debug(f"Could not schedule notification update: {e}")
                _notification_root = None
                return
        except Exception as e:
            # Any other error - clear root and stop updates
            logger.debug(f"Error in _update_notifications: {e}")
            _notification_root = None

def show_teams_notification_window(sender: str, message: str, duration: int = 5):
    """
    Show a Teams-style notification window (auto-updates for new messages)
    
    Args:
        sender: Sender name
        message: Message content
        duration: Duration in seconds (default 5)
    """
    try:
        root = _get_notification_root()
        if root is None:
            logger.error("Could not create notification root")
            return False
        
        notification = TeamsNotificationWindow(sender, message, duration)
        notification.show(root)
        
        # Add to active notifications list
        global _active_notifications
        _active_notifications.append(notification)
        
        # Update root to show notification immediately
        root.update()
        
        logger.info(f"Showing Teams notification: {sender} - {message[:50]}")
        return True
    except Exception as e:
        logger.error(f"Error showing Teams notification window: {e}")
        return False

