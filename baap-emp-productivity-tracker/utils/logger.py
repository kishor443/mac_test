import logging
import os
import sys
import shutil
import io

APP_NAME = "ProductivityTracker"

# Ensure UTF-8 encoding with error replacement for stdout/stderr
# Note: If streams are closed, we skip wrapping and rely on SafeStreamHandler
os.environ["PYTHONUTF8"] = "1"
try:
    if hasattr(sys.stdout, 'buffer'):
        try:
            # Test if buffer is accessible
            _ = sys.stdout.buffer
            if not isinstance(sys.stdout, io.TextIOWrapper):
                sys.stdout = io.TextIOWrapper(
                    sys.stdout.buffer,
                    encoding="utf-8",
                    errors="replace"
                )
        except (ValueError, OSError, AttributeError, TypeError):
            # stdout buffer is closed or invalid - SafeStreamHandler will handle it
            pass
    if hasattr(sys.stderr, 'buffer'):
        try:
            # Test if buffer is accessible
            _ = sys.stderr.buffer
            if not isinstance(sys.stderr, io.TextIOWrapper):
                sys.stderr = io.TextIOWrapper(
                    sys.stderr.buffer,
                    encoding="utf-8",
                    errors="replace"
                )
        except (ValueError, OSError, AttributeError, TypeError):
            # stderr buffer is closed or invalid - SafeStreamHandler will handle it
            pass
except Exception:
    # Ultimate fallback - SafeStreamHandler will handle Unicode errors
    pass

def get_app_data_dir():
    """Return the writable app data directory for the current user."""
    appdata = os.getenv("APPDATA")  # e.g. C:\Users\<user>\AppData\Roaming
    app_dir = os.path.join(appdata, APP_NAME)

    # Create the main folder and subfolders if needed
    os.makedirs(app_dir, exist_ok=True)

    # Create data subfolder
    data_dir = os.path.join(app_dir, "data")
    os.makedirs(data_dir, exist_ok=True)

    return data_dir


def setup_logger(name=APP_NAME):
    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)

    data_dir = get_app_data_dir()
    log_file = os.path.join(data_dir, "app.log")

    # File handler
    fh = logging.FileHandler(log_file, encoding='utf-8')
    fh.setLevel(logging.DEBUG)
    file_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    fh.setFormatter(file_formatter)

    # Console handler (for terminal output) - with safe encoding
    class SafeStreamHandler(logging.StreamHandler):
        """StreamHandler that safely handles Unicode encoding errors."""
        def format(self, record):
            """Override format to sanitize Unicode before formatting."""
            try:
                # Get the original formatted message
                msg = super().format(record)
                # Sanitize to ASCII-safe
                return msg.encode("ascii", "replace").decode("ascii")
            except Exception:
                # If formatting fails, return safe fallback
                try:
                    return f"{record.levelname} - [LogError: Could not format message]"
                except:
                    return "[LogError]"
        
        def emit(self, record):
            """Emit a record, with full Unicode protection."""
            try:
                # Format with Unicode sanitization
                msg = self.format(record)
                stream = self.stream
                # Write with error handling
                try:
                    stream.write(msg + self.terminator)
                    self.flush()
                except (UnicodeEncodeError, UnicodeDecodeError):
                    # If write fails, try ASCII-only
                    try:
                        safe_msg = msg.encode("ascii", "replace").decode("ascii")
                        stream.write(safe_msg + self.terminator)
                        self.flush()
                    except:
                        pass  # Silent fail
            except Exception:
                # Silent fail to prevent EXE crash
                pass
    
    ch = SafeStreamHandler(sys.stdout)
    # Show INFO, WARNING, and ERROR logs in console. Full detail still goes to file.
    ch.setLevel(logging.INFO)
    console_formatter = logging.Formatter('%(levelname)s - %(message)s')
    ch.setFormatter(console_formatter)

    if not logger.handlers:
        logger.addHandler(fh)
        logger.addHandler(ch)

    return logger


logger = setup_logger()
