from __future__ import annotations

from datetime import datetime
from io import BytesIO

try:
    from PIL import ImageGrab  # type: ignore
except Exception:  # pragma: no cover
    ImageGrab = None

from utils.capture_types import CaptureArtifact


def capture_screenshot(base_dir: str | None = None) -> CaptureArtifact | None:
    """
    Capture a screenshot and return it as an in-memory artifact.
    base_dir is kept for backward compatibility but no longer used for storage.
    """
    if ImageGrab is None:
        return None
    try:
        img = ImageGrab.grab()
        buffer = BytesIO()
        img.save(buffer, format="PNG")
        filename = datetime.now().strftime("%Y%m%d-%H%M%S") + ".png"
        return CaptureArtifact(filename=filename, data=buffer.getvalue(), mimetype="image/png")
    except Exception:
        return None
