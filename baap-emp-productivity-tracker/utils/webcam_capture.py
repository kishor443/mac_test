from __future__ import annotations

from datetime import datetime
from typing import Optional

try:
    import cv2  # type: ignore
except Exception:  # pragma: no cover
    cv2 = None

from utils.capture_types import CaptureArtifact
from utils.logger import logger


def capture_webcam_photo(base_dir: str | None, timestamp: Optional[datetime] = None) -> Optional[CaptureArtifact]:
    """
    Capture a single frame from the default webcam without writing to disk.
    Returns an in-memory artifact or None if capture failed/unavailable.
    """
    if cv2 is None:
        logger.warning("Webcam capture skipped: opencv-python not available")
        return None

    ts = timestamp or datetime.now()

    cap = None
    try:
        cap = cv2.VideoCapture(0, cv2.CAP_DSHOW)
        if not cap.isOpened():
            logger.warning("Webcam capture skipped: device not available")
            return None
        cap.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
        cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 360)
        ret, frame = cap.read()
        if not ret or frame is None:
            logger.warning("Webcam capture skipped: read() returned no frame")
            return None

        # Overlay timestamp for quick reference
        overlay = ts.strftime("%Y-%m-%d %H:%M:%S")
        try:
            cv2.putText(
                frame,
                overlay,
                (10, 30),
                cv2.FONT_HERSHEY_SIMPLEX,
                0.7,
                (0, 0, 0),
                3,
                cv2.LINE_AA,
            )
            cv2.putText(
                frame,
                overlay,
                (10, 30),
                cv2.FONT_HERSHEY_SIMPLEX,
                0.7,
                (255, 255, 255),
                1,
                cv2.LINE_AA,
            )
        except Exception:
            pass  # timestamp overlay failure isn't fatal

        success, encoded = cv2.imencode(".jpg", frame)
        if not success:
            logger.warning("Webcam capture skipped: encoding failed")
            return None
        filename = ts.strftime("%Y%m%d-%H%M%S") + ".jpg"
        return CaptureArtifact(filename=filename, data=encoded.tobytes(), mimetype="image/jpeg")
    except Exception as exc:  # pragma: no cover
        logger.error("Webcam capture failed: %s", exc, exc_info=True)
        return None
    finally:
        if cap is not None:
            try:
                cap.release()
            except Exception:
                pass