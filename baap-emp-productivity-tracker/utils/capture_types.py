from __future__ import annotations

from dataclasses import dataclass


@dataclass(slots=True)
class CaptureArtifact:
    """
    Represents an in-memory capture (screenshot or webcam photo).

    Attributes:
        filename: Suggested filename (for uploads/logging).
        data: Raw binary content of the image.
        mimetype: MIME type hint (e.g., image/png, image/jpeg).
    """

    filename: str
    data: bytes
    mimetype: str

    def has_payload(self) -> bool:
        return bool(self.data)

