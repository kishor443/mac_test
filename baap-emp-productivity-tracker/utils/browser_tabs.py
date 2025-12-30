from __future__ import annotations

import time
from datetime import datetime, timezone
from typing import Dict, List, Optional

from utils.logger import logger

try:
    import uiautomation as auto  # type: ignore
except Exception:  # pragma: no cover
    auto = None


def _detect_browser(window) -> Optional[str]:
    try:
        name = (window.Name or "").lower()
        cls = (window.ClassName or "").lower()
    except Exception:
        return None

    if "edge" in name:
        return "edge"
    if "brave" in name:
        return "brave"
    if "chrome" in name or cls.startswith("chrome_widgetwin"):
        return "chrome"
    if "firefox" in name or "mozilla" in cls:
        return "firefox"
    return None


def _iter_descendants(control, max_depth: int, deadline: float):
    if max_depth < 0 or time.perf_counter() > deadline or control is None:
        return
    try:
        child = control.GetFirstChildControl()
    except Exception:
        return
    while child:
        yield child
        yield from _iter_descendants(child, max_depth - 1, deadline)
        try:
            child = child.GetNextSiblingControl()
        except Exception:
            break


def _tab_items_for_window(window, max_depth: int, deadline: float):
    for ctrl in _iter_descendants(window, max_depth, deadline):
        try:
            if ctrl.ControlTypeName == "TabItemControl":
                yield ctrl
        except Exception:
            continue


def _read_tab_url(tab) -> str:
    if tab is None:
        return ""
    try:
        legacy = tab.GetLegacyIAccessiblePattern()
        if legacy:
            value = legacy.Value or ""
            if isinstance(value, str):
                return value
    except Exception:
        pass
    return ""


def _read_tab_title(tab) -> str:
    try:
        return tab.Name or ""
    except Exception:
        return ""


def collect_browser_tabs(user_id: Optional[str] = None, max_tabs: int = 50, timeout_seconds: float = 2.0) -> Dict[str, object]:
    snapshot = {
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "user_id": user_id or "",
        "total_tabs": 0,
        "tabs": [],
    }
    if auto is None:
        logger.debug("Browser tab capture skipped: uiautomation not available")
        return snapshot

    tabs: List[Dict[str, str]] = []
    deadline = time.perf_counter() + timeout_seconds
    try:
        root = auto.GetRootControl()
        windows = root.GetChildren()
    except Exception as exc:
        logger.error("Browser tab capture failed: %s", exc, exc_info=True)
        return snapshot

    for window in windows:
        if time.perf_counter() > deadline or len(tabs) >= max_tabs:
            break
        browser_name = _detect_browser(window)
        if not browser_name:
            continue

        try:
            tab_items = list(_tab_items_for_window(window, max_depth=4, deadline=deadline))
        except Exception:
            tab_items = []

        if not tab_items:
            # fallback: active window only via address bar
            url = ""
            for ctrl in _iter_descendants(window, max_depth=4, deadline=deadline):
                try:
                    if ctrl.ControlTypeName != "EditControl":
                        continue
                    name = (ctrl.Name or "").lower()
                    cls = (ctrl.ClassName or "").lower()
                    if "address" in name or "omnibox" in cls or "search" in name:
                        pattern = ctrl.GetValuePattern()
                        if pattern:
                            url = pattern.Value or ""
                            break
                except Exception:
                    continue
            tabs.append(
                {
                    "browser": browser_name,
                    "title": getattr(window, "Name", "") or "",
                    "url": url,
                }
            )
            continue

        for tab in tab_items:
            if len(tabs) >= max_tabs or time.perf_counter() > deadline:
                break
            title = _read_tab_title(tab)
            url = _read_tab_url(tab)
            tabs.append(
                {
                    "browser": browser_name,
                    "title": title,
                    "url": url,
                }
            )

    snapshot["tabs"] = tabs
    snapshot["total_tabs"] = len(tabs)
    return snapshot

