from __future__ import annotations

from datetime import datetime, timedelta, date
from pathlib import Path
from typing import Iterable
import threading
import shutil

from config import (
    DATA_RETENTION_DAYS,
    SCREENSHOTS_DIR,
    WEBCAM_PHOTOS_DIR,
)
from utils.excel_storage import (
    purge_activity_before,
    read_local_storage,
    write_local_storage,
)


def _cleanup_media_dir(base_path: str, cutoff_date: date) -> int:
    removed = 0
    path = Path(base_path)
    if not path.exists():
        return removed
    for child in path.iterdir():
        if not child.is_dir():
            continue
        try:
            child_date = date.fromisoformat(child.name)
        except ValueError:
            continue
        if child_date < cutoff_date:
            try:
                shutil.rmtree(child, ignore_errors=True)
                removed += 1
            except Exception:
                continue
    return removed


def _purge_local_storage(cutoff_date: date) -> int:
    data = read_local_storage()
    changed = False

    def _filter_date_list(items: Iterable[dict], key: str) -> list[dict]:
        filtered = []
        for item in items or []:
            if not isinstance(item, dict):
                continue
            value = item.get(key)
            try:
                item_date = date.fromisoformat(value)
            except Exception:
                continue
            if item_date >= cutoff_date:
                filtered.append(item)
        return filtered

    history = data.get("history", [])
    filtered_history = []
    for record in history:
        if not isinstance(record, dict):
            continue
        try:
            record_date = date.fromisoformat(record.get("date", ""))
        except Exception:
            continue
        if record_date >= cutoff_date:
            filtered_history.append(record)
    if len(filtered_history) != len(history):
        data["history"] = filtered_history
        changed = True

    sleep_events = data.get("sleep_events", [])
    filtered_sleep = _filter_date_list(sleep_events, "date")
    if len(filtered_sleep) != len(sleep_events):
        data["sleep_events"] = filtered_sleep
        changed = True

    if changed:
        write_local_storage(data)
    return (len(history) - len(filtered_history)) + (len(sleep_events) - len(filtered_sleep))


def enforce_data_retention(retention_days: int | None = None) -> None:
    days = retention_days or DATA_RETENTION_DAYS
    cutoff_dt = datetime.now() - timedelta(days=days)
    cutoff_date = cutoff_dt.date()

    try:
        purge_activity_before(cutoff_dt)
    except Exception:
        pass

    try:
        _purge_local_storage(cutoff_date)
    except Exception:
        pass

    for media_dir in (SCREENSHOTS_DIR, WEBCAM_PHOTOS_DIR):
        try:
            _cleanup_media_dir(media_dir, cutoff_date)
        except Exception:
            pass


def enforce_data_retention_async(retention_days: int | None = None) -> None:
    thread = threading.Thread(
        target=enforce_data_retention,
        kwargs={"retention_days": retention_days},
        daemon=True,
    )
    thread.start()

