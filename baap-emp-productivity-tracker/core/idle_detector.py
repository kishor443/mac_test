import threading
import time

class IdleDetector:
    def __init__(self, activity_tracker, on_idle=None, on_active=None, idle_timeout=300):
        self.activity_tracker = activity_tracker
        self.on_idle = on_idle
        self.on_active = on_active
        self.idle_timeout = idle_timeout
        self.is_idle = False

    def start(self):
        threading.Thread(target=self._monitor, daemon=True).start()

    def _monitor(self):
        while True:
            if not self.is_idle and self.activity_tracker.get_idle_time() > self.idle_timeout:
                self.is_idle = True
                if self.on_idle:
                    self.on_idle()
            elif self.is_idle and self.activity_tracker.get_idle_time() < self.idle_timeout / 2:
                self.is_idle = False
                if self.on_active:
                    self.on_active()
            time.sleep(10)
