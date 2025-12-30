import tkinter as tk
from tkinter import ttk, messagebox
from gui.theme import apply_theme, spacing


class ShiftSelectDialog:
    def __init__(self, shifts_payload):
        self.shifts_payload = shifts_payload or {}
        self.result_shift_id = None
        self.window = None

    def _flatten_shifts(self):
        items = []
        data = self.shifts_payload
        if isinstance(data, dict):
            possible_lists = [
                data.get("items"),
                data.get("data"),
                data.get("shifts"),
                data.get("rows"),
                data.get("result"),
            ]
            for lst in possible_lists:
                if isinstance(lst, list) and lst:
                    items = lst
                    break
            if not items:
                items = [data]
        elif isinstance(data, list):
            items = data
        return items

    @staticmethod
    def _extract_name_and_id(item: dict):
        # Try multiple common fields for name
        name_fields = [
            "shift_type",  # preferred for shifts payload
            "name", "shift_name", "shiftName", "title", "shift_title", "shiftTitle",
            "display_name", "displayName", "label", "shift_type_name", "shiftTypeName",
            "shift",  # could be nested dict with its own name
        ]
        name = None
        for key in name_fields:
            val = item.get(key)
            if isinstance(val, str) and val.strip():
                name = val.strip()
                break
            if isinstance(val, dict):
                # look for nested name keys
                for subk in ("name", "shift_name", "title", "display_name", "label", "shift_type"):
                    subval = val.get(subk)
                    if isinstance(subval, str) and subval.strip():
                        name = subval.strip()
                        break
            if name:
                break
        # Try location/department nested names as helpful context
        if not name and isinstance(item.get("location_data"), dict):
            loc_name = item.get("location_data", {}).get("name")
            if isinstance(loc_name, str) and loc_name.strip():
                name = loc_name.strip()
        # Last-resort heuristic: pick any reasonable string field that looks like a name
        if not name:
            id_like_keys = {"id", "shift_id", "uuid", "_id", "shiftId"}
            candidates = []
            for k, v in item.items():
                if isinstance(v, str):
                    vs = v.strip()
                    if vs and k not in id_like_keys and len(vs) <= 100:
                        weight = 2 if ("name" in k.lower() or "title" in k.lower()) else 1
                        candidates.append((weight, k, vs))
            if candidates:
                candidates.sort(reverse=True)
                name = candidates[0][2]
        # Fallback: build from time range if available
        if not name:
            start = item.get("start_time") or item.get("startTime") or item.get("in_time")
            end = item.get("end_time") or item.get("endTime") or item.get("out_time")
            if start or end:
                name = f"{(start or '?')} - {(end or '?')}"
        if not name:
            name = "Unnamed"

        # Try multiple id fields
        id_fields = ["id", "shift_id", "uuid", "_id", "shiftId"]
        sid = None
        for key in id_fields:
            val = item.get(key)
            if isinstance(val, (str, int)) and str(val).strip():
                sid = str(val).strip()
                break
            if isinstance(val, dict):
                # nested id
                for subk in ("id", "uuid", "_id"):
                    subval = val.get(subk)
                    if isinstance(subval, (str, int)) and str(subval).strip():
                        sid = str(subval).strip()
                        break
            if sid:
                break
        return name, sid

    def show(self):
        self.window = tk.Tk()
        self.window.title("Select Shift")
        self.window.geometry("420x360")
        try:
            self.window.minsize(420, 360)
        except Exception:
            pass
        apply_theme(self.window)

        container = ttk.Frame(self.window, padding=(spacing["xl"], spacing["lg"]))
        container.pack(fill=tk.BOTH, expand=True)

        ttk.Label(container, text="Choose your ERP Shift").pack(anchor=tk.W)

        list_frame = ttk.Frame(container, style="Card.TFrame", padding=(spacing["md"], spacing["sm"]))
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(spacing["sm"], spacing["md"]))

        self.tree = ttk.Treeview(list_frame, columns=("name", "id"), show="headings")
        self.tree.heading("name", text="Name")
        self.tree.heading("id", text="ID")
        self.tree.column("name", width=240)
        self.tree.column("id", width=140)
        self.tree.pack(fill=tk.BOTH, expand=True)

        for item in self._flatten_shifts():
            if not isinstance(item, dict):
                continue
            name, sid = self._extract_name_and_id(item)
            if sid:
                self.tree.insert("", tk.END, values=(name, sid))

        actions = ttk.Frame(container)
        actions.pack(fill=tk.X)
        ttk.Button(actions, text="Select", style="Primary.TButton", command=self._select).pack(side=tk.RIGHT)

        self.window.mainloop()
        return self.result_shift_id

    def _select(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Select Shift", "Please select a shift to continue.")
            return
        values = self.tree.item(sel[0], "values")
        if len(values) >= 2:
            self.result_shift_id = values[1]
        self.window.destroy()


