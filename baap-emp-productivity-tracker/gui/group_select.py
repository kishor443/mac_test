import tkinter as tk
from tkinter import ttk, messagebox
from gui.theme import apply_theme, spacing


class GroupSelectDialog:
    def __init__(self, clients_payload):
        self.clients_payload = clients_payload or {}
        self.result_client_id = None
        self.window = None

    def _flatten_clients(self):
        items = []
        if isinstance(self.clients_payload, dict):
            possible_lists = [
                self.clients_payload.get("items"),
                self.clients_payload.get("data"),
                self.clients_payload.get("clients"),
            ]
            for lst in possible_lists:
                if isinstance(lst, list) and lst:
                    items = lst
                    break
            if not items:
                items = [self.clients_payload]
        elif isinstance(self.clients_payload, list):
            items = self.clients_payload
        return items

    def show(self):
        self.window = tk.Tk()
        self.window.title("Select Group / Client")
        self.window.geometry("420x360")
        try:
            self.window.minsize(420, 360)
        except Exception:
            pass
        apply_theme(self.window)

        container = ttk.Frame(self.window, padding=(spacing["xl"], spacing["lg"]))
        container.pack(fill=tk.BOTH, expand=True)

        ttk.Label(container, text="Choose your ERP Group / Client").pack(anchor=tk.W)

        list_frame = ttk.Frame(container, style="Card.TFrame", padding=(spacing["md"], spacing["sm"]))
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(spacing["sm"], spacing["md"]))

        self.tree = ttk.Treeview(list_frame, columns=("name", "id"), show="headings")
        self.tree.heading("name", text="Name")
        self.tree.heading("id", text="ID")
        self.tree.column("name", width=320, anchor=tk.W)
        self.tree.column("id", width=0, stretch=False)
        self.tree["displaycolumns"] = ("name",)
        self.tree.pack(fill=tk.BOTH, expand=True)

        for item in self._flatten_clients():
            if not isinstance(item, dict):
                continue
            # Prefer organization name from orgn_details -> orgn_name
            orgn_name = None
            orgn_details = item.get("orgn_details")
            if isinstance(orgn_details, list) and orgn_details and isinstance(orgn_details[0], dict):
                orgn_name = orgn_details[0].get("orgn_name")
            # Also consider primary_info.short_name
            short_name = None
            primary_info = item.get("primary_info")
            if isinstance(primary_info, list) and primary_info and isinstance(primary_info[0], dict):
                short_name = primary_info[0].get("short_name")
            name = (
                orgn_name
                or item.get("name")
                or item.get("client_name")
                or item.get("group_name")
                or short_name
                or "Unnamed"
            )
            cid = item.get("id") or item.get("client_id") or item.get("group_id")
            if cid:
                self.tree.insert("", tk.END, values=(name, cid))

        actions = ttk.Frame(container)
        actions.pack(fill=tk.X)
        ttk.Button(actions, text="Select", style="Primary.TButton", command=self._select).pack(side=tk.RIGHT)

        self.window.mainloop()
        return self.result_client_id

    def _select(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Select Group", "Please select a group/client to continue.")
            return
        values = self.tree.item(sel[0], "values")
        if len(values) >= 2:
            self.result_client_id = values[1]
        self.window.destroy()


