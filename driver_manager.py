"""
Driver Manager - Portable Windows Driver Export/Import Tool
Requires Python 3.8+ and PyInstaller for EXE build
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import subprocess
import json
import os
import threading
import ctypes
import sys
import re


# ─── Admin helpers ────────────────────────────────────────────────────────────

def is_admin() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def relaunch_as_admin():
    exe = sys.executable
    args = " ".join(f'"{a}"' for a in sys.argv)
    ctypes.windll.shell32.ShellExecuteW(None, "runas", exe, args, None, 1)
    sys.exit(0)


# ─── Driver data collectors ───────────────────────────────────────────────────

PS_TIMEOUT = 30  # seconds


def _run_ps(script: str) -> str:
    result = subprocess.run(
        ["powershell.exe", "-NonInteractive", "-NoProfile",
         "-ExecutionPolicy", "Bypass", "-Command", script],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        timeout=PS_TIMEOUT,
        creationflags=subprocess.CREATE_NO_WINDOW,
    )
    return result.stdout.strip()


def get_oem_drivers() -> list[dict]:
    """OEM (3rd-party) drivers via Win32_PnPSignedDriver, filter InfName = oem*.inf"""
    script = """
$drivers = Get-WmiObject Win32_PnPSignedDriver |
    Where-Object { $_.InfName -match '^oem\\d+\\.inf$' } |
    Select-Object DeviceName, DriverVersion, Manufacturer, InfName,
                  DriverDate, DeviceClass, Signer |
    Sort-Object DeviceName
ConvertTo-Json -InputObject @($drivers) -Depth 2
"""
    raw = _run_ps(script)
    try:
        rows = json.loads(raw) if raw else []
    except Exception:
        rows = []

    result = []
    for r in rows:
        if not isinstance(r, dict):
            continue
        # DriverDate is WMI date string like "20230615000000.000000+000"
        date_raw = r.get("DriverDate") or ""
        date_str = date_raw[:8]
        if len(date_str) == 8 and date_str.isdigit():
            date_str = f"{date_str[6:8]}.{date_str[4:6]}.{date_str[:4]}"
        else:
            date_str = ""

        result.append({
            "driver_type": "OEM",
            "name": (r.get("DeviceName") or "").strip() or r.get("InfName", ""),
            "inf": (r.get("InfName") or "").strip(),
            "version": (r.get("DriverVersion") or "").strip(),
            "date": date_str,
            "provider": (r.get("Manufacturer") or "").strip(),
            "class": (r.get("DeviceClass") or "").strip(),
            "signer": (r.get("Signer") or "").strip(),
        })
    return result


def get_system_drivers() -> list[dict]:
    """System (kernel) drivers via Win32_SystemDriver"""
    script = """
$drivers = Get-WmiObject Win32_SystemDriver |
    Select-Object DisplayName, Name, PathName, State, Started, ServiceType |
    Sort-Object DisplayName
ConvertTo-Json -InputObject @($drivers) -Depth 2
"""
    raw = _run_ps(script)
    try:
        rows = json.loads(raw) if raw else []
    except Exception:
        rows = []

    result = []
    for r in rows:
        if not isinstance(r, dict):
            continue
        path = (r.get("PathName") or "").strip()
        fname = os.path.basename(path) if path else ""
        result.append({
            "driver_type": "Системный",
            "name": (r.get("DisplayName") or r.get("Name") or "").strip(),
            "inf": (r.get("Name") or "").strip(),
            "version": "",
            "date": "",
            "provider": "",
            "class": (r.get("ServiceType") or "").strip(),
            "state": (r.get("State") or "").strip(),
            "path": path,
            "filename": fname,
        })
    return result


# ─── Main Application ─────────────────────────────────────────────────────────

class DriverManagerApp:

    COL_DEFS = [
        ("check",    "✓",              35,  35),
        ("type",     "Тип",            90,  70),
        ("name",     "Имя устройства", 240, 120),
        ("inf",      "INF / Модуль",   130, 90),
        ("version",  "Версия",         115, 80),
        ("date",     "Дата",           90,  70),
        ("provider", "Поставщик",      170, 90),
        ("class",    "Класс",          130, 70),
    ]

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Driver Manager  —  Управление драйверами")
        self.root.geometry("1100x680")
        self.root.minsize(820, 480)

        self._all_drivers: list[dict] = []
        self._selected: set[str] = set()   # set of iid strings

        self._apply_style()
        self._build_ui()
        self._load_drivers()

    # ── Style ──────────────────────────────────────────────────────────────────

    def _apply_style(self):
        s = ttk.Style()
        s.theme_use("clam")
        s.configure("TButton",   padding=(8, 4))
        s.configure("TLabel",    padding=(2, 1))
        s.configure("Accent.TButton", foreground="white", background="#0078d4")
        s.map("Accent.TButton",
              background=[("active", "#005fa3"), ("disabled", "#aac8e4")])
        s.configure("Warn.TButton",   foreground="white", background="#c84b1a")
        s.map("Warn.TButton",
              background=[("active", "#a03010"), ("disabled", "#e0a080")])

    # ── UI ─────────────────────────────────────────────────────────────────────

    def _build_ui(self):
        # ── Top toolbar ──────────────────────────────────────────────────────
        toolbar = ttk.Frame(self.root, padding=(6, 4))
        toolbar.pack(fill=tk.X)

        ttk.Button(toolbar, text="↻  Обновить",
                   command=self._load_drivers).pack(side=tk.LEFT, padx=3)

        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(
            side=tk.LEFT, fill=tk.Y, padx=6, pady=2)

        ttk.Button(toolbar, text="📤  Экспорт выбранных",
                   style="Accent.TButton",
                   command=self._export_drivers).pack(side=tk.LEFT, padx=3)

        ttk.Button(toolbar, text="📥  Импорт драйверов",
                   style="Warn.TButton",
                   command=self._import_drivers).pack(side=tk.LEFT, padx=3)

        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(
            side=tk.LEFT, fill=tk.Y, padx=6, pady=2)

        ttk.Button(toolbar, text="Выбрать все OEM",
                   command=lambda: self._select_by_type("OEM", True)
                   ).pack(side=tk.LEFT, padx=3)
        ttk.Button(toolbar, text="Снять всё",
                   command=self._deselect_all).pack(side=tk.LEFT, padx=3)

        # Admin badge
        self._admin_lbl = ttk.Label(
            toolbar,
            text="⚠ Не администратор — установка недоступна" if not is_admin() else "✔ Администратор",
            foreground="#c84b1a" if not is_admin() else "#1a7a1a",
        )
        self._admin_lbl.pack(side=tk.RIGHT, padx=10)

        # ── Filter bar ───────────────────────────────────────────────────────
        fbar = ttk.LabelFrame(self.root, text="Фильтр / Поиск", padding=(8, 4))
        fbar.pack(fill=tk.X, padx=6, pady=(0, 4))

        self._show_oem = tk.BooleanVar(value=True)
        self._show_sys = tk.BooleanVar(value=True)

        ttk.Checkbutton(fbar, text="OEM (сторонние)",
                        variable=self._show_oem,
                        command=self._filter).pack(side=tk.LEFT, padx=8)
        ttk.Checkbutton(fbar, text="Системные",
                        variable=self._show_sys,
                        command=self._filter).pack(side=tk.LEFT, padx=8)

        ttk.Separator(fbar, orient=tk.VERTICAL).pack(
            side=tk.LEFT, fill=tk.Y, padx=8, pady=2)

        ttk.Label(fbar, text="🔍").pack(side=tk.LEFT)
        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", lambda *_: self._filter())
        ttk.Entry(fbar, textvariable=self._search_var, width=36).pack(
            side=tk.LEFT, padx=4)

        self._count_lbl = ttk.Label(fbar, text="")
        self._count_lbl.pack(side=tk.RIGHT, padx=8)

        # ── Treeview ─────────────────────────────────────────────────────────
        tree_frame = ttk.Frame(self.root)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=(0, 4))

        cols = [c[0] for c in self.COL_DEFS]
        self._tree = ttk.Treeview(
            tree_frame, columns=cols, show="headings", selectmode="none")

        for cid, heading, width, minw in self.COL_DEFS:
            self._tree.heading(cid, text=heading,
                               command=lambda c=cid: self._sort_by(c))
            self._tree.column(cid, width=width, minwidth=minw,
                              stretch=(cid == "name"))

        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL,
                            command=self._tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL,
                            command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self._tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        # Row colours
        self._tree.tag_configure("oem",          background="#eef5ff")
        self._tree.tag_configure("sys",          background="#f5f5f5")
        self._tree.tag_configure("oem_sel",      background="#90c8f8", foreground="#000")
        self._tree.tag_configure("sys_sel",      background="#a8e0a8", foreground="#000")

        self._tree.bind("<Button-1>", self._on_click)

        # ── Status bar ───────────────────────────────────────────────────────
        self._status_var = tk.StringVar(value="Готов")
        ttk.Label(self.root, textvariable=self._status_var,
                  relief=tk.SUNKEN, anchor=tk.W, padding=(6, 2)
                  ).pack(fill=tk.X, side=tk.BOTTOM)

    # ── Loading ────────────────────────────────────────────────────────────────

    def _load_drivers(self):
        self._set_status("Загрузка списка драйверов…")
        self._tree.delete(*self._tree.get_children())
        self._selected.clear()
        threading.Thread(target=self._load_thread, daemon=True).start()

    def _load_thread(self):
        oem = get_oem_drivers()
        sys_d = get_system_drivers()
        self._all_drivers = oem + sys_d
        self.root.after(0, self._populate)

    def _populate(self):
        self._tree.delete(*self._tree.get_children())
        self._selected.clear()

        query = self._search_var.get().lower()
        show_oem = self._show_oem.get()
        show_sys = self._show_sys.get()

        shown = 0
        for d in self._all_drivers:
            dtype = d["driver_type"]
            if dtype == "OEM" and not show_oem:
                continue
            if dtype == "Системный" and not show_sys:
                continue

            searchable = " ".join([
                d.get("name", ""), d.get("inf", ""),
                d.get("version", ""), d.get("provider", ""),
                d.get("class", ""),
            ]).lower()
            if query and query not in searchable:
                continue

            tag = "oem" if dtype == "OEM" else "sys"
            iid = self._tree.insert("", "end", values=(
                "☐", dtype,
                d.get("name", ""),
                d.get("inf", ""),
                d.get("version", ""),
                d.get("date", ""),
                d.get("provider", ""),
                d.get("class", ""),
            ), tags=(tag,))
            # store driver dict reference via iid
            self._tree.set(iid, "check", "☐")
            shown += 1

        total_oem = sum(1 for d in self._all_drivers if d["driver_type"] == "OEM")
        total_sys = sum(1 for d in self._all_drivers if d["driver_type"] == "Системный")
        self._count_lbl.config(
            text=f"OEM: {total_oem}  |  Системных: {total_sys}  |  Показано: {shown}")
        self._set_status(f"Загружено: {len(self._all_drivers)} драйверов  (OEM: {total_oem}, Системных: {total_sys})")

    def _filter(self):
        self._populate()

    # ── Selection ──────────────────────────────────────────────────────────────

    def _on_click(self, event: tk.Event):
        col = self._tree.identify_column(event.x)
        iid = self._tree.identify_row(event.y)
        if iid and col == "#1":
            self._toggle(iid)

    def _toggle(self, iid: str):
        tags = list(self._tree.item(iid, "tags"))
        base_tag = next((t for t in tags if t in ("oem", "sys")), "sys")

        if iid in self._selected:
            self._selected.discard(iid)
            self._tree.item(iid, tags=(base_tag,))
            self._tree.set(iid, "check", "☐")
        else:
            self._selected.add(iid)
            sel_tag = "oem_sel" if base_tag == "oem" else "sys_sel"
            self._tree.item(iid, tags=(sel_tag,))
            self._tree.set(iid, "check", "☑")

        n = len(self._selected)
        self._set_status(f"Выбрано: {n} драйвер(ов)")

    def _select_by_type(self, dtype: str, select: bool):
        for iid in self._tree.get_children():
            vals = self._tree.item(iid, "values")
            if vals[1] == dtype:
                if select and iid not in self._selected:
                    self._toggle(iid)
                elif not select and iid in self._selected:
                    self._toggle(iid)

    def _deselect_all(self):
        for iid in list(self._selected):
            self._toggle(iid)

    # ── Sorting ────────────────────────────────────────────────────────────────

    _sort_state: dict = {}

    def _sort_by(self, col: str):
        col_idx = [c[0] for c in self.COL_DEFS].index(col)
        reverse = self._sort_state.get(col, False)
        data = [(self._tree.set(iid, col), iid)
                for iid in self._tree.get_children()]
        data.sort(key=lambda x: x[0].lower(), reverse=reverse)
        for idx, (_, iid) in enumerate(data):
            self._tree.move(iid, "", idx)
        self._sort_state[col] = not reverse

    # ── Export ─────────────────────────────────────────────────────────────────

    def _export_drivers(self):
        sel = list(self._selected)
        if not sel:
            messagebox.showwarning(
                "Ничего не выбрано",
                "Отметьте OEM-драйверы галочкой (☐) в первой колонке.")
            return

        oem_sel = [iid for iid in sel
                   if self._tree.item(iid, "values")[1] == "OEM"]
        sys_sel = [iid for iid in sel
                   if self._tree.item(iid, "values")[1] != "OEM"]

        msgs = []
        if sys_sel:
            msgs.append(f"Системных драйверов (не экспортируются): {len(sys_sel)}")
        if not oem_sel:
            messagebox.showwarning(
                "Только OEM",
                "Экспорт возможен только для OEM-драйверов.\n" +
                "\n".join(msgs))
            return

        folder = filedialog.askdirectory(title="Папка для экспорта драйверов")
        if not folder:
            return

        self._set_status("Экспорт…")
        threading.Thread(
            target=self._export_thread,
            args=(oem_sel, folder),
            daemon=True,
        ).start()

    def _export_thread(self, iids: list, folder: str):
        ok, fail = 0, 0
        errors: list[str] = []

        for iid in iids:
            vals = self._tree.item(iid, "values")
            inf = vals[3]   # INF column
            name = vals[2]  # Name column

            if not inf:
                fail += 1
                errors.append(f"(нет INF): {name}")
                continue

            dest = os.path.join(folder, re.sub(r"\.inf$", "", inf, flags=re.I))
            os.makedirs(dest, exist_ok=True)

            try:
                r = subprocess.run(
                    ["pnputil.exe", "/export-driver", inf, dest],
                    capture_output=True, text=True, encoding="cp866",
                    errors="replace",
                    creationflags=subprocess.CREATE_NO_WINDOW,
                    timeout=60,
                )
                if r.returncode == 0:
                    ok += 1
                else:
                    fail += 1
                    detail = (r.stderr or r.stdout or "").strip().splitlines()
                    errors.append(f"{inf}: {detail[0] if detail else 'ошибка'}")
            except Exception as exc:
                fail += 1
                errors.append(f"{inf}: {exc}")

        msg = f"Экспортировано: {ok}\nОшибок: {fail}"
        if errors:
            msg += "\n\nПодробности:\n" + "\n".join(errors[:15])

        self.root.after(0, lambda: self._set_status(
            f"Экспорт завершён: {ok} успешно, {fail} ошибок"))
        self.root.after(0, lambda: messagebox.showinfo("Экспорт завершён", msg))

    # ── Import ─────────────────────────────────────────────────────────────────

    def _import_drivers(self):
        if not is_admin():
            ans = messagebox.askyesno(
                "Нужны права администратора",
                "Установка драйверов требует прав администратора.\n\n"
                "Перезапустить программу от имени администратора?")
            if ans:
                relaunch_as_admin()
            return

        folder = filedialog.askdirectory(title="Папка с драйверами (.inf файлы)")
        if not folder:
            return

        inf_files: list[str] = []
        for root_dir, _, files in os.walk(folder):
            for f in files:
                if f.lower().endswith(".inf"):
                    inf_files.append(os.path.join(root_dir, f))

        if not inf_files:
            messagebox.showwarning("Не найдено",
                                   f"INF-файлы не найдены в:\n{folder}")
            return

        if not messagebox.askyesno(
            "Подтверждение установки",
            f"Найдено INF-файлов: {len(inf_files)}\n\nУстановить все драйверы?",
        ):
            return

        self._set_status("Установка драйверов…")
        threading.Thread(
            target=self._import_thread, args=(inf_files,), daemon=True).start()

    def _import_thread(self, inf_files: list[str]):
        ok, fail = 0, 0
        errors: list[str] = []

        for path in inf_files:
            try:
                r = subprocess.run(
                    ["pnputil.exe", "/add-driver", path, "/install"],
                    capture_output=True, text=True, encoding="cp866",
                    errors="replace",
                    creationflags=subprocess.CREATE_NO_WINDOW,
                    timeout=120,
                )
                if r.returncode == 0:
                    ok += 1
                else:
                    fail += 1
                    detail = (r.stderr or r.stdout or "").strip().splitlines()
                    errors.append(f"{os.path.basename(path)}: {detail[0] if detail else 'ошибка'}")
            except Exception as exc:
                fail += 1
                errors.append(f"{os.path.basename(path)}: {exc}")

        msg = f"Установлено: {ok}\nОшибок: {fail}"
        if errors:
            msg += "\n\nПодробности:\n" + "\n".join(errors[:15])

        self.root.after(0, lambda: self._set_status(
            f"Установка завершена: {ok} успешно, {fail} ошибок"))
        self.root.after(0, lambda: messagebox.showinfo("Установка завершена", msg))
        self.root.after(0, self._load_drivers)

    # ── Helpers ────────────────────────────────────────────────────────────────

    def _set_status(self, text: str):
        self._status_var.set(text)


# ─── Entry point ──────────────────────────────────────────────────────────────

def main():
    if not is_admin():
        # Offer UAC elevation; if user declines, app still works (read-only mode)
        try:
            ans = ctypes.windll.user32.MessageBoxW(
                0,
                "Для установки драйверов нужны права администратора.\n\n"
                "Запустить с правами администратора?",
                "Driver Manager",
                0x34,   # MB_ICONQUESTION | MB_YESNO | MB_TOPMOST
            )
            if ans == 6:  # IDYES
                relaunch_as_admin()
        except Exception:
            pass

    root = tk.Tk()
    app = DriverManagerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
