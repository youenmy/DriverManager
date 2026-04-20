"""
Driver Manager — Portable Windows Driver Export/Import Tool
Requires Python 3.10+ and PyInstaller for EXE build.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import subprocess
import json
import os
import sys
import threading
import ctypes
import re

# ─── Admin helpers ─────────────────────────────────────────────────────────

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


# ─── Subprocess wrapper (safe in windowed PyInstaller) ─────────────────────

_COMMON_KW = dict(
    stdin=subprocess.DEVNULL,
    capture_output=True,
    text=True,
    encoding="utf-8",
    errors="replace",
    creationflags=subprocess.CREATE_NO_WINDOW,
)


def _run_ps(script: str, timeout: int = 60) -> str:
    try:
        r = subprocess.run(
            ["powershell.exe", "-NonInteractive", "-NoProfile",
             "-ExecutionPolicy", "Bypass",
             "-OutputFormat", "Text",
             "-Command", script],
            timeout=timeout,
            **_COMMON_KW,
        )
        out = r.stdout or ""
    except Exception as exc:
        return ""
    # Strip possible BOM
    return out.lstrip("\ufeff").strip()


def _run_cmd(args: list[str], encoding: str = "cp866", timeout: int = 60):
    kw = dict(_COMMON_KW)
    kw["encoding"] = encoding
    return subprocess.run(args, timeout=timeout, **kw)


# ─── Driver data collectors ────────────────────────────────────────────────

def _driver_store_paths() -> dict:
    """Map InfName → driver-store folder via `pnputil /enum-drivers`."""
    try:
        r = _run_cmd(["pnputil.exe", "/enum-drivers"])
    except Exception:
        return {}
    if r.returncode != 0:
        return {}

    keys_pub = ("published name", "опубликованное имя")
    keys_store = ("driver store path", "путь к хранилищу драйверов",
                  "путь хранилища драйверов")
    result, inf = {}, None
    for line in (r.stdout or "").splitlines():
        s = line.strip()
        if not s:
            inf = None
            continue
        if ":" not in s:
            continue
        key, _, val = s.partition(":")
        key = key.strip().lower()
        val = val.strip()
        if key in keys_pub:
            inf = val
        elif inf and key in keys_store:
            result[inf] = val
    return result


def get_oem_drivers() -> list[dict]:
    """OEM (3rd-party) drivers via Win32_PnPSignedDriver (InfName = oem*.inf)."""
    script = r"""
try {
    $drivers = Get-CimInstance Win32_PnPSignedDriver -ErrorAction Stop |
        Where-Object { $_.InfName -match '^oem\d+\.inf$' } |
        Select-Object DeviceName, DriverVersion, Manufacturer, InfName,
                      DriverDate, DeviceClass, Signer, Location |
        Sort-Object DeviceName
    $json = ConvertTo-Json -InputObject @($drivers) -Depth 2 -Compress
    Write-Output $json
} catch {
    Write-Output "[]"
}
"""
    raw = _run_ps(script) or "[]"
    try:
        rows = json.loads(raw)
        if not isinstance(rows, list):
            rows = [rows]
    except Exception:
        rows = []

    store = _driver_store_paths()

    out = []
    for r in rows:
        if not isinstance(r, dict):
            continue
        inf = (r.get("InfName") or "").strip()
        date_raw = r.get("DriverDate") or ""
        if isinstance(date_raw, dict):   # CIM may return @{DateTime=...}
            date_raw = date_raw.get("DateTime", "") or ""
        m = re.search(r"(\d{4})(\d{2})(\d{2})", str(date_raw))
        date_str = f"{m.group(3)}.{m.group(2)}.{m.group(1)}" if m else ""
        path = store.get(inf) or (f"C:\\Windows\\INF\\{inf}" if inf else "")
        out.append({
            "driver_type": "OEM",
            "name": (r.get("DeviceName") or "").strip() or inf,
            "inf": inf,
            "version": (r.get("DriverVersion") or "").strip(),
            "date": date_str,
            "provider": (r.get("Manufacturer") or "").strip(),
            "class": (r.get("DeviceClass") or "").strip(),
            "path": path,
        })
    return out


def get_printer_drivers() -> list[dict]:
    """Printer drivers via Get-PrinterDriver (catches Kyocera, HP, Canon etc)."""
    script = r"""
try {
    $drivers = Get-PrinterDriver -ErrorAction Stop |
        Select-Object Name, Manufacturer, DriverVersion, InfPath,
                      MajorVersion, PrinterEnvironment
    $json = ConvertTo-Json -InputObject @($drivers) -Depth 2 -Compress
    if ([string]::IsNullOrEmpty($json)) { Write-Output "[]" } else { Write-Output $json }
} catch {
    Write-Output "[]"
}
"""
    raw = _run_ps(script) or "[]"
    try:
        rows = json.loads(raw)
        if not isinstance(rows, list):
            rows = [rows]
    except Exception:
        rows = []

    out = []
    seen = set()
    for r in rows:
        if not isinstance(r, dict):
            continue
        inf_path = (r.get("InfPath") or "").strip()
        inf_name = os.path.basename(inf_path) if inf_path else ""

        # Driver version is a packed 64-bit int in WMI
        version = r.get("DriverVersion")
        if isinstance(version, (int, float)) and version:
            v = int(version)
            version = (f"{(v >> 48) & 0xFFFF}.{(v >> 32) & 0xFFFF}."
                       f"{(v >> 16) & 0xFFFF}.{v & 0xFFFF}")
        elif version is None:
            version = ""

        env = (r.get("PrinterEnvironment") or "").strip()
        name = (r.get("Name") or "").strip()
        # dedupe by (name, inf)
        key = (name.lower(), inf_name.lower())
        if key in seen:
            continue
        seen.add(key)

        out.append({
            "driver_type": "Принтер",
            "name": name,
            "inf": inf_name,
            "version": str(version).strip(),
            "date": "",
            "provider": (r.get("Manufacturer") or "").strip(),
            "class": f"Printer ({env})" if env else "Printer",
            "path": inf_path,
        })
    return out


def get_system_drivers() -> list[dict]:
    """System (kernel) drivers via Win32_SystemDriver."""
    script = r"""
try {
    $drivers = Get-CimInstance Win32_SystemDriver -ErrorAction Stop |
        Select-Object DisplayName, Name, PathName, State, Started, ServiceType |
        Sort-Object DisplayName
    $json = ConvertTo-Json -InputObject @($drivers) -Depth 2 -Compress
    Write-Output $json
} catch {
    Write-Output "[]"
}
"""
    raw = _run_ps(script) or "[]"
    try:
        rows = json.loads(raw)
        if not isinstance(rows, list):
            rows = [rows]
    except Exception:
        rows = []

    out = []
    for r in rows:
        if not isinstance(r, dict):
            continue
        path = (r.get("PathName") or "").strip().replace("\\??\\", "")
        out.append({
            "driver_type": "Системный",
            "name": (r.get("DisplayName") or r.get("Name") or "").strip(),
            "inf": (r.get("Name") or "").strip(),
            "version": "",
            "date": "",
            "provider": "",
            "class": (r.get("ServiceType") or "").strip(),
            "path": path,
        })
    return out


# ─── Icon embedded as base64 ───────────────────────────────────────────────

def _icon_path_for_tk() -> str | None:
    """Return path to icon.ico if bundled by PyInstaller."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    for name in ("icon.ico", "icon.png"):
        p = os.path.join(base, name)
        if os.path.exists(p):
            return p
    return None


# ─── UI ────────────────────────────────────────────────────────────────────

BG      = "#f3f5f9"
CARD    = "#ffffff"
ACCENT  = "#2563eb"
ACCENT2 = "#1d4ed8"
DANGER  = "#dc2626"
DANGER2 = "#b91c1c"
PRINT   = "#7c3aed"
TEXT    = "#0f172a"
MUTED   = "#64748b"
ROW_OEM = "#eff6ff"
ROW_SYS = "#f8fafc"
ROW_PRN = "#faf5ff"
ROW_OEM_ALT = "#ffffff"
ROW_SYS_ALT = "#ffffff"
ROW_PRN_ALT = "#ffffff"


class DriverManagerApp:

    COL_DEFS = [
        ("type",     "Тип",            90,  70,  False),
        ("name",     "Имя устройства", 240, 140, True),
        ("inf",      "INF / Модуль",   130, 90,  False),
        ("version",  "Версия",         115, 80,  False),
        ("date",     "Дата",           90,  70,  False),
        ("provider", "Поставщик",      170, 90,  False),
        ("class",    "Класс",          110, 70,  False),
        ("path",     "Путь к файлам",  260, 140, True),
    ]

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Driver Manager — Управление драйверами Windows")
        self.root.geometry("1240x720")
        self.root.minsize(900, 500)
        self.root.configure(bg=BG)

        # Icon
        ico = _icon_path_for_tk()
        if ico:
            try:
                if ico.endswith(".ico"):
                    self.root.iconbitmap(ico)
                else:
                    self._icon_img = tk.PhotoImage(file=ico)
                    self.root.iconphoto(True, self._icon_img)
            except Exception:
                pass

        self._all_drivers: list[dict] = []
        self._loading = False

        self._apply_style()
        self._build_ui()
        # Start load slightly after mainloop kicks in → no first-open race
        self.root.after(50, self._load_drivers)

    # ── Style ─────────────────────────────────────────────────────────────

    def _apply_style(self):
        s = ttk.Style()
        try:
            s.theme_use("vista")
        except tk.TclError:
            s.theme_use("clam")

        s.configure("TFrame", background=BG)
        s.configure("Card.TFrame", background=CARD)
        s.configure("TLabel", background=BG, foreground=TEXT)
        s.configure("Card.TLabel", background=CARD, foreground=TEXT)
        s.configure("Muted.TLabel", background=BG, foreground=MUTED)
        s.configure("Title.TLabel", background=BG, foreground=TEXT,
                    font=("Segoe UI Semibold", 14))
        s.configure("Sub.TLabel", background=BG, foreground=MUTED,
                    font=("Segoe UI", 9))

        s.configure("TButton", padding=(10, 6), font=("Segoe UI", 9))
        s.configure("Accent.TButton", foreground="white", background=ACCENT,
                    padding=(12, 7), font=("Segoe UI Semibold", 9),
                    borderwidth=0)
        s.map("Accent.TButton",
              background=[("active", ACCENT2), ("disabled", "#9fbcf0")])
        s.configure("Danger.TButton", foreground="white", background=DANGER,
                    padding=(12, 7), font=("Segoe UI Semibold", 9),
                    borderwidth=0)
        s.map("Danger.TButton",
              background=[("active", DANGER2), ("disabled", "#f1a1a1")])
        s.configure("Ghost.TButton", padding=(10, 6), font=("Segoe UI", 9))

        s.configure("Treeview",
                    background=CARD, fieldbackground=CARD,
                    foreground=TEXT, rowheight=24,
                    font=("Segoe UI", 9), borderwidth=0)
        s.configure("Treeview.Heading",
                    background="#e2e8f0", foreground=TEXT,
                    font=("Segoe UI Semibold", 9), padding=(6, 6),
                    relief="flat")
        s.map("Treeview.Heading",
              background=[("active", "#cbd5e1")])
        s.map("Treeview",
              background=[("selected", "#2563eb")],
              foreground=[("selected", "white")])

        s.configure("TCheckbutton", background=BG, foreground=TEXT,
                    font=("Segoe UI", 9))
        s.configure("Card.TCheckbutton", background=CARD, foreground=TEXT,
                    font=("Segoe UI", 9))
        s.configure("TEntry", padding=4)
        s.configure("TLabelframe", background=BG, foreground=TEXT)
        s.configure("TLabelframe.Label", background=BG, foreground=MUTED,
                    font=("Segoe UI Semibold", 9))

    # ── UI layout ─────────────────────────────────────────────────────────

    def _build_ui(self):
        # ── Header ─────────────────────────────────────────────────────────
        header = ttk.Frame(self.root, padding=(14, 12, 14, 6))
        header.pack(fill=tk.X)

        ttk.Label(header, text="⚙  Driver Manager",
                  style="Title.TLabel").pack(side=tk.LEFT)
        ttk.Label(header,
                  text="Просмотр, экспорт и импорт драйверов Windows",
                  style="Sub.TLabel").pack(side=tk.LEFT, padx=(10, 0), pady=(3, 0))

        admin_text = "✔ Администратор" if is_admin() else "⚠ Без прав администратора"
        admin_fg = "#15803d" if is_admin() else "#c2410c"
        tk.Label(header, text=admin_text, bg=BG, fg=admin_fg,
                 font=("Segoe UI Semibold", 9)).pack(side=tk.RIGHT)

        # ── Toolbar ────────────────────────────────────────────────────────
        bar = ttk.Frame(self.root, padding=(14, 4, 14, 8))
        bar.pack(fill=tk.X)

        # tk.Button (not ttk) — vista theme ignores bg on ttk.Button
        tk.Button(bar, text="  📤  Экспорт выбранных  ",
                  bg=ACCENT, fg="white",
                  activebackground=ACCENT2, activeforeground="white",
                  relief="flat", bd=0, cursor="hand2",
                  font=("Segoe UI Semibold", 9), padx=4, pady=7,
                  command=self._action_export).pack(side=tk.LEFT, padx=(0, 6))
        tk.Button(bar, text="  📥  Импорт из папки  ",
                  bg=DANGER, fg="white",
                  activebackground=DANGER2, activeforeground="white",
                  relief="flat", bd=0, cursor="hand2",
                  font=("Segoe UI Semibold", 9), padx=4, pady=7,
                  command=self._action_import).pack(side=tk.LEFT, padx=(0, 6))

        ttk.Separator(bar, orient=tk.VERTICAL).pack(
            side=tk.LEFT, fill=tk.Y, padx=8, pady=3)

        ttk.Button(bar, text="↻  Обновить",
                   style="Ghost.TButton",
                   command=self._load_drivers).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(bar, text="Выделить все OEM",
                   style="Ghost.TButton",
                   command=self._select_all_oem).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(bar, text="Снять выделение",
                   style="Ghost.TButton",
                   command=lambda: self._tree.selection_remove(
                       self._tree.selection())).pack(side=tk.LEFT)

        # ── Filter strip ───────────────────────────────────────────────────
        fstrip = ttk.Frame(self.root, padding=(14, 0, 14, 8))
        fstrip.pack(fill=tk.X)

        ttk.Label(fstrip, text="Фильтр:", style="Muted.TLabel").pack(
            side=tk.LEFT, padx=(0, 6))

        self._show_oem = tk.BooleanVar(value=True)
        self._show_sys = tk.BooleanVar(value=True)
        self._show_prn = tk.BooleanVar(value=True)
        ttk.Checkbutton(fstrip, text="OEM (установленные)",
                        variable=self._show_oem,
                        command=self._filter).pack(side=tk.LEFT, padx=4)
        ttk.Checkbutton(fstrip, text="Системные",
                        variable=self._show_sys,
                        command=self._filter).pack(side=tk.LEFT, padx=4)
        ttk.Checkbutton(fstrip, text="Принтеры",
                        variable=self._show_prn,
                        command=self._filter).pack(side=tk.LEFT, padx=4)

        ttk.Separator(fstrip, orient=tk.VERTICAL).pack(
            side=tk.LEFT, fill=tk.Y, padx=10, pady=2)

        ttk.Label(fstrip, text="🔍", style="Muted.TLabel").pack(side=tk.LEFT)
        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", lambda *_: self._filter())
        ttk.Entry(fstrip, textvariable=self._search_var,
                  width=38).pack(side=tk.LEFT, padx=(4, 0))

        self._counts_lbl = ttk.Label(fstrip, text="", style="Muted.TLabel")
        self._counts_lbl.pack(side=tk.RIGHT)

        # ── Table card ─────────────────────────────────────────────────────
        card = tk.Frame(self.root, bg=CARD, highlightthickness=1,
                        highlightbackground="#e2e8f0", highlightcolor="#e2e8f0")
        card.pack(fill=tk.BOTH, expand=True, padx=14, pady=(0, 4))

        tree_frame = ttk.Frame(card, style="Card.TFrame", padding=1)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        cols = [c[0] for c in self.COL_DEFS]
        self._tree = ttk.Treeview(
            tree_frame, columns=cols, show="headings",
            selectmode="extended",
        )
        for cid, heading, width, minw, stretch in self.COL_DEFS:
            self._tree.heading(cid, text=heading,
                               command=lambda c=cid: self._sort_by(c))
            self._tree.column(cid, width=width, minwidth=minw, stretch=stretch)

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

        # Alternating row colours
        self._tree.tag_configure("oem",  background=ROW_OEM)
        self._tree.tag_configure("oem2", background=ROW_OEM_ALT)
        self._tree.tag_configure("sys",  background=ROW_SYS)
        self._tree.tag_configure("sys2", background=ROW_SYS_ALT)
        self._tree.tag_configure("prn",  background=ROW_PRN)
        self._tree.tag_configure("prn2", background=ROW_PRN_ALT)

        # ── Context menu ───────────────────────────────────────────────────
        self._ctx = tk.Menu(self.root, tearoff=0,
                            bg=CARD, fg=TEXT,
                            activebackground=ACCENT, activeforeground="white",
                            font=("Segoe UI", 9), bd=0)
        self._ctx.add_command(label="  📤  Экспортировать…",
                              command=self._action_export)
        self._ctx.add_command(label="  📁  Открыть расположение",
                              command=self._ctx_open_folder)
        self._ctx.add_separator()
        self._ctx.add_command(label="  🗑  Удалить драйвер",
                              command=self._ctx_delete)

        self._tree.bind("<Button-3>", self._on_right_click)
        self._tree.bind("<<TreeviewSelect>>", self._on_select_change)
        self._tree.bind("<Control-a>", self._on_ctrl_a)
        self._tree.bind("<Return>", lambda e: self._ctx_open_folder())
        self._tree.bind("<Delete>", lambda e: self._ctx_delete())

        # ── Status bar ─────────────────────────────────────────────────────
        self._status_var = tk.StringVar(value="Готов")
        statusbar = tk.Frame(self.root, bg="#e2e8f0", height=22)
        statusbar.pack(fill=tk.X, side=tk.BOTTOM)
        tk.Label(statusbar, textvariable=self._status_var,
                 bg="#e2e8f0", fg=TEXT, anchor=tk.W,
                 font=("Segoe UI", 9), padx=10).pack(side=tk.LEFT, fill=tk.X, expand=True)

        self._sel_lbl = tk.Label(statusbar, text="", bg="#e2e8f0", fg=MUTED,
                                 font=("Segoe UI", 9), padx=10)
        self._sel_lbl.pack(side=tk.RIGHT)

    # ── Loading ───────────────────────────────────────────────────────────

    def _load_drivers(self):
        if self._loading:
            return
        self._loading = True
        self._set_status("Загрузка списка драйверов… это может занять до 30 секунд")
        self._tree.delete(*self._tree.get_children())
        self._counts_lbl.configure(text="")
        threading.Thread(target=self._load_thread, daemon=True).start()

    def _load_thread(self):
        oem = get_oem_drivers()
        sys_d = get_system_drivers()
        prn = get_printer_drivers()

        # Retry once if everything is empty — WMI first-call hiccup
        if not oem and not sys_d and not prn:
            import time
            time.sleep(0.4)
            oem = get_oem_drivers()
            sys_d = get_system_drivers()
            prn = get_printer_drivers()

        self._all_drivers = oem + prn + sys_d
        self.root.after(0, self._populate_done)

    def _populate_done(self):
        self._loading = False
        self._populate()

    # ── Populate / filter ──────────────────────────────────────────────────

    def _populate(self):
        self._tree.delete(*self._tree.get_children())
        query = self._search_var.get().lower()
        show_oem = self._show_oem.get()
        show_sys = self._show_sys.get()

        show_prn = self._show_prn.get()
        shown = 0
        oem_i = sys_i = prn_i = 0
        for d in self._all_drivers:
            dtype = d["driver_type"]
            if dtype == "OEM" and not show_oem:
                continue
            if dtype == "Системный" and not show_sys:
                continue
            if dtype == "Принтер" and not show_prn:
                continue

            hay = " ".join(str(d.get(k, "")) for k in
                           ("name", "inf", "version", "provider", "class", "path")).lower()
            if query and query not in hay:
                continue

            if dtype == "OEM":
                tag = "oem" if oem_i % 2 == 0 else "oem2"
                oem_i += 1
            elif dtype == "Принтер":
                tag = "prn" if prn_i % 2 == 0 else "prn2"
                prn_i += 1
            else:
                tag = "sys" if sys_i % 2 == 0 else "sys2"
                sys_i += 1

            self._tree.insert("", "end", values=(
                dtype,
                d.get("name", ""),
                d.get("inf", ""),
                d.get("version", ""),
                d.get("date", ""),
                d.get("provider", ""),
                d.get("class", ""),
                d.get("path", ""),
            ), tags=(tag,))
            shown += 1

        total_oem = sum(1 for d in self._all_drivers if d["driver_type"] == "OEM")
        total_sys = sum(1 for d in self._all_drivers if d["driver_type"] == "Системный")
        total_prn = sum(1 for d in self._all_drivers if d["driver_type"] == "Принтер")
        self._counts_lbl.configure(
            text=f"OEM: {total_oem}   •   Принтеров: {total_prn}   •   "
                 f"Системных: {total_sys}   •   Показано: {shown}")

        if not self._all_drivers:
            self._set_status("⚠ Не удалось получить список драйверов. Нажмите «Обновить».")
        else:
            self._set_status(f"Загружено {len(self._all_drivers)} драйверов")

    def _filter(self):
        self._populate()

    # ── Selection helpers ─────────────────────────────────────────────────

    def _on_select_change(self, _event=None):
        n = len(self._tree.selection())
        self._sel_lbl.configure(
            text=f"Выделено: {n}" if n else "")

    def _on_ctrl_a(self, _event=None):
        self._tree.selection_set(self._tree.get_children())
        return "break"

    def _select_all_oem(self):
        oem = [iid for iid in self._tree.get_children()
               if self._tree.item(iid, "values")[0] == "OEM"]
        self._tree.selection_set(oem)

    def _sort_by(self, col: str):
        reverse = getattr(self, "_sort_rev", {}).get(col, False)
        data = [(self._tree.set(iid, col), iid)
                for iid in self._tree.get_children()]
        data.sort(key=lambda x: x[0].lower(), reverse=reverse)
        for idx, (_, iid) in enumerate(data):
            self._tree.move(iid, "", idx)
        rev = getattr(self, "_sort_rev", {})
        rev[col] = not reverse
        self._sort_rev = rev

    # ── Context menu ──────────────────────────────────────────────────────

    def _on_right_click(self, event):
        iid = self._tree.identify_row(event.y)
        if iid and iid not in self._tree.selection():
            self._tree.selection_set(iid)
        if not self._tree.selection():
            return
        try:
            self._ctx.tk_popup(event.x_root, event.y_root)
        finally:
            self._ctx.grab_release()

    def _ctx_open_folder(self):
        sel = self._tree.selection()
        if not sel:
            return
        # Take first selected row
        vals = self._tree.item(sel[0], "values")
        path = vals[7]  # path column
        if not path:
            messagebox.showinfo("Путь не найден",
                                "Для этого драйвера путь недоступен.")
            return
        # Select the file in Explorer if it exists, else open dir
        if os.path.exists(path):
            # Normalise slashes
            subprocess.Popen(
                ["explorer.exe", f"/select,{os.path.normpath(path)}"])
        else:
            parent = os.path.dirname(path)
            if parent and os.path.exists(parent):
                try:
                    os.startfile(parent)
                except Exception:
                    subprocess.Popen(["explorer.exe", parent])
            else:
                messagebox.showwarning(
                    "Путь недоступен",
                    f"Не удалось найти:\n{path}")

    def _ctx_delete(self):
        sel = self._tree.selection()
        if not sel:
            return

        deletable = [iid for iid in sel
                     if self._tree.item(iid, "values")[0] in ("OEM", "Принтер")]
        non_del = len(sel) - len(deletable)

        if not deletable:
            messagebox.showwarning(
                "Только OEM и принтеры",
                "Удалять можно только OEM-драйверы и драйверы принтеров.\n"
                "Системные драйверы трогать нельзя.")
            return

        if not is_admin():
            if messagebox.askyesno(
                "Нужны права администратора",
                "Удаление требует прав администратора.\n\n"
                "Перезапустить программу от имени администратора?"):
                relaunch_as_admin()
            return

        names = "\n".join("• " + self._tree.item(i, "values")[1]
                          for i in deletable[:8])
        if len(deletable) > 8:
            names += f"\n… и ещё {len(deletable) - 8}"

        extra = f"\n\n(Пропущено системных: {non_del})" if non_del else ""

        if not messagebox.askyesno(
            "Подтверждение удаления",
            f"Удалить {len(deletable)} драйвер(ов)?\n\n{names}{extra}\n\n"
            "Операция необратима — драйверы будут удалены из системы."):
            return

        self._set_status("Удаление драйверов…")
        threading.Thread(
            target=self._delete_thread, args=(deletable,), daemon=True).start()

    def _delete_thread(self, iids: list):
        ok, fail = 0, 0
        errors: list[str] = []
        for iid in iids:
            vals = self._tree.item(iid, "values")
            dtype, name, inf = vals[0], vals[1], vals[2]
            try:
                if dtype == "Принтер":
                    # Remove printer driver via PS cmdlet
                    escaped = name.replace("'", "''")
                    script = (f"Remove-PrinterDriver -Name '{escaped}' "
                              f"-ErrorAction Stop; Write-Output 'OK'")
                    out = _run_ps(script, timeout=60)
                    if out.strip().endswith("OK"):
                        ok += 1
                    else:
                        fail += 1
                        errors.append(f"{name}: {out.splitlines()[0] if out else 'ошибка'}")
                else:
                    if not inf:
                        fail += 1
                        errors.append(f"{name}: нет INF имени")
                        continue
                    r = _run_cmd(
                        ["pnputil.exe", "/delete-driver", inf, "/uninstall", "/force"],
                        timeout=120)
                    if r.returncode == 0:
                        ok += 1
                    else:
                        fail += 1
                        msg = (r.stderr or r.stdout or "").strip().splitlines()
                        errors.append(f"{inf}: {msg[0] if msg else 'ошибка'}")
            except Exception as exc:
                fail += 1
                errors.append(f"{name}: {exc}")

        text = f"Удалено: {ok}\nОшибок: {fail}"
        if errors:
            text += "\n\nПодробности:\n" + "\n".join(errors[:12])
        self.root.after(0, lambda: self._set_status(
            f"Удаление завершено: {ok} ✓ / {fail} ✗"))
        self.root.after(0, lambda: messagebox.showinfo("Удаление завершено", text))
        self.root.after(0, self._load_drivers)

    # ── Export ────────────────────────────────────────────────────────────

    def _action_export(self):
        sel = self._tree.selection()
        if not sel:
            messagebox.showwarning(
                "Ничего не выбрано",
                "Выделите драйверы в списке (Ctrl+клик / Shift+клик).")
            return

        exportable = [iid for iid in sel
                      if self._tree.item(iid, "values")[0] in ("OEM", "Принтер")]
        if not exportable:
            messagebox.showwarning(
                "Только OEM и принтеры",
                "Экспорт работает только для OEM-драйверов и драйверов принтеров.\n"
                "Системные драйверы нельзя выгрузить отдельно.")
            return

        folder = filedialog.askdirectory(title="Папка для экспорта драйверов")
        if not folder:
            return

        self._set_status("Экспорт…")
        threading.Thread(
            target=self._export_thread,
            args=(exportable, folder), daemon=True).start()

    def _export_thread(self, iids: list, folder: str):
        import shutil
        ok, fail = 0, 0
        errors: list[str] = []
        for iid in iids:
            vals = self._tree.item(iid, "values")
            dtype, name, inf, path = vals[0], vals[1], vals[2], vals[7]
            if not inf and not path:
                fail += 1
                errors.append(f"{name}: нет INF")
                continue
            safe = re.sub(r"[^\w.-]+", "_", name)[:50] or inf or "driver"
            dest = os.path.join(folder, f"{(inf or 'drv').replace('.inf','')}_{safe}")
            os.makedirs(dest, exist_ok=True)

            exported = False
            err_msg = "не удалось экспортировать"
            # Try pnputil first — works for all drivers published in DriverStore
            if inf:
                try:
                    r = _run_cmd(
                        ["pnputil.exe", "/export-driver", inf, dest],
                        timeout=60)
                    if r.returncode == 0:
                        ok += 1
                        exported = True
                    else:
                        lines = (r.stderr or r.stdout or "").strip().splitlines()
                        err_msg = lines[0] if lines else "pnputil error"
                except Exception as exc:
                    err_msg = str(exc)

            # Fallback: copy DriverStore folder if path is real
            if not exported and path and os.path.isfile(path):
                src_dir = os.path.dirname(path)
                if "DriverStore" in src_dir and os.path.isdir(src_dir):
                    try:
                        shutil.copytree(src_dir, dest, dirs_exist_ok=True)
                        ok += 1
                        exported = True
                    except Exception as exc:
                        err_msg = f"copy failed: {exc}"

            if not exported:
                fail += 1
                errors.append(f"{inf or name}: {err_msg}")

        text = f"Экспортировано: {ok}\nОшибок: {fail}\n\nПапка:\n{folder}"
        if errors:
            text += "\n\nПодробности:\n" + "\n".join(errors[:12])
        self.root.after(0, lambda: self._set_status(
            f"Экспорт завершён: {ok} ✓ / {fail} ✗"))
        self.root.after(0, lambda: messagebox.showinfo("Экспорт завершён", text))

    # ── Import ────────────────────────────────────────────────────────────

    def _action_import(self):
        if not is_admin():
            if messagebox.askyesno(
                "Нужны права администратора",
                "Установка драйверов требует прав администратора.\n\n"
                "Перезапустить программу от имени администратора?"):
                relaunch_as_admin()
            return

        folder = filedialog.askdirectory(title="Папка с драйверами (.inf файлы)")
        if not folder:
            return

        infs: list[str] = []
        for root_dir, _, files in os.walk(folder):
            for f in files:
                if f.lower().endswith(".inf"):
                    infs.append(os.path.join(root_dir, f))

        if not infs:
            messagebox.showwarning("INF не найдены",
                                   f"Не найдено .inf файлов в:\n{folder}")
            return

        if not messagebox.askyesno(
            "Подтверждение установки",
            f"Найдено INF-файлов: {len(infs)}\n\nУстановить все?"):
            return

        self._set_status("Установка драйверов…")
        threading.Thread(
            target=self._import_thread, args=(infs,), daemon=True).start()

    def _import_thread(self, infs: list[str]):
        ok, fail = 0, 0
        errors: list[str] = []
        for p in infs:
            try:
                r = _run_cmd(
                    ["pnputil.exe", "/add-driver", p, "/install"],
                    timeout=120)
                if r.returncode == 0:
                    ok += 1
                else:
                    fail += 1
                    msg = (r.stderr or r.stdout or "").strip().splitlines()
                    errors.append(f"{os.path.basename(p)}: {msg[0] if msg else 'ошибка'}")
            except Exception as exc:
                fail += 1
                errors.append(f"{os.path.basename(p)}: {exc}")

        text = f"Установлено: {ok}\nОшибок: {fail}"
        if errors:
            text += "\n\nПодробности:\n" + "\n".join(errors[:12])
        self.root.after(0, lambda: self._set_status(
            f"Установка завершена: {ok} ✓ / {fail} ✗"))
        self.root.after(0, lambda: messagebox.showinfo("Установка завершена", text))
        self.root.after(0, self._load_drivers)

    # ── Helpers ───────────────────────────────────────────────────────────

    def _set_status(self, text: str):
        self._status_var.set(text)


# ─── Entry point ──────────────────────────────────────────────────────────

def main():
    root = tk.Tk()
    DriverManagerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
