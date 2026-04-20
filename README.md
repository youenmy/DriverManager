# Driver Manager

Portable Windows utility for listing, exporting, and installing device drivers.
Designed for one simple workflow: **pull drivers from one PC, install them on another.**

![Driver Manager](icon.png)

---

## Features

- **List all drivers** installed in Windows, with real device names, versions, dates, providers, and file paths
- **Two categories** with toggleable filters:
  - **OEM** — third-party drivers (`oem*.inf`) stored in the Driver Store
  - **System** — kernel-mode drivers from `Win32_SystemDriver`
- **Path column** shows the actual driver store folder (OEM) or `.sys` file path (system)
- **File-manager style multi-select** — `Click` / `Ctrl+Click` / `Shift+Click` / `Ctrl+A`
- **Right-click context menu**:
  - 📤 Export selected drivers…
  - 📁 Open driver location in Explorer
  - 🗑 Uninstall driver (OEM only)
- **Live search** across name, INF, version, provider, class, path
- **Portable single-file EXE** — no install, no dependencies, no admin required to browse
- Auto-elevation via UAC when installing / uninstalling drivers

## Download

Grab the latest EXE from [Releases](https://github.com/youenmy/DriverManager/releases/latest):

**→ [DriverManager.exe](https://github.com/youenmy/DriverManager/releases/latest/download/DriverManager.exe)** (~10 MB)

Just double-click. Accept the UAC prompt if you intend to install or uninstall drivers.

## Usage

### Export drivers from PC A

1. Launch `DriverManager.exe`
2. Filter to **OEM** in the top strip (only OEM drivers can be exported)
3. Select the drivers you need (`Ctrl+Click` for multiple, `Shift+Click` for range, `Ctrl+A` for all)
4. Click **📤 Экспорт выбранных** or right-click → **Экспортировать…**
5. Choose a destination folder (e.g. `D:\MyDrivers`)
6. Done — each driver goes into its own subfolder with the `.inf`, `.sys` and related files

### Import drivers on PC B

1. Launch `DriverManager.exe` (UAC will ask for admin rights)
2. Click **📥 Импорт из папки**
3. Point it at the folder containing `.inf` files (recursively scanned)
4. Confirm — all found drivers are installed with `pnputil /add-driver /install`

### Uninstall an OEM driver

1. Select the OEM driver(s)
2. Right-click → **🗑 Удалить драйвер**
3. Confirm — runs `pnputil /delete-driver <inf> /uninstall /force`

## Under the hood

| Operation | Tool |
|---|---|
| Enumerate OEM drivers | `Get-CimInstance Win32_PnPSignedDriver` + `pnputil /enum-drivers` |
| Enumerate system drivers | `Get-CimInstance Win32_SystemDriver` |
| Export | `pnputil /export-driver <inf> <dest>` |
| Import | `pnputil /add-driver <inf> /install` |
| Uninstall | `pnputil /delete-driver <inf> /uninstall /force` |
| Open location | `explorer.exe /select,<path>` |

GUI is plain Python `tkinter` + `ttk` (Vista theme) — no extra runtime deps.

## Build from source

Requires Python 3.10+ and PyInstaller.

```bat
pip install pyinstaller pillow
python make_icon.py
pyinstaller --onefile --windowed --name DriverManager ^
    --icon icon.ico --add-data "icon.ico;." ^
    --uac-admin driver_manager.py
```

Or just run the included [build_exe.bat](build_exe.bat).

The resulting `dist\DriverManager.exe` is fully portable — copy it anywhere.

## Repo layout

```
driver_manager.py   — main application (single file)
make_icon.py        — regenerates icon.ico with Pillow
build_exe.bat       — one-click PyInstaller build script
icon.ico / icon.png — app icon
```

## Notes

- Only **OEM** drivers can be exported, uninstalled, or re-installed. Windows system
  drivers are tied to the OS and shouldn't be extracted.
- If the list is empty on first open: click **↻ Обновить** — some WMI queries need a
  warm-up call.
- Some OEM drivers (especially those with additional installer packages like NVIDIA
  GeForce Experience) may require the original vendor installer on the target PC to
  be fully functional; `pnputil` only transfers the kernel driver itself.

## License

MIT
