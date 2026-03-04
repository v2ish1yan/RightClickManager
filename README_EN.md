# RightClickManager

A Windows context menu manager built with PySide6.  
It helps you scan, filter, sort, edit, and batch-manage common right-click menu entries.

## Features

- Full scan across common locations in `HKCU` and `HKLM`
  - `shell` entries
  - `shellex\ContextMenuHandlers` (read-only display)
- Search and filter with input debounce for smoother interaction
- Clickable table-header sorting (ascending/descending)
- Single-item actions
  - Update
  - Enable/Disable
  - Delete
- Batch actions (multi-select)
  - Batch enable
  - Batch disable
  - Batch delete
- Action history
  - Supports undo/redo for writable registry changes
- JSON import/export backup
  - Export skips non-exportable submenu placeholder items
  - Import is idempotent (same key is updated instead of duplicated repeatedly)
  - Supports import diff preview (dry-run) before applying changes
- Third-party menu filtering view
  - Switch between `All / Third-Party Only / System Only`
- Modern rounded UI
  - Card-style layout
  - High-contrast header
  - Smoother scrolling behavior

## Requirements

- Windows 10 / 11
- Python 3.9+

## Quick Start

### 1. Install dependencies

```powershell
py -3 -m pip install PySide6
```

### 2. Run

```powershell
py -3 .\context_menu_manager.py
```

## Build EXE

### Recommended: build script

```powershell
.\build_exe.ps1
```

Default output:

`dist\RightClickManager.exe`

Specify output name:

```powershell
.\build_exe.ps1 -AppName RightClickManager_v6
```

If the old EXE is still running (file locked), the script automatically falls back to a timestamped filename.

### Manual build

```powershell
py -3 -m pip install pyinstaller
py -3 -m PyInstaller --noconfirm --clean --onefile --windowed --uac-admin --name "RightClickManager" .\context_menu_manager.py
```

## Managed Registry Scope

By default, the app scans these paths under both HKCU and HKLM:

- `Software\Classes\*\shell`
- `Software\Classes\AllFilesystemObjects\shell`
- `Software\Classes\Directory\shell`
- `Software\Classes\Directory\Background\shell`
- `Software\Classes\Folder\shell`
- `Software\Classes\Drive\shell`
- `Software\Classes\DesktopBackground\shell`
- Their corresponding `shellex\ContextMenuHandlers`

Notes:

- `shell` entries support create/update/enable-disable/delete.
- `ContextMenuHandlers` entries are currently shown for visibility only (read-only).

## JSON Import/Export Format

Example export structure:

```json
{
  "version": 1,
  "exported_at": "2026-02-28T18:00:00",
  "count": 2,
  "skipped_unexportable": 1,
  "items": [
    {
      "scope_id": "hkcu_file_shell",
      "display_name": "Open with Notepad",
      "key_name": "OpenWithNotepad",
      "command": "notepad.exe \"%1\"",
      "enabled": true
    }
  ]
}
```

Field descriptions:

- `scope_id`: target scope for import
- `display_name`: visible menu text
- `key_name`: registry key name (stable naming recommended)
- `command`: command to execute
- `enabled`: whether this item is enabled

## Permissions and Safety

- Reading HKLM is usually allowed, but writing HKLM typically requires Administrator privileges.
- Delete is irreversible; export a JSON backup first.
- The app now defaults to Administrator launch (UAC prompt expected).

## FAQ

### 1. Why are some entries read-only?

- They may be `ContextMenuHandlers` (display-only in this version), or
- The target is in HKLM and your current process lacks write permission.

### 2. Why do changes not apply immediately?

- Windows Explorer may cache context menu data. Try restarting Explorer or signing out and back in.

### 3. Will repeated imports create duplicates forever?

- No. Import uses idempotent updates (same key is overwritten, not endlessly duplicated).

## Project Structure

```text
.
├─ context_menu_manager.py   # Main app (UI + registry logic)
├─ build_exe.ps1             # Build script
├─ README.md                 # Chinese documentation
└─ README_EN.md              # English documentation
```

## Roadmap

- [x] Action history (undo/redo)
- [x] Import diff preview (dry-run)
- [x] Third-party menu filtering view

## License

Licensed under the **MIT License**. See [LICENSE](./LICENSE).

Additional notes:

- You can use, modify, and redistribute this project (including commercial usage).
- Keep the original license notice in redistributions.
- The software is provided "AS IS", without warranty of any kind.
