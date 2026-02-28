# RightClickManager

一个面向 Windows 的右键菜单管理器（PySide6 桌面应用），支持扫描、筛选、排序、编辑与批量管理常见右键菜单项。

## 功能特性

- 全量扫描常见位置（`HKCU` / `HKLM`）
  - `shell`
  - `shellex\\ContextMenuHandlers`（只读展示）
- 搜索过滤（输入防抖，减少卡顿）
- 列表表头排序（升序/降序）
- 单条操作
  - 更新
  - 启用/禁用
  - 删除
- 批量操作（多选）
  - 批量启用
  - 批量禁用
  - 批量删除
- 导入/导出 JSON 备份
  - 导出自动跳过不可导出的子菜单占位项
  - 导入支持幂等更新（同 key 覆盖更新，不无限重复创建）
- 现代化圆角 UI
  - 卡片布局
  - 高对比头部
  - 平滑滚动优化

## 支持环境

- Windows 10 / 11
- Python 3.9+（当前已在 3.14 验证）

## 快速开始

### 1. 安装依赖

```powershell
py -3 -m pip install PySide6
```

### 2. 运行

```powershell
py -3 .\context_menu_manager.py
```

## 打包 EXE

### 推荐方式：脚本打包

```powershell
.\build_exe.ps1
```

默认输出：

`dist\RightClickManager.exe`

可指定输出名：

```powershell
.\build_exe.ps1 -AppName RightClickManager_v6
```

如果旧 exe 正在运行、文件被占用，脚本会自动切换到时间戳文件名继续打包。

### 手动打包

```powershell
py -3 -m pip install pyinstaller
py -3 -m PyInstaller --noconfirm --clean --onefile --windowed --name "RightClickManager" .\context_menu_manager.py
```

## 管理范围说明

默认扫描以下路径（HKCU + HKLM）：

- `Software\Classes\*\shell`
- `Software\Classes\AllFilesystemObjects\shell`
- `Software\Classes\Directory\shell`
- `Software\Classes\Directory\Background\shell`
- `Software\Classes\Folder\shell`
- `Software\Classes\Drive\shell`
- `Software\Classes\DesktopBackground\shell`
- 对应 `shellex\ContextMenuHandlers`

说明：

- `shell` 条目支持新增/更新/启用禁用/删除。
- `ContextMenuHandlers` 当前仅用于补全可见范围（只读展示）。

## JSON 导入导出格式

导出结构示例：

```json
{
  "version": 1,
  "exported_at": "2026-02-28T18:00:00",
  "count": 2,
  "skipped_unexportable": 1,
  "items": [
    {
      "scope_id": "hkcu_file_shell",
      "display_name": "用 Notepad 打开",
      "key_name": "OpenWithNotepad",
      "command": "notepad.exe \"%1\"",
      "enabled": true
    }
  ]
}
```

字段说明：

- `scope_id`：导入目标位置
- `display_name`：菜单显示名称
- `key_name`：键名（建议稳定）
- `command`：执行命令
- `enabled`：是否启用

## 权限与安全

- 读取 HKLM 一般可行，但写入 HKLM 通常需要管理员权限。
- 删除操作不可撤销，建议先导出 JSON 备份。
- 本工具不会主动提升权限，请按需用管理员身份运行。

## 常见问题

### 1. 为什么有些菜单项显示为只读？

- 可能是 `ContextMenuHandlers`（本版本仅展示）
- 或目标位于 HKLM 且当前权限不足

### 2. 右键菜单改完没立即生效？

- Windows 资源管理器可能有缓存，可尝试重启资源管理器或重新登录。

### 3. 重复导入会不会越导越多？

- 目前导入为幂等更新（同 key 覆盖），不会无上限重复堆叠。

## 项目结构

```text
.
├─ context_menu_manager.py   # 主程序（UI + 注册表逻辑）
├─ build_exe.ps1             # 打包脚本
└─ README.md                 # 项目说明
```

## 路线图

- [ ] 操作历史（撤销/重做）
- [ ] 导入前差异预览（dry-run）
- [ ] 第三方菜单过滤视图
- [ ] 英文文档（README_EN.md）

## License

本项目基于 **MIT License** 开源，详见 [LICENSE](./LICENSE)。

补充说明：

- 你可以自由使用、修改、分发本项目代码（含商用）。
- 分发时请保留原始许可证声明。
- 本项目按 “AS IS” 提供，不附带任何担保。
