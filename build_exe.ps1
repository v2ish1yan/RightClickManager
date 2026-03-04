[CmdletBinding()]
param(
    [string]$AppName = "RightClickManager"
)

$ErrorActionPreference = "Stop"

Write-Host "==> Checking PyInstaller..."
$pyiVersion = py -3 -m PyInstaller --version 2>$null
if (-not $pyiVersion) {
    Write-Host "==> Installing PyInstaller..."
    py -3 -m pip install pyinstaller
}

Write-Host "==> Cleaning old build artifacts..."
if (Test-Path ".\build") { Remove-Item ".\build" -Recurse -Force }
if (-not (Test-Path ".\dist")) { New-Item ".\dist" -ItemType Directory | Out-Null }

$targetName = $AppName
$targetExe = ".\dist\$targetName.exe"
if (Test-Path $targetExe) {
    try {
        Remove-Item $targetExe -Force -ErrorAction Stop
    } catch {
        $suffix = Get-Date -Format "yyyyMMdd_HHmmss"
        $targetName = "${AppName}_${suffix}"
        Write-Host "==> Existing exe is in use, switching output name to: $targetName"
    }
}

Write-Host "==> Building exe..."
py -3 -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --uac-admin `
    --name "$targetName" `
    .\context_menu_manager.py

Write-Host "==> Done: dist\$targetName.exe"
