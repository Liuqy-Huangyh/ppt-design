<#
  卸载 PPT Studio 桌面快捷方式（保留 server\store 下的档案/deck/日志等数据）。

  用法：
    powershell -NoProfile -ExecutionPolicy Bypass -File scripts\uninstall-desktop.ps1
    powershell -NoProfile -ExecutionPolicy Bypass -File scripts\uninstall-desktop.ps1 -StartMenu
    powershell -NoProfile -ExecutionPolicy Bypass -File scripts\uninstall-desktop.ps1 -WhatIf:$true
#>
param(
  [switch]$StartMenu,
  [string]$DesktopDir = "",
  [object]$WhatIf = $null
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [Text.Encoding]::UTF8

$previewMode = $false
if ($null -ne $WhatIf) {
  $text = ([string]$WhatIf).Trim().ToLowerInvariant()
  if ($text -notmatch '^(\$)?(false|0|no)$') { $previewMode = $true }
}

$Root = Split-Path -Parent $PSScriptRoot
$DesktopDir = if ($DesktopDir) { $DesktopDir } else { [Environment]::GetFolderPath("Desktop") }
$iconCopyDir = Join-Path $Root "server\store\desktop"
$iconCopy = Join-Path $iconCopyDir "icon.ico"

$targets = New-Object System.Collections.Generic.List[string]
$targets.Add((Join-Path $DesktopDir "PPT Studio.lnk"))
$targets.Add((Join-Path $Root "desktop-fallback\PPT Studio.lnk"))
if ($StartMenu) {
  $startMenuDir = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs"
  $targets.Add((Join-Path $startMenuDir "PPT Studio.lnk"))
}

$removed = New-Object System.Collections.Generic.List[string]
$planned = New-Object System.Collections.Generic.List[string]
foreach ($path in $targets) {
  if (-not (Test-Path $path)) { continue }
  if ($previewMode) { $planned.Add($path); continue }
  Remove-Item -Path $path -Force
  $removed.Add($path)
}
if (Test-Path $iconCopy) {
  if ($previewMode) { $planned.Add($iconCopy) }
  else {
    Remove-Item -Path $iconCopy -Force
    $removed.Add($iconCopy)
  }
}
if ((-not $previewMode) -and (Test-Path $iconCopyDir)) {
  $left = Get-ChildItem -Path $iconCopyDir -Force -ErrorAction SilentlyContinue
  if (-not $left) {
    Remove-Item -Path $iconCopyDir -Force -ErrorAction SilentlyContinue
  }
}

Write-Host "PPT Studio 桌面卸载完成" -ForegroundColor Cyan
foreach ($path in $removed) { Write-Host "  已删除：$path" }
foreach ($path in $planned) { Write-Host "  将删除（-WhatIf 预览）：$path" -ForegroundColor DarkGray }
if ($removed.Count -eq 0 -and $planned.Count -eq 0) {
  Write-Host "  没有找到需要删除的快捷方式" -ForegroundColor DarkGray
}
Write-Host "  数据保留：$(Join-Path $Root 'server\store')（档案 / deck / 日志不受影响）"
