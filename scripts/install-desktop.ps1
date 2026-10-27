<#
  安装 PPT Studio 桌面快捷方式。

  用法：
    powershell -NoProfile -ExecutionPolicy Bypass -File scripts\install-desktop.ps1
    powershell -NoProfile -ExecutionPolicy Bypass -File scripts\install-desktop.ps1 -StartMenu
    powershell -NoProfile -ExecutionPolicy Bypass -File scripts\install-desktop.ps1 -WhatIf:$true   # 只看计划

  说明：
    * 快捷方式名 “PPT Studio”，指向 scripts\launch-ppt-studio.ps1，图标用自绘 assets\icon.ico；
    * 图标会复制到 server\store\desktop\icon.ico，卸载时删除副本，不动其他数据；
    * DSH 等文件沙箱会拒绝写桌面：此时脚本在工作区 desktop-fallback\ 生成一份副本并明确提示，
      在普通 PowerShell（非沙箱）中重新运行即可装到真实桌面。
#>
param(
  [switch]$StartMenu,
  [string]$DesktopDir = "",
  [string]$IconSource = "",
  [object]$WhatIf = $null
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [Text.Encoding]::UTF8

# -File 方式会把 `-WhatIf:$false` 作为字符串传进来（PowerShell 的 -File 不解析表达式），
# 这里统一兼容：null=执行；"$false"/"false"/"0"/"no"=执行；其余非空=只看计划。
$previewMode = $false
if ($null -ne $WhatIf) {
  $text = ([string]$WhatIf).Trim().ToLowerInvariant()
  if ($text -notmatch '^(\$)?(false|0|no)$') { $previewMode = $true }
}

$Root = Split-Path -Parent $PSScriptRoot
$LaunchScript = Join-Path $PSScriptRoot "launch-ppt-studio.ps1"
$IconSource = if ($IconSource) { $IconSource } else { Join-Path $Root "assets\icon.ico" }
$DesktopDir = if ($DesktopDir) { $DesktopDir } else { [Environment]::GetFolderPath("Desktop") }

if (-not (Test-Path $LaunchScript)) { throw "缺少启动脚本：$LaunchScript" }
if (-not (Test-Path $IconSource)) {
  throw "缺少图标：$IconSource（可运行 python scripts\make-icon.py 生成）"
}

$iconCopyDir = Join-Path $Root "server\store\desktop"
$iconCopy = Join-Path $iconCopyDir "icon.ico"
if (-not $previewMode) {
  New-Item -ItemType Directory -Force -Path $iconCopyDir | Out-Null
  Copy-Item -Path $IconSource -Destination $iconCopy -Force
}

$powershellExe = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
if (-not (Test-Path $powershellExe)) {
  $powershellExe = (Get-Command powershell.exe -ErrorAction Stop).Source
}

function New-StudioShortcut([string]$LinkPath) {
  $parent = Split-Path -Parent $LinkPath
  if ($parent) { New-Item -ItemType Directory -Force -Path $parent | Out-Null }
  $shell = New-Object -ComObject WScript.Shell
  $shortcut = $shell.CreateShortcut($LinkPath)
  $shortcut.TargetPath = $powershellExe
  $shortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$LaunchScript`""
  $shortcut.WorkingDirectory = $Root
  $shortcut.IconLocation = "$iconCopy,0"
  $shortcut.Description = "PPT Studio · 双击启动本地 PPT 制作服务"
  $shortcut.WindowStyle = 1
  $shortcut.Save()
}

$targets = New-Object System.Collections.Generic.List[object]
$targets.Add([pscustomobject]@{ Kind = "桌面"; Path = (Join-Path $DesktopDir "PPT Studio.lnk") })
if ($StartMenu) {
  $startMenuDir = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs"
  $targets.Add([pscustomobject]@{ Kind = "开始菜单"; Path = (Join-Path $startMenuDir "PPT Studio.lnk") })
}

$fallbackAllowed = [bool]$env:DSH_SHELL
$installed = New-Object System.Collections.Generic.List[string]
$planned = New-Object System.Collections.Generic.List[string]
$warnings = New-Object System.Collections.Generic.List[string]

foreach ($target in $targets) {
  if ($previewMode) {
    $planned.Add($target.Path)
    continue
  }
  try {
    New-StudioShortcut $target.Path
    $installed.Add($target.Path)
  } catch {
    if (-not $fallbackAllowed) { throw }
    $fallbackDir = Join-Path $Root "desktop-fallback"
    New-Item -ItemType Directory -Force -Path $fallbackDir | Out-Null
    $fallback = Join-Path $fallbackDir "PPT Studio.lnk"
    New-StudioShortcut $fallback
    $installed.Add($fallback)
    $warnings.Add("无法写入$($target.Kind)（$($_.Exception.Message)）。已生成工作区副本：$fallback。请在没有沙箱限制的普通 PowerShell 中重新运行本脚本，即可安装到$($target.Kind)。")
  }
}

Write-Host "PPT Studio 桌面安装完成" -ForegroundColor Cyan
Write-Host "  启动脚本：$LaunchScript"
Write-Host "  图标：$iconCopy"
foreach ($path in $installed) { Write-Host "  快捷方式：$path" }
foreach ($path in $planned) { Write-Host "  将创建（-WhatIf 预览）：$path" -ForegroundColor DarkGray }
foreach ($warning in $warnings) { Write-Host "  注意：$warning" -ForegroundColor Yellow }
Write-Host "  卸载：powershell -NoProfile -ExecutionPolicy Bypass -File scripts\uninstall-desktop.ps1"
