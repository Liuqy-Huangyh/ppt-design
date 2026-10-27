<#
  PPT Studio 启动器看门狗。

  启动器用 Job Object 保证 python 随宿主一起退出；本脚本只负责两件事：
    * 宿主窗口被关闭（进程直接消失，来不及写日志）时，补一条退出原因；
    * 万一 Job Object 不可用，兜底 kill 掉残留的服务进程。

  正常情况下启动器会先写日志并删除 state 文件，看门狗检测不到 state 就静默退出。
#>
param(
  [int]$LauncherPid,
  [int]$ServerPid,
  [string]$LogPath,
  [string]$StatePath
)

$ErrorActionPreference = "SilentlyContinue"

while (Get-Process -Id $LauncherPid -ErrorAction SilentlyContinue) {
  Start-Sleep -Milliseconds 500
}
Start-Sleep -Milliseconds 700

if ($LauncherPid -le 0) { exit 0 }
if (-not (Test-Path $StatePath)) { exit 0 }

$server = Get-Process -Id $ServerPid -ErrorAction SilentlyContinue
if ($server) {
  Stop-Process -Id $ServerPid -Force -ErrorAction SilentlyContinue
}
$line = "[{0}] 退出原因：控制台窗口已关闭（看门狗清理 server pid {1}）" -f `
  (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $ServerPid
try {
  [IO.File]::AppendAllText($LogPath, $line + [Environment]::NewLine,
    (New-Object Text.UTF8Encoding($false)))
} catch { }
Remove-Item $StatePath -Force -ErrorAction SilentlyContinue
exit 0
