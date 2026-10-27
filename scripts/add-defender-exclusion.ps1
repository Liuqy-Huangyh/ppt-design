# 把桌面启动器所在目录加入 Defender/SmartScreen 排除列表（需要管理员）
$ErrorActionPreference = 'Stop'
$Root = 'C:\Users\91125\Desktop\PPT'
$paths = @((Join-Path $Root 'bin'), (Join-Path $Root 'vendor\webview2'))
foreach ($p in $paths) {
    if (Test-Path $p) {
        Add-MpPreference -ExclusionPath $p
        Write-Host "[OK] 已排除：$p"
    }
}
$current = (Get-MpPreference).ExclusionPath
New-Item -ItemType Directory -Force -Path (Join-Path $Root 'server\store') | Out-Null
@("time=$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')", "exclusions:") + ($current | ForEach-Object { "  $_" }) |
    Set-Content -Path (Join-Path $Root 'server\store\mp-exclusion.txt') -Encoding UTF8
Write-Host "完成，结果已写入 server\store\mp-exclusion.txt"
Start-Sleep -Seconds 2
