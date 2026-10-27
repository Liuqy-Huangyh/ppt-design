<#
  PPT 工作区总入口。

  常用：
    .\scripts\ppt.ps1 doctor                       环境自检
    .\scripts\ppt.ps1 web                          启动 PPT Studio 服务（默认 127.0.0.1:8765）
    .\scripts\ppt.ps1 web -Host 0.0.0.0 -Port 8765 局域网可访问
    .\scripts\ppt.ps1 build .\templates\academic-defense.md -Open
    .\scripts\ppt.ps1 build .\templates\editorial-share.md -Theme editorial
    .\scripts\ppt.ps1 html  .\templates\editorial-share.md -Open
    .\scripts\ppt.ps1 qa    .\out\academic-defense.pptx
    .\scripts\ppt.ps1 visual .\out\design-showcase.pptx     # 渲染后逐页视觉/缺陷评审
    .\scripts\ppt.ps1 render .\out\academic-defense.pptx     # 出 PDF + 每页 PNG
    .\scripts\ppt.ps1 md    .\examples\marp-demo.md -Pptx    # marp：md → 原生 pptx
    .\scripts\ppt.ps1 tomd  .\旧PPT.pptx                     # 反解成 Markdown
    .\scripts\ppt.ps1 serve .\templates                      # marp 实时预览
#>
param(
  [Parameter(Position = 0)][string]$Command = "help",
  [Parameter(Position = 1)][string]$Path,
  [string]$Out,
  [ValidateSet("academic", "editorial", "minimal")][string]$Theme,
  [switch]$Open,
  [switch]$Pdf,
  [switch]$Pptx,
  [switch]$Html,
  [switch]$NoPng,
  [int]$Width = 1600,
  [Alias("Host")][string]$Bind = "127.0.0.1",
  [int]$Port = 8765
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$env:PYTHONIOENCODING = "utf-8"
$Root = Split-Path -Parent $PSScriptRoot
$Kit = Join-Path $Root "kit"
$NodeDir = "C:\Users\$env:USERNAME\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\node\bin"
$NodeExe = Join-Path $NodeDir "node.exe"
$BundledPython = "C:\Users\$env:USERNAME\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe"
$Edge = @(
  "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
  "C:\Program Files\Microsoft\Edge\Application\msedge.exe"
) | Where-Object { Test-Path $_ } | Select-Object -First 1

function Get-Python {
  $cands = @($env:PPT_PYTHON, $BundledPython, (Get-Command python -ErrorAction SilentlyContinue).Source)
  foreach ($c in $cands) {
    if ($c -and (Test-Path $c)) {
      & $c -c "import pptx" 2>$null
      if ($LASTEXITCODE -eq 0) { return $c }
    }
  }
  throw "找不到可用的 Python（需要 python-pptx）。可设置环境变量 PPT_PYTHON 指向解释器。"
}

$Python = Get-Python

function Invoke-Kit {
  param([string[]]$KitArgs)
  $env:PYTHONPATH = $Kit
  & $Python -m pptxkit @KitArgs
  if ($LASTEXITCODE -ne 0) { throw "pptxkit 执行失败（退出码 $LASTEXITCODE）" }
}

function Resolve-Out([string]$default) {
  if ($Out) { return (Join-Path (Get-Location) $Out) }
  return $default
}

function Assert-WritableTarget([string]$target) {
  # 测试样例固定在 tests\fixtures\，任何构建命令都不允许覆盖（评审 F2）。
  $fixtureDir = [IO.Path]::GetFullPath((Join-Path $Root "tests\fixtures"))
  $targetFull = [IO.Path]::GetFullPath($target)
  $sep = [IO.Path]::DirectorySeparatorChar
  if ($targetFull -eq $fixtureDir -or
      $targetFull.StartsWith($fixtureDir + $sep, [StringComparison]::OrdinalIgnoreCase)) {
    $stem = [IO.Path]::GetFileNameWithoutExtension($Path)
    throw "拒绝写入只读样例目录 tests\fixtures\：$targetFull。请用 -Out 指定其他输出路径（例如 .\out\$stem.pptx）。"
  }
}

function Open-File([string]$p) {
  if ($Open -and $p -and (Test-Path $p)) { Start-Process $p | Out-Null }
}

function Invoke-Marp([string]$Mode) {
  $marp = Join-Path $Root "node_modules\.bin\marp.CMD"
  if (-not (Test-Path $marp)) { throw "marp-cli 未安装，请先在 $Root 运行 pnpm install" }
  $src = (Resolve-Path $Path).Path
  $base = [IO.Path]::GetFileNameWithoutExtension($src)
  $dir = Join-Path $Root "out"
  New-Item -ItemType Directory -Force -Path $dir | Out-Null
  $env:PATH = "$NodeDir;$env:PATH"
  $marpArgs = @($src, "--allow-local-files")
  $themeSet = Join-Path $Root "templates\marp-theme.css"
  if (Test-Path $themeSet) { $marpArgs += @("--theme-set", $themeSet) }
  switch ($Mode) {
    "pptx" { $marpArgs += @("--pptx", "--output", (Join-Path $dir "$base.pptx")) }
    "pdf" { $marpArgs += @("--pdf", "--output", (Join-Path $dir "$base.pdf")) }
    "html" { $marpArgs += @("--html", "--output", (Join-Path $dir "$base.marp.html")) }
  }
  if ($Edge) { $marpArgs += @("--browser", "edge") }
  & $marp @marpArgs
  if ($LASTEXITCODE -ne 0) { throw "marp 执行失败（退出码 $LASTEXITCODE）" }
  Write-Host "✓ 输出目录：$dir"
}

switch ($Command.ToLower()) {
  "doctor" {
    Write-Host "工作区：$Root"
    Write-Host "Python：$Python"
    & $Python -c "import pptx, PIL, sys; print('  python', sys.version.split()[0], '| python-pptx', pptx.__version__, '| Pillow', PIL.__version__)"
    $node = if (Test-Path $NodeExe) { $NodeExe } else { (Get-Command node -ErrorAction SilentlyContinue).Source }
    if ($node) { Write-Host "Node：$node"; & $node -v | ForEach-Object { Write-Host "  $_" } } else { Write-Host "Node：未找到（marp/slidev 不可用）" }
    $pkgs = @("@marp-team/marp-cli", "pptxgenjs", "reveal.js")
    foreach ($p in $pkgs) {
      $mark = if (Test-Path (Join-Path $Root "node_modules\$p")) { "✓" } else { "✗" }
      Write-Host "  $mark $p"
    }
    $edgeLabel = if ($Edge) { $Edge } else { "未找到（marp 导出 PDF/PPTX 需要浏览器）" }
    Write-Host ("Edge： " + $edgeLabel)
    $ppt = $null
    try { $ppt = New-Object -ComObject PowerPoint.Application; Write-Host "PowerPoint：✓ 可用（可导出 PDF/PNG）" } catch { Write-Host "PowerPoint：✗ 未检测到 COM 接口（PDF/PNG 渲染不可用）" }
    if ($ppt) { $ppt.Quit() | Out-Null; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) }
    Write-Host "字体目录：$env:WINDIR\Fonts"
  }
  "build" {
    if (-not $Path) { throw "用法：ppt.ps1 build <文件.md|.json> [-Theme academic|editorial|minimal] [-Out 输出.pptx] [-Open]" }
    $src = (Resolve-Path $Path).Path
    $target = Resolve-Out ((Join-Path $Root ("out\" + [IO.Path]::GetFileNameWithoutExtension($src) + ".pptx")))
    Assert-WritableTarget $target
    $a = @("build", $src, "-o", $target)
    if ($Theme) { $a += @("--theme", $Theme) }
    Invoke-Kit -KitArgs $a
    Open-File $target
  }
  "html" {
    if (-not $Path) { throw "用法：ppt.ps1 html <文件.md> [-Theme ...] [-Out 输出.html] [-Open]" }
    $src = (Resolve-Path $Path).Path
    $target = Resolve-Out ((Join-Path $Root ("out\" + [IO.Path]::GetFileNameWithoutExtension($src) + ".html")))
    Assert-WritableTarget $target
    $a = @("html", $src, "-o", $target)
    if ($Theme) { $a += @("--theme", $Theme) }
    Invoke-Kit -KitArgs $a
    Open-File $target
  }
  "qa" {
    if (-not $Path) { throw "用法：ppt.ps1 qa <文件.pptx>" }
    Invoke-Kit -KitArgs @("qa", (Resolve-Path $Path).Path)
  }
  "visual" {
    # 渲染后评审：先 WPS 导 PNG，再按「视觉 / 缺陷」两条闸门逐页判 PASS / NEEDS_REVISION
    if (-not $Path) { throw "用法：ppt.ps1 visual <文件.pptx> [-Out 输出目录] [-Width 1600]" }
    $src = (Resolve-Path $Path).Path
    $stem = [IO.Path]::GetFileNameWithoutExtension($src)
    $visualDir = if ($Out) { Join-Path (Get-Location) $Out } else { Join-Path $Root "out\$stem-visual" }
    New-Item -ItemType Directory -Force -Path $visualDir | Out-Null
    $pngDir = Join-Path $visualDir "$stem-png"
    $rendered = $false
    try {
      & (Join-Path $PSScriptRoot "render-pptx.ps1") -Path $src -OutDir $visualDir -NoPdf -Width $Width
      if ($LASTEXITCODE -eq 0 -and (Test-Path $pngDir) -and
          (Get-ChildItem $pngDir -Filter *.png -ErrorAction SilentlyContinue)) {
        $rendered = $true
      }
    } catch {
      Write-Host "渲染失败，降级为结构检查：$($_.Exception.Message)"
    }
    if (-not $rendered) {
      Write-Host "提示：未渲染（缺少 Office/WPS 或渲染失败），仅执行结构检查。"
    }
    $env:PYTHONPATH = $Kit
    $cliArgs = @("-m", "pptxkit.inspect", "--visual", $src)
    if ($rendered) { $cliArgs += @("--png-dir", $pngDir) }
    & $Python @cliArgs
    $code = $LASTEXITCODE
    if ($code -ne 0) { exit $code }
  }
  "tomd" {
    if (-not $Path) { throw "用法：ppt.ps1 tomd <文件.pptx> [-Out 输出.md]" }
    $src = (Resolve-Path $Path).Path
    $target = Resolve-Out ([IO.Path]::ChangeExtension($src, ".md"))
    Invoke-Kit -KitArgs @("to-md", $src, "-o", $target)
  }
  "themes" { Invoke-Kit -KitArgs @("themes") }
  "profile" {
    # 案例风格档案：pptx 读几何/配色/字体，png|jpg 只读版式骨架；打印档案 JSON
    if (-not $Path) { throw "用法：ppt.ps1 profile <文件.pptx|.png|.jpg> [-Out 输出.json]" }
    $src = (Resolve-Path $Path).Path
    $a = @("profile", $src)
    if ($Out) { $a += @("-o", (Join-Path (Get-Location) $Out)) }
    Invoke-Kit -KitArgs $a
  }
  "render" {
    if (-not $Path) { throw "用法：ppt.ps1 render <文件.pptx> [-NoPng] [-Width 1920]" }
    $src = (Resolve-Path $Path).Path
    $a = @{ Path = $src; Width = $Width }
    if ($NoPng) { $a["NoPng"] = $true }
    & (Join-Path $PSScriptRoot "render-pptx.ps1") @a
  }
  "md" {
    if (-not $Path) { throw "用法：ppt.ps1 md <文件.md> [-Pptx|-Pdf|-Html]" }
    $mode = if ($Pptx) { "pptx" } elseif ($Pdf) { "pdf" } else { "html" }
    Invoke-Marp $mode
    Open-File (Join-Path $Root ("out\" + [IO.Path]::GetFileNameWithoutExtension((Resolve-Path $Path).Path) + "." + $(if ($mode -eq "pptx") { "pptx" } elseif ($mode -eq "pdf") { "pdf" } else { "marp.html" })))
  }
  "serve" {
    $marp = Join-Path $Root "node_modules\@marp-team\marp-cli\marp-cli.js"
    $target = if ($Path) { (Resolve-Path $Path).Path } else { $Root }
    $env:PATH = "$NodeDir;$env:PATH"
    Write-Host "Marp 预览服务启动中……（Ctrl+C 退出）"
    & $NodeExe $marp -s $target --allow-local-files
  }
  "web" {
    # 浏览器可用的 PPT 制作服务：生成 / 案例库 / 反馈
    $localUrl = "http://127.0.0.1:$Port"
    Write-Host "PPT Studio 服务启动中……（Ctrl+C 停止）"
    Write-Host "  本机访问：$localUrl"
    if ($Bind -in @("0.0.0.0", "::", "*")) {
      Write-Host "  局域网访问：http://<本机局域网 IP>:$Port （用 ipconfig 查看本机 IP）"
    } elseif ($Bind -ne "127.0.0.1") {
      Write-Host "  访问地址：http://$($Bind):$Port"
    }
    $env:PYTHONPATH = $Root
    & $Python -m server.app --host $Bind --port $Port
    if ($LASTEXITCODE -ne 0) { throw "PPT Studio 服务退出码 $LASTEXITCODE" }
  }
  "open" {
    if (-not $Path) { throw "用法：ppt.ps1 open <文件>" }
    Start-Process (Resolve-Path $Path).Path | Out-Null
  }
  default {
    Write-Host @"
PPT 工作区命令

  build   <md|json>  Markdown → 可编辑 .pptx        [-Theme academic|editorial|minimal] [-Out] [-Open]
  html    <md>       Markdown → 单文件 HTML 演示    [-Theme ...] [-Out] [-Open]
  web               启动 PPT Studio 服务（浏览器生成/案例/反馈） [-Host 0.0.0.0] [-Port 8765]
  md      <md>       Markdown → pptx/pdf/html（marp 引擎，原生形状） [-Pptx|-Pdf]
  render  <pptx>     PPTX → PDF + 每页 PNG（走本机 WPS/PowerPoint）
  visual  <pptx>     渲染后逐页评审：视觉焦点/留白 + 越界溢出/重叠/字号/图片拉伸
  qa      <pptx>     体检：字数、溢出、字体、备注
  tomd    <pptx>     反解成 Markdown（图片导出到 media/）
  serve   [dir]      Marp 实时预览（改 md 自动刷新）
  themes             列出内置主题
  doctor             环境自检

示例
  .\scripts\ppt.ps1 build .\templates\academic-defense.md -Open
  .\scripts\ppt.ps1 web
  .\scripts\ppt.ps1 web -Host 0.0.0.0 -Port 8765
  .\scripts\ppt.ps1 visual .\out\design-showcase.pptx
  .\scripts\ppt.ps1 render .\out\academic-defense.pptx
"@
  }
}
