<#
  用本机 WPS 或 PowerPoint 把 .pptx 导出成 PDF 和每页 PNG（不上传任何文件）。
  默认先用 WPS（COM 自动化更稳，不会弹窗），不行再回退 PowerPoint。

  用法：
    .\render-pptx.ps1 -Path .\out\demo.pptx
    .\render-pptx.ps1 -Path .\out\demo.pptx -NoPng -Width 1920
    .\render-pptx.ps1 -Path .\out\demo.pptx -Engine PowerPoint.Application
#>
param(
  [Parameter(Mandatory = $true)][string]$Path,
  [string]$OutDir,
  [switch]$NoPng,
  [switch]$NoPdf,
  [int]$Width = 1600,
  [string[]]$Engine = @("Kwpp.Application", "PowerPoint.Application")
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [Text.Encoding]::UTF8

$src = (Resolve-Path $Path).Path
if (-not $OutDir) { $OutDir = Split-Path $src -Parent }
$base = [IO.Path]::GetFileNameWithoutExtension($src)
$pdfPath = Join-Path $OutDir "$base.pdf"
$pngDir = Join-Path $OutDir "$base-png"

function Export-WithEngine([string]$progId) {
  $app = $null
  $pres = $null
  try {
    $app = New-Object -ComObject $progId
    try { $app.Visible = $true } catch {}
    $pres = $app.Presentations.Open($src, -1, 0, -1)

    # 有些版本刚 Open 时属性还读不到，重试几次
    $count = 0
    for ($t = 0; $t -lt 16 -and $count -lt 1; $t++) {
      Start-Sleep -Milliseconds 200
      try { $count = [int]$pres.Slides.Count } catch { $count = 0 }
    }
    if ($count -lt 1) { throw "打开后读不到页数" }

    Write-Host "引擎：$progId ｜ $count 页"

    if (-not $NoPdf) {
      if (Test-Path $pdfPath) { Remove-Item $pdfPath -Force }
      $pres.SaveAs($pdfPath, 32)          # ppSaveAsPDF
      Write-Host "✓ PDF：$pdfPath"
    }

    if (-not $NoPng) {
      New-Item -ItemType Directory -Force -Path $pngDir | Out-Null
      Get-ChildItem $pngDir -Filter *.png -ErrorAction SilentlyContinue | Remove-Item -Force
      $h = [int]($Width * 9 / 16)
      $ok = $false
      try {
        $pres.Export($pngDir, "PNG", $Width, $h)
        $ok = (Get-ChildItem $pngDir -Filter *.png -ErrorAction SilentlyContinue).Count -gt 0
      } catch { $ok = $false }
      if (-not $ok) {
        for ($i = 1; $i -le $count; $i++) {
          $file = Join-Path $pngDir ("slide-{0:D2}.png" -f $i)
          $pres.Slides.Item($i).Export($file, "PNG", $Width, $h)
        }
      }
      # 统一命名：幻灯片1.PNG → slide-01.png
      $files = @(Get-ChildItem $pngDir -Filter *.png -ErrorAction SilentlyContinue)
      if ($files.Count -gt 0 -and -not ($files | Where-Object { $_.Name -like "slide-*" })) {
        foreach ($f in $files) {
          $num = ($f.BaseName -replace '\D', '')
          if ($num) {
            $target = Join-Path $pngDir ("slide-{0:D2}.png" -f [int]$num)
            if ($f.FullName -ne $target) { Move-Item $f.FullName $target -Force }
          }
        }
      }
      $n = (Get-ChildItem $pngDir -Filter *.png -ErrorAction SilentlyContinue).Count
      Write-Host "✓ 图片：$pngDir（$n 张，宽 $Width px）"
    }
    return $true
  }
  catch {
    Write-Verbose ("$progId 不可用：" + $_.Exception.Message)
    return $false
  }
  finally {
    if ($pres) { try { $pres.Close() } catch {} }
    if ($app) {
      try { $app.Quit() } catch {}
      try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
  }
}

$done = $false
foreach ($progId in $Engine) {
  if (Export-WithEngine $progId) { $done = $true; break }
}

if (-not $done) {
  throw "WPS / PowerPoint 都没能通过 COM 打开这个文件。可以手动打开后用「另存为 PDF」导出。"
}
