<#
  PPT Studio 桌面启动器。

  双击桌面快捷方式时运行；做四件事：
    1) 单实例：已在运行时只把浏览器带到前台，不再起第二个服务；
    2) 端口：从 8765 起找空闲端口（最多顺延到 8799）；
    3) 后台起 `python -m server.app`，轮询 /api/health（≤5s）成功后才打开浏览器；
    4) 保持控制台窗口：“关闭此窗口即退出”——Job Object 保证子进程随窗口一起结束。

  日志：server\store\logs\startup-YYYYMMDD.log
  测试参数：-NoBrowser（不打开浏览器）、-NoWatchdog（不启动清理看门狗）、-Port 起始端口。
#>
[CmdletBinding()]
param(
  [int]$Port = 8765,
  [switch]$NoBrowser,
  [switch]$NoWatchdog,
  [string]$LogDir = ""
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$env:PYTHONIOENCODING = "utf-8"

$Root = Split-Path -Parent $PSScriptRoot
$LogDir = if ($LogDir) { $LogDir } else { Join-Path $Root "server\store\logs" }
New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
$LogPath = Join-Path $LogDir ("startup-" + (Get-Date -Format "yyyyMMdd") + ".log")
$StatePath = Join-Path $Root "server\store\launch.json"
$MutexName = "Local\PPTStudioLauncher"
$PortStart = if ($Port -ge 1024) { $Port } else { 8765 }
$PortEnd = [Math]::Min($PortStart + 34, 8799)

$script:reason = "启动中"
$script:chosenPort = 0
$script:server = $null
$script:stateWritten = $false
$script:ownsMutex = $false
$script:mutex = $null
$script:exitCode = 0

function Write-Log([string]$Message) {
  $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
  try {
    [IO.File]::AppendAllText($LogPath, $line + [Environment]::NewLine,
      (New-Object Text.UTF8Encoding($false)))
  } catch { }
}

function Find-Python {
  $candidates = @(
    $env:PPT_PYTHON,
    (Join-Path $env:USERPROFILE ".dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe"),
    (Get-Command python -ErrorAction SilentlyContinue).Source
  )
  foreach ($candidate in $candidates) {
    if ($candidate -and (Test-Path $candidate)) {
      try {
        & $candidate -c "import pptx" 2>$null
        if ($LASTEXITCODE -eq 0) { return $candidate }
      } catch { }
    }
  }
  throw "找不到可用的 Python（需要 python-pptx）。可设置环境变量 PPT_PYTHON 指向解释器。"
}

function Test-PortFree([int]$Candidate) {
  $listener = $null
  try {
    $listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, $Candidate)
    $listener.Start()
    return $true
  } catch {
    return $false
  } finally {
    if ($listener) { try { $listener.Stop() } catch { } }
  }
}

function Test-Health([string]$BaseUrl) {
  try {
    $response = Invoke-WebRequest -UseBasicParsing -Uri "$BaseUrl/api/health" -TimeoutSec 1
    return ($response.StatusCode -eq 200)
  } catch {
    return $false
  }
}

function Find-RunningServer {
  for ($candidate = 8765; $candidate -le 8799; $candidate++) {
    if (Test-Health ("http://127.0.0.1:" + $candidate)) { return $candidate }
  }
  return 0
}

function Open-Studio([string]$Url) {
  if ($NoBrowser) { return }
  try { Start-Process $Url | Out-Null } catch { Write-Log ("打开浏览器失败：" + $_.Exception.Message) }
}

function Initialize-KillJob {
  # Job Object + KILL_ON_JOB_CLOSE：窗口/宿主无论怎么退出，python 子进程都会被一起回收。
  try {
    if (-not ("PptStudioKillJob" -as [type])) {
      Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public static class PptStudioKillJob
{
    [StructLayout(LayoutKind.Sequential)]
    public struct IO_COUNTERS
    {
        public ulong ReadOperationCount, WriteOperationCount, OtherOperationCount;
        public ulong ReadTransferCount, WriteTransferCount, OtherTransferCount;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct BASIC_LIMIT
    {
        public long PerProcessUserTimeLimit;
        public long PerJobUserTimeLimit;
        public uint LimitFlags;
        public UIntPtr MinimumWorkingSetSize;
        public UIntPtr MaximumWorkingSetSize;
        public uint ActiveProcessLimit;
        public UIntPtr Affinity;
        public uint PriorityClass;
        public uint SchedulingClass;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct EXTENDED_LIMIT
    {
        public BASIC_LIMIT BasicLimitInformation;
        public IO_COUNTERS IoInfo;
        public UIntPtr ProcessMemoryLimit;
        public UIntPtr JobMemoryLimit;
        public UIntPtr PeakProcessMemoryUsed;
        public UIntPtr PeakJobMemoryUsed;
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr CreateJobObject(IntPtr attributes, string name);
    [DllImport("kernel32.dll")]
    private static extern bool SetInformationJobObject(IntPtr job, int infoClass, IntPtr info, uint length);
    [DllImport("kernel32.dll")]
    private static extern bool AssignProcessToJobObject(IntPtr job, IntPtr process);
    [DllImport("kernel32.dll")]
    private static extern IntPtr OpenProcess(uint access, bool inherit, int pid);
    [DllImport("kernel32.dll")]
    private static extern bool CloseHandle(IntPtr handle);

    private static IntPtr job = IntPtr.Zero;

    public static bool Init()
    {
        job = CreateJobObject(IntPtr.Zero, null);
        if (job == IntPtr.Zero) return false;
        var info = new EXTENDED_LIMIT();
        info.BasicLimitInformation.LimitFlags = 0x2000; // JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
        int size = Marshal.SizeOf(info);
        IntPtr buffer = Marshal.AllocHGlobal(size);
        try
        {
            Marshal.StructureToPtr(info, buffer, false);
            return SetInformationJobObject(job, 9, buffer, (uint)size);
        }
        finally { Marshal.FreeHGlobal(buffer); }
    }

    public static bool Assign(IntPtr processHandle)
    {
        return job != IntPtr.Zero && AssignProcessToJobObject(job, processHandle);
    }

    public static bool AssignPid(int pid)
    {
        if (job == IntPtr.Zero) return false;
        // PROCESS_SET_QUOTA(0x0100) | PROCESS_TERMINATE(0x0001) | PROCESS_QUERY_INFORMATION(0x0400)
        IntPtr handle = OpenProcess(0x0501, false, pid);
        if (handle == IntPtr.Zero) return false;
        try { return AssignProcessToJobObject(job, handle); }
        finally { CloseHandle(handle); }
    }
}
'@ -ErrorAction Stop
    }
    return [PptStudioKillJob]::Init()
  } catch {
    Write-Log ("Job Object 不可用：" + $_.Exception.Message)
    return $false
  }
}

try {
  $createdNew = $false
  $script:mutex = New-Object System.Threading.Mutex($true, $MutexName, [ref]$createdNew)
  $script:ownsMutex = [bool]$createdNew
  if (-not $script:ownsMutex) {
    # 第二个实例：等已有服务就绪（最多 5s），然后只打开浏览器。
    $existing = Find-RunningServer
    $deadline = (Get-Date).AddSeconds(5)
    while (-not $existing -and (Get-Date) -lt $deadline) {
      Start-Sleep -Milliseconds 250
      $existing = Find-RunningServer
    }
    if ($existing) {
      $url = "http://127.0.0.1:$existing"
      Write-Log "已有实例运行于 $url，本次只打开浏览器（未启动新进程）"
      Write-Host "PPT Studio 已在运行：$url"
      Open-Studio $url
      $script:reason = "已有实例，仅打开浏览器"
      $script:exitCode = 0
    } else {
      Write-Log "检测到启动锁，但 5 秒内没有可用的健康检查；放弃启动"
      Write-Host "PPT Studio 正在启动或启动失败，请稍后重试（日志：$LogPath）" -ForegroundColor Yellow
      $script:reason = "已有实例但未就绪"
      $script:exitCode = 1
    }
    exit $script:exitCode
  }

  $python = Find-Python
  $chosen = 0
  for ($candidate = $PortStart; $candidate -le $PortEnd; $candidate++) {
    if (Test-PortFree $candidate) { $chosen = $candidate; break }
  }
  if (-not $chosen) {
    throw "端口 $PortStart-$PortEnd 都被占用，无法启动。可关闭占用端口的程序，或用 -Port 指定其他起始端口。"
  }
  $script:chosenPort = $chosen
  Write-Log "启动：端口 $chosen，python=$python，日志目录=$LogDir"

  $env:PYTHONPATH = $Root
  $serverOut = Join-Path $LogDir "server-$chosen.log"
  $serverErr = Join-Path $LogDir "server-$chosen.err.log"
  $startedAt = Get-Date
  $server = Start-Process -FilePath $python `
    -ArgumentList @("-m", "server.app", "--host", "127.0.0.1", "--port", "$chosen") `
    -WorkingDirectory $Root -PassThru -WindowStyle Hidden `
    -RedirectStandardOutput $serverOut -RedirectStandardError $serverErr
  $script:server = $server

  if (Initialize-KillJob) {
    try {
      $assigned = [PptStudioKillJob]::AssignPid($server.Id)
      if (-not $assigned) {
        Write-Log "警告：无法把服务进程加入 Job Object，将依赖退出时主动 kill"
      }
    } catch {
      Write-Log ("警告：AssignProcessToJobObject 失败：" + $_.Exception.Message)
    }
  }

  $url = "http://127.0.0.1:$chosen"
  $healthy = $false
  $deadline = (Get-Date).AddSeconds(5)
  while ((Get-Date) -lt $deadline) {
    if ($server.HasExited) { break }
    if (Test-Health $url) { $healthy = $true; break }
    Start-Sleep -Milliseconds 200
  }
  if (-not $healthy) {
    throw "服务在 5 秒内未就绪（端口 $chosen）。Python 输出见 $serverOut / $serverErr。"
  }
  $elapsed = [int]((Get-Date) - $startedAt).TotalMilliseconds

  $state = @{
    pid         = $server.Id
    launcher_pid = $PID
    port        = $chosen
    url         = $url
    started_at  = (Get-Date).ToString("o")
  } | ConvertTo-Json
  [IO.File]::WriteAllText($StatePath, $state, (New-Object Text.UTF8Encoding($false)))
  $script:stateWritten = $true

  if (-not $NoWatchdog) {
    $watchdog = Join-Path $PSScriptRoot "watch-ppt-studio.ps1"
    if (Test-Path $watchdog) {
      $watchArgs = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$watchdog`"" +
        " -LauncherPid $PID -ServerPid $($server.Id) -LogPath `"$LogPath`" -StatePath `"$StatePath`""
      try {
        Start-Process -FilePath "powershell.exe" -ArgumentList $watchArgs -WindowStyle Hidden | Out-Null
      } catch {
        Write-Log ("启动看门狗失败：" + $_.Exception.Message)
      }
    }
  }

  try { $Host.UI.RawUI.WindowTitle = "PPT Studio" } catch { }
  Write-Host ""
  Write-Host "PPT Studio · 服务运行中 · 关闭此窗口即退出" -ForegroundColor Cyan
  Write-Host "  地址：$url"
  Write-Host "  端口：$chosen（启动耗时 $elapsed ms）"
  Write-Host "  日志：$LogPath"
  Write-Host ""
  Write-Log "启动成功：端口 $chosen，启动耗时 $elapsed ms，地址 $url"
  Open-Studio $url

  $script:reason = "服务进程已退出"
  try {
    $server.WaitForExit()
    if ($server.HasExited) {
      try {
        $script:reason = "服务进程已退出（退出码 $($server.ExitCode)）"
      } catch { }
    }
  } catch {
    Write-Log ("等待服务退出时被中断：" + $_.Exception.Message)
    $script:reason = "用户中断（Ctrl+C/窗口关闭）"
  }
} catch {
  $script:reason = "启动失败：" + $_.Exception.Message
  $script:exitCode = 1
  Write-Log $script:reason
  Write-Host "PPT Studio 启动失败：$($_.Exception.Message)" -ForegroundColor Red
  Write-Host "日志：$LogPath"
} finally {
  if ($script:server -and -not $script:server.HasExited) {
    try { Stop-Process -Id $script:server.Id -Force -ErrorAction SilentlyContinue } catch { }
  }
  if ($script:stateWritten -and (Test-Path $StatePath)) {
    try {
      $stateOnDisk = Get-Content $StatePath -Raw -Encoding UTF8 | ConvertFrom-Json
      if ($stateOnDisk.launcher_pid -eq $PID) { Remove-Item $StatePath -Force -ErrorAction SilentlyContinue }
    } catch { }
  }
  if ($script:ownsMutex -and $script:mutex) {
    try { $script:mutex.ReleaseMutex() } catch { }
    try { $script:mutex.Dispose() } catch { }
  }
  if ($script:chosenPort) {
    Write-Log "退出原因：$($script:reason)；端口 $($script:chosenPort) 已释放"
  } else {
    Write-Log "退出原因：$($script:reason)"
  }
}

exit $script:exitCode
