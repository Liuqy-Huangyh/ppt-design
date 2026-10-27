# PPT Studio 桌面版

把联网版 PPT 服务包装成「双击就用」的桌面应用：一个带图标的快捷方式，双击启动本地服务并自动打开浏览器；关闭控制台窗口即完全退出。

## 1. 安装（一次）

在普通 PowerShell 里运行（不需要管理员）：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\install-desktop.ps1
```

会做两件事：

1. 把 `assets\icon.ico`（自绘多尺寸图标）复制到 `server\store\desktop\icon.ico`；
2. 在桌面创建 **`PPT Studio.lnk`**，指向 `scripts\launch-ppt-studio.ps1`，图标使用上面的 `.ico`。

需要开始菜单也放一份：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\install-desktop.ps1 -StartMenu
```

只看计划不落盘：加 `-WhatIf:$true`（PowerShell 的 `-File` 会把 `-WhatIf:$false` 当字符串传入，脚本已兼容）。
指定其他桌面目录（测试/沙箱用）：`-DesktopDir .\desktop-fallback`。

> 图标可用 `python scripts\make-icon.py` 重新生成（256/128/64/48/32/16 多尺寸）。

## 2. 使用

双击桌面「PPT Studio」：

- 自动从 **8765** 起找空闲端口（最多顺延到 **8799**），`8765` 被占用也不会失败；
- 后台启动 `python -m server.app --host 127.0.0.1 --port <端口>`；
- 轮询 `/api/health`（最多 5s），**确认服务就绪后**才打开默认浏览器；
- 控制台显示：

  ```text
  PPT Studio · 服务运行中 · 关闭此窗口即退出
    地址：http://127.0.0.1:8765
    端口：8765（启动耗时 812 ms）
    日志：...\server\store\logs\startup-20261027.log
  ```

**关闭这个控制台窗口（或 Ctrl+C）即退出**：启动器用 Windows Job Object 把 python 子进程绑定到自己的生命周期，窗口一关，服务进程同步结束、端口释放（另有看门狗兜底清理并补写退出原因）。

单实例：已经运行时再次双击不会起第二个服务，只会把浏览器打开到现有地址；启动器用命名 Mutex + 端口健康探测判断。

## 3. 端口

默认从 `8765` 开始顺延；如果整段 `8765–8799` 都被占用会报错并提示。想换起始端口：

```powershell
# 手动运行启动器时
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\launch-ppt-studio.ps1 -Port 9000

# 或改快捷方式「属性 → 目标」在参数末尾加 -Port 9000
```

## 4. 日志

- 启动器日志：`server\store\logs\startup-YYYYMMDD.log`，包含 **端口、启动耗时、退出原因**；
- 服务本身 stdout/stderr：`server\store\logs\server-<端口>.log` 与 `server-<端口>.err.log`；
- 运行状态（供第二个实例/排错）：`server\store\launch.json`（退出时会删除）。

## 5. 卸载

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\uninstall-desktop.ps1
# 如果安装时用了 -StartMenu，卸载也加
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\uninstall-desktop.ps1 -StartMenu
```

只删除桌面/开始菜单快捷方式、`server\store\desktop\icon.ico` 图标副本；**不动** `server\store` 里的档案、deck、反馈与日志。

## 6. 排错

| 现象 | 处理 |
|---|---|
| 双击后一闪而过 / 没打开浏览器 | 看 `server\store\logs\startup-<日期>.log` 与 `server-<端口>.err.log`；常见原因是 Python 缺 `python-pptx`（脚本会提示设置 `PPT_PYTHON`）。 |
| 提示端口都被占用 | `netstat -ano | findstr 8765` 找到占用进程；或改快捷方式参数 `-Port 9000`。 |
| 已在运行时再次双击没反应 | 设计如此：只把浏览器打开到现有服务；如果浏览器没弹出，手动访问日志里的地址。 |
| 关了窗口但端口仍被占用 | 正常路径会同时结束 python；若系统阻止 Job Object（日志有 “Job Object 不可用”），看门狗会兜底 kill；仍残留时用任务管理器结束对应 `python.exe`。 |
| `visual` 评审提示「未渲染」 | 本机没有可用的 WPS/Office COM：渲染闸门自动降级为结构检查（文字越界/溢出、重叠、字号、留白等仍会检查）。安装 WPS 或 Office 后重新双击启动即可恢复渲染评审。 |
| 快捷方式图标是白纸 | 图标副本丢失：重新运行 `scripts\install-desktop.ps1`；或确认 `server\store\desktop\icon.ico` 存在。 |
| 沙箱/受限环境无法写桌面 | `install-desktop.ps1` 会检测到写桌面被拒，在工作区 `desktop-fallback\PPT Studio.lnk` 生成副本并提示；在普通 PowerShell（非沙箱）重新运行即可装到真实桌面。 |

## 7. 文件清单

| 文件 | 作用 |
|---|---|
| `scripts\launch-ppt-studio.ps1` | 启动器：单实例、端口顺延、健康检查、开浏览器、生命周期管理、日志 |
| `scripts\watch-ppt-studio.ps1` | 看门狗：宿主窗口被关闭时补写退出原因并兜底清理服务进程 |
| `scripts\install-desktop.ps1` | 安装快捷方式 + 复制图标（可选 `-StartMenu`） |
| `scripts\uninstall-desktop.ps1` | 删除快捷方式与图标副本（保留数据） |
| `scripts\make-icon.py` | 用 PIL 生成 `assets\icon.ico`（多尺寸） |
| `assets\icon.ico` | 深蓝底 + 白色 P / 幻灯片图标（256/128/64/48/32/16） |
