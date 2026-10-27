# PPT 工作区

本地做 PPT 的一体化工作台：**一份 Markdown**，可以同时产出「可编辑的 .pptx」「网页演示」「PDF 和每页 PNG」。
不联网、不上传、不依赖 Office 之外的服务；在这台机器上已经装好并跑通。

```powershell
# 1) 写一页 Markdown（或直接改 templates 里的模板）
# 2) 生成可编辑 PPTX（显式 -Out，避免覆盖任何既有产物）
.\ppt.cmd build .\templates\academic-defense.md -Out .\out\答辩示例.pptx -Open

# 3) 同时想要网页版 / 设计感版本
.\ppt.cmd html .\templates\editorial-share.md -Open

# 4) 导出 PDF + 每页图片（给别人预览、或者在微信里发）
.\ppt.cmd render .\out\答辩示例.pptx
```

---

## 1. 为什么是这套组合

| 需求 | 方案 | 产物 |
|---|---|---|
| 交付后还要改、领导要动 | **pptxkit**（python-pptx 驱动） | 原生 .pptx，形状/文字/表格都能编辑 |
| 好看、有设计感、线上分享 | **pptxkit html** / Marp / reveal.js | 单文件 HTML，键盘翻页，可托管 |
| 论文、答辩、学术汇报 | **academic 主题** + refs/timeline/表格版式 | .pptx + PDF |
| 已有旧 PPT 想重排 | **tomd** 反解成 Markdown | .md + 图片 |
| 图表/数据驱动批量生成 | python-pptx 脚本 / PptxGenJS | 每次都能重跑 |

一句话：**内容写在 Markdown 里，版式交给主题，导出交给命令**。改内容不用碰排版，换主题不用改内容。

---

## 2. 目录结构

```
PPT/
├─ ppt.cmd                  ← 总入口（双击/命令行都行，已绕过执行策略限制）
├─ scripts/
│   ├─ ppt.ps1              命令分发：build / html / md / render / qa / tomd / serve / doctor
│   ├─ render-pptx.ps1      pptx → PDF + 每页 PNG（走本机 WPS/PowerPoint COM）
│   ├─ make-assets.py       生成模板示意图（可替换成自己的图）
│   └─ contact-sheet.py     把导出的 PNG 拼成一张总览图，用来检查整份 deck
├─ kit/pptxkit/             Python 工具包（零第三方依赖，只用内置 python-pptx）
│   ├─ themes.py            三套主题：academic / editorial / minimal
│   ├─ primitives.py        绘图原语：文本框、色块、图片、表格、页码
│   ├─ layouts.py           11 种版式
│   ├─ markdown.py          Markdown → 幻灯片结构（front matter / 分页 / 备注）
│   ├─ builder.py           → .pptx
│   ├─ html_deck.py         → 单文件 HTML 演示
│   ├─ inspect.py           PPTX 体检 + 反解 Markdown
│   └─ cli.py               python -m pptxkit
├─ templates/
│   ├─ academic-defense.md  论文答辩模板（17 页，含表格/图/引用/参考文献）
│   ├─ editorial-share.md   杂志风分享模板（12 页）
│   ├─ business-minimal.md  商务复盘模板（9 页）
│   ├─ marp-demo.md         Marp 示例（md → pptx/pdf/html）
│   └─ marp-theme.css       Marp 主题（与 academic 同配色）
├─ assets/figs/             模板示意图（Pillow 生成）
├─ out/                     所有产物：pptx / pdf / html / *-png/
├─ node_modules/            marp-cli、pptxgenjs、reveal.js（已安装）
└─ README.md
```

---

## 3. 四条工作流

### A. Markdown → 可编辑 PPTX（主力）

```powershell
.\ppt.cmd build .\templates\academic-defense.md -Out .\out\答辩示例.pptx   # 显式 -Out；默认目标也是 out\
.\ppt.cmd build .\我的稿子.md -Theme editorial -Out .\out\分享版.pptx -Open
```

> 测试用的样例固定在 `tests\fixtures\`（只读），`build` / `html` 拒绝把输出写进该目录（评审 F2）。
> 平时构建请始终用 `-Out` 指定输出，避免和样例或其他产物重名。

会自动：补封面、按内容推断版式、超出就自动降字号、把 `::: notes` 写进演讲者备注、加页码和页脚。

### B. Markdown → 单文件 HTML 演示（设计感 / 线上分享）

```powershell
.\ppt.cmd html .\templates\editorial-share.md -Open
```

- 生成一个 HTML，图片字体全在本地，可直接丢到任何静态托管（或发给别人双击）
- 快捷键：`← →` 翻页、`O` 缩略图总览、`S` 讲者备注、`F` 全屏、`P` 打印成 PDF、`?` 帮助
- 地址栏 `#/3` 直接定位到第 3 页

### C. 已有 PPT 反解 → 改内容 → 重新出片

```powershell
.\ppt.cmd tomd .\旧PPT.pptx            # → 旧PPT.md + media/ 里的图片
# 改 Markdown，然后
.\ppt.cmd build .\旧PPT.md -Theme minimal -Out .\out\旧PPT-简约.pptx
```

反解会保留：标题、要点层级、表格、图片（导出到 `media/`）、演讲者备注。

### D. Marp（生态路线，可选）

```powershell
.\ppt.cmd md .\templates\marp-demo.md -Html     # 已验证可用
.\ppt.cmd md .\templates\marp-demo.md -Pptx     # md → 原生 pptx（需要浏览器）
```

> Marp 的 pptx/pdf 导出要靠无头浏览器渲染。如果报 `spawn EPERM`（受限/沙箱环境常见），改用 `build` 出 pptx，或在本机普通 PowerShell 窗口里运行。

---

## 4. Markdown 写法

```markdown
---
title: 论文题目
subtitle: 副标题
author: 张三
org: 某某大学
date: 2026 年 10 月
theme: academic          # academic | editorial | minimal
kicker: 博士学位论文答辩
footer: 论文短标题（每页左下角）
---

[01]                     # 章节页的大号编号
# 研究背景与问题          # 一页的开始
## 检测需求正在下沉        # 本页副标题（章节页会用到）

---

# 研究背景

- **需求侧**：工业质检要求毫秒级响应
- 普通要点
  - 二级要点（缩进两格）

| 方法 | mAP | 参数量 |
|---|---|---|
| YOLOv8-s | 44.9 | 11.2 |

![图 1　整体框架](../assets/figs/framework.png)

> 一句引言
> —— 出处

::: notes
这里是演讲者备注，只会出现在备注栏和 HTML 的 S 键面板里。
:::
```

**版式控制**（不写就自动推断）：

```markdown
<!-- layout: section -->     章节页          <!-- layout: two-col -->   两栏
<!-- layout: stats -->       三个数字卡片     <!-- layout: timeline -->  流程
<!-- layout: quote -->       引用            <!-- layout: refs -->      参考文献
<!-- layout: table -->       表格            <!-- layout: image -->     图片
<!-- layout: bullets -->     要点            <!-- layout: end -->       致谢页
```

两栏用 `### 左栏标题` / `### 右栏标题` 切换；数字卡片用 `- 38 FPS | 实时帧率 | 备注`；流程用 `- 名称 | 说明`。

---

## 5. 三套主题

| 主题 | 气质 | 适合 |
|---|---|---|
| `academic` | 白底 + 深蓝 + 金色短线，克制、信息密度高 | 论文答辩、组会、基金汇报 |
| `editorial` | 米白 + 克莱因蓝 + 宋体大标题 + 线框圆 | 公开分享、设计感、线上传播 |
| `minimal` | 纯白 + 一点橙，留白最多 | 产品/业务复盘、路演、正式场合 |

换主题只改 front matter 的 `theme:` 或命令加 `-Theme`，内容一个字都不用动。
字体固定用系统自带的「微软雅黑 / 宋体」，换机器不会跑版。

---

## 6. 命令速查

| 命令 | 作用 |
|---|---|
| `.\ppt.cmd build <md>` | Markdown → 可编辑 .pptx（`-Theme` `-Out` `-Open`） |
| `.\ppt.cmd html <md>` | Markdown → 单文件 HTML 演示 |
| `.\ppt.cmd md <md> -Pptx/-Pdf/-Html` | 用 Marp 引擎转换 |
| `.\ppt.cmd render <pptx>` | PPTX → PDF + 每页 PNG（`-Width 1920`） |
| `.\ppt.cmd qa <pptx>` | 体检：字数、溢出风险、字体缺失、无备注页 |
| `.\ppt.cmd tomd <pptx>` | PPTX → Markdown（图片导出到 media/） |
| `.\ppt.cmd serve <dir>` | Marp 实时预览，改 md 自动刷新 |
| `.\ppt.cmd themes` | 列出内置主题 |
| `.\ppt.cmd doctor` | 环境自检（Python / Node / 组件 / 浏览器 / Office） |

底层也可以直接用 Python：

```powershell
$env:PYTHONPATH = ".\kit"
python -m pptxkit build .\templates\academic-defense.md -o .\out\demo.pptx
```

---

## 7. 环境说明

| 组件 | 版本 / 位置 | 用途 |
|---|---|---|
| Python | 3.12（随 DSH 运行时） | 跑 pptxkit |
| python-pptx | 1.0.2（已内置） | 生成/编辑 .pptx |
| Pillow | 12.3.0 | 图片处理、示意图、总览拼图 |
| Node | 24.x（随 DSH 运行时） | 跑 marp / pptxgenjs |
| @marp-team/marp-cli | 4.5.1 | Markdown → pptx/pdf/html |
| pptxgenjs | 4.0.1 | JS 侧生成 pptx（工作流 C 的备选） |
| reveal.js | 6.0.2 | 需要自定义交互页面时直接用 |
| WPS Office 12.1 / Office LTSC 2021 | 本机 | 导出 PDF/PNG（render 命令自动挑能用的那个） |

> 脚本已按 **UTF-8 BOM** 保存，用编辑器改 `scripts/*.ps1` 时请保留 BOM，否则中文会乱码。

---

## 8. 常见问题

**导出时报「文件被占用 / Permission denied」**
`out\xxx.pptx` 正被 PowerPoint / WPS 打开。关掉再生成，或加 `-Out` 换个文件名。

**`render` 报找不到 COM**
本机没有 WPS/PowerPoint 时无法渲染。可手动打开文件「另存为 PDF」，或只交付 .pptx。

**中文变成方块 / 版式跑位**
目标机器缺「微软雅黑」。这两个字体是 Windows 自带，Mac 上建议换成「苹方」，改 `kit/pptxkit/themes.py` 里的字体常量即可。

**marp 导出 pptx 报 `spawn EPERM`**
受限环境不允许起无头浏览器。用 `build` 出 pptx，或在普通 PowerShell 窗口里跑。

**想加自己的模板**
把校徽、母版配色写进 `themes.py` 的 `THEMES`；或在 `layouts.py` 里加一个版式函数并注册到 `LAYOUTS`，Markdown 里就能用 `<!-- layout: 你的名字 -->`。

---

## 9. 后续可升级（需要联网时装）

```powershell
# Slidev：Vue 组件化 + 大量社区主题，适合做技术分享
pnpm add -D @slidev/cli @slidev/theme-default

# reveal.js 主题与插件（已装 reveal.js，可直接在 HTML 里引 node_modules/reveal.js/dist）
# Marp 社区主题：把 css 放进 templates/，用 --theme-set 加载
```

当前这套已经覆盖：**可编辑交付、设计感/线上分享、学术答辩**三个主场景，且不依赖网络。

---

## 10. 联网服务（PPT Studio）

```powershell
.\ppt.cmd web                          # 127.0.0.1:8765
.\ppt.cmd web -Host 0.0.0.0 -Port 8765 # 局域网（启动会警告：仅可信网络）
```

浏览器打开后三条链路：生成（主题/提示词 → 可编辑 pptx + 网页版）→ 上传自己的 PPT 当案例（提取配色/字体/密度 = 风格档案）→ 打分反馈（1–5 星，自动调密度与标题比例）。

API 见 docs/API.md，接口契约见 docs/CONTRACT.md；`GET /api/health` 的 llm 字段表示是否接了大模型（未配置 `PPT_LLM_*` 时纯离线可用）。

## 11. 字体与设计系统

```powershell
python scripts/install-fonts.py          # 安装思源黑体/思源宋体/Inter（用户级，免管理员）
python scripts/install-fonts.py --check  # 查看已装
```

- 字体：Noto Sans SC / Noto Serif SC / Inter（OFL 许可，可商用）
- 设计系统：kit/pptxkit/design.py —— OKLCH 配色推导、1.25 字号阶梯、12 栏网格、WCAG 对比度
- 样张：scripts/design-showcase.py → out/design-showcase.pptx（6 页编排对照）

> 视觉升级目前只用在样张与主题字体上；把网格/字号接进 11 个版式（cover/section/stats/quote/timeline）是下一迭代的第一项。

## 12. 验收与信任

- 四套测试：tests/test_server.py（26）、test_profile.py（39）、test_e2e.py（7）、test_web.py（29）
- 三轮评审：F1 任意图片读取、F2 测试基线漂移、F11 档案写入非原子并发 500 —— 全部修复并复验
- 当前字节的复验结论见 docs/复验报告.md（含 sha256 绑定），评审全文见 docs/评审结论.md
