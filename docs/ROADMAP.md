# 改进路线（Phase 2 候选）— 调研于 2026-10-27

> 结论先行：**不要为了"AI 感"堆依赖**。当前系统的短板不是"不会生成"，而是"案例理解太浅"（只学了配色/字体/密度）和"迭代信号太粗"（只有 1–5 星）。
> Phase 2 优先做四件**纯自研、零新依赖**的事，把"色彩迁移"升级成"版式骨架迁移 + 可度量的质量分"；HTML→PPTX、图片→可编辑这类能力先做技术验证再决定。

## 1. 候选项目盘点

| 项目 | Star | 许可 | 能补什么 | 结论 |
|---|---|---|---|---|
| [para-lost/AutoPresent](https://github.com/para-lost/AutoPresent)（CVPR 2025） | – | 研究代码 | 从零"设计"版式：先规划视觉结构再渲染，正好对应我们的 11 种版式不够用的问题 | **参考思路**，不直接引依赖 |
| [atharva9167j/dom-to-pptx](https://github.com/atharva9167j/dom-to-pptx) | 383 | MIT | 浏览器端 DOM → 真可编辑 pptx（保留渐变/圆角/阴影），能把我们的 HTML deck 变成可交付 pptx | P1 验证 |
| [Hasasasa/html-to-editable-pptx](https://github.com/Hasasasa/html-to-editable-pptx) | 93 | MIT | 同上，Python CLI + agent skill，文本保持原生文本框 | P1 备选 |
| [pipipi-pikachu/pptxtojson](https://github.com/pipipi-pikachu/pptxtojson) | 455 | MIT | pptx → 结构化 JSON，解析力比 python-pptx 高一档（元素几何、主题、备注） | **自研替代**：直接用 python-pptx 读更多属性即可，不必引 JS 依赖 |
| [iOfficeAI/OfficeCLI](https://github.com/iOfficeAI/OfficeCLI) | 33k | 待核 | 声称不装 Office 也能创建/检查/渲染/自动化 PPTX | P1 验证（可替掉对 WPS/PowerPoint COM 的依赖） |
| [ningzimu/image-to-editable-ppt-skill](https://github.com/ningzimu/image-to-editable-ppt-skill) | 3.1k | 待核 | 图片/PDF → 可编辑 deck（OCR + 重排） | P2，依赖重 |
| [icip-cas/PPTAgent](https://github.com/icip-cas/PPTAgent) | 5.2k | 待核 | 反思式生成：生成→自检→重写的循环 | **借鉴**其"自检-重写"结构，不引代码 |
| [DECKBench](https://dl.acm.org/doi/10.1145/3770855.3817525) / [SlidesGen-Bench](https://ar5iv.labs.arxiv.org/html/2601.09487) | – | 论文 | 幻灯片质量的量化指标：信息密度、版式一致性、图文比例、可读性 | **采用指标口径**，自研打分器 |

## 2. P0：Phase 2 立即做（零新依赖，2–3 个任务）

### P0-1 案例理解升级：从"颜色字体"到"版式骨架"

现在 `analyze()` 只取主题配色、字体、字数、要点数。补四项**几何指标**（python-pptx 直接可读）：

- 标题位置比 `title_top / slide_h`、左边距比 `margin_left / slide_w`（决定"标题在天上还是居中"）
- 每页平均形状数与图片面积占比（决定"图文型"还是"纯文字型"）
- 正文字号阶梯：`{h1, h2, body}` 的比例（决定层级感强弱）
- 强调色使用频率：强调色出现在多少个形状上（决定"克制"还是"活泼"）

→ 产出 `profile.layout` 字段；`apply()` 时据此选**版式族**（偏学术的白底左对齐 / 偏杂志的居中大字 / 偏商务的窄栏），而不只是换色。

### P0-2 案例检索：给提示词找"最像的那份案例"

纯标准库 TF-IDF（中文按 2-gram 切）在 `profiles/*.json` 的 `metrics + outline` 上建索引：

```python
def retrieve(prompt: str, profiles: list[dict], k: int = 3) -> list[tuple[dict, float]]
```

生成时若未指定 `profile_id`，自动取 top-1 并回一句 `notes: "自动匹配案例：<name> (0.72)"`。
作用：让"上传多份案例"真正有用，而不是只能手动选一份。

### P0-3 客观质量分：把 1–5 星变成可解释信号

新增 `kit/pptxkit/quality.py`，输出 0–100 分 + 四项子分：

| 维度 | 算法（可复现） |
|---|---|
| 信息密度 | 每页字数落在 40–140 字得分最高，超出线性衰减 |
| 版式一致性 | 全 deck 标题字号种类数、左边距方差；越少越高分 |
| 配色克制 | 不同色相数 ≤ 4 得满分；强调色占比 5%–15% 最佳 |
| 可读性 | 最小字号 ≥ 14pt、行距 ≥ 1.2、对比度（WCAG）达 AA |

接入点：`/api/decks/{id}/feedback` 允许只传 `{"auto": true}`，用打分器自动产出 rating，
`learn()` 用**子分最低的那一项**决定调参方向（比"一刀切降密度"精确）。
同时 `qa` 命令打印这四项分数，让用户看见"为什么不好看"。

### P0-4 大纲结构模板化

`outline.py` 现在按 kind 选结构。再补一层：从**案例反解出的页序**（`profile.outline`：每页 layout + 标题动词）里抽"叙事骨架"，
生成时按骨架排页（问题→证据→方案→收益）。这比固定模板像"学了这个人的讲法"。

## 3. P1：需要 spike（各 0.5–1 天验证，不通过就放弃）

1. **HTML deck → 可编辑 PPTX**（dom-to-pptx / html-to-editable-pptx）
   - 价值：网页版的"设计感"能直接变成可交付 pptx，不用两套维护
   - 风险：需要无头浏览器；沙箱里 `spawn EPERM`，只能本机验证
   - 验收：把 `out/editorial-share.html` 转成 pptx，文字仍是文本框（不是图片），配色不跑
2. **免 Office 渲染**（OfficeCLI）
   - 价值：去掉 WPS/PowerPoint COM 依赖，服务器/CI 也能出 PDF/PNG
   - 风险：许可证、二进制体积、渲染一致性
   - 验收：同一 pptx 用 COM 与新方案各渲染一次，逐页像素差异 < 2%
3. **图片/PDF → 可编辑 deck**
   - 价值：用户给的是截图版案例时也能学
   - 风险：依赖 PaddleOCR/多模态模型，重
   - 结论：先不做，用"上传 .pptx 案例"覆盖 80% 场景

## 4. 明确不引入

- 星数 < 100 的 HTML→PPTX 早期项目：接口不稳，自己 200 行能覆盖的部分不背依赖
- 需要 playwright/puppeteer 常驻的重型前端方案：服务端无浏览器环境会直接崩
- AGPL 项目的代码级复用（PPTist 等）：商用有传染风险，只借鉴交互
- "端到端生成 pptx 的大模型"：不可控、不可编辑、无法迭代，和本项目的"可编辑"前提冲突

## 5. 与 Phase 1 的衔接

Phase 1（T1–T5）落地后，Phase 2 按 `P0-1 → P0-3 → P0-2 → P0-4` 顺序排，每项独立可验收：

- P0-1、P0-3 改 `kit/`（新增文件 + `profile`/`inspect` 少量接线）
- P0-2 新增 `kit/pptxkit/retrieval.py`，`server` 只加一个分支
- P0-4 改 `outline.py`，用 P0-1 产出的 `profile.outline`

数据来源：GitHub star 数经 shields.io 于 2026-10-27 读取；论文指标来自 DECKBench / SlidesGen-Bench 摘要页。

---

## 6. 已完成（2026-10-27，本轮）

| 事项 | 产物 | 说明 |
|---|---|---|
| 开源字体安装 | `scripts/install-fonts.py` | Noto Sans SC（Regular/Medium/Bold）、Noto Serif SC（Regular/Bold）、Inter（Regular/Medium/Bold），按当前用户装到 `%LOCALAPPDATA%\\Microsoft\\Windows\\Fonts` 并写 HKCU 注册表，无需管理员 |
| 设计系统 | `kit/pptxkit/design.py` | OKLCH 配色推导、1.25 字号阶梯、12 栏网格、WCAG 对比度计算 |
| 编排样张 | `scripts/design-showcase.py` → `out/design-showcase.pptx` | 6 页：出血色块封面、KPI 三栏、编辑风双栏、时间线、引言、收尾；用于对照旧版式 |
| 主题字体切换 | `kit/pptxkit/themes.py` | academic/editorial/minimal 全部改用思源黑体/思源宋体，模板已重新出片 |
| 颜色/图标工具 | pip: `coloraide`、`svglib`、`reportlab`；npm: `lucide-static` | 图标暂只用于 HTML 路径；PPTX 侧留待 P1 |

**下一步（未做）**：把 `design.py` 的网格/字号/配色接进 `layouts.py`，重做 cover/section/stats/quote/timeline 五个版式；
给 `html_deck.build()` 加 `theme_obj`（svc 已提出，属于 T1 遗留缺口）。

---

## 7. 版式骨架（`profile.layout`）已落地（t15）

`profile.analyze()` 现在除了配色/字体/密度，还输出一个 `layout` 块（10 个比例 + `source` / `confidence`）；
`profile.analyze_image()` 可以从一张幻灯片图片（png/jpg/webp）用纯 PIL + numpy 估计同一结构；
`builder.build(..., skeleton=profile.layout_hints(pf))` 会让 cover / section / bullets / image 四个版式
按骨架微调标题纵向位置与左右边距（限幅）。不传 `skeleton` 时输出与旧版**逐字节一致**
（已用 17 页 academic deck 对比全部 slide XML 验证）。

### 7.1 字段含义（坐标均为相对页面的 0–1 比例）

| 字段 | 含义 | 测量方式（pptx） | 估计方式（图片） |
|---|---|---|---|
| `title_top_ratio` | 标题框上沿 | 文字形状里字号最大者（top < 0.6）的 `top/H`，全 deck 中位数 | 顶部最密集文字带的上沿 |
| `title_height_ratio` | 标题框高度 | 同上的 `(bottom-top)/H` | 该文字带的高度 |
| `margin_left_ratio` | 内容左边距 | 所有文字/图片形状（排除页脚）的 `min(left)/W` | 墨迹 bbox 左边界 / 宽 |
| `margin_right_ratio` | 内容右边距 | `1 - max(right)/W` | `1 - 墨迹 bbox 右边界 / 宽` |
| `margin_top_ratio` | 内容上边距 | `min(top)/H` | 墨迹 bbox 上边界 / 高 |
| `margin_bottom_ratio` | 内容下边距 | `1 - max(bottom)/H` | `1 - 墨迹 bbox 下边界 / 高` |
| `image_area_ratio` | 图片面积占比 | 图片形状（`shape_type == PICTURE`）面积和 | 16px 图块中「颜色方差 > 260 且彩色像素占比 > 45%」的占比 |
| `text_area_ratio` | 文字面积占比 | 文字形状面积和（可能重叠，故封顶 1.0） | 内容 bbox 面积 − 图片面积 |
| `alignment` | 左对齐 / 居中 | 标题段落的 `alignment`，取不到时看标题框中心是否在 0.5 ± 0.04 | 标题带墨迹左右边界是否居中且宽度 ≤ 0.7 页宽 |
| `columns` | 1 栏 / 2 栏 | 正文文字框左右两簇不重叠 → 2 | 标题带以下列投影的最长空档 ≥ 6% 内容宽 → 2 |
| `source` | `pptx` / `image` / `default` | — | — |
| `confidence` | 0–1 | pptx 固定 1.0 | 由墨迹量/标题带/图片块/栏数命中累加，上限 0.9 |

数值取全 deck 的**中位数**（对齐/栏数取众数）；页脚（top > 0.88）与无文字的形状不参与统计。

### 7.2 生成时怎么用

```python
profile = pptxkit.profile.analyze("案例.pptx")        # 或 analyze_image("slide.png")
key = pptxkit.profile.apply(profile, "academic")      # 配色/字体（原有能力）
hints = pptxkit.profile.layout_hints(profile)         # 骨架 → 微调量
builder.build(doc, "out.pptx", theme_obj=key, skeleton=hints)
```

`layout_hints()` 返回 `title_shift`（英寸，限幅 ±0.15 页高）、`margin_left` / `margin_right`
（限幅 ±0.10 页宽）以及原始 `layout` / `source` / `confidence`。各版式内部还有安全区间
（例如正文页标题带限制在 0.55–2.20in），避免极端案例把标题推到页脚或压住 kicker。
默认档案 / 无档案时三个位移都是 0，版式走原分支，输出与旧版逐字节一致。

### 7.3 误判边界（请按“估计”使用，不要当精确测量）

- **不做文字识别**：图片分析没有 OCR、没有任何模型，只给比例；文字内容、字号、字体一律拿不到。
- **图片面积**只看「色方差大 + 彩色像素多」的块：纯色色块、低饱和图表、黑白照片会被低估；大面积渐变可能被高估。
- **墨迹阈值**固定为「与四边中位色的灰度差 > 28」：白底白字、浅色文字/细线会漏检；深色底整体按绝对差处理。
- **标题带**取「顶部 50% 内最密的一段文字墨迹」，小字号 kicker 与主标题相邻时可能合并；居中判断基于标题带墨迹中心 ±0.04 页宽，块状标题容易被判成居中。
- **栏数**只输出 1 或 2：三栏/多栏、卡片网格、时间线都可能被判成 2；有通栏正文时可能被判成 1。
- **页脚/装饰**：pptx 侧用 `top > 0.88` 排除页脚；图片侧用「贴边、几乎整页高的窄竖线」排除装饰色条，其余装饰（角标、页码）仍会进入 bbox。
- **比例精度**：pptx 侧是形状 bbox（不是文字墨迹），同一段文字换行后 bbox 变高，`title_height_ratio` 会随之变化；图片侧分辨率与抗锯齿会影响 ±0.01–0.03。
- 图片分析的 `palette` / `fonts` / `metrics` 保持默认值——**只有 layout 有意义**。

### 7.4 验证

- `tests/test_layout_profile.py`（16 项）：合成 pptx（标题 0.12 页高、边距 0.08 页宽、左右两栏）与合成图（白底 + 标题条 + 高方差图片块 + 双栏图）的断言；骨架限幅；默认档案/无档案构建的逐字节一致；CLI `profile` 对 pptx 与 png 都输出 JSON。
- 既有四套测试（test_profile 39 / test_e2e 7 / test_server 32 / test_web 29）全绿；`ppt.cmd build templates/academic-defense.md` 页数不变（17 页）、qa 无新增告警。
