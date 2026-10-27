# 视觉升级 v2：五个版式接上设计系统

把 `kit/pptxkit/design.py` 的 12 栏网格、1.25 字号阶梯与 OKLCH 配色推导，接进
`layouts.py` 的 **cover / section / stats / quote / timeline** 五个版式。
其余 6 个版式（bullets / two-col / table / image / refs / end）的绘制代码没有改，
只共享同一个版心常量 `MARGIN_X`（0.95 → 1.0，见文末「已知影响」）。

- 改动文件：`kit/pptxkit/design.py`、`kit/pptxkit/themes.py`、`kit/pptxkit/layouts.py`
- 对照图：本目录 `before/`（改动前构建）与 `after/`（改动后构建），
  `contact-sheet-academic.png` 是 academic 五版式的上下对照总表
- 重建的三份 deck：`decks/academic.pptx`（17 页）、`decks/editorial.pptx`（12 页）、
  `decks/business.pptx`（9 页）

## 1. 新增/统一的设计常量

| 项 | 值 | 来源 |
|---|---|---|
| 版心 | `MARGIN_X = 1.0`、`CONTENT_W = 11.333` | `design.MARGIN_X` / `design.content_w()` |
| 12 栏 | `col_w(1)=0.688`、`col_w(4)=3.591`、`col_w(8)=7.462`、`col_w(12)=11.333` | `col_x(i) = 1.0 + i*0.96775`，`COL_GUTTER = 0.28` |
| 字号 | `T_CAPTION 11.5 / T_SMALL 14.4 / T_BODY 18 / T_H4 22.5 / T_H3 28.1 / T_H2 35.2 / T_H1 43.9 / T_DISPLAY 54.9` | `design.type_size(step)`，比率 1.25 |
| 装饰字号 | `T_DECOR 107.3`（引号）、`T_GHOST 209.5`（章节大编号） | 同一条阶梯的 8 / 11 级，替代 110 / 210 这类魔数 |
| 纵向节奏 | `GAP = COL_GUTTER = 0.28`，行位置 `row_y(n) = 0.85 + n` | `design.row_y` |

新增小工具：`design.lines_height()`（n 行文字高度）、`design.grid_box()`、
`design.text_width()/fits_one_line()`（Pillow 量真实字宽，字体缺失时退回字数估算）。
`themes.title_scale(theme)` 返回「当前 title_size / 内置基准」，让 `profile.apply()`
的 `title_scale` 仍然等比作用于封面/章节标题。

## 2. 五个版式逐项改动（旧 → 新）

### cover

| 元素 | 旧 | 新 | 理由 |
|---|---|---|---|
| kicker | y 1.68–1.98，11/12pt | `row_y(0.90)=1.75`，`T_CAPTION 11.5` | 三条基线对齐网格 |
| 标题 | y 1.95/2.20/2.30（按主题），44/38/33pt，宽 11.13 | `row_y(1.30)=2.15`，`T_H1 43.9`（长标题 `T_H2 35.2`），宽 `col_w(12)=11.333` / editorial 用 `col_w(8)=7.462` | 统一基线；editorial 让出右侧 4.59in 给色块 |
| 标题下短线 | y 2.94–3.88，宽 1.05，2.6pt | y = 2.15 + 两行高 + 0.28 = 3.87，宽 `col_w(1)×1.5 = 1.03` | 标题区固定预留两行，短标题也不贴线 |
| 副标题 | 18pt，y 3.24–4.18，宽 9.6 | `T_BODY 18`，y 4.15，宽 `col_w(6)=5.527` | 与署名同栏宽，短行不散 |
| 署名 | 12.5pt，y 5.95，宽 9.0 | `T_SMALL 14.4`，`row_y(5.1)=5.95`，宽 `col_w(6)` | 正文最小档也保证 14pt 可读 |
| editorial 装饰 | 右上角 4.5in 线框圆 + 顶部横线 | 第 8 栏起 4.591in 宽出血色块 + 2pt 强调竖线 | 网格对齐，色块与文字栏不重叠 |
| academic 装饰 | 左侧 3pt 竖条（0.95 起） | 同一竖条，文字从 1.0 起 | 与网格左缘对齐 |
| minimal 装饰 | 顶部横线 y 1.85 | `row_y(0.45)=1.30`，满版心宽 | 横线与 kicker 间距固定 0.45 |

### section

| 元素 | 旧 | 新 | 理由 |
|---|---|---|---|
| 背景大编号 | x 8.30/8.80/9.05，y 0.25/0.45/0.70，字号 235/210/180 | 右对齐到版心右缘 12.333（盒从 `col_x(4)=4.871` 起，宽 `col_w(8)=7.462`），y 0.30，`T_GHOST 209.5` | 三个主题同一字号/同一右缘；编号不再被幻灯片边缘裁切 |
| 左侧竖条 | (0.95, 2.02, 2.32) | (1.0, 2.0, 2.25) = `row_y(1.15)…row_y(3.40)` | 起止都落在网格行上 |
| PART 标签 | 11pt @ (1.27, 1.99) | 11.5pt @ (1.28, 1.99) = `col_x(0)+GAP` | 与标题同一左缘 |
| 标题 | 38pt @ (1.27, 2.46, 7.5) | `T_H2 35.2` @ (1.28, 2.45, `col_w(8)` 7.462) | 落回阶梯；宽度按栏 |
| 副标题 | 15pt @ (1.27, 4.35, 8.2) | `T_BODY 18` @ (1.28, 4.35, 7.462) | 正文档 18pt；标题换行时自动下移 |

### stats

| 元素 | 旧 | 新 | 理由 |
|---|---|---|---|
| 卡片 | 宽 `(11.433-0.55×2)/3`、间距 0.55 | 3 项各占 4 栏（`col_w(4)=3.591`），x = `col_x(0/4/8)` | 数字卡与网格同宽，多卡时 12/n 栏自动分配 |
| 数字上短线 | 0.66in，2.6pt | `col_w(1)=0.688`，2.2pt | 一栏宽的装饰线 |
| 数字 | 54 / 42 / 34pt | `T_DISPLAY 54.9` / `T_H1 43.9` / `T_H2 35.2` | 只按阶梯降档 |
| 标签 | 14pt | `T_H4 22.5` | 数据卡的二级信息也要在 3m 外可读 |
| 注释 | 11pt | `T_SMALL 14.4` | 正文下限 14pt |
| 纵向 | y = 标题底 + 0.30，行距写死 | y = 标题底 + `GAP`，块间距 `GAP/2`，高度用 `lines_height` | 与标题、卡间距同一节奏 |

### quote

| 元素 | 旧 | 新 | 理由 |
|---|---|---|---|
| 引号 | 110pt @ (0.89, 1.28) | `T_DECOR 107.3` @ (1.0, 1.25) | 去掉 110 这个魔数；与正文左缘对齐 |
| 引文 | 30/26pt @ (1.07, 2.42, 10.83) | 在 `T_H2 35.2 / T_H3 28.1 / T_H4 22.5` 里选「一行放得下」的最大档；@ (1.0, 2.45, `col_w(12)` 11.333) | 用 Pillow 量真实字宽，避免出现只剩两个字的孤行；亦不因字号过大溢出 |
| 出处短线 | (1.07, 5.52, 0.72)，2pt | (1.0, 5.55, `col_w(1)` 0.688)，2pt | 与左边距、一栏宽对齐 |
| 出处 | 13pt @ (1.07, 5.68) | `T_SMALL 14.4` @ (1.0, 5.71) | 统一小字档 |

### timeline

| 元素 | 旧 | 新 | 理由 |
|---|---|---|---|
| 列宽 | `(CONTENT_W - 0.34×(n-1))/n` | `col_w(12//n)`：4 步 `col_w(3)=2.623`，3 步 `col_w(4)=3.591`，x=`col_x(i*span)` | 与 12 栏严格对齐 |
| 主线 | y = 标题底 + 1.05 | `max(row_y(2.05)=2.90, 标题底+3×GAP)` | 标题换行时整体下移，不压标题 |
| 圆点 | 0.15in | `DOT = 0.12`（`GAP×0.43`） | 与 2pt 线宽的视觉重量匹配 |
| 序号 / 阶段 / 说明 | 11 / 16 / 11.5pt | `T_CAPTION 11.5` / `T_H4 22.5` / `T_SMALL 14.4` | 三档全部来自阶梯 |
| 说明文字 | 直接渲染，`月份 \| 阶段 \| 说明` 的第二个 `\|` 会画出来 | 归一成「阶段：说明」 | 三列写法在版式里不再漏出分隔符 |

## 3. 三个主题的配色（design.palette 推导）

`themes.py` 不再手写六组色值，改为 `design.palette(强调色, mode=...)`：
academic/minimal 用 `clean`，editorial 用 `paper`。`Theme` 字段（bg/ink/muted/accent/
accent_soft/rule/font_*/title_size/body_size/kicker_size）一个没动，`profile.apply()`
仍然可以整组覆盖。

| 主题 | 角色 | 旧 | 新 | 对比度（新） |
|---|---|---|---|---|
| academic (`clean`) | bg / ink / muted / accent / soft / rule | `#FFFFFF` `#14213D` `#5C6B82` `#1B3A6B` `#E9EEF7` `#D5DCE8` | `#FFFFFF` `#161E2B` `#69717D` `#1B3A6B` `#E8F1FF` `#DDE0E4` | ink 16.7:1 / muted 4.9:1 / accent 11.3:1 |
| editorial (`paper`) | 同上 | `#F7F3EA` `#141414` `#6E6A63` `#1F3BFF` `#E7EBFF` `#D6CFC0` | `#F9F7FB` `#181E2B` `#6A707D` `#1F3BFF` `#E5EDFF` `#DADCE2` | ink 15.7:1 / muted 4.7:1 / accent 6.3:1 |
| minimal (`clean`) | 同上 | `#FFFFFF` `#111111` `#8A8A8A` `#E4572E` `#FDEDE8` `#E6E6E6` | `#FFFFFF` `#2A1915` `#7C6C67` `#C54E2C` `#FEEBE6` `#E3DEDD` | ink 16.8:1 / muted 5.0:1 / accent 4.7:1 |

- 强调色与底色对比 < 6:1 时，`design.palette()` 会把强调色压暗一档：minimal 的橙
  `#E4572E → #C54E2C`（4.7:1，AA 正文达标、大字号更宽裕）；editorial 的克莱因蓝
  6.3:1 保持原值；academic 深蓝 11.3:1 保持原值。
- editorial 的纸感底色由强调色色相推导（`paper`），从暖米色变成极浅的冷调纸白，
  这是「配色由单一强调色推导」的直接结果，不是手调。

## 4. 验证

| 项 | 结果 |
|---|---|
| 页数 | academic 17 / editorial 12 / business 9（与改动前一致） |
| `ppt.cmd qa` 提示数 | 14 / 8 / 10 条，与改动前完全相同；**全部是「该页没有讲者备注」**，0 条溢出、0 条密度告警 |
| 溢出检查 | 改动前后均为 0；`inspect.qa` 的 `_fit_ratio` 未触发 |
| 字体 | 三份 deck 均为 Noto Sans SC / Noto Serif SC，未报缺失 |
| 单元测试 | `python tests/test_profile.py`：39 个用例全绿 |
| CLI 回归 | `ppt.cmd build templates/academic-defense.md -Out out/_v2.pptx` 退出码 0，17 页 |

## 5. 已知影响

- 重建产物放在本目录 `decks/`，没有覆盖 `out/academic-defense.pptx` 等历史文件——
  那些文件正好作为改动前的对照基线；需要刷新主产物时重新跑
  `ppt.cmd build templates/<name>.md` 即可。
- `themes.MARGIN_X` 从 0.95 改为 1.0（= `design.MARGIN_X`），版心宽度 11.433 → 11.333。
  六个未改动版式的**代码**没变，但左缘统一右移 0.05in、宽度收窄 0.1in；qa 与 PNG
  对照显示没有新增溢出，正文页与新五版式现在共用同一条左缘。
- 封面/章节标题按 `T_H1/T_H2` 固定阶梯，正文页标题仍按 `theme.title_size` 自适应；
  风格档案的 `title_scale` 通过 `themes.title_scale()` 等比作用到这两处（默认主题比例 1.0）。
- `stats`/`timeline`/`quote` 在内容不足时下半页留白较多，这是版式「一页一个论点」的
  刻意留白；若需要填满，请补充内容而不是缩小字号（改字号会破坏阶梯）。
