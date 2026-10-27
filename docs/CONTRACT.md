# PPT Studio 契约（团队共同遵守）

> 本文件是评审通过前的设计基线。所有成员按此实现，不得擅自改接口；确实需要改，先在输出里说明。

## 1. 目标

在现有 `C:\Users\91125\Desktop\PPT` 工作区里，长出一个**可联网的 PPT 制作服务**：

1. 浏览器打开就能用：填主题/提示词 → 生成**可编辑 .pptx** + 网页版 → 下载；
2. **上传自己的 PPT 作为案例**，系统分析其配色、字体、信息密度，生成"风格档案"；
3. 之后生成时选用该档案 → 产出的 deck 向案例靠拢；
4. 生成后打分反馈 → 档案自动微调（密度/标题比例）→ 下次生成更接近预期；
5. 服务可绑定 `0.0.0.0` 供局域网访问；可选接入 OpenAI 兼容 LLM 做大纲扩写。

## 2. 硬约束（先看这条，避免返工）

- **不能 pip install**（沙箱拒绝写临时目录）。只用 Python 标准库 + 工作区已有依赖（python-pptx、Pillow）。
- **不要用 subprocess 抓子进程输出**（沙箱里会 EPERM）。测试要起服务，就在**同进程内**起 `ThreadingHTTPServer`（端口传 0）。
- 脚本/文件一律 **UTF-8**；`.ps1` 必须是 **UTF-8 with BOM**。
- Python 用工作区自带的解释器，命令里写：
  `& "$env:USERPROFILE\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe"`
- 不联网也能跑通全部测试；LLM 只是可选增强，缺 key 必须优雅降级。
- 上传文件按**不可信输入**处理：限制 25 MB、只收 `.pptx/.ppt`、文件名只做展示，落盘用生成的 id，**绝不能**把用户文件名拼进路径。

## 3. 目录与归属

```
PPT/
├─ server/                    # 服务端（svc 负责）
│   ├─ __init__.py
│   ├─ app.py                 # 路由 + 静态文件 + 启动入口
│   ├─ jobs.py                # 生成任务、deck 元数据、反馈落盘
│   └─ store/                 # 运行时数据（git 忽略）
│       ├─ cases/<profile_id>/source.pptx
│       ├─ profiles/<profile_id>.json
│       ├─ decks/<deck_id>/{deck.pptx,deck.html,outline.md,meta.json}
│       └─ feedback.jsonl
├─ web/                       # 前端（web 负责）
│   ├─ index.html
│   ├─ app.js
│   └─ style.css
├─ kit/pptxkit/               # 生成内核（style 负责扩展）
│   ├─ profile.py             # 新增：案例分析 / 档案存取 / 反馈调参 / 主题覆盖
│   ├─ outline.py             # 新增：提示词 → Markdown 大纲（离线启发式 + 可选 LLM）
│   ├─ builder.py             # 改：build(..., theme_obj=Theme) 支持主题对象
│   └─ themes.py              # 改：Theme 加 replace 友好的副本工具（可选）
├─ tests/
│   ├─ test_profile.py        # style 负责
│   ├─ test_server.py         # svc 负责
│   └─ test_e2e.py            # qa 负责
├─ docs/
│   ├─ CONTRACT.md            # 本文件
│   └─ API.md                 # svc 负责更新
└─ scripts/ppt.ps1            # svc 增加 `web` 子命令
```

## 4. 数据模型

### 4.1 风格档案（profiles/<id>.json）

```json
{
  "id": "pf-20261027-ab12cd34",
  "name": "我的答辩模板",
  "source": "source.pptx",
  "created_at": "2026-10-27T19:02:11+08:00",
  "slide_size": [13.33, 7.5],
  "palette": {"bg": "#FFFFFF", "ink": "#14213D", "muted": "#5C6B82",
               "accent": "#1B3A6B", "accent_soft": "#E9EEF7", "rule": "#D5DCE8"},
  "fonts": {"title": "微软雅黑", "body": "微软雅黑"},
  "metrics": {"slides": 17, "chars_per_slide": 96.5, "bullets_per_slide": 3.2,
               "title_pt": 31.0, "body_pt": 19.0, "image_ratio": 0.12},
  "style": {"density": 1.0, "title_scale": 1.0, "bullet_char": "▪"},
  "learn": {"rounds": 0, "history": []}
}
```

- `palette/fonts` 从 `ppt/theme/theme1.xml`（`a:clrScheme`、`a:fontScheme`）与正文 run 的字体/字号统计得出；取不到就用安全默认值。
- `density` 单位：0.6–1.4，1.0 为基准；`title_scale` 0.85–1.25。
- 颜色统一大写 `#RRGGBB`。
- `id` 形如 `pf-YYYYMMDD-<8 位以上随机 hex>`：随机空间足够大，`save()` 写前检查冲突，
  同日多次上传不会互相覆盖（评审 F9）。

### 4.2 deck 元数据（decks/<id>/meta.json）

```json
{"id":"dk-20261027-1f2a3b4c","created_at":"...","title":"...","theme":"academic",
 "profile_id":"pf-...","pages":10,"files":{"pptx":"deck.pptx","html":"deck.html"},
 "outline_chars": 1830, "llm_used": false}
```

## 5. HTTP API（全部 JSON，UTF-8；错误统一 `{"error":{"code":"...","message":"..."}}`）

| 方法 | 路径 | 请求 | 响应 |
|---|---|---|---|
| GET | `/api/health` | – | `{"status":"ok","version":"1.0.0","llm":false}` |
| GET | `/api/themes` | – | `[{"key":"academic","label":"学术答辩"}, ...]` |
| GET | `/api/profiles` | – | `[档案摘要]`（不含 palette 以外的冗余） |
| POST | `/api/profiles` | body=原始 .pptx 字节，头 `X-Filename` | `{"id":"pf-...", "profile":{...}}` |
| GET | `/api/profiles/{id}` | – | 完整档案 |
| POST | `/api/generate` | `{"title","subtitle","author","org","date","theme","profile_id","prompt","outline","pages"}` | `{"id":"dk-...","title":"...","pages":n,"files":{"pptx":"/files/dk-.../deck.pptx","html":"/files/dk-.../deck.html"},"llm_used":false,"notes":["..."]}` |
| GET | `/api/decks` | – | 最近 50 条 deck 摘要 |
| GET | `/api/decks/{id}` | – | meta.json |
| POST | `/api/decks/{id}/feedback` | `{"rating":1-5,"notes":""}` | `{"ok":true,"profile_id":"pf-...","style":{...},"message":"已调整：密度 0.95"}` |
| GET | `/files/{deck_id}/{name}` | – | 二进制下载（`name` 只允许 `deck.pptx/deck.html`） |
| GET | `/` `/app.js` `/style.css` | – | web/ 静态文件 |

生成规则（优先级从高到低）：
1. 传了 `outline`（Markdown）→ 直接用；
2. 传了 `prompt` 且 LLM 可用 → LLM 生成 Markdown 大纲（失败自动降级到 3）；
3. 传了 `prompt` → `outline.build_outline(prompt, pages=…, kind=…, profile=…)` 离线启发式；
4. 都没有 → 400 `missing_content`。

响应里 `notes` 用于告诉前端"用了哪个档案/是否降级"。`title` 是服务端最终采用的标题
（payload 没给时取提示词/首标题），前端应优先使用它。

资源与安全上限（超出直接 400，不落盘）：
- `outline` 原始字节 ≤ 64 KB；由 outline 派生的页数（含自动封面）≤ 60 页
  → 400 `outline_too_large`。
- outline 里的图片只允许工作区 `assets/` 下的相对路径；绝对路径、盘符、UNC、`..` 逃逸、
  不存在/非图片后缀 → 400 `invalid_outline`（评审 F1）。`deck.html` 的图片一律占位。
- 默认同源（不发 CORS 头）；`PPT_ALLOW_ORIGIN` 可显式指定一个跨域来源（`*` 被忽略）。
  绑定非回环地址时启动日志会提示「仅可信网络」（评审 F7）。

### 5.1 定稿：网页版与 PPTX 同一主题（推翻早期"html 固定基础主题"的退让）

`POST /api/generate` 带 `profile_id` 时，`deck.pptx` 与 `deck.html` **必须**使用同一个 `Theme`（颜色、字体都一致）；
`html_deck.build()` 需接受 `theme_obj`，由服务端传入 `profile.apply()` 的结果。

**前置条件（安全）**：`html_deck` 把颜色与字体写进 `<style>` 前必须做白名单校验——颜色仅接受 `^#[0-9A-Fa-f]{6}$`，
字体名只允许已知字体（Noto Sans SC / Noto Serif SC / Inter / 微软雅黑 等）或按 CSS 字符串转义，禁止把档案里的原始字符串直接拼接。
该要求来自评审 F10（档案可带回 `</style><script>`，启用 theme_obj 后会变成存储型 XSS）。

证据要求：同一档案下 `deck.pptx` 的显式色与 `deck.html` 的 `--accent` 十六进制一致；恶意字体名档案生成的 html 不含可执行脚本。

### 5.2 样例资产只读（评审 F2）

- 回归样例固定在 `tests/fixtures/case-academic.pptx`、`case-editorial.pptx`，只读；
  `tests/test_profile.py` 与 `tests/test_e2e.py` 只读取它们，不再依赖 `out/`。
- `scripts/ppt.ps1` 的 `build` / `html` 拒绝把输出写进 `tests/fixtures/`；
  README 的示例命令一律带 `-Out`，避免默认输出覆盖样例或既有产物。
- 重建样例：`python tests/fixtures/make_fixtures.py`（手工执行，测试不会写样例）。

## 6. 内核接口（冻结签名）

```python
# kit/pptxkit/profile.py
def analyze(pptx_path: str | Path, *, name: str = "", source: str = "") -> dict
def save(profile: dict, store_dir: str | Path) -> dict      # 返回带 id 的档案，落盘 cases/+profiles/
def load(store_dir: str | Path, profile_id: str) -> dict
def list_all(store_dir: str | Path) -> list[dict]
def apply(profile: dict, base_theme: str | None = None) -> "Theme"   # 覆盖配色/字体/标题比例
def learn(profile: dict, rating: int, notes: str = "") -> dict        # 就地调整 style 并记 history

# kit/pptxkit/outline.py
def llm_available() -> bool
def build_outline(prompt: str, *, pages: int = 10, kind: str = "auto",
                  profile: dict | None = None, title: str = "") -> str

# kit/pptxkit/builder.py（新增的都是可选、向后兼容的扩展参数）
def build(doc, out, *, theme: str | None = None, theme_obj: Theme | None = None,
          base_dir=None, asset_root=None) -> Path
# asset_root: 额外的图片白名单目录（CLI 用同级 assets/ 支持模板的 ../assets 引用）；
#             不传时只允许 base_dir 内的相对路径。

# kit/pptxkit/html_deck.py（同上，保持向后兼容）
def build(md_path, out_path=None, *, theme: str | None = None,
          title: str | None = None, theme_obj: Theme | None = None,
          image_mode: str = "relative") -> Path
# theme_obj: 风格档案 Theme（优先于 theme 名）；
# image_mode: "relative"（保留安全相对路径）| "placeholder"（图片渲染成占位，
#             服务端生成 deck 时使用，避免回显本机路径）。
```

`profile.apply()` 返回一个 `Theme` 实例（`themes.Theme` 是 frozen dataclass，用 `dataclasses.replace` 生成副本），`title_size` 乘 `title_scale`。

## 7. 迭代规则（反馈如何生效）

- `rating >= 4`：`density *= 1.03`（内容可再多一点）
- `rating == 3`：不变
- `rating <= 2`：`density *= 0.90`，`title_scale *= 0.97`（更少字、标题更克制）
- 每轮 clamp 到 4.1 节区间；`learn.rounds += 1`，history 追加 `{at, rating, density, title_scale, notes}`
- 下一次生成用新参数：`density` 影响大纲每页要点数（`round(base * density)`）与正文目标字数，`title_scale` 影响标题字号。

## 8. 验收口径

| 任务 | 必须能跑通 |
|---|---|
| T1 服务 | `python tests/test_server.py` 全绿；`.\ppt.cmd web` 能起服务并打印地址 |
| T2 内核 | `python tests/test_profile.py` 全绿（只读样例 `tests/fixtures/case-academic.pptx` / `case-editorial.pptx`） |
| T3 前端 | 页面可加载；`/api/health` 有响应；上传+生成+下载在浏览器里走通 |
| T4 验证 | `python tests/test_e2e.py` 全绿：上传案例→出档案→生成→校验 pptx 页数/可打开→反馈→档案变化 |
| T5 评审 | 逐条核对本文档 1–7 节，给出 pass / needs_revision + findings |

## 9. 本地运行

```powershell
.\ppt.cmd web                        # 默认 127.0.0.1:8765
.\ppt.cmd web -Host 0.0.0.0 -Port 8765   # 局域网可访问
```

可选 LLM（OpenAI 兼容）：

```
PPT_LLM_BASE_URL=https://api.deepseek.com/v1
PPT_LLM_API_KEY=sk-...
PPT_LLM_MODEL=deepseek-chat
```

未配置时 `/api/health` 返回 `llm:false`，生成走离线启发式，功能不受影响。
