# PPT Studio HTTP API

本文件描述 `server/` 暴露的接口，与 `docs/CONTRACT.md` 第 4–5 节一致。
服务只用 Python 标准库，**不联网也能完整使用**；LLM 只是可选增强。

## 1. 启动

```powershell
.\ppt.cmd web                              # 默认 127.0.0.1:8765
.\ppt.cmd web -Host 0.0.0.0 -Port 8765     # 局域网可访问
```

等价的直接入口（`-Host` 对应 `--host`）：

```powershell
python -m server.app --host 127.0.0.1 --port 8765
```

启动后终端会打印可访问地址，例如 `PPT Studio 服务已启动：http://127.0.0.1:8765`。
浏览器打开该地址即进入 `web/` 里的单页界面（`GET /`）。

测试/嵌入式用法（同一进程内起线程，端口传 0 由系统分配）：

```python
from server.app import create_server

httpd, base_url = create_server("127.0.0.1", 0, store_dir="...")
# threading.Thread(target=httpd.serve_forever).start()
```

## 2. 通用约定

- 所有响应体 UTF-8；JSON 接口返回 `application/json; charset=utf-8`。
- 错误统一为：

  ```json
  {"error": {"code": "missing_content", "message": "请至少提供 outline 或 prompt"}}
  ```

- 默认**只允许同源**：web/ 页面与接口同源，不需要 CORS，响应里不带
  `Access-Control-Allow-Origin`。需要跨域调试时用环境变量 `PPT_ALLOW_ORIGIN`
  指定精确来源（如 `http://127.0.0.1:5173`）；`*` 会被忽略，避免任意网页借用户
  浏览器读写本机档案。`OPTIONS` 预检返回 204。
- 请求体上限：上传案例 25 MB（超出 413），其他 JSON 接口 1 MB；
  `outline` 另受 64 KB 上限约束（见 4.6）。
- 不缓存接口响应（`Cache-Control: no-store`）。

### 错误码

| HTTP | code | 触发条件 |
|---|---|---|
| 400 | `missing_content` | `/api/generate` 既没有 `outline` 也没有 `prompt` |
| 400 | `invalid_json` | 请求体不是合法 JSON |
| 400 | `bad_request` | JSON 不是对象、Content-Length 非法等 |
| 400 | `invalid_filename` | 缺少 `X-Filename`、扩展名不是 `.pptx/.ppt`、文件名含路径分隔符或 `..` |
| 400 | `empty_body` | 上传内容为空 |
| 400 | `invalid_pptx` | 不是有效的 OOXML `.pptx`（或缺 `ppt/presentation.xml`）；旧版 `.ppt` 无法离线解析 |
| 400 | `invalid_theme` | 未知主题名 |
| 400 | `invalid_outline` | outline 里的图片路径不是工作区 `assets/` 下的相对路径（绝对路径/盘符/UNC/`..`/越界/不存在/非图片） |
| 400 | `outline_too_large` | outline 超过 64 KB，或派生页数超过 60 页 |
| 400 | `empty_outline` | 大纲里没有可渲染的页面 |
| 400 | `invalid_rating` | 反馈评分不是 1–5 的整数 |
| 400 | `no_profile` | 对没有关联档案的 deck 提交反馈 |
| 404 | `not_found` | 路由/文件不存在 |
| 404 | `profile_not_found` | 档案 id 不存在 |
| 404 | `deck_not_found` | deck id 不存在 |
| 405 | `method_not_allowed` | 只支持 GET 的接口收到 POST |
| 413 | `payload_too_large` | 请求体超过上限 |
| 500 | `internal` | 服务端异常（细节只进日志，不回栈） |

## 3. 数据模型

### 3.1 风格档案 `profiles/<profile_id>.json`

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

- `id` 形如 `pf-YYYYMMDD-<8 位以上随机 hex>`（如 `pf-20261027-ab12cd34`），
  8 位 hex 的随机空间用于避免同日多次上传互相覆盖。
- `density ∈ [0.6, 1.4]`，`title_scale ∈ [0.85, 1.25]`，颜色统一大写 `#RRGGBB`。
- 档案不对外暴露内部字段（下划线开头）。

### 3.2 deck 元数据 `decks/<deck_id>/meta.json`

```json
{"id": "dk-20261027-1f2a3b4c", "created_at": "2026-10-27T19:10:03+08:00",
 "title": "季度复盘", "theme": "academic", "profile_id": "pf-20261027-ab12cd34",
 "pages": 10, "files": {"pptx": "deck.pptx", "html": "deck.html"},
 "outline_chars": 1830, "llm_used": false}
```

### 3.3 存储布局

```
server/store/                     # 默认数据目录，可用 create_server(store_dir=...) 改
├─ cases/<profile_id>/source.pptx # 上传案例原文（落盘名固定，与用户文件名无关）
├─ profiles/<profile_id>.json     # 风格档案
├─ decks/<deck_id>/               # deck.pptx / deck.html / outline.md / meta.json
├─ feedback.jsonl                 # 追加式反馈流水
└─ tmp/                           # 上传临时文件（处理完即删除）
```

## 4. 接口

### 4.1 `GET /api/health`

```json
{"status": "ok", "version": "1.0.0", "llm": false}
```

`llm` 为 `true` 表示配置了 `PPT_LLM_BASE_URL` + `PPT_LLM_API_KEY`（只做配置判断，不发网络请求）。

### 4.2 `GET /api/themes`

```json
[{"key": "academic", "label": "学术答辩"},
 {"key": "editorial", "label": "杂志编辑风"},
 {"key": "minimal", "label": "商务简约"}]
```

### 4.3 `GET /api/profiles`

档案摘要列表，按 `created_at` 倒序：

```json
[{"id": "pf-20261027-ab12cd34", "name": "我的答辩模板", "source": "source.pptx",
  "created_at": "...", "slide_size": [13.33, 7.5],
  "palette": {"...": "..."}, "fonts": {"...": "..."},
  "style": {"density": 1.0, "title_scale": 1.0, "bullet_char": "▪"},
  "slides": 17, "learn": {"rounds": 0}}]
```

### 4.4 `POST /api/profiles` —— 上传案例，提取风格档案

- 请求体：`.pptx` 的**原始字节**（不是 multipart）。
- 请求头：`X-Filename` 为原始文件名，**只作展示**。非 ASCII 名可用 percent-encoding
  （如 `%E6%A1%88%E4%BE%8B.pptx`），服务端会还原成中文展示名。
- 限制：25 MB、扩展名 `.pptx/.ppt`，落盘固定为 `cases/<新 id>/source.pptx`。
  文件名含 `/`、`\`、`..`、控制字符一律 400。
- 档案 id 使用 `pf-YYYYMMDD-<8 位以上随机 hex>`；同日多次上传会得到不同 id，
  绝不覆盖已有档案（`profile.save` 写前检查冲突并重新生成）。
- 旧版二进制 `.ppt`（OLE）无法离线解析，返回 400 `invalid_pptx`，请先另存为 `.pptx`。

响应：

```json
{"id": "pf-20261027-ab12cd34", "profile": { ...完整档案... }}
```

### 4.5 `GET /api/profiles/{profile_id}`

返回完整档案；不存在 → 404 `profile_not_found`。

### 4.6 `POST /api/generate` —— 生成 deck

请求（全部可选，按需组合）：

```json
{"title": "季度复盘", "subtitle": "2026 Q3", "author": "张三", "org": "产品部",
 "date": "2026-10-27", "theme": "academic", "profile_id": "pf-20261027-ab12cd34",
 "prompt": "给团队讲我们怎么把 PPT 制作自动化", "outline": "# 标题\n- 要点", "pages": 10}
```

生成优先级（高 → 低）：

1. `outline`（Markdown）→ 直接使用；
2. `prompt` + LLM 可用 → 调 OpenAI 兼容接口扩写；调用失败自动降级到 3；
3. `prompt` → `outline.build_outline(...)` 离线启发式；
4. 都没有 → 400 `missing_content`。

约束（超出直接 400，不落盘）：

- `outline` 原始字节 ≤ 64 KB；
- 由 outline 派生的页数（含自动封面）≤ 60 页；请求参数 `pages` 只影响离线大纲，范围 2–40；
- `outline` 中的图片只能是工作区 `assets/` 下的相对路径（绝对路径、盘符、UNC、`..` 逃逸、
  不存在的文件、非图片后缀一律 400 `invalid_outline`）；
- 生成出的 `deck.html` 里图片一律渲染为占位，不回显服务端本机路径。

响应：

```json
{"id": "dk-20261027-1f2a3b4c", "title": "季度复盘", "pages": 10,
 "files": {"pptx": "/files/dk-20261027-1f2a3b4c/deck.pptx",
           "html": "/files/dk-20261027-1f2a3b4c/deck.html"},
 "llm_used": false, "profile_id": "pf-20261027-ab12cd34", "theme": "academic",
 "notes": ["使用请求提供的大纲", "套用档案「我的答辩模板」的配色与字号"]}
```

- `title` 是服务端最终采用的标题（`meta.json` 的 `title`）：payload 没给标题时取提示词/首标题，
  前端应优先使用它。
- `pages` 是实际页数（含封面），与 `meta.json` 一致。
- `notes` 告诉前端：用了哪个档案、是否降级、是否用了 LLM。无 key 时 `llm_used` 恒为 `false` 且不报错。
- 生成产物：`outline.md`、`deck.pptx`（可编辑）、`deck.html`（单文件网页版）、`meta.json`。
- 带 `profile_id` 时 `deck.pptx` 与 `deck.html` 使用**同一个档案 Theme**（配色、字体一致），
  见契约 §5.1；`deck.html` 的 CSS 值经过白名单校验，恶意字体名只会回退为默认字体。

### 4.7 `GET /api/decks`

最近 50 条 deck 摘要（按时间倒序）：

```json
[{"id": "dk-20261027-1f2a3b4c", "created_at": "...", "title": "季度复盘",
  "theme": "academic", "profile_id": "...", "pages": 10, "llm_used": false}]
```

### 4.8 `GET /api/decks/{deck_id}`

返回 `meta.json` 原文；不存在 → 404 `deck_not_found`。

### 4.9 `POST /api/decks/{deck_id}/feedback` —— 打分并微调档案

```json
{"rating": 5, "notes": "内容可以再密一点"}
```

- `rating` 必须是 1–5 的整数；`notes` 可省略。
- 规则（契约第 7 节）：`rating ≥ 4` → `density *= 1.03`；`rating == 3` 不变；
  `rating ≤ 2` → `density *= 0.90`、`title_scale *= 0.97`；随后 clamp 回区间，
  `learn.rounds += 1`，history 追加一条。
- 档案立即落盘（下次生成直接生效），并追加一行 `feedback.jsonl`。

响应：

```json
{"ok": true, "profile_id": "pf-20261027-ab12cd34",
 "style": {"density": 1.03, "title_scale": 1.0, "bullet_char": "▪"},
 "message": "已调整：密度 1.03"}
```

- deck 没有关联档案 → 400 `no_profile`；deck 不存在 → 404 `deck_not_found`。

### 4.10 `GET /files/{deck_id}/{name}`

二进制下载，`name` 只允许 `deck.pptx`（`attachment`）与 `deck.html`（`inline`）。
其他 name、非法 deck id、`..` 等路径穿越一律 404 `not_found`。

### 4.11 静态页面

| 路径 | 内容 |
|---|---|
| `GET /` | `web/index.html`（前端未就绪时返回一张说明页） |
| `GET /app.js`、`/style.css`、`/assets/*` | `web/` 下的静态文件 |
| 其他 | `web/` 内查找，找不到 404 |

静态文件解析后必须落在 `web/` 目录内，`..` 穿越被拒绝。

## 5. 可选 LLM（OpenAI 兼容）

```powershell
$env:PPT_LLM_BASE_URL = "https://api.deepseek.com/v1"
$env:PPT_LLM_API_KEY  = "sk-..."
$env:PPT_LLM_MODEL    = "deepseek-chat"     # 可省略，默认 deepseek-chat
```

- 未配置时 `/api/health` 的 `llm` 为 `false`，生成走离线启发式，功能不受影响。
- 配置后仅在 `prompt` 路径（且没有显式 `outline`）时调用；超时 20 秒。
  网络错误、响应缺字段、内容不含 `#` 标题都会**自动降级**到离线大纲，并在 `notes` 里说明。

## 6. 安全边界

- 上传文件按不可信输入处理：大小、扩展名、OOXML 结构三重校验；落盘文件名由服务端生成，
  用户文件名只进响应和档案的 `name` 字段，绝不参与路径拼接。
- 所有 id 必须匹配 `^[A-Za-z0-9][A-Za-z0-9_-]{0,63}$` 才允许访问磁盘。
- outline 图片路径按白名单解析（只允许工作区 `assets/` 下的相对路径），绝对路径、盘符、
  UNC、`..` 逃逸、越界、不存在、非图片后缀一律 400，且不会创建 deck 目录；
  `builder._resolve` 在库层再次校验，防止 CLI 绕过。
- `deck.html` 里的 CSS 变量做白名单/转义：颜色只接受 `#RRGGBB`，字体名只允许
  字母/数字/空格/下划线/连字符/中文，超出范围回退默认字体——档案来自用户上传，
  不能把原始字符串拼进 `<style>`。
- 静态文件与下载文件都做了解析后目录包含检查。
- 默认同源（不发 CORS 头）；绑定 `0.0.0.0` 时启动日志会警告「仅可信网络」。
  服务没有鉴权，不要把端口暴露到不可信网络。
- 服务端异常只回 `500 internal`，traceback 只写日志。

## 7. 已知限制 / 契约缺口

1. `deck.html` 与 `deck.pptx` 共用同一档案 Theme（契约 §5.1）；但网页版里的图片一律渲染为
   占位（服务端 deck 目录不携带图片文件），pptx 会把 `assets/` 下的图片嵌入。若要让网页版
   显示图片，需要额外把图片复制到 deck 目录并提供下载路由，属后续迭代。
2. 旧版 `.ppt`（OLE 二进制）无法离线解析，会返回 400 `invalid_pptx` 提示另存为 `.pptx`。
3. 服务端对 `kit/pptxkit/profile.py`、`outline.py` 做了适配层（`server/kitlink.py`）：
   优先调用 kit 的契约实现；模块缺位或导入失败时用内置兜底实现（默认配色、启发式大纲），
   保证服务在离线/内核未就绪时仍可用。kit 存在时不会使用兜底。
4. LLM 配置下，服务端先自行调用一次 OpenAI 兼容接口；失败后调用内核离线大纲时会临时隐藏
   LLM 环境变量，避免为同一个请求重复等待超时。
5. `html_deck.build` 的图片模式默认 `relative`（保留安全的相对路径，CLI 用）；
   服务端生成 deck 时传 `placeholder`。
