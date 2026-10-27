"""生成任务编排、deck 元数据、案例/档案与反馈落盘。

本模块不关心 HTTP，只处理数据；`app.py` 负责路由。所有落盘路径都来自服务端
生成的 id，用户提供的文件名只作为展示名出现（见 `_safe_filename`）。
"""
from __future__ import annotations

import contextlib
import datetime as _dt
import json
import os
import re
import threading
import time
import uuid
from pathlib import Path
from urllib import request as _urlrequest

from . import kitlink

VERSION = "1.0.0"

MAX_UPLOAD_BYTES = 25 * 1024 * 1024
MAX_JSON_BYTES = 1 * 1024 * 1024
MAX_OUTLINE_BYTES = 64 * 1024          # 用户 outline 的字节上限（评审 F4）
MAX_OUTLINE_PAGES = 60                 # 由 outline 派生的页数硬上限
ALLOWED_UPLOAD_EXT = {".pptx", ".ppt"}
DECK_FILES = ("deck.pptx", "deck.html")
DEFAULT_PAGES = 10
MAX_DECKS_LISTED = 50
LLM_TIMEOUT_SECONDS = 20.0
ID_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_-]{0,63}$")

WORKSPACE_ROOT = Path(__file__).resolve().parent.parent
ASSETS_DIR = WORKSPACE_ROOT / "assets"
IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp"}
_IMG_LINE_RE = re.compile(r"^!\[[^\]]*\]\((.+?)\)\s*$")
_DRIVE_RE = re.compile(r"^[A-Za-z]:")


class ApiError(Exception):
    """带 HTTP 语义的接口错误，统一渲染成 {"error":{"code","message"}}。"""

    def __init__(self, status: int, code: str, message: str):
        super().__init__(message)
        self.status = int(status)
        self.code = code
        self.message = message


# --------------------------------------------------------------------------- #
# 存储
# --------------------------------------------------------------------------- #

class Store:
    """server/store 目录结构（契约第 3 节）。"""

    def __init__(self, root):
        self.root = Path(root)
        self.cases = self.root / "cases"
        self.profiles = self.root / "profiles"
        self.decks = self.root / "decks"
        self.tmp = self.root / "tmp"
        self.feedback_path = self.root / "feedback.jsonl"
        for folder in (self.cases, self.profiles, self.decks, self.tmp):
            folder.mkdir(parents=True, exist_ok=True)
        self._lock = threading.Lock()
        self._profile_locks: dict[str, threading.RLock] = {}
        self._profile_locks_guard = threading.Lock()

    # -- 并发安全 ---------------------------------------------------------- #

    def profile_lock(self, profile_id: str) -> threading.RLock:
        """同一个档案的读-改-写用同一把锁串行化（评审 F6/F11，RLock 可重入）。"""
        with self._profile_locks_guard:
            lock = self._profile_locks.get(profile_id)
            if lock is None:
                lock = threading.RLock()
                self._profile_locks[profile_id] = lock
            return lock

    def write_json(self, path: Path, payload: dict) -> None:
        """原子写：同目录临时文件 + os.replace（评审 F11）。

        先截断再写的做法会让并发的 `get_profile` / 生成读到半截 JSON 而 500；
        这里与 kit `profile.save()` 一致：写完再原子替换，读方永远看到完整文件。
        """
        path.parent.mkdir(parents=True, exist_ok=True)
        text = json.dumps(payload, ensure_ascii=False, indent=2) + "\n"
        tmp_path = path.with_name(f".{path.name}.{uuid.uuid4().hex}.tmp")
        with self._lock:
            try:
                tmp_path.write_text(text, encoding="utf-8")
                # Windows 上目标文件被短暂打开时 os.replace 可能抛 PermissionError，
                # 重试几次即可；读方仍只会看到旧/新完整文件。
                for attempt in range(6):
                    try:
                        os.replace(tmp_path, path)
                        break
                    except PermissionError:
                        if attempt == 5:
                            raise
                        time.sleep(0.02)
            finally:
                try:
                    tmp_path.unlink()
                except OSError:
                    pass

    def append_feedback(self, record: dict) -> None:
        line = json.dumps(record, ensure_ascii=False) + "\n"
        with self._lock:
            with self.feedback_path.open("a", encoding="utf-8") as handle:
                handle.write(line)

    # -- 路径安全 ---------------------------------------------------------- #

    def deck_dir(self, deck_id: str) -> Path:
        if not valid_id(deck_id):
            raise ApiError(404, "not_found", "文件不存在")
        return self.decks / deck_id

    def profile_path(self, profile_id: str) -> Path:
        if not valid_id(profile_id):
            raise ApiError(404, "profile_not_found", "档案不存在")
        return self.profiles / f"{profile_id}.json"


def valid_id(value: str) -> bool:
    return bool(value) and bool(ID_RE.fullmatch(str(value)))


def _now_iso() -> str:
    return _dt.datetime.now().astimezone().isoformat(timespec="seconds")


def _new_id(prefix: str) -> str:
    """deck/兜底 id：<prefix>-YYYYMMDD-<8 位随机 hex>（评审 F9）。"""
    return f"{prefix}-{_dt.datetime.now():%Y%m%d}-{uuid.uuid4().hex[:8]}"


def _fold_field(value) -> str:
    """front matter 字段：折叠所有空白/控制字符为单空格，避免换行注入（评审 F5）。"""
    text = "" if value is None else str(value)
    return re.sub(r"[\s\x00-\x1f\x7f]+", " ", text).strip()


def _text(value, limit: int = 0) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r\n", "\n").replace("\r", "\n").strip()
    if limit and len(text) > limit:
        text = text[:limit]
    return text


def _int(value, default: int, low: int, high: int) -> int:
    try:
        number = int(value)
    except (TypeError, ValueError):
        return default
    return max(low, min(high, number))


def normalize_filename(raw: str) -> str:
    """把 X-Filename 还原成可读展示名。

    浏览器里非 ASCII 文件名不能直接放进请求头，前端可能用 percent-encoding；
    个别客户端会直接塞 UTF-8 字节（被 http.server 按 latin-1 解码成乱码），
    这里两种都尝试还原。还原只影响展示名，不参与任何路径拼接。
    """
    name = _text(raw, 300).strip().strip('"').strip("'")
    if not name:
        return ""
    if "%" in name:
        try:
            from urllib.parse import unquote

            name = unquote(name)
        except Exception:
            pass
    try:
        name = name.encode("latin-1").decode("utf-8")
    except (UnicodeEncodeError, UnicodeDecodeError):
        pass
    return name.strip()


def _safe_filename(raw: str) -> tuple[str, str]:
    """X-Filename 只做展示：返回 (展示名, 扩展名)。

    任何可能被拼进路径的东西（分隔符、`..`、控制字符）一律拒绝。
    """
    name = normalize_filename(raw)
    if not name:
        raise ApiError(400, "invalid_filename", "缺少 X-Filename 头（原始文件名）")
    if "\x00" in name or "/" in name or "\\" in name or ".." in name:
        raise ApiError(400, "invalid_filename", "文件名非法：不允许路径分隔符或 ..")
    ext = Path(name).suffix.lower()
    if ext not in ALLOWED_UPLOAD_EXT:
        raise ApiError(400, "invalid_filename",
                       f"只支持 {'/'.join(sorted(ALLOWED_UPLOAD_EXT))} 文件")
    return name, ext


# --------------------------------------------------------------------------- #
# outline 安全：图片只允许工作区 assets/ 下的相对路径（评审 F1）
# --------------------------------------------------------------------------- #

def _safe_outline_image(raw: str) -> Path | None:
    text = str(raw or "").strip().strip('"').strip("'")
    if not text or "\x00" in text:
        return None
    normalised = text.replace("\\", "/")
    if normalised.startswith("/") or _DRIVE_RE.match(normalised):
        return None
    parts = [part for part in normalised.split("/") if part not in ("", ".")]
    if not parts or any(part == ".." for part in parts):
        return None
    if parts[0] != "assets":          # 白名单：只允许工作区 assets/ 下的图片
        return None
    try:
        resolved = (WORKSPACE_ROOT / normalised).resolve()
        assets = ASSETS_DIR.resolve()
    except OSError:
        return None
    if resolved != assets and assets not in resolved.parents:
        return None
    if resolved.suffix.lower() not in IMAGE_SUFFIXES:
        return None
    if not resolved.is_file():
        return None
    return resolved


def _validate_outline_images(text: str) -> None:
    for line in (text or "").splitlines():
        match = _IMG_LINE_RE.match(line.strip())
        if not match:
            continue
        raw = match.group(1)
        if _safe_outline_image(raw) is None:
            raise ApiError(
                400, "invalid_outline",
                f"大纲中的图片路径不允许：{raw}（只允许工作区 assets/ 下的相对路径）")


# --------------------------------------------------------------------------- #
# 档案：上传分析 / 列表 / 读取 / 反馈
# --------------------------------------------------------------------------- #

def _is_ooxml_pptx(path: Path) -> bool:
    """真伪校验：必须是含 ppt/presentation.xml 的 zip 包。

    kit 的 `profile.analyze` 对损坏文件会退回「默认档案」而不报错，服务端必须
    自己拦下垃圾字节，避免把无效文件当案例落盘。
    """
    import zipfile

    try:
        if not zipfile.is_zipfile(path):
            return False
        with zipfile.ZipFile(path) as archive:
            return "ppt/presentation.xml" in set(archive.namelist())
    except Exception:
        return False


def create_profile(store: Store, data: bytes, filename: str) -> dict:
    display, ext = _safe_filename(filename)
    if not data:
        raise ApiError(400, "empty_body", "上传内容为空")
    if len(data) > MAX_UPLOAD_BYTES:
        raise ApiError(413, "payload_too_large",
                       f"文件超过 {MAX_UPLOAD_BYTES // (1024 * 1024)} MB 上限")

    tmp_path = store.tmp / f"{uuid.uuid4().hex}{ext}"
    try:
        tmp_path.write_bytes(data)
        if not _is_ooxml_pptx(tmp_path):
            if ext == ".ppt":
                raise ApiError(400, "invalid_pptx",
                               "旧版 .ppt 无法离线解析，请用 PowerPoint / WPS 另存为 .pptx 后上传")
            raise ApiError(400, "invalid_pptx",
                           "文件不是有效的 .pptx（缺少 ppt/presentation.xml）")
        try:
            profile = kitlink.analyze(str(tmp_path), name=Path(display).stem,
                                      source="source.pptx")
        except ApiError:
            raise
        except Exception as exc:
            raise ApiError(400, "invalid_pptx", f"无法解析该 PPT：{exc}") from None
        if not isinstance(profile, dict):
            raise ApiError(400, "invalid_pptx", "PPT 分析结果为空")
        profile.setdefault("name", Path(display).stem or "我的案例")
        profile["source"] = "source.pptx"
        profile["_source_path"] = str(tmp_path)
        saved = kitlink.save_profile(profile, store.root)
        if not isinstance(saved, dict):
            saved = profile
        profile_id = str(saved.get("id") or profile.get("id") or _new_id("pf"))
        saved["id"] = profile_id
        saved.pop("_source_path", None)
        saved.setdefault("created_at", _now_iso())
        saved.setdefault("source", "source.pptx")

        case_path = store.cases / profile_id / "source.pptx"
        case_path.parent.mkdir(parents=True, exist_ok=True)
        if not case_path.is_file():
            case_path.write_bytes(data)
        store.write_json(store.profiles / f"{profile_id}.json", saved)
        return saved
    finally:
        try:
            tmp_path.unlink()
        except OSError:
            pass


def get_profile(store: Store, profile_id: str) -> dict:
    path = store.profile_path(profile_id)
    # 先判存在再取锁：对不存在的 id 探测不留下永久锁（避免锁表随探测无限增长）。
    if not path.is_file():
        raise ApiError(404, "profile_not_found", f"找不到档案：{profile_id}")
    # 与 add_feedback 共用同一把 per-profile RLock（RLock 可重入）：
    # 反馈的「读→改→原子写」进行中时，生成/详情接口不会读到中间态（评审 F11）。
    with store.profile_lock(str(profile_id)):
        if not path.is_file():      # 锁内复查：并发删除时仍返回 404 而不是 500
            raise ApiError(404, "profile_not_found", f"找不到档案：{profile_id}")
        try:
            profile = kitlink.load_profile(store.root, profile_id)
        except FileNotFoundError:
            raise ApiError(404, "profile_not_found", f"找不到档案：{profile_id}") from None
        except Exception as exc:
            raise ApiError(500, "internal", f"读取档案失败：{exc}") from None
    if not isinstance(profile, dict):
        raise ApiError(500, "internal", "档案内容损坏")
    return profile


def _profile_summary(profile: dict) -> dict:
    summary: dict = {}
    for key in ("id", "name", "source", "created_at", "slide_size", "palette",
                "fonts", "style"):
        if key in profile:
            summary[key] = profile[key]
    metrics = profile.get("metrics") or {}
    if isinstance(metrics, dict):
        summary["slides"] = metrics.get("slides")
    learn = profile.get("learn") or {}
    if isinstance(learn, dict):
        summary["learn"] = {"rounds": learn.get("rounds", 0)}
    return summary


def list_profile_summaries(store: Store, limit: int = 100) -> list[dict]:
    try:
        profiles = kitlink.list_profiles(store.root)
    except Exception as exc:
        raise ApiError(500, "internal", f"读取档案列表失败：{exc}") from None
    items = [_profile_summary(p) for p in profiles if isinstance(p, dict)]
    items.sort(key=lambda p: str(p.get("created_at") or ""), reverse=True)
    return items[:limit]


def _persist_profile(store: Store, profile: dict) -> None:
    profile_id = str(profile.get("id") or "")
    if not valid_id(profile_id):
        raise ApiError(500, "internal", "档案缺少合法 id")
    payload = {k: v for k, v in profile.items() if not k.startswith("_")}
    store.write_json(store.profiles / f"{profile_id}.json", payload)


# --------------------------------------------------------------------------- #
# LLM（可选增强，OpenAI 兼容）
# --------------------------------------------------------------------------- #

def llm_config() -> dict | None:
    base = _text(os.environ.get("PPT_LLM_BASE_URL"), 300).rstrip("/")
    key = _text(os.environ.get("PPT_LLM_API_KEY"), 300)
    model = _text(os.environ.get("PPT_LLM_MODEL"), 100) or "deepseek-chat"
    if base and key:
        return {"base": base, "key": key, "model": model}
    return None


def llm_configured() -> bool:
    return llm_config() is not None


def _strip_fence(text: str) -> str:
    text = (text or "").strip()
    match = re.search(r"```(?:markdown|md)?\s*(.+?)```", text, re.S)
    if match:
        text = match.group(1).strip()
    return text


def llm_outline(prompt: str, *, pages: int = DEFAULT_PAGES, title: str = "",
                profile: dict | None = None, cfg: dict | None = None) -> str:
    cfg = cfg or llm_config()
    if cfg is None:
        raise RuntimeError("LLM 未配置")
    density = 1.0
    if isinstance(profile, dict):
        try:
            density = float((profile.get("style") or {}).get("density", 1.0))
        except (TypeError, ValueError):
            density = 1.0
    system = (
        "你是资深演示文稿策划。只输出 Markdown 大纲：每一页用一级标题 `# ` 开头，"
        "要点用 `- ` 列表，不要输出解释、不要用代码围栏、不要写 front matter。"
    )
    user = (
        f"主题：{title or prompt}\n"
        f"目标页数：约 {pages} 页（不含封面）\n"
        f"信息密度：{density:.2f}（1.0 为基准，越大每页要点越多）\n"
        f"用户提示词：{prompt}\n"
    )
    payload = {
        "model": cfg["model"],
        "temperature": 0.7,
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ],
    }
    request = _urlrequest.Request(
        cfg["base"] + "/chat/completions",
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "Authorization": "Bearer " + cfg["key"],
        },
        method="POST",
    )
    with _urlrequest.urlopen(request, timeout=LLM_TIMEOUT_SECONDS) as response:
        data = json.loads(response.read().decode("utf-8"))
    choices = data.get("choices") or []
    if not choices:
        raise RuntimeError("LLM 响应缺少 choices")
    content = (choices[0].get("message") or {}).get("content") or ""
    text = _strip_fence(content)
    if "#" not in text:
        raise RuntimeError("LLM 返回内容不是 Markdown 大纲")
    return text


# --------------------------------------------------------------------------- #
# 生成
# --------------------------------------------------------------------------- #

def _has_front_matter(text: str) -> bool:
    lines = text.split("\n")
    if not lines or lines[0].strip() not in ("---", "+++"):
        return False
    for line in lines[1:41]:
        stripped = line.strip()
        if stripped in ("---", "+++"):
            return True
        if stripped and ":" not in stripped:
            return False
    return False


def _compose_markdown(body: str, meta: dict, theme_key: str) -> str:
    body = (body or "").replace("\r\n", "\n").strip()
    if _has_front_matter(body):
        return body + "\n"
    header = []
    for key in ("title", "subtitle", "author", "org", "date"):
        value = _fold_field(meta.get(key))
        if value:
            header.append(f"{key}: {value}")
    header.append(f"theme: {theme_key}")
    return "---\n" + "\n".join(header) + "\n---\n\n" + body + "\n"


def _offline_outline(prompt: str, *, pages: int, profile: dict | None,
                     title: str) -> tuple[str, str | None]:
    """返回 (Markdown, 降级说明或 None)。

    调用 kit 的离线大纲时临时隐藏 LLM 环境变量：LLM 已经由服务端先试过一次，
    这里再试一次只会白白多等一个超时，而且会让 `llm_used` 与事实不一致。
    """
    try:
        with _llm_muted():
            return kitlink.build_outline(prompt, pages=pages, profile=profile,
                                         title=title), None
    except Exception as exc:
        outline = kitlink.local_build_outline(prompt, pages=pages, profile=profile,
                                              title=title)
        return outline, f"大纲内核异常，已使用内置大纲（{type(exc).__name__}）"


_LLM_ENV_KEYS = ("PPT_LLM_BASE_URL", "PPT_LLM_API_KEY", "PPT_LLM_MODEL",
                 "PPT_LLM_TIMEOUT")
_LLM_ENV_LOCK = threading.Lock()


@contextlib.contextmanager
def _llm_muted():
    """临时移除 LLM 配置（进程级环境变量，用锁串行化这段极短的窗口）。"""
    with _LLM_ENV_LOCK:
        saved = {key: os.environ.pop(key) for key in _LLM_ENV_KEYS
                 if key in os.environ}
        try:
            yield
        finally:
            os.environ.update(saved)


def _build_pptx(doc: dict, out: Path, theme_key: str, theme_obj,
                notes: list[str]) -> None:
    from pptxkit import builder

    if theme_obj is not None:
        try:
            builder.build(doc, out, theme_obj=theme_obj, base_dir=WORKSPACE_ROOT)
            return
        except TypeError as exc:
            if "theme_obj" not in str(exc):
                raise
            notes.append("生成内核暂不支持主题对象，档案仅影响大纲密度")
    builder.build(doc, out, theme=theme_key, base_dir=WORKSPACE_ROOT)


def _build_html(outline_path: Path, out: Path, theme_key: str, title: str,
                theme_obj=None) -> None:
    """网页版与 pptx 使用同一个 Theme（契约 §5.1）；图片只用占位。"""
    from pptxkit import html_deck

    try:
        html_deck.build(outline_path, out, theme=theme_key, title=title or None,
                        theme_obj=theme_obj, image_mode="placeholder")
    except TypeError:
        html_deck.build(outline_path, out, theme=theme_key, title=title or None)


def _parse_outline(markdown_text: str) -> dict:
    from pptxkit import markdown as md_parser

    return md_parser.parse(markdown_text)


def generate(store: Store, payload: dict | None) -> dict:
    payload = payload if isinstance(payload, dict) else {}
    title = _fold_field(payload.get("title"))[:120]
    subtitle = _fold_field(payload.get("subtitle"))[:200]
    author = _fold_field(payload.get("author"))[:80]
    org = _fold_field(payload.get("org"))[:120]
    date = _fold_field(payload.get("date"))[:60]
    prompt = _text(payload.get("prompt"), 4000)
    raw_outline = payload.get("outline")
    outline = "" if raw_outline is None else str(raw_outline)
    if outline:
        if len(outline.encode("utf-8")) > MAX_OUTLINE_BYTES:
            raise ApiError(400, "outline_too_large",
                           f"outline 超过 {MAX_OUTLINE_BYTES // 1024} KB 上限")
        outline = _text(outline, MAX_OUTLINE_BYTES)
    theme_key = (_text(payload.get("theme"), 40) or "academic").lower()
    pages = _int(payload.get("pages"), DEFAULT_PAGES, 2, 40)

    try:
        theme_key = kitlink.get_theme(theme_key).key
    except KeyError:
        raise ApiError(400, "invalid_theme",
                       f"未知主题：{payload.get('theme')}") from None

    profile = None
    profile_id = _text(payload.get("profile_id"), 80)
    if profile_id:
        profile = get_profile(store, profile_id)

    notes: list[str] = []
    cfg = llm_config()
    llm_used = False
    body = ""
    if outline:
        body = outline
        notes.append("使用请求提供的大纲")
    elif prompt:
        if cfg is not None:
            try:
                body = llm_outline(prompt, pages=pages, title=title,
                                   profile=profile, cfg=cfg)
                llm_used = True
                notes.append(f"已用 {cfg['model']} 扩写大纲")
            except Exception as exc:
                notes.append(f"LLM 调用失败，已降级为离线大纲（{type(exc).__name__}）")
        if not body:
            body, degrade_note = _offline_outline(prompt, pages=pages,
                                                  profile=profile, title=title)
            if degrade_note:
                notes.append(degrade_note)
            if cfg is None:
                notes.append("未配置 LLM，使用离线启发式大纲")
    else:
        raise ApiError(400, "missing_content", "请至少提供 outline 或 prompt")

    markdown_text = _compose_markdown(body, {
        "title": title, "subtitle": subtitle, "author": author,
        "org": org, "date": date,
    }, theme_key)
    _validate_outline_images(markdown_text)
    doc = _parse_outline(markdown_text)
    meta_parsed = doc.get("meta") or {}
    slides = doc.get("slides") or []
    pages_actual = len(slides) + (1 if meta_parsed.get("title") else 0)
    if pages_actual <= 0:
        raise ApiError(400, "empty_outline", "大纲里没有可用页面（请用 # 标题 分页）")
    if pages_actual > MAX_OUTLINE_PAGES:
        raise ApiError(
            400, "outline_too_large",
            f"大纲派生 {pages_actual} 页，超过上限 {MAX_OUTLINE_PAGES} 页")
    deck_title = (_text(meta_parsed.get("title"), 120) or title
                  or _first_heading(markdown_text) or prompt[:40] or "未命名演示")
    if pages_actual > 40:
        notes.append(f"大纲共 {pages_actual} 页（上限 {MAX_OUTLINE_PAGES} 页）")

    theme_obj = None
    if profile is not None:
        try:
            theme_obj = kitlink.apply_profile(profile, theme_key)
        except Exception as exc:
            notes.append(f"档案套用失败，回退到主题 {theme_key}（{type(exc).__name__}）")
            theme_obj = None
        else:
            notes.append(f"套用档案「{profile.get('name') or profile_id}」的配色与字号")

    deck_id = _new_id("dk")
    deck_dir = store.decks / deck_id
    deck_dir.mkdir(parents=True, exist_ok=True)
    outline_path = deck_dir / "outline.md"
    outline_path.write_text(markdown_text, encoding="utf-8")

    pptx_path = deck_dir / "deck.pptx"
    _build_pptx(doc, pptx_path, theme_key, theme_obj, notes)
    _build_html(outline_path, deck_dir / "deck.html", theme_key, deck_title, theme_obj)

    meta = {
        "id": deck_id,
        "created_at": _now_iso(),
        "title": deck_title,
        "theme": theme_key,
        "profile_id": profile_id or None,
        "pages": pages_actual,
        "files": {"pptx": "deck.pptx", "html": "deck.html"},
        "outline_chars": len(markdown_text),
        "llm_used": llm_used,
    }
    store.write_json(deck_dir / "meta.json", meta)

    return {
        "id": deck_id,
        "title": deck_title,
        "pages": pages_actual,
        "files": {
            "pptx": f"/files/{deck_id}/deck.pptx",
            "html": f"/files/{deck_id}/deck.html",
        },
        "llm_used": llm_used,
        "profile_id": profile_id or None,
        "theme": theme_key,
        "notes": notes,
    }


def _first_heading(text: str) -> str:
    match = re.search(r"^#\s+(.+)$", text or "", re.M)
    return _text(match.group(1), 120) if match else ""


# --------------------------------------------------------------------------- #
# deck 列表 / 读取 / 下载
# --------------------------------------------------------------------------- #

def list_decks(store: Store, limit: int = MAX_DECKS_LISTED) -> list[dict]:
    items: list[dict] = []
    if not store.decks.is_dir():
        return items
    for meta_path in store.decks.glob("*/meta.json"):
        try:
            meta = json.loads(meta_path.read_text(encoding="utf-8"))
        except Exception:
            continue
        if not isinstance(meta, dict):
            continue
        items.append({
            "id": meta.get("id") or meta_path.parent.name,
            "created_at": meta.get("created_at"),
            "title": meta.get("title"),
            "theme": meta.get("theme"),
            "profile_id": meta.get("profile_id"),
            "pages": meta.get("pages"),
            "llm_used": meta.get("llm_used"),
        })
    items.sort(key=lambda m: str(m.get("created_at") or ""), reverse=True)
    return items[:limit]


def get_deck(store: Store, deck_id: str) -> dict:
    meta_path = store.deck_dir(deck_id) / "meta.json"
    if not meta_path.is_file():
        raise ApiError(404, "deck_not_found", f"找不到 deck：{deck_id}")
    try:
        return json.loads(meta_path.read_text(encoding="utf-8"))
    except Exception as exc:
        raise ApiError(500, "internal", f"deck 元数据损坏：{exc}") from None


def deck_file_path(store: Store, deck_id: str, name: str) -> Path:
    if name not in DECK_FILES or not valid_id(deck_id):
        raise ApiError(404, "not_found", "文件不存在")
    path = store.decks / deck_id / name
    if not path.is_file():
        raise ApiError(404, "not_found", "文件不存在")
    return path


# --------------------------------------------------------------------------- #
# 反馈闭环
# --------------------------------------------------------------------------- #

def _feedback_message(before: dict, after: dict) -> str:
    try:
        density_before = float(before.get("density", 1.0))
    except (TypeError, ValueError):
        density_before = 1.0
    try:
        density_after = float(after.get("density", 1.0))
    except (TypeError, ValueError):
        density_after = 1.0
    try:
        title_before = float(before.get("title_scale", 1.0))
    except (TypeError, ValueError):
        title_before = 1.0
    try:
        title_after = float(after.get("title_scale", 1.0))
    except (TypeError, ValueError):
        title_after = 1.0
    parts = []
    if abs(density_after - density_before) > 1e-9:
        parts.append(f"密度 {density_after:.2f}")
    if abs(title_after - title_before) > 1e-9:
        parts.append(f"标题比例 {title_after:.2f}")
    if not parts:
        return f"已记录评分，本次无需调整（密度 {density_after:.2f}）"
    return "已调整：" + "，".join(parts)


def add_feedback(store: Store, deck_id: str, rating: int, notes: str = "") -> dict:
    meta = get_deck(store, deck_id)
    profile_id = meta.get("profile_id")
    if not profile_id:
        raise ApiError(400, "no_profile", "该 deck 没有关联风格档案，无法调整")
    profile_id = str(profile_id)
    # 读-改-写整体持锁：并发评分不能互相覆盖（评审 F6）。
    with store.profile_lock(profile_id):
        profile = get_profile(store, profile_id)
        before_style = dict(profile.get("style") or {})
        updated = kitlink.learn_profile(profile, int(rating), notes)
        if not isinstance(updated, dict):
            updated = profile
        _persist_profile(store, updated)
        style = dict(updated.get("style") or {})
        store.append_feedback({
            "at": _now_iso(),
            "deck_id": deck_id,
            "profile_id": profile_id,
            "rating": int(rating),
            "notes": notes,
            "style": style,
        })
    return {
        "ok": True,
        "profile_id": profile_id,
        "style": style,
        "message": _feedback_message(before_style, style),
    }
