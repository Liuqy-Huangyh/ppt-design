"""PPT Studio HTTP 服务：路由、静态托管、启动入口。

* 只用标准库（ThreadingHTTPServer），不联网也能跑。
* `create_server(host, port, store_dir)` 返回 `(httpd, base_url)`，端口传 0 时由系统分配，
  方便测试在同一进程的线程里起服务。
* 正式入口：`python -m server.app --host 0.0.0.0 --port 8765`
"""
from __future__ import annotations

import argparse
import json
import mimetypes
import os
import sys
import traceback
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote, urlsplit

from .jobs import (
    DECK_FILES,
    MAX_JSON_BYTES,
    MAX_UPLOAD_BYTES,
    VERSION,
    ApiError,
    Store,
    add_feedback,
    create_profile,
    deck_file_path,
    generate,
    get_deck,
    get_profile,
    list_decks,
    list_profile_summaries,
    llm_configured,
    normalize_filename,
)
from . import kitlink

ROOT = Path(__file__).resolve().parent.parent
WEB_DIR = ROOT / "web"

PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
MIME_TYPES = {
    ".html": "text/html; charset=utf-8",
    ".js": "application/javascript; charset=utf-8",
    ".mjs": "application/javascript; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".json": "application/json; charset=utf-8",
    ".svg": "image/svg+xml",
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".gif": "image/gif",
    ".webp": "image/webp",
    ".ico": "image/x-icon",
    ".woff": "font/woff",
    ".woff2": "font/woff2",
    ".txt": "text/plain; charset=utf-8",
    ".md": "text/markdown; charset=utf-8",
    ".pdf": "application/pdf",
}

PLACEHOLDER_PAGE = """<!DOCTYPE html>
<html lang="zh-CN">
<head><meta charset="utf-8"><title>PPT Studio</title>
<style>body{font:16px/1.7 "Microsoft YaHei",system-ui,sans-serif;max-width:640px;margin:12vh auto;padding:0 24px;color:#14213D}
code{background:#EEF2F8;padding:2px 6px;border-radius:4px}h1{font-size:28px}.muted{color:#5C6B82}</style></head>
<body>
<h1>PPT Studio 服务已启动</h1>
<p class="muted">前端页面（<code>web/index.html</code>）尚未就绪，接口已经可用：</p>
<ul>
  <li><code>GET /api/health</code> 健康检查</li>
  <li><code>GET /api/themes</code> 主题列表</li>
  <li><code>POST /api/profiles</code> 上传案例（原始 pptx 字节 + <code>X-Filename</code>）</li>
  <li><code>POST /api/generate</code> 生成 deck</li>
</ul>
<p class="muted">把前端文件放进 <code>web/</code> 目录后刷新本页即可使用。</p>
</body></html>
"""


class Handler(BaseHTTPRequestHandler):
    server_version = "PPTStudio/1.0"
    sys_version = ""

    # -- 入口 -------------------------------------------------------------- #

    def do_GET(self):  # noqa: N802 - http.server 命名约定
        self._dispatch("GET")

    def do_HEAD(self):  # noqa: N802 - 与 GET 一致，但只回响应头
        self._dispatch("GET")

    def do_POST(self):  # noqa: N802
        self._dispatch("POST")

    def do_OPTIONS(self):  # noqa: N802
        self.send_response(204)
        self._cors_headers()
        self.send_header("Allow", "GET, POST, OPTIONS")
        self.send_header("Content-Length", "0")
        self.end_headers()

    def log_message(self, fmt, *args):  # 静默访问日志，避免刷屏
        pass

    # -- 分发 -------------------------------------------------------------- #

    def _dispatch(self, method: str) -> None:
        try:
            raw_path = urlsplit(self.path).path or "/"
            segments = [unquote(part) for part in raw_path.split("/") if part]
            if any(part in (".", "..") for part in segments):
                raise ApiError(404, "not_found", "路径不存在")
            if method == "GET":
                self._route_get(raw_path, segments)
            else:
                self._route_post(segments)
        except ApiError as exc:
            self._send_error(exc.status, exc.code, exc.message)
        except (BrokenPipeError, ConnectionResetError):
            pass
        except Exception:
            traceback.print_exc()
            try:
                self._send_error(500, "internal", "服务端异常，请查看服务日志")
            except (BrokenPipeError, ConnectionResetError):
                pass

    def _route_get(self, raw_path: str, segments: list[str]) -> None:
        if not segments:
            self._serve_static("index.html")
            return
        head = segments[0]
        if head == "api":
            self._get_api(segments)
            return
        if head == "files":
            if len(segments) != 3:
                raise ApiError(404, "not_found", "文件不存在")
            self._serve_deck_file(segments[1], segments[2])
            return
        self._serve_static("/".join(segments))

    def _get_api(self, segments: list[str]) -> None:
        api = segments[1:]
        if api == ["health"]:
            self._send_json(200, {
                "status": "ok",
                "version": VERSION,
                "llm": llm_configured(),
            })
            return
        if api == ["themes"]:
            self._send_json(200, kitlink.theme_keys())
            return
        if api == ["profiles"]:
            self._send_json(200, list_profile_summaries(self.server.store))
            return
        if len(api) == 2 and api[0] == "profiles":
            self._send_json(200, get_profile(self.server.store, api[1]))
            return
        if api == ["decks"]:
            self._send_json(200, list_decks(self.server.store))
            return
        if len(api) == 2 and api[0] == "decks":
            self._send_json(200, get_deck(self.server.store, api[1]))
            return
        raise ApiError(404, "not_found", f"接口不存在：/{'/'.join(segments)}")

    def _route_post(self, segments: list[str]) -> None:
        if len(segments) < 2 or segments[0] != "api":
            raise ApiError(404, "not_found", "接口不存在")
        api = segments[1:]
        store: Store = self.server.store
        if api == ["profiles"]:
            filename = normalize_filename(self.headers.get("X-Filename") or "")
            self._validate_filename(filename)
            data = self._read_body(MAX_UPLOAD_BYTES)
            profile = create_profile(store, data, filename)
            self._send_json(200, {"id": profile.get("id"), "profile": profile})
            return
        if api == ["generate"]:
            payload = self._read_json()
            self._send_json(200, generate(store, payload))
            return
        if len(api) == 3 and api[0] == "decks" and api[2] == "feedback":
            deck_id = api[1]
            get_deck(store, deck_id)  # 先确认 deck 存在（404）
            payload = self._read_json()
            rating = payload.get("rating")
            if isinstance(rating, bool) or not isinstance(rating, int) or not 1 <= rating <= 5:
                raise ApiError(400, "invalid_rating", "rating 必须是 1-5 的整数")
            notes = payload.get("notes")
            notes = "" if notes is None else str(notes)[:1000]
            self._send_json(200, add_feedback(store, deck_id, rating, notes))
            return
        if api and api[0] in ("health", "themes"):
            raise ApiError(405, "method_not_allowed", "该接口只支持 GET")
        raise ApiError(404, "not_found", f"接口不存在：/{'/'.join(segments)}")

    # -- 静态文件 ---------------------------------------------------------- #

    def _serve_static(self, relative: str) -> None:
        base = WEB_DIR.resolve()
        candidate = (base / relative).resolve() if relative else base / "index.html"
        if base not in candidate.parents and candidate != base:
            raise ApiError(404, "not_found", "文件不存在")
        if candidate.is_dir():
            candidate = candidate / "index.html"
        if candidate.is_file():
            self._send_bytes(200, candidate.read_bytes(),
                             MIME_TYPES.get(candidate.suffix.lower())
                             or mimetypes.guess_type(candidate.name)[0]
                             or "application/octet-stream")
            return
        if relative in ("", "index.html"):
            self._send_bytes(200, PLACEHOLDER_PAGE.encode("utf-8"),
                             "text/html; charset=utf-8")
            return
        raise ApiError(404, "not_found", "文件不存在")

    def _serve_deck_file(self, deck_id: str, name: str) -> None:
        if name not in DECK_FILES:
            raise ApiError(404, "not_found", "文件不存在")
        path = deck_file_path(self.server.store, deck_id, name)
        if name.endswith(".pptx"):
            content_type = PPTX_MIME
            disposition = f'attachment; filename="{name}"'
        else:
            content_type = "text/html; charset=utf-8"
            disposition = "inline"
        self._send_bytes(200, path.read_bytes(), content_type, {
            "Content-Disposition": disposition,
        })

    # -- 响应helpers ------------------------------------------------------- #

    def _cors_headers(self, extra: dict | None = None) -> None:
        """默认同源：只有显式配置了允许来源才发 CORS 头（评审 F7）。

        web/ 页面与接口同源，不需要 CORS；`*` 会让任意网页借用户浏览器读写
        本机档案，默认关闭。需要跨域调试时用 PPT_ALLOW_ORIGIN 精确指定。
        """
        origin = getattr(self.server, "allow_origin", None)
        if origin:
            self.send_header("Access-Control-Allow-Origin", origin)
            self.send_header("Vary", "Origin")
            self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
            self.send_header("Access-Control-Allow-Headers", "Content-Type, X-Filename")
            self.send_header("Access-Control-Max-Age", "86400")
        for key, value in (extra or {}).items():
            self.send_header(key, value)

    def _send_json(self, status: int, payload, extra: dict | None = None) -> None:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self._send_bytes(status, data, "application/json; charset=utf-8", extra)

    def _send_error(self, status: int, code: str, message: str) -> None:
        self._send_json(status, {"error": {"code": code, "message": message}})

    def _send_bytes(self, status: int, data: bytes, content_type: str,
                    extra: dict | None = None) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self._cors_headers(extra)
        self.end_headers()
        if self.command != "HEAD" and data:
            self.wfile.write(data)

    # -- 请求体 ------------------------------------------------------------ #

    def _read_body(self, limit: int) -> bytes:
        raw_length = self.headers.get("Content-Length")
        if raw_length is None:
            if (self.headers.get("Transfer-Encoding") or "").lower().strip():
                raise ApiError(400, "bad_request", "不支持分块传输的请求体")
            return b""
        try:
            length = int(raw_length)
        except (TypeError, ValueError):
            raise ApiError(400, "bad_request", "Content-Length 非法") from None
        if length < 0:
            raise ApiError(400, "bad_request", "Content-Length 非法")
        if length > limit:
            self.close_connection = True
            raise ApiError(413, "payload_too_large",
                           f"请求体超过 {limit // (1024 * 1024)} MB 上限")
        chunks: list[bytes] = []
        remaining = length
        while remaining > 0:
            chunk = self.rfile.read(min(remaining, 65536))
            if not chunk:
                break
            chunks.append(chunk)
            remaining -= len(chunk)
        return b"".join(chunks)

    def _read_json(self) -> dict:
        raw = self._read_body(MAX_JSON_BYTES)
        if not raw.strip():
            return {}
        try:
            payload = json.loads(raw.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise ApiError(400, "invalid_json", f"请求体不是合法 JSON：{exc}") from None
        if not isinstance(payload, dict):
            raise ApiError(400, "bad_request", "请求体必须是 JSON 对象")
        return payload

    def _validate_filename(self, filename: str) -> None:
        name = (filename or "").strip().strip('"').strip("'")
        if not name:
            raise ApiError(400, "invalid_filename", "缺少 X-Filename 头（原始文件名）")
        if "\x00" in name or "/" in name or "\\" in name or ".." in name:
            raise ApiError(400, "invalid_filename", "文件名非法：不允许路径分隔符或 ..")
        ext = Path(name).suffix.lower()
        if ext not in (".pptx", ".ppt"):
            raise ApiError(400, "invalid_filename", "只支持 .pptx/.ppt 文件")


class StudioHTTPServer(ThreadingHTTPServer):
    daemon_threads = True
    allow_reuse_address = True

    def __init__(self, address, handler, store: Store, allow_origin: str | None = None):
        self.store = store
        self.allow_origin = allow_origin
        super().__init__(address, handler)


def _normalise_origin(value: str | None) -> str | None:
    origin = (value or "").strip().rstrip("/")
    if not origin:
        return None
    if origin == "*":
        # 通配 CORS + 无鉴权的本地服务风险过高，显式拒绝（见 docs/API.md §6）。
        return None
    return origin


def create_server(host: str = "127.0.0.1", port: int = 8765,
                  store_dir=None, allow_origin: str | None = None
                  ) -> tuple[StudioHTTPServer, str]:
    """创建（但不启动）HTTP 服务，返回 (httpd, base_url)。

    测试里传 port=0 让系统分配端口，然后在线程中 `httpd.serve_forever()`。
    allow_origin 默认从环境变量 PPT_ALLOW_ORIGIN 读取；未配置时只允许同源。
    """
    store = Store(store_dir or (Path(__file__).resolve().parent / "store"))
    if allow_origin is None:
        allow_origin = os.environ.get("PPT_ALLOW_ORIGIN")
    httpd = StudioHTTPServer((host, int(port)), Handler, store,
                             allow_origin=_normalise_origin(allow_origin))
    actual_port = httpd.server_address[1]
    display_host = "127.0.0.1" if host in ("", "0.0.0.0", "::") else host
    return httpd, f"http://{display_host}:{actual_port}"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="python -m server.app",
        description="PPT Studio 本地服务：生成 deck、管理风格档案、打分反馈",
    )
    parser.add_argument("--host", default="127.0.0.1", help="监听地址（0.0.0.0 供局域网访问）")
    parser.add_argument("--port", type=int, default=8765, help="监听端口")
    parser.add_argument("--store", default=None, help="数据目录（默认 server/store）")
    args = parser.parse_args(argv)

    httpd, base_url = create_server(args.host, args.port, args.store)
    print(f"PPT Studio 服务已启动：{base_url}", flush=True)
    print(f"  数据目录：{httpd.store.root}", flush=True)
    print(f"  LLM：{'已配置' if llm_configured() else '未配置（离线启发式大纲）'}", flush=True)
    if getattr(httpd, "allow_origin", None):
        print(f"  允许跨域来源：{httpd.allow_origin}", flush=True)
    if args.host not in ("127.0.0.1", "localhost", "::1"):
        print("  ⚠ 已绑定非回环地址：服务没有鉴权，任何能访问该端口的主机都可读写"
              "档案与 deck，请仅在可信网络使用。", flush=True)
    print("  按 Ctrl+C 停止", flush=True)
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n已停止。", flush=True)
    finally:
        httpd.server_close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
