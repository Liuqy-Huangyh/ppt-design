"""T3 前端静态验收：三件套、零外链、表单字段、元素 id、渲染安全。

只读 web/ 下三个文件，不启动服务、不联网、不依赖 server/ 与 kit/。

运行：
    & "$env:USERPROFILE\\.dsh\\dsh-runtimes\\dsh-primary-runtime\\dependencies\\python\\python.exe" tests\\test_web.py
"""
from __future__ import annotations

import re
import unittest
from html.parser import HTMLParser
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
WEB = ROOT / "web"
HTML_PATH = WEB / "index.html"
JS_PATH = WEB / "app.js"
CSS_PATH = WEB / "style.css"

VOID_TAGS = {
    "area", "base", "br", "col", "embed", "hr", "img", "input",
    "link", "meta", "param", "source", "track", "wbr",
}

EXTERNAL_PATTERNS = (
    re.compile(r"https?://"),
    re.compile(r"""["'(]//[A-Za-z0-9.-]+\.[A-Za-z]{2,}"""),
    re.compile(r"fonts\.googleapis", re.I),
    re.compile(r"cdn\.", re.I),
    re.compile(r"@import", re.I),
)


class PageIndex(HTMLParser):
    """收集 id -> (tag, attrs)、form 成员、script/link 标签。"""

    def __init__(self, text: str):
        super().__init__(convert_charrefs=True)
        self.ids: dict[str, tuple[str, dict]] = {}
        self.duplicate_ids: list[str] = []
        self.tags: list[tuple[str, dict]] = []
        self.form_members: dict[str, set[str]] = {}
        self._open: list[tuple[str, str | None]] = []
        self.feed(text)
        self.close()

    def _active_forms(self) -> list[str]:
        return [form for _, form in self._open if form]

    def _register(self, tag: str, attrs: dict) -> None:
        self.tags.append((tag, attrs))
        if "id" in attrs:
            if attrs["id"] in self.ids:
                self.duplicate_ids.append(attrs["id"])
            self.ids[attrs["id"]] = (tag, attrs)
            for form_id in self._active_forms():
                self.form_members.setdefault(form_id, set()).add(attrs["id"])
        if tag == "form" and "id" in attrs:
            self.form_members.setdefault(attrs["id"], set()).add(attrs["id"])

    def handle_starttag(self, tag, attrs):
        clean = {key: (value if value is not None else "") for key, value in attrs}
        self._register(tag, clean)
        if tag == "form" and "id" in clean:
            self._open.append((tag, clean["id"]))
        elif tag not in VOID_TAGS:
            self._open.append((tag, None))

    def handle_startendtag(self, tag, attrs):
        clean = {key: (value if value is not None else "") for key, value in attrs}
        self._register(tag, clean)

    def handle_endtag(self, tag):
        for index in range(len(self._open) - 1, -1, -1):
            if self._open[index][0] == tag:
                del self._open[index:]
                return


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


class WebAssetsTest(unittest.TestCase):
    """三件套存在、可解码、非空、零外链、零构建。"""

    @classmethod
    def setUpClass(cls):
        cls.html = read_text(HTML_PATH)
        cls.js = read_text(JS_PATH)
        cls.css = read_text(CSS_PATH)
        cls.page = PageIndex(cls.html)

    def test_three_files_exist(self):
        for path, minimum in ((HTML_PATH, 3000), (JS_PATH, 5000), (CSS_PATH, 5000)):
            self.assertTrue(path.is_file(), f"缺少 {path}")
            self.assertGreater(path.stat().st_size, minimum, f"{path.name} 内容过少")

    def test_files_are_utf8_without_bom(self):
        for path in (HTML_PATH, JS_PATH, CSS_PATH):
            raw = path.read_bytes()
            self.assertFalse(raw.startswith(b"\xef\xbb\xbf"), f"{path.name} 不应带 BOM")
            raw.decode("utf-8")

    def test_no_external_resources(self):
        for name, text in (("index.html", self.html), ("app.js", self.js), ("style.css", self.css)):
            for pattern in EXTERNAL_PATTERNS:
                match = pattern.search(text)
                self.assertIsNone(match, f"{name} 出现外链/外部资源：{match.group(0) if match else ''}")

    def test_no_frontend_framework(self):
        lowered = (self.js + self.css + self.html).lower()
        for forbidden in ("react", "vue.js", "vue@", "jquery", "bootstrap", "tailwind", "cdn"):
            self.assertNotIn(forbidden, lowered, f"不应引入前端框架/CDN：{forbidden}")
        self.assertNotIn("import ", self.js)
        self.assertNotIn("export ", self.js)

    def test_local_assets_only(self):
        scripts = [attrs for tag, attrs in self.page.tags if tag == "script"]
        self.assertEqual(len(scripts), 1, "只应有一个本地脚本")
        self.assertEqual(scripts[0].get("src"), "app.js")
        self.assertIn("defer", scripts[0])
        self.assertNotEqual(scripts[0].get("type", ""), "module")
        styles = [attrs for tag, attrs in self.page.tags
                  if tag == "link" and attrs.get("rel") == "stylesheet"]
        self.assertEqual(len(styles), 1)
        self.assertEqual(styles[0].get("href"), "style.css")

    def test_html_basics(self):
        self.assertIn('<html lang="zh-CN">', self.html)
        self.assertIn('<meta charset="utf-8">', self.html)
        self.assertIn('name="viewport"', self.html)
        self.assertIn("<title>", self.html)
        self.assertEqual(self.page.duplicate_ids, [], "存在重复 id")
        self.assertIn('href="#panel-generate"', self.html)


class GenerateFormTest(unittest.TestCase):
    """生成表单：字段、类型、约束。"""

    @classmethod
    def setUpClass(cls):
        cls.page = PageIndex(read_text(HTML_PATH))
        cls.members = cls.page.form_members.get("deck-form", set())

    def test_form_contains_all_fields(self):
        expected = {
            "f-title", "f-subtitle", "f-author", "f-date",
            "f-theme", "f-prompt", "f-outline", "f-pages", "f-profile",
        }
        missing = expected - self.members
        self.assertEqual(missing, set(), f"生成表单缺少字段：{sorted(missing)}")

    def test_text_fields(self):
        for field_id, name in (
            ("f-title", "title"), ("f-subtitle", "subtitle"),
            ("f-author", "author"), ("f-date", "date"),
        ):
            tag, attrs = self.page.ids[field_id]
            self.assertEqual(tag, "input")
            self.assertEqual(attrs.get("type"), "text")
            self.assertEqual(attrs.get("name"), name)

    def test_pages_is_bounded_number(self):
        tag, attrs = self.page.ids["f-pages"]
        self.assertEqual(tag, "input")
        self.assertEqual(attrs.get("type"), "number")
        self.assertEqual(attrs.get("name"), "pages")
        self.assertEqual(attrs.get("min"), "2")
        self.assertEqual(attrs.get("max"), "40")
        self.assertEqual(attrs.get("value"), "10")

    def test_theme_and_profile_are_selects(self):
        tag, attrs = self.page.ids["f-theme"]
        self.assertEqual(tag, "select")
        self.assertEqual(attrs.get("name"), "theme")
        tag, attrs = self.page.ids["f-profile"]
        self.assertEqual(tag, "select")
        self.assertEqual(attrs.get("name"), "profile_id")

    def test_prompt_and_outline_are_textareas(self):
        for field_id, name in (("f-prompt", "prompt"), ("f-outline", "outline")):
            tag, attrs = self.page.ids[field_id]
            self.assertEqual(tag, "textarea")
            self.assertEqual(attrs.get("name"), name)
        self.assertIn("mode-prompt", self.members)
        self.assertIn("mode-outline", self.members)

    def test_submit_and_result_surfaces(self):
        self.assertEqual(self.page.ids["gen-submit"][0], "button")
        self.assertEqual(self.page.ids["gen-submit"][1].get("type"), "submit")
        self.assertEqual(self.page.ids["gen-preview"][0], "iframe")
        preview_attrs = self.page.ids["gen-preview"][1]
        self.assertEqual(preview_attrs.get("width"), "1280")
        self.assertEqual(preview_attrs.get("height"), "720")
        self.assertIn("gen-pages", self.page.ids)
        pptx_tag, pptx_attrs = self.page.ids["gen-pptx"]
        self.assertEqual(pptx_tag, "a")
        self.assertIn("download", pptx_attrs)
        html_tag, html_attrs = self.page.ids["gen-html"]
        self.assertEqual(html_tag, "a")
        self.assertEqual(html_attrs.get("target"), "_blank")
        self.assertIn("noopener", html_attrs.get("rel", ""))
        self.assertEqual(self.page.ids["gen-error"][1].get("role"), "alert")


class CasesAndFeedbackTest(unittest.TestCase):
    """案例库上传与反馈表单。"""

    @classmethod
    def setUpClass(cls):
        cls.page = PageIndex(read_text(HTML_PATH))

    def test_upload_accepts_pptx(self):
        tag, attrs = self.page.ids["file-input"]
        self.assertEqual(tag, "input")
        self.assertEqual(attrs.get("type"), "file")
        self.assertIn(".pptx", attrs.get("accept", ""))
        self.assertIn(".ppt", attrs.get("accept", ""))
        self.assertIn("file-input", self.page.ids)
        self.assertEqual(self.page.ids["dropzone"][1].get("role"), "button")

    def test_profile_card_slots_exist(self):
        for element_id in ("profile-list", "profiles-empty", "upload-status", "upload-error"):
            self.assertIn(element_id, self.page.ids)
        self.assertIn("每页字数", read_text(JS_PATH))
        self.assertIn("每页要点", read_text(JS_PATH))
        self.assertIn("swatches", read_text(JS_PATH))
        self.assertIn("fonts", read_text(JS_PATH))

    def test_feedback_form_fields(self):
        members = self.page.form_members.get("feedback-form", set())
        self.assertEqual(self.page.ids["fb-deck"][0], "select")
        self.assertIn("fb-deck", members)
        for value in range(1, 6):
            element_id = f"fb-rating-{value}"
            self.assertIn(element_id, self.page.ids)
            tag, attrs = self.page.ids[element_id]
            self.assertEqual(tag, "input")
            self.assertEqual(attrs.get("type"), "radio")
            self.assertEqual(attrs.get("name"), "rating")
            self.assertEqual(attrs.get("value"), str(value))
        self.assertIn("fb-notes", members)
        self.assertEqual(self.page.ids["fb-notes"][0], "textarea")
        self.assertEqual(self.page.ids["fb-error"][1].get("role"), "alert")

    def test_state_placeholders_exist(self):
        for element_id in ("gen-empty", "gen-status", "fb-empty", "fb-message", "fb-style",
                           "preview-frame"):
            self.assertIn(element_id, self.page.ids)


class ScriptContractTest(unittest.TestCase):
    """app.js 是否按契约调用接口、是否只用 textContent 渲染、是否处理错误。"""

    @classmethod
    def setUpClass(cls):
        cls.js = read_text(JS_PATH)
        cls.html = read_text(HTML_PATH)
        cls.page = PageIndex(cls.html)

    def test_api_endpoints(self):
        for endpoint in ("/api/health", "/api/themes", "/api/profiles", "/api/generate", "/api/decks"):
            self.assertIn(endpoint, self.js, f"缺少接口调用：{endpoint}")
        self.assertIn("/feedback", self.js)
        self.assertIn("/files/", self.js)
        self.assertIn("encodeURIComponent", self.js)

    def test_generate_payload_fields(self):
        for key in ("'title'", "'subtitle'", "'author'", "'date'", "profile_id", "payload.pages", "payload[mode]"):
            self.assertIn(key, self.js, f"生成请求缺少字段处理：{key}")

    def test_upload_sends_raw_bytes_with_filename_header(self):
        self.assertIn("'X-Filename'", self.js)
        self.assertIn("body: file", self.js)
        self.assertIn("method: 'POST'", self.js)
        self.assertIn("/api/profiles", self.js)
        self.assertIn("MAX_UPLOAD_BYTES", self.js)

    def test_feedback_posts_rating(self):
        self.assertIn("rating: rating", self.js)
        self.assertIn("method: 'POST'", self.js)
        self.assertIn("'Content-Type': 'application/json'", self.js)

    def test_renders_with_textcontent_only(self):
        lowered = self.js.lower()
        for forbidden in ("innerhtml", "outerhtml", "insertadjacenthtml", "document.write", "eval("):
            self.assertNotIn(forbidden, lowered, f"禁止使用 {forbidden} 注入用户输入")
        self.assertGreaterEqual(self.js.count("textContent"), 25, "应统一用 textContent 渲染文本")
        self.assertIn("textContent", self.js)

    def test_error_message_is_surfaced(self):
        self.assertIn("payload.error", self.js)
        self.assertIn("detail.message", self.js)
        self.assertIn("error.message", self.js)
        self.assertIn("setError", self.js)

    def test_loading_and_empty_states(self):
        for phrase in ("正在生成", "正在分析", "生成中…", "提交中…"):
            self.assertIn(phrase, self.js)
        self.assertIn("aria-busy", self.js)
        self.assertIn("hidden", self.js)
        self.assertIn("profiles-empty", self.js)

    def test_mode_toggle_controls_content_source(self):
        self.assertIn("content-mode", self.js)
        self.assertIn("prompt-panel", self.js)
        self.assertIn("outline-panel", self.js)

    def test_referenced_ids_exist_in_html(self):
        patterns = (
            re.compile(r"""\$\(\s*['"]#([A-Za-z0-9_-]+)['"]"""),
            re.compile(r"""getElementById\(\s*['"]([A-Za-z0-9_-]+)['"]"""),
            re.compile(r"""querySelector\(\s*['"]#([A-Za-z0-9_-]+)['"]"""),
        )
        referenced = set()
        for pattern in patterns:
            referenced.update(pattern.findall(self.js))
        # '#f-' + key 这类动态前缀不是元素 id，跳过
        referenced = {element_id for element_id in referenced if not element_id.endswith("-")}
        self.assertTrue(referenced, "应从 app.js 中解析出元素 id")
        missing = sorted(element_id for element_id in referenced if element_id not in self.page.ids)
        self.assertEqual(missing, [], f"app.js 引用了 HTML 里不存在的 id：{missing}")

    def test_preview_uses_files_html(self):
        self.assertIn("files.html", self.js)
        self.assertIn("gen-preview", self.js)
        self.assertIn("preview.src", self.js)
        self.assertIn("preview-frame", self.js)
        self.assertIn("--preview-scale", self.js)


class StyleTest(unittest.TestCase):
    """杂志感视觉与响应式回落的可核对项。"""

    @classmethod
    def setUpClass(cls):
        cls.css = read_text(CSS_PATH)

    def test_editorial_tokens(self):
        for token in ("--paper:", "--klein:", "--serif:", "--rule:"):
            self.assertIn(token, self.css)
        self.assertIn("serif", self.css)
        self.assertIn("#002FA7", self.css)
        self.assertIn("Songti", self.css)

    def test_states_are_styled(self):
        for selector in (".error", ".empty", ".status", ".profile-card", ".swatch-chip",
                         ".rating-item", ".preview-frame iframe", "[hidden]"):
            self.assertIn(selector, self.css, f"缺少样式：{selector}")

    def test_responsive_and_motion(self):
        self.assertIn("@media (max-width: 900px)", self.css)
        self.assertIn("prefers-reduced-motion", self.css)
        self.assertIn("aspect-ratio", self.css)


if __name__ == "__main__":
    unittest.main(verbosity=2)
