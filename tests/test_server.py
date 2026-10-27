"""T1 服务端自测：同进程线程起服务，不 spawn 子进程、不联网。

运行：
    & "$env:USERPROFILE\\.dsh\\dsh-runtimes\\dsh-primary-runtime\\dependencies\\python\\python.exe" tests\\test_server.py
"""
from __future__ import annotations

import contextlib
import http.client
import io
import json
import os
import re
import shutil
import sys
import threading
import unittest
import urllib.error
import urllib.parse
import urllib.request
import uuid
import zipfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from pptx import Presentation                      # noqa: E402
from pptx.util import Inches, Pt                    # noqa: E402

from server.app import create_server                # noqa: E402
from server import jobs                             # noqa: E402

FIXTURE_ACADEMIC = ROOT / "tests" / "fixtures" / "case-academic.pptx"
ASSET_IMAGE = ROOT / "assets" / "figs" / "framework.png"


def make_sample_pptx(path: Path) -> Path:
    """用 python-pptx 造一个带主题/字体/图片的样例 deck（不依赖 kit、不依赖 out/）。"""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    cover = prs.slides.add_slide(prs.slide_layouts[0])
    cover.shapes.title.text = "案例模板 · 风格取样"
    cover.placeholders[1].text = "用于测试的示例演示文稿"

    for index in range(1, 4):
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = f"第 {index} 节 · 内容页"
        body = slide.placeholders[1].text_frame
        body.text = f"{index} 个要点：说明当前进展"
        for bullet in ("第一，梳理现状", "第二，拆解问题", "第三，给出方案"):
            paragraph = body.add_paragraph()
            paragraph.text = bullet
            paragraph.level = 0
            for run in paragraph.runs:
                run.font.size = Pt(19)
        for run in slide.shapes.title.text_frame.paragraphs[0].runs:
            run.font.size = Pt(31)

    prs.save(str(path))
    return path


def build_review_deck(path: Path, kind: str) -> Path:
    """构造用于 visual 评审闸门的坏/好样例（不依赖 WPS，纯 python-pptx）。"""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    def slide():
        return prs.slides.add_slide(prs.slide_layouts[6])

    def text(slide_obj, left, top, width, height, content, size):
        box = slide_obj.shapes.add_textbox(Inches(left), Inches(top),
                                           Inches(width), Inches(height))
        box.text_frame.word_wrap = True
        box.text_frame.text = content
        for paragraph in box.text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(size)
        return box

    if kind == "overflow":
        text(slide(), 1, 1, 2.0, 0.3,
             "这是一段明显超过文本框高度的正文，用来触发文字溢出闸门。" * 3, 20)
    elif kind == "out-of-bounds":
        text(slide(), 12.6, 2, 2.0, 0.8, "越界文本", 20)
    elif kind == "overlap":
        current = slide()
        text(current, 2.0, 2, 4, 1, "重叠的两个内容块", 20)
        text(current, 2.6, 2, 4, 1, "重叠的两个内容块", 20)
    elif kind == "small-font":
        text(slide(), 2, 2.5, 8, 1.2,
             "这是一段正文，字号只有十磅，明显低于十四磅的下限，阅读起来会比较吃力。" * 2, 10)
    elif kind == "blank":
        text(slide(), 1, 1, 1.2, 0.4, "空白", 16)
    elif kind == "stretch":
        from PIL import Image

        image_path = path.with_name("square.png")
        Image.new("RGB", (200, 200), (200, 60, 60)).save(image_path)
        slide().shapes.add_picture(str(image_path), Inches(1), Inches(1),
                                   Inches(5), Inches(1))
    elif kind == "good":
        current = slide()
        text(current, 1, 1.2, 8, 1, "季度复盘", 32)
        text(current, 1, 2.6, 10, 2,
             "本季度完成率 92%，平均交付周期 6.5 天，下一步聚焦三个关键动作。", 20)
    elif kind == "render":
        text(slide(), 1, 1, 1.2, 0.4, "空白", 16)
        current = slide()
        text(current, 1, 1.2, 8, 1, "季度复盘", 32)
        text(current, 1, 2.6, 10, 2,
             "本季度完成率 92%，平均交付周期 6.5 天，下一步聚焦三个关键动作。", 20)
    else:
        raise ValueError(f"未知样例类型：{kind}")
    prs.save(str(path))
    return path


class ServerTestCase(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        # 确保离线路径：测试期间不配置 LLM
        for key in ("PPT_LLM_BASE_URL", "PPT_LLM_API_KEY", "PPT_LLM_MODEL"):
            os.environ.pop(key, None)
        # CORS 默认同源的断言不能被外部环境变量影响
        os.environ.pop("PPT_ALLOW_ORIGIN", None)
        # 临时目录放在工作区 .tmp 下（系统 temp / mkdtemp 在沙箱里 ACL 受限）
        tmp_root = ROOT / ".tmp"
        tmp_root.mkdir(parents=True, exist_ok=True)
        cls.tmpdir = tmp_root / f"server-test-{uuid.uuid4().hex[:8]}"
        cls.tmpdir.mkdir(parents=True, exist_ok=True)
        cls.store_dir = cls.tmpdir / "store"
        cls.httpd, cls.base_url = create_server("127.0.0.1", 0, cls.store_dir)
        cls.port = cls.httpd.server_address[1]
        cls.thread = threading.Thread(target=cls.httpd.serve_forever,
                                      name="ppt-studio-test", daemon=True)
        cls.thread.start()
        cls.sample = make_sample_pptx(cls.tmpdir / "案例.pptx")
        cls.profile_id = None
        cls.profile_deck_id = None
        cls.outline_deck_id = None

    @classmethod
    def tearDownClass(cls):
        cls.httpd.shutdown()
        cls.httpd.server_close()
        cls.thread.join(timeout=5)
        shutil.rmtree(cls.tmpdir, ignore_errors=True)

    # -- helpers ---------------------------------------------------------- #

    def request(self, method, path, data=None, headers=None, timeout=30):
        url = self.base_url + path
        request = urllib.request.Request(url, data=data, method=method,
                                         headers=headers or {})
        try:
            with urllib.request.urlopen(request, timeout=timeout) as response:
                return response.status, response.read(), dict(response.headers)
        except urllib.error.HTTPError as exc:
            return exc.code, exc.read(), dict(exc.headers)

    def get_json(self, path):
        status, body, _ = self.request("GET", path)
        self.assertEqual(status, 200, f"GET {path} -> {status}: {body[:300]!r}")
        return json.loads(body.decode("utf-8"))

    def post_json(self, path, payload, headers=None):
        head = {"Content-Type": "application/json; charset=utf-8"}
        head.update(headers or {})
        status, body, _ = self.request(
            "POST", path, data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
            headers=head)
        return status, json.loads(body.decode("utf-8"))

    def assert_error(self, status, payload, expected_status, expected_code):
        self.assertEqual(status, expected_status, payload)
        self.assertIn("error", payload)
        error = payload["error"]
        self.assertEqual(set(error), {"code", "message"})
        self.assertEqual(error["code"], expected_code, error)
        self.assertTrue(error["message"], "错误信息不应为空")

    # -- 01-03：契约规定的只读接口 ---------------------------------------- #

    def test_01_health(self):
        status, body, headers = self.request("GET", "/api/health")
        self.assertEqual(status, 200)
        self.assertIn("application/json", headers.get("Content-Type", ""))
        payload = json.loads(body.decode("utf-8"))
        self.assertEqual(payload["status"], "ok")
        self.assertEqual(payload["version"], "1.0.0")
        self.assertIs(payload["llm"], False)

    def test_02_themes(self):
        themes = self.get_json("/api/themes")
        self.assertIsInstance(themes, list)
        self.assertGreaterEqual(len(themes), 3)
        for item in themes:
            self.assertEqual(set(item), {"key", "label"})
        keys = {item["key"] for item in themes}
        self.assertIn("academic", keys)
        self.assertIn("editorial", keys)
        labels = {item["key"]: item["label"] for item in themes}
        self.assertEqual(labels["academic"], "学术答辩")

    def test_03_profiles_initially_empty(self):
        self.assertEqual(self.get_json("/api/profiles"), [])

    # -- 04-06：生成的三条路径 --------------------------------------------- #

    def test_04_generate_from_outline(self):
        outline = ("# 季度复盘\n- 三件事：进度、问题、下一步\n- 用数据说话\n\n"
                   "# 数据回顾\n- 完成率 92%\n- 平均交付周期 6.5 天\n\n"
                   "# 下一步\n- 聚焦三个关键动作")
        status, payload = self.post_json("/api/generate", {
            "title": "季度复盘",
            "subtitle": "2026 Q3 交付回顾",
            "author": "张三",
            "org": "产品部",
            "date": "2026-10-27",
            "theme": "academic",
            "outline": outline,
        })
        self.assertEqual(status, 200, payload)
        self.assertTrue(payload["id"].startswith("dk-"), payload)
        self.assertEqual(payload["llm_used"], False)
        self.assertGreaterEqual(payload["pages"], 3)
        self.assertEqual(payload["files"]["pptx"], f"/files/{payload['id']}/deck.pptx")
        self.assertEqual(payload["files"]["html"], f"/files/{payload['id']}/deck.html")
        self.assertIn("notes", payload)
        self.assertTrue(any("大纲" in note for note in payload["notes"]), payload["notes"])
        type(self).outline_deck_id = payload["id"]

        deck_dir = self.store_dir / "decks" / payload["id"]
        for name in ("deck.pptx", "deck.html", "outline.md", "meta.json"):
            self.assertTrue((deck_dir / name).is_file(), f"缺少 {name}")
        meta = json.loads((deck_dir / "meta.json").read_text(encoding="utf-8"))
        self.assertEqual(meta["id"], payload["id"])
        self.assertEqual(meta["pages"], payload["pages"])
        self.assertEqual(meta["theme"], "academic")
        self.assertEqual(meta["files"], {"pptx": "deck.pptx", "html": "deck.html"})
        self.assertEqual(meta["llm_used"], False)
        self.assertGreater(meta["outline_chars"], 0)

        # 返回的下载地址真的能下到文件
        status, body, headers = self.request("GET", payload["files"]["pptx"])
        self.assertEqual(status, 200)
        self.assertTrue(body.startswith(b"PK"), body[:16])
        self.assertIn("presentationml", headers.get("Content-Type", ""))
        status, body, _ = self.request("GET", payload["files"]["html"])
        self.assertEqual(status, 200)
        self.assertIn(b"<html", body.lower())

        # pptx 必须真的能打开，且页数与接口返回一致
        prs = Presentation(str(deck_dir / "deck.pptx"))
        self.assertEqual(len(prs.slides), payload["pages"])

    def test_05_generate_from_prompt_offline(self):
        status, payload = self.post_json("/api/generate", {
            "title": "AI 与产品效率",
            "prompt": "给团队讲我们怎么把 PPT 制作自动化，强调可编辑与风格一致",
            "theme": "minimal",
            "pages": 8,
        })
        self.assertEqual(status, 200, payload)
        self.assertEqual(payload["llm_used"], False)
        self.assertTrue(any("离线" in note for note in payload["notes"]),
                        payload["notes"])
        self.assertGreaterEqual(payload["pages"], 2)
        deck_dir = self.store_dir / "decks" / payload["id"]
        outline_text = (deck_dir / "outline.md").read_text(encoding="utf-8")
        self.assertIn("# ", outline_text)
        meta = json.loads((deck_dir / "meta.json").read_text(encoding="utf-8"))
        self.assertEqual(meta["theme"], "minimal")
        self.assertTrue((deck_dir / "deck.pptx").stat().st_size > 0)

    def test_06_outline_wins_over_prompt(self):
        status, payload = self.post_json("/api/generate", {
            "title": "优先级测试",
            "prompt": "PROMPT-SHOULD-NOT-BE-USED",
            "outline": "# PRIORITY-MARKER\n- 只用 outline\n- 忽略 prompt",
        })
        self.assertEqual(status, 200, payload)
        self.assertFalse(payload["llm_used"])
        outline_text = (self.store_dir / "decks" / payload["id"] / "outline.md").read_text(
            encoding="utf-8")
        self.assertIn("PRIORITY-MARKER", outline_text)
        self.assertNotIn("PROMPT-SHOULD-NOT-BE-USED", outline_text)
        self.assertTrue((self.store_dir / "decks" / payload["id"] / "deck.html").is_file())

    # -- 07-08：案例上传 → 档案 → 带档案生成 → 反馈 ------------------------ #

    def test_07_profile_upload_and_use(self):
        raw = self.sample.read_bytes()
        encoded = urllib.parse.quote("案例.pptx")
        status, body, _ = self.request(
            "POST", "/api/profiles", data=raw,
            headers={
                "Content-Type": "application/octet-stream",
                "X-Filename": encoded,
            })
        self.assertEqual(status, 200, body[:400])
        payload = json.loads(body.decode("utf-8"))
        self.assertTrue(payload["id"].startswith("pf-"), payload)
        profile = payload["profile"]
        self.assertEqual(profile["id"], payload["id"])
        # §4.1 数据模型必备字段
        for key in ("id", "name", "created_at", "slide_size", "palette", "fonts",
                    "metrics", "style", "learn"):
            self.assertIn(key, profile, key)
        self.assertIn("案例", profile["name"])
        self.assertEqual(set(profile["palette"]),
                         {"bg", "ink", "muted", "accent", "accent_soft", "rule"})
        for key, value in profile["palette"].items():
            self.assertRegex(value, r"^#[0-9A-F]{6}$", key)
        self.assertGreaterEqual(profile["metrics"]["slides"], 4)
        # 落盘：case 原文 + profile.json
        case_path = self.store_dir / "cases" / payload["id"] / "source.pptx"
        self.assertTrue(case_path.is_file())
        self.assertEqual(case_path.read_bytes(), raw)
        profile_path = self.store_dir / "profiles" / f"{payload['id']}.json"
        self.assertTrue(profile_path.is_file())
        on_disk = json.loads(profile_path.read_text(encoding="utf-8"))
        self.assertEqual(on_disk["id"], payload["id"])
        self.assertEqual(on_disk["source"], "source.pptx")
        type(self).profile_id = payload["id"]

        # GET 单个 / 列表
        full = self.get_json(f"/api/profiles/{payload['id']}")
        self.assertEqual(full["id"], payload["id"])
        summaries = self.get_json("/api/profiles")
        self.assertEqual(len(summaries), 1)
        self.assertEqual(summaries[0]["id"], payload["id"])
        self.assertIn("palette", summaries[0])
        self.assertNotIn("metrics", summaries[0])

        # 用档案生成：pptx 仍然出得来（kit 未支持 theme_obj 时也能降级）
        status, generated = self.post_json("/api/generate", {
            "title": "案例驱动的 deck",
            "prompt": "按照上传案例的风格讲一讲下季度的计划",
            "profile_id": payload["id"],
            "theme": "editorial",
        })
        self.assertEqual(status, 200, generated)
        self.assertEqual(generated["profile_id"], payload["id"])
        self.assertEqual(generated["theme"], "editorial")
        self.assertTrue(any("档案" in note for note in generated["notes"]),
                        generated["notes"])
        meta = json.loads((self.store_dir / "decks" / generated["id"] / "meta.json")
                          .read_text(encoding="utf-8"))
        self.assertEqual(meta["profile_id"], payload["id"])
        type(self).profile_deck_id = generated["id"]

    def test_08_feedback_adjusts_profile(self):
        self.assertIsNotNone(self.profile_deck_id, "test_07 应先创建带档案的 deck")
        before = self.get_json(f"/api/profiles/{self.profile_id}")
        density_before = before["style"]["density"]

        status, payload = self.post_json(f"/api/decks/{self.profile_deck_id}/feedback",
                                         {"rating": 5, "notes": "内容再密一点"})
        self.assertEqual(status, 200, payload)
        self.assertTrue(payload["ok"])
        self.assertEqual(payload["profile_id"], self.profile_id)
        self.assertIn("style", payload)
        self.assertTrue(payload["message"].startswith("已调整"), payload["message"])
        density_up = payload["style"]["density"]
        self.assertGreater(density_up, density_before)

        # 落盘已更新
        updated = self.get_json(f"/api/profiles/{self.profile_id}")
        self.assertAlmostEqual(updated["style"]["density"], density_up, places=3)
        self.assertEqual(updated["learn"]["rounds"], 1)
        self.assertEqual(len(updated["learn"]["history"]), 1)
        self.assertEqual(updated["learn"]["history"][0]["rating"], 5)

        status, payload = self.post_json(f"/api/decks/{self.profile_deck_id}/feedback",
                                         {"rating": 1, "notes": "字太多了"})
        self.assertEqual(status, 200, payload)
        density_down = payload["style"]["density"]
        self.assertLess(density_down, density_up)
        self.assertLess(payload["style"]["title_scale"], 1.0)
        self.assertIn("密度", payload["message"])

        feedback_path = self.store_dir / "feedback.jsonl"
        self.assertTrue(feedback_path.is_file())
        lines = [line for line in feedback_path.read_text(encoding="utf-8").splitlines()
                 if line.strip()]
        self.assertEqual(len(lines), 2)
        self.assertEqual(json.loads(lines[-1])["rating"], 1)

    # -- 09-13：错误请求 --------------------------------------------------- #

    def test_09_missing_content(self):
        status, payload = self.post_json("/api/generate", {})
        self.assert_error(status, payload, 400, "missing_content")
        status, payload = self.post_json("/api/generate", {"title": "只有标题"})
        self.assert_error(status, payload, 400, "missing_content")

    def test_10_invalid_json(self):
        status, body, _ = self.request(
            "POST", "/api/generate", data=b"{not json",
            headers={"Content-Type": "application/json"})
        payload = json.loads(body.decode("utf-8"))
        self.assert_error(status, payload, 400, "invalid_json")

        status, body, _ = self.request(
            "POST", "/api/generate", data=b"[]",
            headers={"Content-Type": "application/json"})
        payload = json.loads(body.decode("utf-8"))
        self.assert_error(status, payload, 400, "bad_request")

    def test_11_upload_validation(self):
        raw = self.sample.read_bytes()

        status, body, _ = self.request("POST", "/api/profiles", data=raw)
        self.assert_error(status, json.loads(body.decode("utf-8")), 400, "invalid_filename")

        status, body, _ = self.request(
            "POST", "/api/profiles", data=raw,
            headers={"X-Filename": urllib.parse.quote("案例.txt")})
        self.assert_error(status, json.loads(body.decode("utf-8")), 400, "invalid_filename")

        status, body, _ = self.request(
            "POST", "/api/profiles", data=raw,
            headers={"X-Filename": "../../evil.pptx"})
        self.assert_error(status, json.loads(body.decode("utf-8")), 400, "invalid_filename")

        status, body, _ = self.request(
            "POST", "/api/profiles", data=b"",
            headers={"X-Filename": "empty.pptx"})
        self.assert_error(status, json.loads(body.decode("utf-8")), 400, "empty_body")

        status, body, _ = self.request(
            "POST", "/api/profiles", data=b"this is not a pptx",
            headers={"X-Filename": "broken.pptx"})
        self.assert_error(status, json.loads(body.decode("utf-8")), 400, "invalid_pptx")

        # 失败的上传不应留下档案
        self.assertEqual(len(self.get_json("/api/profiles")), 1)

    def test_12_upload_too_large(self):
        connection = http.client.HTTPConnection("127.0.0.1", self.port, timeout=15)
        try:
            connection.putrequest("POST", "/api/profiles")
            connection.putheader("X-Filename", "big.pptx")
            connection.putheader("Content-Length", str(26 * 1024 * 1024))
            connection.putheader("Content-Type", "application/octet-stream")
            connection.endheaders()
            response = connection.getresponse()
            body = response.read()
            self.assertEqual(response.status, 413, body[:200])
            payload = json.loads(body.decode("utf-8"))
            self.assert_error(413, payload, 413, "payload_too_large")
        finally:
            connection.close()

    def test_13_files_and_paths(self):
        deck_id = self.outline_deck_id
        self.assertIsNotNone(deck_id)
        for name in ("outline.md", "meta.json", "source.pptx", "evil.pptx"):
            status, body, _ = self.request("GET", f"/files/{deck_id}/{name}")
            self.assertEqual(status, 404, name)
            self.assertEqual(json.loads(body.decode("utf-8"))["error"]["code"], "not_found")
        status, body, _ = self.request("GET", f"/files/dk-20000101-dead/deck.pptx")
        self.assertEqual(status, 404)
        status, body, _ = self.request("GET", f"/files/{deck_id}/..%2F..%2Fmeta.json")
        self.assertEqual(status, 404)
        for path, code in (
            ("/api/decks/dk-20000101-dead", "deck_not_found"),
            ("/api/profiles/pf-20000101-dead", "profile_not_found"),
            ("/api/unknown-endpoint", "not_found"),
            ("/no-such-asset.js", "not_found"),
        ):
            status, body, _ = self.request("GET", path)
            self.assertEqual(status, 404, path)
            payload = json.loads(body.decode("utf-8"))
            self.assertEqual(payload["error"]["code"], code, path)

    # -- 14-16：列表 / 校验 / 静态托管 ------------------------------------ #

    def test_14_decks_list_and_meta(self):
        decks = self.get_json("/api/decks")
        self.assertIsInstance(decks, list)
        self.assertGreaterEqual(len(decks), 3)
        ids = [item["id"] for item in decks]
        self.assertIn(self.outline_deck_id, ids)
        for item in decks:
            self.assertEqual(set(item), {"id", "created_at", "title", "theme",
                                         "profile_id", "pages", "llm_used"})
        meta = self.get_json(f"/api/decks/{self.outline_deck_id}")
        self.assertEqual(meta["id"], self.outline_deck_id)
        self.assertIn("outline_chars", meta)

    def test_15_feedback_validation(self):
        deck_id = self.outline_deck_id
        status, payload = self.post_json(f"/api/decks/{deck_id}/feedback",
                                         {"rating": 9})
        self.assert_error(status, payload, 400, "invalid_rating")
        status, payload = self.post_json(f"/api/decks/{deck_id}/feedback", {})
        self.assert_error(status, payload, 400, "invalid_rating")
        status, payload = self.post_json(f"/api/decks/{deck_id}/feedback",
                                         {"rating": 4})
        self.assert_error(status, payload, 400, "no_profile")
        status, payload = self.post_json("/api/decks/dk-20000101-dead/feedback",
                                         {"rating": 4})
        self.assert_error(status, payload, 404, "deck_not_found")

    def test_16_static_and_methods(self):
        status, body, headers = self.request("GET", "/")
        self.assertEqual(status, 200)
        self.assertIn("text/html", headers.get("Content-Type", ""))
        self.assertIn(b"<", body)
        # 只支持 GET 的接口收到 POST → 405，仍然是统一 error 体
        status, body, _ = self.request("POST", "/api/health", data=b"{}")
        self.assertEqual(status, 405)
        self.assertEqual(json.loads(body.decode("utf-8"))["error"]["code"],
                         "method_not_allowed")

    def test_17_base_url(self):
        self.assertRegex(self.base_url, r"^http://127\.0\.0\.1:\d+$")
        self.assertNotEqual(self.port, 0)

    # -- 18-25：修复轮回归（评审 F1/F3/F4/F5/F6/F7/F8/F9/F10） --------------- #

    def fetch(self, url, method="GET"):
        request = urllib.request.Request(url, method=method)
        try:
            with urllib.request.urlopen(request, timeout=20) as response:
                return response.status, response.read(), dict(response.headers)
        except urllib.error.HTTPError as exc:
            return exc.code, exc.read(), dict(exc.headers)

    def test_18_outline_images_are_sandboxed(self):
        """F1：绝对路径 / .. 逃逸 → 400 invalid_outline 且不落盘；assets 相对路径可用。"""
        decks_dir = self.store_dir / "decks"
        before = {p.name for p in decks_dir.iterdir()} if decks_dir.is_dir() else set()

        cases = {
            "absolute": "# 图片页\n- 要点\n\n![](C:/Windows/Web/Wallpaper/Spotlight/img14.jpg)",
            "unc": "# 图片页\n- 要点\n\n![](//server/share/x.png)",
            "dotdot": "# 图片页\n- 要点\n\n![](../secret.png)",
            "outside_assets": "# 图片页\n- 要点\n\n![](out/academic-defense.pptx)",
            "missing": "# 图片页\n- 要点\n\n![](assets/figs/nope.png)",
        }
        for label, outline in cases.items():
            status, payload = self.post_json("/api/generate", {"title": label, "outline": outline})
            self.assert_error(status, payload, 400, "invalid_outline")
        after = {p.name for p in decks_dir.iterdir()} if decks_dir.is_dir() else set()
        self.assertEqual(before, after, "非法图片请求不应创建 deck 目录")

        self.assertTrue(ASSET_IMAGE.is_file(), f"缺少测试用图片：{ASSET_IMAGE}")
        status, payload = self.post_json("/api/generate", {
            "title": "合法图片",
            "outline": "# 图片页\n- 要点\n\n![框架图](assets/figs/framework.png)",
        })
        self.assertEqual(status, 200, payload)
        deck_dir = decks_dir / payload["id"]
        with zipfile.ZipFile(deck_dir / "deck.pptx") as archive:
            media = [name for name in archive.namelist()
                     if name.startswith("ppt/media/")]
        self.assertTrue(media, "合法相对路径图片应嵌入 pptx")
        html = (deck_dir / "deck.html").read_text(encoding="utf-8")
        self.assertNotIn("C:/", html)
        self.assertNotIn("C:\\", html)
        self.assertNotIn("assets/figs/framework.png", html)
        self.assertIn("图片未随本网页提供", html)

    def test_19_outline_limits(self):
        """F4：页数 ≤60、outline ≤64KB，超出 400 outline_too_large。"""
        eight_hundred = "\n\n".join(f"# 第{i}页\n- 要点" for i in range(800))
        status, payload = self.post_json("/api/generate", {"title": "800", "outline": eight_hundred})
        self.assert_error(status, payload, 400, "outline_too_large")

        sixty_one = "\n\n".join(f"# 第{i}页\n- 要点" for i in range(61))
        status, payload = self.post_json("/api/generate", {"title": "61", "outline": sixty_one})
        self.assert_error(status, payload, 400, "outline_too_large")

        fifty_nine = "\n\n".join(f"# 第{i}页\n- 要点" for i in range(59))
        status, payload = self.post_json("/api/generate", {"title": "边界", "outline": fifty_nine})
        self.assertEqual(status, 200, payload)
        self.assertEqual(payload["pages"], 60)
        self.assertTrue(any("60" in note for note in payload["notes"]), payload["notes"])

        huge = "# 大\n- " + "a" * (70 * 1024)
        status, payload = self.post_json("/api/generate", {"title": "huge", "outline": huge})
        self.assert_error(status, payload, 400, "outline_too_large")

    def test_20_front_matter_fields_are_folded(self):
        """F5：标题里的换行/`---` 不得注入 front matter 结构。"""
        injected = "A\n---\n# 注入页\n- x"
        status, payload = self.post_json("/api/generate", {
            "title": injected,
            "outline": "# 正文\n- 一",
        })
        self.assertEqual(status, 200, payload)
        self.assertEqual(payload["title"], "A --- # 注入页 - x")
        self.assertEqual(payload["pages"], 2, payload)
        outline = (self.store_dir / "decks" / payload["id"] / "outline.md").read_text(
            encoding="utf-8")
        self.assertNotRegex(outline, r"(?m)^#\s*注入页")
        self.assertIn("title: A --- # 注入页 - x", outline)

    def test_21_feedback_concurrent_updates(self):
        """F6：并发评分不得丢更新（每轮 learn.rounds 都要 +1）。"""
        self.assertIsNotNone(self.profile_deck_id, "test_07 应先创建带档案的 deck")
        before = self.get_json(f"/api/profiles/{self.profile_id}")
        rounds_before = before["learn"]["rounds"]
        history_before = len(before["learn"]["history"])
        density_before = before["style"]["density"]

        def submit(_):
            return self.post_json(f"/api/decks/{self.profile_deck_id}/feedback",
                                  {"rating": 5, "notes": "并发"})

        with ThreadPoolExecutor(max_workers=8) as pool:
            results = list(pool.map(submit, range(10)))
        for status, payload in results:
            self.assertEqual(status, 200, payload)
            self.assertTrue(payload["ok"])

        after = self.get_json(f"/api/profiles/{self.profile_id}")
        self.assertEqual(after["learn"]["rounds"], rounds_before + 10)
        self.assertEqual(len(after["learn"]["history"]), history_before + 10)
        self.assertGreater(after["style"]["density"], density_before)

    def test_22_cors_defaults_same_origin(self):
        """F7：默认不发 CORS 头；只有显式配置来源才允许跨域。"""
        status, body, headers = self.request("GET", "/api/health")
        self.assertEqual(status, 200)
        self.assertNotIn("Access-Control-Allow-Origin", headers)
        status, body, headers = self.request("OPTIONS", "/api/generate")
        self.assertNotIn("Access-Control-Allow-Origin", headers)

        origin = "http://127.0.0.1:1234"
        extra_dir = self.store_dir.parent / "cors-store"
        httpd, base_url = create_server("127.0.0.1", 0, extra_dir, allow_origin=origin)
        thread = threading.Thread(target=httpd.serve_forever, daemon=True)
        thread.start()
        try:
            status, body, headers = self.fetch(base_url + "/api/health")
            self.assertEqual(status, 200)
            self.assertEqual(headers.get("Access-Control-Allow-Origin"), origin)
            status, body, headers = self.fetch(base_url + "/api/health", method="OPTIONS")
            self.assertEqual(headers.get("Access-Control-Allow-Origin"), origin)
        finally:
            httpd.shutdown()
            httpd.server_close()
            thread.join(timeout=5)

    def test_23_generate_returns_title_and_unique_ids(self):
        """F8/F9：响应带 title；同日多次上传 id 唯一且为 8+ 位 hex。"""
        status, payload = self.post_json("/api/generate", {
            "prompt": "没有标题时用提示词首行做标题",
        })
        self.assertEqual(status, 200, payload)
        self.assertIn("title", payload)
        self.assertTrue(payload["title"])
        meta = self.get_json(f"/api/decks/{payload['id']}")
        self.assertEqual(payload["title"], meta["title"])
        self.assertRegex(payload["id"], r"^dk-\d{8}-[0-9a-f]{8,}$")

        self.assertTrue(FIXTURE_ACADEMIC.is_file(), f"缺少只读样例：{FIXTURE_ACADEMIC}")
        raw = FIXTURE_ACADEMIC.read_bytes()
        ids = []
        for _ in range(12):
            status, body, _ = self.request(
                "POST", "/api/profiles", data=raw,
                headers={"X-Filename": "case-academic.pptx"})
            self.assertEqual(status, 200, body[:300])
            profile = json.loads(body.decode("utf-8"))
            ids.append(profile["id"])
            self.assertRegex(profile["id"], r"^pf-\d{8}-[0-9a-f]{8,}$")
            self.assertTrue((self.store_dir / "profiles" / f"{profile['id']}.json").is_file())
            self.assertTrue((self.store_dir / "cases" / profile["id"] / "source.pptx").is_file())
        self.assertEqual(len(ids), len(set(ids)), f"同日上传出现重复 id：{ids}")

    def test_24_html_uses_profile_theme_and_sanitises_fonts(self):
        """F3+F10：deck.html 与 pptx 同一档案 Theme；恶意字体名不得跳出 <style>。"""
        self.assertTrue(FIXTURE_ACADEMIC.is_file())
        raw = FIXTURE_ACADEMIC.read_bytes()
        status, body, _ = self.request(
            "POST", "/api/profiles", data=raw,
            headers={"X-Filename": "case-academic.pptx"})
        self.assertEqual(status, 200, body[:300])
        profile_id = json.loads(body.decode("utf-8"))["id"]

        status, payload = self.post_json("/api/generate", {
            "title": "档案配色",
            "outline": "# 一页\n- 要点",
            "profile_id": profile_id,
            "theme": "editorial",
        })
        self.assertEqual(status, 200, payload)
        deck_dir = self.store_dir / "decks" / payload["id"]
        html = (deck_dir / "deck.html").read_text(encoding="utf-8")
        self.assertIn("--accent:#1B3A6B", html)
        self.assertNotIn("--accent:#1F3BFF", html)
        with zipfile.ZipFile(deck_dir / "deck.pptx") as archive:
            slides_xml = "".join(
                archive.read(name).decode("utf-8", "replace")
                for name in archive.namelist()
                if re.fullmatch(r"ppt/slides/slide\d+\.xml", name))
        self.assertIn("1B3A6B", slides_xml, "pptx 显式色应来自同一档案")

        evil_id = "pf-20261027-deadbeef"
        evil = {
            "id": evil_id, "name": "evil", "source": "source.pptx",
            "created_at": "2026-10-27T19:00:00+08:00", "slide_size": [13.33, 7.5],
            "palette": {"bg": "#FFFFFF", "ink": "#111111", "muted": "#555555",
                        "accent": "#123456", "accent_soft": "#EEEEEE", "rule": "#DDDDDD"},
            "fonts": {"title": "</style><script>alert(1)</script>", "body": "微软雅黑"},
            "metrics": {"slides": 1},
            "style": {"density": 1.0, "title_scale": 1.0, "bullet_char": "▪"},
            "learn": {"rounds": 0, "history": []},
        }
        evil_path = self.store_dir / "profiles" / f"{evil_id}.json"
        evil_path.write_text(json.dumps(evil, ensure_ascii=False), encoding="utf-8")

        status, payload = self.post_json("/api/generate", {
            "title": "xss", "outline": "# 一页\n- 要点",
            "profile_id": evil_id, "theme": "editorial",
        })
        self.assertEqual(status, 200, payload)
        html = (self.store_dir / "decks" / payload["id"] / "deck.html").read_text(
            encoding="utf-8")
        self.assertIn("--accent:#123456", html)
        self.assertNotIn("</style><script>", html)
        self.assertNotIn("<script>alert(1)</script>", html)
        self.assertIn('--font-title:"Microsoft YaHei"', html)

    def test_25_builder_and_html_deck_guards(self):
        """F1/F10 内核侧：builder 拒绝越界图片；html_deck 不回显绝对路径。"""
        kit_dir = ROOT / "kit"
        if str(kit_dir) not in sys.path:
            sys.path.insert(0, str(kit_dir))
        from pptxkit import builder as kit_builder
        from pptxkit import html_deck as kit_html

        base = self.store_dir.parent / "imgbase"
        base.mkdir(parents=True, exist_ok=True)
        legal = kit_builder._resolve("assets/x.png", base)
        self.assertTrue(legal.endswith("x.png"))
        with self.assertRaises(ValueError):
            kit_builder._resolve("../escape.png", base)
        with self.assertRaises(ValueError):
            kit_builder._resolve("C:/Windows/x.png", base)
        with self.assertRaises(ValueError):
            kit_builder._resolve("//server/share/x.png", base)

        work = self.store_dir.parent / "html-guard"
        work.mkdir(parents=True, exist_ok=True)
        md_path = work / "deck.md"
        md_path.write_text("# 图片页\n- 要点\n\n![](C:/secret/x.png)\n", encoding="utf-8")
        out = work / "deck.html"
        kit_html.build(md_path, out)
        html = out.read_text(encoding="utf-8")
        self.assertNotIn("C:/secret/x.png", html)
        self.assertIn("图片未随本网页提供", html)

        md_path.write_text("# 图片页\n- 要点\n\n![](assets/figs/a.png)\n", encoding="utf-8")
        kit_html.build(md_path, out)
        self.assertIn('src="assets/figs/a.png"', out.read_text(encoding="utf-8"))

    def test_26_concurrent_feedback_and_generate_never_500(self):
        """F11：并发 feedback（写档案）+ generate（读档案）不得出现任何 5xx。"""
        self.assertIsNotNone(self.profile_id, "test_07 应先创建档案")
        self.assertIsNotNone(self.profile_deck_id, "test_07 应先创建带档案的 deck")

        def worker(seed):
            results = []
            for index in range(5):
                if (seed + index) % 2:
                    status, payload = self.post_json("/api/generate", {
                        "title": f"并发 {seed}-{index}",
                        "outline": "# 一页\n- 要点",
                        "profile_id": self.profile_id,
                    })
                else:
                    status, payload = self.post_json(
                        f"/api/decks/{self.profile_deck_id}/feedback",
                        {"rating": 5, "notes": "F11 并发混合"})
                results.append((status, payload))
            return results

        bad = []
        with ThreadPoolExecutor(max_workers=6) as pool:
            for results in pool.map(worker, range(6)):
                for status, payload in results:
                    if status != 200:
                        bad.append((status, payload))
        self.assertEqual(bad, [], f"并发 feedback+generate 出现非 200：{bad[:3]}")

        # 档案仍然完整可读，且反馈轮次正常累计
        profile = self.get_json(f"/api/profiles/{self.profile_id}")
        self.assertIn("style", profile)
        self.assertGreaterEqual(profile["learn"]["rounds"], 1)
        self.assertTrue((self.store_dir / "profiles" / f"{self.profile_id}.json").is_file())

    def test_27_missing_profiles_do_not_leak_locks(self):
        """锁表残留：探测不存在的档案 id 不应创建 per-profile 锁。"""
        store = self.httpd.store
        before = set(store._profile_locks)
        missing = [f"pf-20261027-{uuid.uuid4().hex[:8]}" for _ in range(120)]

        def probe(profile_id):
            return self.request("GET", f"/api/profiles/{profile_id}")[0]

        with ThreadPoolExecutor(max_workers=8) as pool:
            statuses = list(pool.map(probe, missing))
        self.assertEqual(set(statuses), {404}, statuses)
        after = set(store._profile_locks)
        self.assertEqual(after - before, set(),
                         f"不存在的 id 不应留下锁：{sorted(after - before)[:5]}")

    def test_28_head_health_matches_get(self):
        """HEAD /api/health 与 GET 一致：200 + 相同 Content-Length，无响应体。"""
        status, body, headers = self.request("GET", "/api/health")
        self.assertEqual(status, 200)
        connection = http.client.HTTPConnection("127.0.0.1", self.port, timeout=10)
        try:
            connection.request("HEAD", "/api/health")
            response = connection.getresponse()
            head_body = response.read()
            self.assertEqual(response.status, 200)
            self.assertEqual(head_body, b"")
            self.assertEqual(response.getheader("Content-Length"), str(len(body)))
            self.assertIn("application/json", response.getheader("Content-Type", ""))
        finally:
            connection.close()

    def test_29_learn_history_capped_at_200(self):
        """learn.history 最多 200 条（丢最旧），rounds 继续累加。"""
        kit_dir = ROOT / "kit"
        if str(kit_dir) not in sys.path:
            sys.path.insert(0, str(kit_dir))
        from pptxkit import profile as profiler

        profile = profiler.analyze(FIXTURE_ACADEMIC)
        for index in range(250):
            profiler.learn(profile, 5, notes=f"r{index}")
        learn = profile["learn"]
        self.assertEqual(learn["rounds"], 250)
        self.assertEqual(len(learn["history"]), 200)
        self.assertEqual(learn["history"][0]["notes"], "r50")
        self.assertEqual(learn["history"][-1]["notes"], "r249")
        self.assertAlmostEqual(profile["style"]["density"], 1.4, places=4)

        store_dir = self.store_dir.parent / "history-store"
        saved = profiler.save(profile, store_dir)
        reloaded = profiler.load(store_dir, saved["id"])
        self.assertEqual(len(reloaded["learn"]["history"]), 200)
        self.assertEqual(reloaded["learn"]["rounds"], 250)

    def test_30_visual_gates_trigger(self):
        """渲染前结构闸门：越界/溢出、重叠、小字号、留白、图片拉伸各触发一次。"""
        kit_dir = ROOT / "kit"
        if str(kit_dir) not in sys.path:
            sys.path.insert(0, str(kit_dir))
        from pptxkit import inspect as kit_inspect

        work = self.store_dir.parent / "visual-fixtures"
        work.mkdir(parents=True, exist_ok=True)
        expected = {
            "overflow": "文字溢出",
            "out-of-bounds": "文字越界",
            "overlap": "元素重叠",
            "small-font": "字号过小",
            "blank": "留白过多",
            "stretch": "图片拉伸",
        }
        for kind, issue_kind in expected.items():
            path = build_review_deck(work / f"{kind}.pptx", kind)
            report = kit_inspect.review(path)
            kinds = {issue["kind"] for page in report["pages"]
                     for issue in page["issues"]}
            self.assertIn(issue_kind, kinds, f"{kind} 应触发 {issue_kind}，实际 {kinds}")
            self.assertFalse(report["rendered"])
            self.assertIn("未渲染", " ".join(report["notes"]))
            self.assertIn("未渲染（缺少 Office/WPS）", kit_inspect.format_review(report))

    def test_31_visual_render_gate_and_degrade(self):
        """有渲染 PNG 时叠加视觉闸门；没有时明确降级为结构检查。"""
        kit_dir = ROOT / "kit"
        if str(kit_dir) not in sys.path:
            sys.path.insert(0, str(kit_dir))
        from PIL import Image
        from pptxkit import inspect as kit_inspect

        work = self.store_dir.parent / "visual-fixtures"
        work.mkdir(parents=True, exist_ok=True)
        deck = build_review_deck(work / "render.pptx", "render")

        structural = kit_inspect.review(deck)
        self.assertFalse(structural["rendered"])
        self.assertIn("未渲染（缺少 Office/WPS）", kit_inspect.format_review(structural))

        png_dir = work / "render-png"
        png_dir.mkdir(parents=True, exist_ok=True)
        Image.new("RGB", (800, 450), (255, 255, 255)).save(png_dir / "slide-01.png")
        filled = Image.new("RGB", (800, 450), (255, 255, 255))
        for x in range(80, 500):
            for y in range(100, 150):
                filled.putpixel((x, y), (20, 30, 60))
        filled.save(png_dir / "slide-02.png")

        rendered = kit_inspect.review(deck, png_dir=png_dir)
        self.assertTrue(rendered["rendered"])
        self.assertIn("已分析 2 张渲染 PNG", " ".join(rendered["notes"]))
        page_one = rendered["pages"][0]
        kinds = {issue["kind"] for issue in page_one["issues"]}
        self.assertTrue(kinds & {"留白过多", "无视觉焦点"}, kinds)
        self.assertEqual(rendered["pages"][1]["status"], "PASS",
                         rendered["pages"][1]["issues"])
        output = kit_inspect.format_review(rendered)
        self.assertIn("渲染：已分析逐页 PNG（视觉闸门生效）", output)
        self.assertIn("汇总：PASS 1 / NEEDS_REVISION 1", output)

    def test_32_visual_cli_contract(self):
        """visual CLI：逐页 PASS/NEEDS_REVISION + 汇总；全 PASS 退出码 0。"""
        kit_dir = ROOT / "kit"
        if str(kit_dir) not in sys.path:
            sys.path.insert(0, str(kit_dir))
        from pptxkit import inspect as kit_inspect

        work = self.store_dir.parent / "visual-fixtures"
        work.mkdir(parents=True, exist_ok=True)
        bad = build_review_deck(work / "bad-cli.pptx", "blank")
        buffer = io.StringIO()
        with contextlib.redirect_stdout(buffer):
            code = kit_inspect.main(["--visual", str(bad)])
        output = buffer.getvalue()
        self.assertEqual(code, 1)
        self.assertIn("[NEEDS_REVISION] 第 1 页", output)
        self.assertIn("汇总：PASS 0 / NEEDS_REVISION 1", output)

        good = build_review_deck(work / "good-cli.pptx", "good")
        buffer = io.StringIO()
        with contextlib.redirect_stdout(buffer):
            code = kit_inspect.main(["--visual", str(good)])
        output = buffer.getvalue()
        self.assertEqual(code, 0, output)
        self.assertIn("[PASS] 第 1 页", output)
        self.assertIn("汇总：PASS 1 / NEEDS_REVISION 0", output)


if __name__ == "__main__":
    unittest.main(verbosity=2)
