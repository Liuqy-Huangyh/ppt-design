"""Markdown → 幻灯片中间结构（不依赖任何 Markdown 库）。

约定
----
    ---                 front matter 分隔 / 幻灯片分页
    # 标题               开新的一页
    ## 副标题            本页副标题（正文最前面时）或加粗小标题
    ### 栏目名           两栏版式里切换左/右栏
    - 要点 / 1. 步骤     列表（缩进 2 空格为二级）
    | a | b |            表格（首行表头）
    ![说明](figs/x.png)  图片，说明作为图注
    > 引言 —— 出处       引用页
    <!-- layout: xxx --> 指定版式
    ::: notes ... :::    演讲者备注
"""
from __future__ import annotations

import re

LIST_RE = re.compile(r"^(\s*)([-*+]|\d+[.)])\s+(.*)$")
IMG_RE = re.compile(r"^!\[(.*?)\]\((.+?)\)\s*$")
TABLE_SEP_RE = re.compile(r"^\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?\s*$")
LAYOUT_RE = re.compile(r"<!--\s*layout\s*:\s*([\w\-]+)\s*-->", re.I)
NOTE_OPEN = re.compile(r"^:::\s*notes?\s*$", re.I)
KICKER_RE = re.compile(r"^\[(.*?)\]\s*$")


def _blank() -> dict:
    return {
        "layout": None, "title": "", "subtitle": "", "kicker": "",
        "bullets": [], "left": None, "right": None,
        "table": None, "image": None, "quote": None,
        "stats": None, "steps": None, "refs": None, "note": "",
    }


def _split_row(line: str) -> list[str]:
    line = line.strip()
    if line.startswith("|"):
        line = line[1:]
    if line.endswith("|"):
        line = line[:-1]
    return [c.strip() for c in line.split("|")]


def _is_empty(slide: dict) -> bool:
    return not any(slide[k] for k in ("title", "subtitle", "bullets", "table",
                                      "image", "quote", "stats", "steps", "refs",
                                      "left", "right"))


def parse(text: str) -> dict:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = text.split("\n")
    meta: dict[str, str] = {}

    if lines and lines[0].strip() in ("---", "+++"):
        end = 1
        while end < len(lines) and lines[end].strip() not in ("---", "+++"):
            end += 1
        for raw in lines[1:end]:
            if ":" in raw:
                k, v = raw.split(":", 1)
                meta[k.strip().lower()] = v.strip().strip('"').strip("'")
        lines = lines[end + 1:]

    slides: list[dict] = []
    cur = _blank()
    pending_layout: str | None = None
    pending_kicker: str | None = None
    in_notes = False
    notes_buf: list[str] = []
    column_target: str | None = None   # two-col 模式当前写到哪一栏
    quote_buf: list[str] = []
    table_buf: list[str] = []

    def flush_quote():
        nonlocal quote_buf
        if quote_buf:
            body, cite = [], ""
            for q in quote_buf:
                q = q.strip()
                if q.startswith(("—", "——", "--")):
                    cite = q.lstrip("-— ").strip()
                else:
                    body.append(q)
            text = " ".join(body).strip()
            old = cur.get("quote")
            if old:  # 同一页里正文与出处被空行分开时，合并而不是覆盖
                if text:
                    old["text"] = (old.get("text", "") + " " + text).strip()
                if cite:
                    old["cite"] = cite
            else:
                cur["quote"] = {"text": text, "cite": cite}
            quote_buf = []

    def flush_table():
        nonlocal table_buf
        rows = [r for r in table_buf if not TABLE_SEP_RE.match(r)]
        if len(rows) >= 2:
            cur["table"] = {
                "header": _split_row(rows[0]),
                "rows": [_split_row(r) for r in rows[1:]],
            }
        table_buf = []

    def push_slide():
        nonlocal cur, column_target, pending_kicker
        flush_quote()
        flush_table()
        if not _is_empty(cur) or cur["note"]:
            if pending_layout and not cur["layout"]:
                cur["layout"] = pending_layout
            slides.append(cur)
        cur = _blank()
        column_target = None
        if pending_kicker:
            cur["kicker"] = pending_kicker
            pending_kicker = None

    for raw in lines:
        line = raw.rstrip()
        stripped = line.strip()

        if in_notes:
            if stripped == ":::":
                in_notes = False
                cur["note"] = "\n".join(notes_buf).strip()
                notes_buf = []
            else:
                notes_buf.append(line)
            continue

        if NOTE_OPEN.match(stripped):
            in_notes = True
            continue

        m = LAYOUT_RE.search(stripped)
        if m:
            name = m.group(1).strip().lower()
            if _is_empty(cur) and not cur["note"]:
                pending_layout = name
            else:
                cur["layout"] = name
            continue

        if not stripped:
            flush_quote()
            flush_table()
            continue

        if stripped in ("---", "***", "___") and len(set(stripped)) == 1:
            push_slide()
            continue

        if stripped == ":::":
            continue

        # 表格
        if stripped.startswith("|") and stripped.count("|") >= 2:
            if table_buf is not None:
                table_buf.append(stripped)
                continue

        # 标题
        if stripped.startswith("#"):
            level = len(stripped) - len(stripped.lstrip("#"))
            heading = stripped[level:].strip()
            if level == 1:
                push_slide()
                cur["title"] = heading
                if pending_layout:
                    cur["layout"] = pending_layout
                    pending_layout = None
            elif level == 2:
                if not cur["bullets"] and not cur["subtitle"] and not cur["table"]:
                    cur["subtitle"] = heading
                else:
                    cur["bullets"].append({"text": f"**{heading}**", "level": 0})
            elif level == 3:
                if (cur["layout"] in ("two-col", "two_col") or column_target) and not cur["left"]:
                    cur["left"] = {"title": heading, "bullets": []}
                    column_target = "left"
                elif column_target == "left":
                    cur["right"] = {"title": heading, "bullets": []}
                    column_target = "right"
                else:
                    cur["bullets"].append({"text": f"**{heading}**", "level": 0})
            else:
                cur["bullets"].append({"text": f"**{heading}**", "level": 1})
            continue

        # 图片
        m = IMG_RE.match(stripped)
        if m:
            alt, path = m.group(1).strip(), m.group(2).strip()
            cur["image"] = {"path": path, "caption": alt}
            cur["layout"] = cur["layout"] or "image"
            continue

        # 引用
        if stripped.startswith(">"):
            quote_buf.append(stripped.lstrip(">").strip())
            continue
        flush_quote()

        # 小标签 [标签]：写在 # 之前时留给下一页
        m = KICKER_RE.match(stripped)
        if m and not cur["bullets"] and not cur["kicker"]:
            if _is_empty(cur):
                pending_kicker = m.group(1).strip()
            else:
                cur["kicker"] = m.group(1).strip()
            continue

        # 列表
        m = LIST_RE.match(line)
        if m:
            indent, marker, body = m.group(1), m.group(2), m.group(3).strip()
            level = min(2, len(indent.expandtabs(2)) // 2) if indent else 0
            ordered = marker[0].isdigit()
            ordered_index = marker[:-1] if ordered else ""
            item = {"text": body, "level": level}
            if ordered:
                item["bullet"] = False
                item["text"] = f"{ordered_index}. {body}"
            if column_target in ("left", "right") and cur[column_target] is not None:
                cur[column_target]["bullets"].append(item)
            else:
                cur["bullets"].append(item)
            continue

        # 普通段落
        if column_target in ("left", "right") and cur[column_target] is not None:
            cur[column_target]["bullets"].append({"text": stripped, "level": 0})
        else:
            cur["bullets"].append({"text": stripped, "level": 0})

    flush_quote()
    flush_table()
    if in_notes and notes_buf:
        cur["note"] = "\n".join(notes_buf).strip()
    if not _is_empty(cur) or cur["note"]:
        if pending_layout and not cur["layout"]:
            cur["layout"] = pending_layout
        slides.append(cur)

    _infer_layouts(slides, meta)
    return {"meta": meta, "slides": slides}


def _infer_layouts(slides: list[dict], meta: dict) -> None:
    end_titles = {"谢谢", "致谢", "thank you", "thanks", "q&a", "谢谢聆听", "敬请指正"}
    has_cover = bool((meta or {}).get("title"))
    for idx, s in enumerate(slides):
        if s["layout"]:
            _coerce(s)
            continue
        if s["table"]:
            s["layout"] = "table"
        elif s["image"]:
            s["layout"] = "image"
        elif s["quote"] and not s["bullets"]:
            s["layout"] = "quote"
        elif s["left"] or s["right"]:
            s["layout"] = "two-col"
        elif idx == 0 and not has_cover and not s["bullets"] and not s["table"]:
            s["layout"] = "cover"
        elif s["title"].strip().lower() in end_titles and not s["bullets"]:
            s["layout"] = "end"
        elif not s["bullets"] and s["title"]:
            s["layout"] = "section"
        else:
            s["layout"] = "bullets"
        _coerce(s)


def _coerce(s: dict) -> None:
    """把 bullets 转成 stats / steps / refs 需要的数据结构。"""
    if s["layout"] == "stats" and s["bullets"] and not s["stats"]:
        out = []
        for b in s["bullets"]:
            parts = [p.strip() for p in b["text"].split("|")]
            if len(parts) >= 2:
                item = {"value": parts[0], "label": parts[1]}
                if len(parts) >= 3:
                    item["note"] = parts[2]
                out.append(item)
        s["stats"] = out or None
    if s["layout"] == "timeline" and s["bullets"] and not s["steps"]:
        out = []
        for b in s["bullets"]:
            text = b["text"]
            for sep in ("|", "：", ":"):
                if sep in text:
                    name, desc = text.split(sep, 1)
                    out.append({"name": name.strip(), "desc": desc.strip()})
                    break
            else:
                out.append({"name": text, "desc": ""})
        s["steps"] = out or None
    if s["layout"] == "refs" and s["bullets"] and not s["refs"]:
        s["refs"] = [b["text"] for b in s["bullets"]]
