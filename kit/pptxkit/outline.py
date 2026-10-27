"""提示词 → Markdown 大纲：离线启发式为主，配置了 LLM 时优先尝试调用。

对外接口（docs/CONTRACT.md 第 6 节）
------------------------------------
    llm_available() -> bool
    build_outline(prompt, *, pages=10, kind="auto", profile=None, title="") -> str

输出必须能被 kit/pptxkit/markdown.py 直接解析：
front matter（title/subtitle/...）+ 每页 `# 页标题` + `- 要点` + `::: notes` + `---` 分页。

环境变量（可选，OpenAI 兼容接口）：
    PPT_LLM_BASE_URL / PPT_LLM_API_KEY / PPT_LLM_MODEL / PPT_LLM_TIMEOUT
"""
from __future__ import annotations

import json
import os
import re
from urllib import request as _urlrequest

from . import markdown as _md

KINDS = {
    "academic": "学术答辩",
    "business": "商务汇报",
    "share": "公开分享",
}

_KIND_ALIASES = {
    "学术": "academic", "答辩": "academic", "论文": "academic", "研究": "academic",
    "学术汇报": "academic", "论文答辩": "academic", "academic": "academic",
    "商务": "business", "商业": "business", "汇报": "business", "路演": "business",
    "方案": "business", "产品": "business", "复盘": "business", "工作汇报": "business",
    "business": "business", "pitch": "business",
    "分享": "share", "演讲": "share", "故事": "share", "公开课": "share",
    "share": "share", "talk": "share", "story": "share",
}

_KIND_KEYWORDS = {
    "academic": ("答辩", "论文", "研究", "实验", "学术", "综述", "课题", "算法", "方法研究", "评测"),
    "business": ("汇报", "方案", "商业", "产品", "路演", "复盘", "项目", "业绩", "增长",
                 "运营", "营销", "预算", "季度", "年度", "客户", "市场", "战略"),
    "share": ("分享", "演讲", "故事", "公开课", "心得", "体会", "聊聊", "talk", "复盘会"),
}

_BASE_BULLETS = {"academic": 3, "business": 3, "share": 2}

# 每条要点的目标字数（1.0 密度时）；密度高了写长一点，低了收短一点
_BASE_CHARS = {"academic": 26, "business": 24, "share": 22}
_CUT_PUNCT = "，；。、："

_TITLE_PREFIXES = (
    "请帮我", "帮我", "帮忙", "麻烦你", "麻烦", "我想让", "我想要", "我想", "我需要",
    "给我", "帮我做", "帮我写", "生成一份", "生成一个", "生成", "制作一份", "制作一个",
    "制作", "写一份", "写一个", "写篇", "写", "做一份", "做一个", "做", "来一份",
    "来一个", "出一个", "出一份", "出", "准备一份", "准备",
)
_TITLE_SUFFIXES = (
    "的演示文稿", "演示文稿", "的幻灯片", "幻灯片", "的PPT", "的ppt", "的Ppt",
    "PPT", "ppt", "Ppt", "的汇报材料", "汇报材料", "的汇报", "汇报",
    "的分享材料", "的分享", "分享材料", "的内容",
)
_STOPWORDS = {
    "的", "了", "和", "与", "及", "是", "为", "在", "把", "对", "从", "到", "用",
    "让", "给", "被", "而", "就", "都", "也", "还", "很", "更", "最", "可以",
    "我们", "你们", "他们", "这个", "那个", "一个", "一种", "如何", "什么", "怎么",
    "进行", "通过", "基于", "关于", "以及", "一些", "相关", "内容", "主题",
}

_GENERIC_WORDS = {
    "academic": ["现有方法", "关键环节", "真实数据", "落地链路"],
    "business": ["现有流程", "关键指标", "试点场景", "长期收益"],
    "share": ["手上的事", "关键判断", "真实案例", "长期积累"],
}

_SUBTITLES = {
    "academic": "从问题定义到实验结论的一条完整链路",
    "business": "一页纸讲清楚结论、数据与行动",
    "share": "一条故事线，讲清我是怎么想明白的",
}

_OPENING_NOTES = {
    "academic": "开场 1 分钟：先说清楚问题从哪来，再给结构。不要念稿。",
    "business": "开场 30 秒：先给结论，再解释路径，让听众知道要做什么决定。",
    "share": "开场 1 分钟：用一个具体场景把听众带进来，不要先讲道理。",
}


# ------------------------------------------------------------------ 环境与 LLM


def _env(name: str) -> str:
    return (os.environ.get(name) or "").strip()


def llm_available() -> bool:
    """是否配置了可用的 OpenAI 兼容接口（缺任一变量就视为不可用）。"""
    return bool(_env("PPT_LLM_BASE_URL") and _env("PPT_LLM_API_KEY"))


def _chat_endpoint() -> str:
    base = _env("PPT_LLM_BASE_URL").rstrip("/")
    if base.endswith("/chat/completions"):
        return base
    return base + "/chat/completions"


def _llm_messages(prompt: str, *, pages: int, kind: str, title: str, density: float) -> list[dict]:
    kind_label = KINDS.get(kind, KINDS["business"])
    count = max(2, min(5, round(_BASE_BULLETS.get(kind, 3) * density)))
    user = f"""请为主题「{title}」写一份中文演示大纲，共 {pages} 页。

场景类型：{kind_label}；信息密度：{density:.2f}（1.0 为基准，越大每页要点越多，本份建议每页 {count} 条要点）。

硬性格式要求：
1. 第一段是 front matter，用两行 `---` 包住，至少包含 title、subtitle、author、org、date、kicker、footer；
2. 正文每页以 `# 页标题` 开头，下面 2–5 条 `- 要点`（不要用 `##`）；
3. 每页末尾附 `::: notes` 换行写 1–2 句讲者备注，再用单独一行 `:::` 结束；
4. 页与页之间用单独一行 `---` 分隔；
5. 只输出 Markdown 本身，不要代码块包裹，不要任何解释。

原始需求：{prompt or title}
"""
    return [
        {"role": "system", "content": "你是资深演示文稿策划，擅长把一句话需求变成结构清晰、观点明确的演示大纲。只输出 Markdown。"},
        {"role": "user", "content": user},
    ]


def llm_outline(prompt: str, *, pages: int = 10, kind: str = "auto",
                title: str = "", density: float = 1.0) -> str | None:
    """调用 /chat/completions 生成大纲；未配置或失败时返回 None（调用方负责降级）。"""
    if not llm_available():
        return None
    resolved_kind = resolve_kind(kind, prompt)
    resolved_title = title.strip() or _guess_title(prompt, resolved_kind)
    model = _env("PPT_LLM_MODEL") or "deepseek-chat"
    try:
        timeout = float(_env("PPT_LLM_TIMEOUT") or 30)
    except ValueError:
        timeout = 30.0
    payload = {
        "model": model,
        "messages": _llm_messages(prompt, pages=pages, kind=resolved_kind,
                                  title=resolved_title, density=density),
        "temperature": 0.7,
        "stream": False,
    }
    req = _urlrequest.Request(
        _chat_endpoint(),
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {_env('PPT_LLM_API_KEY')}",
        },
        method="POST",
    )
    try:
        with _urlrequest.urlopen(req, timeout=timeout) as resp:
            raw = resp.read().decode("utf-8", "replace")
        data = json.loads(raw)
        content = data["choices"][0]["message"]["content"]
        return str(content or "").strip() or None
    except Exception:
        return None


# ------------------------------------------------------------------ 文本工具


def _strip_affixes(text: str) -> str:
    word = (text or "").strip()
    changed = True
    while changed and word:
        changed = False
        for prefix in _TITLE_PREFIXES:
            if word.startswith(prefix) and len(word) > len(prefix) + 1:
                word = word[len(prefix):].strip()
                changed = True
                break
        for suffix in _TITLE_SUFFIXES:
            if word.endswith(suffix) and len(word) > len(suffix) + 1:
                word = word[: -len(suffix)].strip()
                changed = True
                break
    word = re.sub(r"^(关于|有关|针对|围绕|以)", "", word).strip()
    word = re.sub(r"(为主题|为题目|主题是|题目是).*$", "", word).strip()
    word = word.strip("，。、；：,.!?！？;:—-…· \t\"'“”‘’《》()（）")
    return word


def _guess_title(prompt: str, kind: str = "business") -> str:
    text = re.sub(r"\r\n?", "\n", str(prompt or "")).strip()
    if not text:
        return {"academic": "研究进展汇报", "business": "本次汇报", "share": "本次分享"}.get(kind, "本次汇报")
    first = re.split(r"[。！？!?\n，,；;]", text, maxsplit=1)[0].strip()
    title = _strip_affixes(first) or _strip_affixes(text) or text
    title = re.sub(r"\s+", " ", title).strip()
    if len(title) > 28:
        title = title[:28].rstrip("，。、；：,.!?！？;:—- ")
    return title or "本次汇报"


def _short_topic(title: str, kind: str) -> str:
    """要点里引用主题时用短称呼：长标题不整段塞进句子。"""
    title = (title or "").strip()
    if len(title) <= 12:
        return title or "这件事"
    trimmed = re.sub(r"^(面向|关于|基于|针对|一种|一个)", "", title).strip()
    if len(trimmed) <= 12:
        return trimmed
    return {"academic": "本课题", "business": "本项目", "share": "这件事"}.get(kind, "本项目")


def _keywords(prompt: str, topic: str, kind: str, limit: int = 4) -> list[str]:
    parts = re.split(
        r"[\s，。、；：:,.!?！？;()（）\[\]【】“”\"'《》/\\|+&—\-]+",
        str(prompt or ""),
    )
    words: list[str] = []
    for part in parts:
        word = _strip_affixes(part)
        word = re.sub(r"[与和的及]$", "", word).strip() if len(word) > 2 else word
        if len(word) < 2 or len(word) > 16:
            continue
        if re.fullmatch(r"(?:[A-Za-z]{1,2}|Q[1-4]|\d{2,4}|第[一二三四1-4]季度?)", word):
            continue
        if word in _STOPWORDS or word == topic:
            continue
        if word in words:
            continue
        words.append(word)
    generic = _GENERIC_WORDS.get(kind, _GENERIC_WORDS["business"])
    for word in generic:
        if len(words) >= limit:
            break
        if word not in words:
            words.append(word)
    return words[:limit]


def resolve_kind(kind: str | None, prompt: str = "") -> str:
    """把 kind（含中文别名 / auto）解析成 academic|business|share。"""
    key = (kind or "auto").strip().lower()
    if key in KINDS:
        return key
    if key in _KIND_ALIASES:
        return _KIND_ALIASES[key]
    # 未知取值（比如「公开分享」「论文答辩」）连同 prompt 一起做关键词判断
    text = f"{kind or ''} {prompt or ''}"
    scores = {name: 0 for name in KINDS}
    for name, words in _KIND_KEYWORDS.items():
        for word in words:
            if word and word in text:
                scores[name] += 1
    best = max(scores, key=lambda name: (scores[name], name == "business"))
    return best if scores[best] > 0 else "business"


# ------------------------------------------------------------------ 结构模板


def _sections(kind: str) -> list[dict]:
    if kind == "academic":
        return [
            {"kicker": "背景", "title": "研究背景与问题",
             "bullets": [
                 "需求侧：{topic} 的落地要求正在变高，场景从实验环境走向真实环境",
                 "困境：主流做法在效果与成本之间被迫二选一",
                 "机会：约束变严的同时，可用的新工具也在变多",
                 "本文目标：在给定约束下把 {topic} 做得更稳、更省",
             ],
             "note": "先讲清楚约束从哪来：三个真实场景足以说明为什么非这样做不可。"},
            {"kicker": "问题", "title": "研究问题与目标",
             "bullets": [
                 "问题定义：把 {topic} 拆成可度量、可验证的子问题",
                 "目标：在 {kw1} 与 {kw2} 之间取得更好的平衡",
                 "边界：不追求通用最优，只解决约束最强的那一段",
             ],
             "note": "这一页只讲一个判断：问题被定义清楚，方法才有比较的意义。"},
            {"kicker": "相关工作", "title": "相关工作与差异",
             "bullets": [
                 "路线一：围绕 {kw1} 做工程化改进，见效快但上限受限",
                 "路线二：围绕 {kw2} 做结构化重构，上限高但代价大",
                 "本文差异：把两条路线放进同一目标函数，联合设计",
             ],
             "note": "对比表不用念，挑一条最接近的工作讲清楚差别即可。"},
            {"kicker": "方法", "title": "方法总览",
             "bullets": [
                 "整体链路：输入 → 特征编码 → 融合与对齐 → 输出，端到端可训练",
                 "关键模块：特征对齐与前向计算解耦，主干保持轻量",
                 "训练策略：分阶段进行，先保效果，再压成本",
             ],
             "note": "用一张框架图讲完整条链路，重点说清楚每个模块解决什么问题。"},
            {"kicker": "方法", "title": "关键设计",
             "bullets": [
                 "设计一：把高开销的中间环节换成更轻的等价结构",
                 "设计二：用输出分布约束替代复杂的后处理链",
                 "设计三：把部署约束前置到训练阶段",
             ],
             "note": "三个设计各用一句话，评委追问时再展开推导细节。"},
            {"kicker": "实验", "title": "实验设置",
             "bullets": [
                 "数据：公开基准 + 自建真实场景数据集",
                 "基线：覆盖三条主流路线的最强公开实现",
                 "指标：效果、成本、稳定性三组，均报告多次运行的均值与方差",
             ],
             "note": "只交代会影响结论可信度的设置，其余放进附录。"},
            {"kicker": "实验", "title": "实验结果",
             "bullets": [
                 "主结果：核心指标相对最强基线有明显提升",
                 "成本：同等效果下资源占用与延迟同步下降",
                 "稳定性：多次运行的方差收敛，长跑不再抖动",
             ],
             "note": "表格只讲两行：最强基线与自己最好的配置，其余留给提问。"},
            {"kicker": "分析", "title": "消融与分析",
             "bullets": [
                 "去掉核心模块后指标下降最多，说明它是主要贡献点",
                 "规模翻倍收益递减，当前配置已接近拐点",
                 "失败案例：极端场景仍是短板，原因已经定位",
             ],
             "note": "诚实讲短板，比硬撑更有说服力；说清楚下一步怎么补。"},
            {"kicker": "结论", "title": "结论与贡献",
             "bullets": [
                 "贡献一：提出一套可复用的方法与工程实现，代价可控",
                 "贡献二：给出{topic}的端到端落地链路与经验",
                 "结论：结构先验比单纯堆叠资源更有效",
             ],
             "note": "结论要能一句话被复述，别在最后引入新信息。"},
            {"kicker": "展望", "title": "不足与展望",
             "bullets": [
                 "不足：验证场景有限，泛化性还需要补实验",
                 "展望：与硬件协同设计，进一步压缩成本",
                 "下一步：补齐评测集，开源实现与复现脚本",
             ],
             "note": "收尾留一个开放问题，方便评委接着问。"},
        ]

    if kind == "share":
        return [
            {"kicker": "开场", "title": "先问一个问题",
             "bullets": [
                 "先问一句：{topic}，我们真的想清楚了吗",
                 "今天不急着给结论，先讲我是怎么走到结论的",
             ],
             "note": "开场 1 分钟：用一个具体场景把听众带进来，不要先讲道理。"},
            {"kicker": "故事", "title": "我的起点",
             "bullets": [
                 "两年前，我在这件事上完全是个外行",
                 "第一次尝试失败，原因很具体",
                 "那次失败改变了我后面所有的选择",
             ],
             "note": "细节比形容词有用：讲清楚当时的时间、资源与判断。"},
            {"kicker": "故事", "title": "转折点",
             "bullets": [
                 "转折来自一次很小的观察",
                 "把 {kw1} 换一种用法，问题突然变简单",
                 "从那之后，我开始有意识地积累方法",
             ],
             "note": "转折点要讲得具体，最好有一个可以复述的动作。"},
            {"kicker": "判断", "title": "三个判断",
             "bullets": [
                 "判断一：产能瓶颈从「做」转移到「审」",
                 "判断二：个人项目死于维护，不死于开发",
                 "判断三：工具放大方法论，而不是替代它",
             ],
             "note": "三个判断逐个停顿两秒，让听众有时间点头或反对。"},
            {"kicker": "案例", "title": "一次真实的取舍",
             "bullets": [
                 "背景：时间只有两周，目标不能降",
                 "取舍：砍掉两个功能，把一件事做到极致",
                 "结果：按时交付，反而多了三个新用户",
             ],
             "note": "讲取舍而不是讲成绩：听众关心的是你怎么做决定。"},
            {"kicker": "方法", "title": "我到底怎么做",
             "bullets": [
                 "先写清楚要解决的问题，再动手",
                 "把重复劳动交给工具，把判断留给自己",
                 "每次复盘只改一个变量",
             ],
             "note": "给方法要配一个能立刻执行的动作，否则听众带不走。"},
            {"kicker": "代价", "title": "代价与教训",
             "bullets": [
                 "代价一：短期效率下降，需要熬过适应期",
                 "代价二：要自己维护工具链",
                 "教训：不要为了自动化而自动化",
             ],
             "note": "讲代价的时候语气放平，别变成诉苦。"},
            {"kicker": "行动", "title": "给你的三件事",
             "bullets": [
                 "今天就能做：把最常做的一件事写成清单",
                 "这周能做：找一个环节试着交给 AI",
                 "这个月能做：做一次完整的复盘",
             ],
             "note": "三件事要足够小，小到听众今晚就能开始。"},
            {"kicker": "结尾", "title": "最后一句话",
             "bullets": [
                 "{topic} 不是一次冲刺，而是一套可以复用的方法",
                 "如果只记住一句：先想清楚，再动手",
             ],
             "note": "停两秒再说谢谢，把最后一个问题留给听众。"},
        ]

    # business（默认）
    return [
        {"kicker": "结论", "title": "结论先行",
         "bullets": [
             "一句话结论：{topic} 值得现在做，窗口期大约 6 个月",
             "三个理由：需求明确、路径可行、投入可控",
             "需要的决策：批准试点范围与资源",
         ],
         "note": "开场 30 秒先给结论，再解释路径，让听众知道要做什么决定。"},
        {"kicker": "现状", "title": "现状与问题",
         "bullets": [
             "现状：{kw1} 已经在跑，但主要靠人堆",
             "问题：响应慢、口径不一致、成本随规模线性上涨",
             "影响：直接影响转化率与交付质量",
         ],
         "note": "问题要具体到能被验证，避免用「效率低」这种形容词。"},
        {"kicker": "判断", "title": "机会与判断",
         "bullets": [
             "判断一：需求会持续，短期没有替代方案",
             "判断二：技术方案的成熟度已经跨过可用门槛",
             "判断三：先做的一方会拿到数据与流程优势",
         ],
         "note": "把判断讲成可验证的假设，后面用数据回应。"},
        {"kicker": "方案", "title": "方案总览",
         "bullets": [
             "一条主线：以 {topic} 为核心，先跑通再扩展",
             "三个模块：接入、处理、反馈，边界清晰",
             "落地方式：小步快跑，每两周一个可见成果",
         ],
         "note": "方案只讲结构，不讲实现细节；细节留给讨论环节。"},
        {"kicker": "数据", "title": "关键数据", "layout": "stats",
         "bullets": [
             "38% | 效率提升 | 试点场景实测（示例数据，请替换）",
             "3 周 | 回本周期 | 按当前人力成本测算（示例数据）",
             "12 人 | 首批范围 | 覆盖一个完整业务小组",
         ],
         "note": "数据页只讲结论与口径，具体口径放进附录备查。"},
        {"kicker": "行动", "title": "行动与里程碑", "layout": "timeline",
         "bullets": [
             "第 1 阶段 | 两周内完成范围确认与指标基线",
             "第 2 阶段 | 一个月内跑通试点，拿到第一组数据",
             "第 3 阶段 | 一个季度内推广到全部团队",
         ],
         "note": "里程碑要有人名与日期，否则只是愿望清单。"},
        {"kicker": "风险", "title": "资源与风险",
         "bullets": [
             "投入：人力 × 3，周期 1 个季度，预算可控",
             "风险一：数据口径不统一，需要先做治理",
             "风险二：依赖外部供应商，需要备选方案",
         ],
         "note": "主动讲风险，比被追问更可信；每条风险给出应对动作。"},
        {"kicker": "收益", "title": "预期收益",
         "bullets": [
             "直接收益：效率提升，交付周期缩短",
             "间接收益：口径统一，决策更快",
             "长期收益：沉淀可复用的流程与数据资产",
         ],
         "note": "收益分三层讲，让不同角色的听众都能对上号。"},
        {"kicker": "下一步", "title": "下一步",
         "bullets": [
             "本周：确认负责人与要看的指标",
             "本月：完成试点，给出复盘",
             "本季度：决定是否全面推广",
         ],
         "note": "结尾落到具体动作与时间点，不要停在口号上。"},
    ]


def _spread(total: int, want: int) -> list[int]:
    if want >= total:
        return list(range(total))
    if want <= 1:
        return [0]
    idx = sorted({int(round(i * (total - 1) / (want - 1))) for i in range(want)})
    i = 0
    while len(idx) < want and i < total:
        if i not in idx:
            idx.append(i)
            idx.sort()
        i += 1
    return idx[:want]


def _pick_sections(kind: str, count: int) -> list[dict]:
    base = _sections(kind)
    if count <= len(base):
        return [base[i] for i in _spread(len(base), count)]
    out = list(base)
    extra = count - len(base)
    for i in range(extra):
        out.append({
            "kicker": "延伸",
            "title": f"延伸讨论 {i + 1}",
            "bullets": [
                "围绕「{topic}」补充一组事实与证据",
                "补充一个与 {kw1} 有关的反例，说明边界",
                "留一个问题给听众，引出下一段",
            ],
            "note": "延伸页按现场时间取舍，时间紧就整页跳过。",
        })
    return out


# ------------------------------------------------------------------ 渲染 Markdown


def _one_line(text: str) -> str:
    return re.sub(r"\s+", " ", str(text or "")).replace("---", "—").strip()


_CJK = "\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff"


def _tidy(text: str) -> str:
    """模板拼接留下的中文空格：汉字之间的空格去掉，标点前不留空格。"""
    text = str(text or "")
    text = re.sub(rf"(?<=[{_CJK}])\s+(?=[{_CJK}])", "", text)
    text = re.sub(rf"\s+([，。、；：！？）」』])", r"\1", text)
    text = re.sub(rf"([（「『])\s+", r"\1", text)
    return re.sub(r"\s{2,}", " ", text).strip()


def _fit(text: str, limit: int) -> str:
    """按密度收短要点：只在标点处断句，绝不把词切开。"""
    if len(text) <= limit:
        return text
    head = text[: limit + 1]
    cut = max(head.rfind(ch) for ch in _CUT_PUNCT)
    if cut >= max(6, int(limit * 0.5)):
        return text[:cut].rstrip()
    return text


def _render(prompt: str, *, pages: int, kind: str, title: str, density: float) -> str:
    resolved_title = _tidy(_one_line(title) or _guess_title(prompt, kind))
    words = _keywords(prompt, resolved_title, kind)
    while len(words) < 4:
        words.append(_GENERIC_WORDS.get(kind, _GENERIC_WORDS["business"])[len(words) % 4])

    fmt = {
        "topic": _short_topic(resolved_title, kind),
        "kw1": words[0],
        "kw2": words[1],
        "kw3": words[2],
        "kw4": words[3],
    }
    total_slides = max(4, int(pages))
    section_count = max(4, total_slides - 1)     # 封面由 front matter 自动生成
    sections = _pick_sections(kind, section_count)

    base_count = _BASE_BULLETS.get(kind, 3)
    bullet_count = max(1, min(6, round(base_count * density)))
    target_chars = max(12, min(60, round(_BASE_CHARS.get(kind, 24) * density)))

    lines: list[str] = [
        "---",
        f"title: {resolved_title}",
        f"subtitle: {_SUBTITLES.get(kind, _SUBTITLES['business'])}",
        "author: ",
        "org: ",
        "date: ",
        f"kicker: {KINDS.get(kind, KINDS['business'])}",
        f"footer: {resolved_title[:18]}",
        "---",
        "",
    ]

    for index, section in enumerate(sections):
        is_layout = bool(section.get("layout"))
        want = len(section["bullets"]) if is_layout else bullet_count
        bullets = [_tidy(b.format(**fmt)) for b in section["bullets"]][:want]
        if not is_layout:
            bullets = [_fit(b, target_chars) for b in bullets]
        while len(bullets) < want:
            bullets.append(_tidy(
                f"补充：关于「{resolved_title}」的一个具体细节（请替换为真实内容 {len(bullets) + 1}）"
            ))
        note = _tidy(section.get("note") or "")
        if index == 0:
            opening = _one_line(_OPENING_NOTES.get(kind, ""))
            if opening and opening not in note and "开场" not in note:
                note = f"{opening} {note}".strip()

        kicker = _one_line(section.get("kicker", ""))
        if kicker:
            lines.append(f"[{kicker}]")
        lines.append(f"# {_tidy(_one_line(section['title']))}")
        lines.append("")
        if section.get("layout"):
            lines.append(f"<!-- layout: {section['layout']} -->")
            lines.append("")
        for bullet in bullets:
            lines.append(f"- {bullet}")
        lines.append("")
        if note:
            lines.append("::: notes")
            lines.append(_one_line(note))
            lines.append(":::")
            lines.append("")
        if index != len(sections) - 1:
            lines.append("---")
            lines.append("")

    return "\n".join(lines).rstrip() + "\n"


def _normalize(text: str) -> str:
    text = str(text or "").replace("\r\n", "\n").replace("\r", "\n")
    text = text.strip()
    if text.startswith("```"):
        text = text.split("\n", 1)[1] if "\n" in text else ""
        if text.rstrip().endswith("```"):
            text = text.rstrip()[:-3]
    text = text.strip("\n")
    return text + "\n"


def _valid_outline(text: str) -> bool:
    try:
        doc = _md.parse(_normalize(text))
    except Exception:
        return False
    slides = doc.get("slides") or []
    titled = [s for s in slides if (s.get("title") or "").strip()]
    if len(titled) < 3:
        return False
    if not (doc.get("meta") or {}).get("title"):
        return False
    return True


# ------------------------------------------------------------------ 对外入口


def build_outline(prompt: str, *, pages: int = 10, kind: str = "auto",
                  profile: dict | None = None, title: str = "") -> str:
    """把一句提示词长成可直接渲染的 Markdown 大纲。

    有 LLM 时先试 LLM（失败/不合格自动降级），始终返回合法 Markdown。
    """
    prompt = str(prompt or "").strip()
    resolved_kind = resolve_kind(kind, prompt)
    try:
        pages = int(pages)
    except (TypeError, ValueError):
        pages = 10
    pages = max(4, min(40, pages))

    style = (profile or {}).get("style") if isinstance(profile, dict) else None
    style = style if isinstance(style, dict) else {}
    try:
        density = float(style.get("density") or 1.0)
    except (TypeError, ValueError):
        density = 1.0
    density = max(0.6, min(1.4, density))

    if llm_available():
        try:
            text = llm_outline(prompt, pages=pages, kind=resolved_kind,
                               title=title, density=density)
        except Exception:
            text = None
        if text and _valid_outline(text):
            return _normalize(text)

    return _render(prompt, pages=pages, kind=resolved_kind,
                   title=title, density=density)
