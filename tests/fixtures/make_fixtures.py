"""生成 tests/fixtures 下的只读样例（评审 F2）。

样例一旦提交就**只读**：tests/test_profile.py 与 tests/test_e2e.py 只读取它们，
任何构建命令都不允许写回本目录（scripts/ppt.ps1 会拒绝）。需要重建样例时手动运行：

    & "$env:USERPROFILE\\.dsh\\dsh-runtimes\\dsh-primary-runtime\\dependencies\\python\\python.exe" tests\\fixtures\\make_fixtures.py

为什么固定字体：样例用于断言「档案能读出案例的字体」，所以样例必须自带
确定的字体（微软雅黑 / 宋体），不能跟随 themes.py 的当前取值漂移。
配色仍用内置 academic / editorial 主题，与契约 §4.1 的示例一致。
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
KIT = ROOT / "kit"
if str(KIT) not in sys.path:
    sys.path.insert(0, str(KIT))

from pptx import Presentation                            # noqa: E402
from pptx.dml.color import RGBColor                      # noqa: E402
from pptx.enum.shapes import MSO_SHAPE                   # noqa: E402
from pptx.util import Inches, Pt                         # noqa: E402

from pptxkit import builder, markdown as md                 # noqa: E402
from pptxkit.themes import clone_with, get_theme            # noqa: E402

FIXTURES = Path(__file__).resolve().parent

ACADEMIC_SECTIONS = [
    ("研究背景与问题", ["检测需求正从云端下沉到边缘设备",
                        "现有大模型在端侧延迟与功耗都不可接受",
                        "本文关注：精度不降的前提下压缩计算量"]),
    ("相关工作与研究缺口", ["两阶段检测器精度高但速度慢",
                            "单阶段检测器速度快但小目标召回不足",
                            "缺少面向边缘芯片的结构化设计方法"]),
    ("研究目标与贡献", ["提出轻量级主干与特征融合结构",
                        "在公开数据集上保持精度并降低 42% 计算量",
                        "给出从训练到端侧部署的完整链路"]),
    ("整体技术路线", ["输入分辨率自适应缩放",
                      "主干网络：深度可分离卷积 + 通道重排",
                      "检测头：解耦分类与回归分支"]),
    ("轻量级主干设计", ["以 MobileNetV3 为基线改造",
                        "引入注意力增强的高效残差块",
                        "通道数按算力预算逐级裁剪"]),
    ("多尺度特征融合", ["自顶向下与自底向上双向融合",
                        "增加 P2 层提升小目标召回",
                        "融合权重随尺度自适应调整"]),
    ("检测头与损失函数", ["解耦头分别处理分类与定位",
                          "分类用焦点损失缓解正负样本失衡",
                          "定位用 CIoU 损失提升回归质量"]),
    ("训练策略与数据增强", ["余弦退火学习率 + 热身",
                            "Mosaic 与 MixUp 组合增强",
                            "知识蒸馏从大模型迁移软标签"]),
    ("实验设置与数据集", ["公开数据集 + 自建边缘场景数据",
                          "输入 640×640，批大小 32",
                          "训练 300 轮，单卡 V100"]),
    ("主要实验结果", ["mAP 提升 1.8 个点，计算量下降 42%",
                      "端侧推理延迟从 46ms 降到 27ms",
                      "参数量压缩到 3.2M"]),
    ("消融实验", ["主干改造贡献 +0.9 mAP",
                  "多尺度融合贡献 +0.7 mAP",
                  "蒸馏贡献 +0.4 mAP"]),
    ("与其他方法对比", ["对比 YOLO 系列与轻量检测器",
                        "同算力下精度领先 1.2 个点",
                        "同精度下延迟降低 35%"]),
    ("边缘设备部署", ["导出 ONNX 并做算子融合",
                      "INT8 量化后精度损失小于 0.5",
                      "在 RK3588 上达到 38 FPS"]),
    ("工程落地与限制", ["复杂光照与遮挡仍是主要失败场景",
                        "量化对极小目标影响明显",
                        "后续引入时序信息改善稳定性"]),
    ("结论", ["轻量化设计可以在端侧保持可用精度",
              "结构化改造 + 蒸馏是性价比最高的组合",
              "完整链路已在真实设备上验证"]),
    ("致谢与讨论", ["感谢实验室同学与算力支持",
                    "欢迎讨论端侧部署的工程细节",
                    "联系方式：demo@example.com"]),
]

EDITORIAL_SECTIONS = [
    ("先问一个问题", ["如果 AI 成为同事，你的工作会变成什么样？",
                      "工具在变，但判断力仍然是稀缺品",
                      "今天聊聊我们真实的协作现场"]),
    ("三个正在发生的变化", ["从写代码到描述意图",
                            "从交付文档到交付可运行的流程",
                            "从个人效率到团队吞吐"]),
    ("把 AI 当同事的四个原则", ["先给它上下文，再给它任务",
                                "小步验证，别一次梭哈",
                                "让产出可检查、可回退",
                                "保留人类做最终判断"]),
    ("一次真实的协作复盘", ["需求拆解：20 分钟定边界",
                            "并行产出：三个人同时推进",
                            "验收阶段：用清单逐条核对"]),
    ("我们踩过的坑", ["提示词太抽象，结果不可用",
                      "没有版本管理，改动互相覆盖",
                      "过度信任，漏掉了边界条件"]),
    ("可复用的工作流", ["模板化输入：结构先于内容",
                        "固定检查项：让质量可衡量",
                        "复盘沉淀：把经验写回模板"]),
    ("对团队的影响", ["新人上手时间缩短一半",
                      "重复劳动被压缩到原来的三分之一",
                      "更多时间花在判断与设计上"]),
    ("下一步计划", ["把工作流沉淀成内部工具",
                    "建立可复用的提示词库",
                    "每季度做一次效果复盘"]),
    ("给听众的三点建议", ["从最小闭环开始",
                          "记录每一次失败的原因",
                          "别把工具当成答案"]),
    ("Q&A", ["欢迎提出具体场景的问题",
             "会后可以一起拆解你的工作流"]),
]


def _markdown(title: str, subtitle: str, sections: list[tuple[str, list[str]]],
              theme: str) -> str:
    lines = [
        "---",
        f"title: {title}",
        f"subtitle: {subtitle}",
        "author: 张 三",
        "org: 某某大学 · 计算机科学与技术学院",
        "date: 2026 年 10 月",
        f"theme: {theme}",
        "---",
        "",
    ]
    for index, (heading, bullets) in enumerate(sections, start=1):
        lines.append(f"# {heading}")
        lines.extend(f"- {bullet}" for bullet in bullets)
        if index != len(sections):
            lines.append("")
    return "\n".join(lines) + "\n"


def main() -> int:
    academic_md = _markdown(
        "面向复杂场景的轻量级目标检测方法研究",
        "从结构设计到边缘部署的完整链路",
        ACADEMIC_SECTIONS, "academic")
    editorial_md = _markdown(
        "把 AI 当同事之后",
        "一次真实的协作复盘",
        EDITORIAL_SECTIONS, "editorial")

    academic_theme = clone_with(
        get_theme("academic"),
        font_title="微软雅黑", font_body="微软雅黑", font_display="微软雅黑")
    editorial_theme = clone_with(
        get_theme("editorial"),
        font_title="宋体", font_body="微软雅黑", font_display="宋体")

    for name, text, theme in (
        ("case-academic.pptx", academic_md, academic_theme),
        ("case-editorial.pptx", editorial_md, editorial_theme),
    ):
        target = FIXTURES / name
        builder.build(md.parse(text), target, theme_obj=theme)
        if name == "case-academic.pptx":
            _add_palette_swatches(target)
        print(f"wrote {target} ({len(Presentation(str(target)).slides)} slides)")
    return 0


def _add_palette_swatches(path: Path) -> None:
    """给学术样例补两块主题色证据（浅底强调色 + 分隔线）。

    profile.analyze 会挑「最接近主题色的实际填充/描边」；纯要点页没有用到
    accent_soft / rule，会让分析结果漂移到近似色。加两个不含文字的装饰形状，
    让样例的真实配色与契约 §4.1 的 academic 取值一致（页数不变）。
    """
    prs = Presentation(str(path))
    slide = prs.slides[1]
    chip = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                  Inches(0.95), Inches(6.95), Inches(0.5), Inches(0.18))
    chip.fill.solid()
    chip.fill.fore_color.rgb = RGBColor(0xE9, 0xEE, 0xF7)
    chip.line.fill.background()
    rule = slide.shapes.add_connector(1, Inches(0.95), Inches(6.85),
                                      Inches(6.0), Inches(6.85))
    rule.line.color.rgb = RGBColor(0xD5, 0xDC, 0xE8)
    rule.line.width = Pt(1)
    prs.save(str(path))


if __name__ == "__main__":
    raise SystemExit(main())
