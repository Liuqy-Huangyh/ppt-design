---
marp: true
theme: wks
paginate: true
size: 16:9
footer: "Marp 引擎 · Markdown 直接产出可编辑 PPTX"
---

<!-- _class: cover -->

# Marp：Markdown 一键出 PPT

## `marp x.md --pptx` → 原生形状、可在 PowerPoint 里继续改

2026 · PPT 工作区示例

---

## 为什么用它

- 写作即排版：只关心内容层级，版式交给主题
- 一份源文件，同时产出 **PPTX / PDF / HTML**
- 版本可 diff，适合多人协作与长期维护
- 代码块、表格、公式、图片都支持

---

## 三种输出

| 命令 | 产物 | 适合 |
|---|---|---|
| `--pptx` | 原生可编辑 PPTX | 需要交付、领导要改 |
| `--pdf` | PDF | 定稿、发送、打印 |
| `--html` | 单页 HTML | 线上分享、嵌入网页 |

```powershell
.\scripts\ppt.ps1 md .\templates\marp-demo.md -Pptx
```

---

## 什么时候别用它

- 需要复杂动画与转场（用 PowerPoint 或 Slidev）
- 需要像素级设计控制（用 HTML 方案或直接在 PPT 里做）
- 图表要可编辑数据（用 python-pptx / PptxGenJS 生成）

> 工具的选择标准只有一个：后期谁维护。
