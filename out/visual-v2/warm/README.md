# editorial 暖纸色修复（t14）

t12 把 `editorial` 的纸底改成了「跟着强调色色相走」的冷调 `#F9F7FB`，丢了杂志风的米色纸感。
这次给 `design.palette()` 增加**暖纸模式**（显式开关、默认关闭），editorial 打开它，
纸底回到暖米色，其余主题与既有调用完全不受影响。

## 1. 参数与做法

```python
# design.py
PAPER_WARM_HUE = 80.0     # 暖纸色相（60–90° 区间内）
PAPER_WARM_L = 0.975      # 明度
PAPER_WARM_C = 0.010      # 低彩度
PAPER_WARM_RULE_L = 0.890 # 分隔线：同色相、更暗一档
PAPER_WARM_RULE_C = 0.012

palette(accent, *, mode="paper", warm_paper=False)   # warm_paper 默认 False
```

- `warm_paper=True` 只影响 `mode="paper"` 的 `bg` 与 `rule`；`ink / muted / accent /
  accent_soft` 仍然按强调色推导，保留「配色由单一强调色推导」的机制。
- `mode="clean"` 会忽略该参数；默认 `warm_paper=False` 时 `palette()` 的返回值与 t12
  逐值一致（`#F9F7FB / #DADCE2`），所以 `scripts/design-showcase.py` 等既有调用不受影响。
- `themes.py` 里只有 editorial 打开：`_colors("#1F3BFF", "paper", warm_paper=True)`；
  academic / minimal 仍是 `clean`。

## 2. 结果（editorial 六种颜色）

| 角色 | t12（冷调 paper） | t14（暖纸） | 说明 |
|---|---|---|---|
| bg | `#F9F7FB` | `#FAF6EF` | OKLCH L 0.9743 / C 0.0102 / H 81.8°，与旧暖米 `#F7F3EA` 的 RGB 差 ≤5 |
| rule | `#DADCE2` | `#DFDAD2` | 同暖色相 H 79.8°、L 0.890、C 0.012 |
| ink | `#181E2B` | `#181E2B` | 不变（强调色色相推导） |
| muted | `#6A707D` | `#6A707D` | 不变 |
| accent | `#1F3BFF` | `#1F3BFF` | 不变（对比 6.3:1，未触发压暗） |
| accent_soft | `#E5EDFF` | `#E5EDFF` | 不变（克莱因蓝浅底，保留纸面撞色） |

academic / minimal 的六种颜色与 t12 **逐值一致**（脚本比对）：

```
academic #FFFFFF #161E2B #69717D #1B3A6B #E8F1FF #DDE0E4   ← 与 t12 相同
minimal  #FFFFFF #2A1915 #7C6C67 #C54E2C #FEEBE6 #E3DEDD   ← 与 t12 相同
```

## 3. 对照物料

- `before/` — t12 的冷调 paper 版本（`out/visual-v2/after/editorial-*.png`）
- `after/` — 本次暖纸版本（同一个 12 页 deck 重新渲染）
- `contact-sheet-editorial-warm.png` — BEFORE / AFTER 两行 × cover / section / stats / quote / timeline
- `colors.txt` — 三主题六色值、OKLCH 参数、默认/暖纸对比、Theme 字段与 `apply/clone_with` 自检结果
- `decks/` — 三份模板重建产物：`academic.pptx` 17 页、`editorial.pptx` 12 页、`business.pptx` 9 页

## 4. 验证

| 项 | 结果 |
|---|---|
| 页数 | 17 / 12 / 9（未变） |
| qa | 14 / 10 / 8 条提示，与 t12 完全一致；全部是「该页没有讲者备注」，0 溢出 / 0 密度 / 0 缺字体 |
| 测试 | `tests/test_profile.py` 39 OK；另回归 `test_e2e.py` 7 OK、`test_server.py` 32 OK、`test_web.py` 29 OK |
| 兼容性 | `palette(mode="paper")` 默认仍返回 `#F9F7FB/#DADCE2`；`clean` 模式忽略 `warm_paper`；`profile.apply()/clone_with()` 未改动，测试全绿 |
| 验收构建 | `ppt.cmd build templates/editorial-share.md -Out out/_editorial-warm.pptx` → 12 页 |
