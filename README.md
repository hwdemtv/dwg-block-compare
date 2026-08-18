# dwg-block-compare

对比新旧两版 AutoCAD DWG 图纸中指定图块(摄像机、音箱、插座等设备符号)的位置与数量差异,并在新版图纸副本上分色圈画标记的 Claude Code 技能。

标记语义:红圈=两版一致 / 黄圈=微调(1~200mm) / 绿圈=新增 / 蓝圈+叉=旧图有新图无。

## 组成

- `SKILL.md` — 技能说明与四步工作流(对比 → 标记 → 核验 → 报告)
- `scripts/dwg_compare.py` — 炸开真值法取符号真实位置 + 一对一匹配(带闭环断言)
- `scripts/dwg_mark.py` — 在新图副本画分色标记圈 + 图例
- `scripts/dwg_verify.py` — 逐图层程序化核验(圈数/半径/圈心配对)
- `scripts/gen_report.py` — 生成 Word 对比方案
- `references/pitfalls.md` — AutoCAD COM + 天正环境的坑清单

## 依赖

Windows + AutoCAD(已验证 2023)+ Python 包 `pywin32`、`python-docx`。

## 安装

将本目录(含 SKILL.md)复制到 `~/.claude/skills/dwg-block-compare/` 即可,Claude Code 会在收到"对比 DWG 图块"类请求时自动触发。
