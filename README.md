# dwg-block-compare — AutoCAD DWG 双版本图块对比与分色标记 | DWG Block Diff & Color Markup

[![AutoCAD](https://img.shields.io/badge/AutoCAD-2023%20%2B%20COM-red)](https://www.autodesk.com/)
[![Python](https://img.shields.io/badge/Python-3.8%2B-blue)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey)](https://github.com/hwdemtv/dwg-block-compare)
[![Claude Code Skill](https://img.shields.io/badge/Claude%20Code-Skill-8A2BE2)](SKILL.md)

对比新旧两版 AutoCAD DWG 图纸中指定图块(摄像机、音箱、插座等设备符号)的**位置与数量差异**,并在新版图纸副本上自动绘制分色标记圈与图例,同时生成 Word 对比方案。Compare device blocks between two DWG revisions, color-mark every difference (added / moved / unchanged / removed) on a copy of the new drawing, and export a Word report.

适用于设计院出图后的**版本校核、图纸会审、变更追溯**——比如弱电深化设计中"这版图纸摄像机比上版多了哪些、挪了哪些"这类问题,几分钟内给出可核验的答案。Works as a Claude Code skill or as standalone Python scripts.

## 标记语义 Markup Legend

| 标记 | 含义 | 判定 |
|---|---|---|
| 🔴 红圈 | 两版位置完全一致 | 偏差 < 1mm |
| 🟡 黄圈 | 位置基本一致(微调) | 偏差 1~200mm |
| 🟢 绿圈 | 本版新增 | 新图有、旧图无 |
| 🔵 蓝圈+叉 | 移位或删除 | 旧图有、新图无,画在旧图原位置 |

每类图块 4 个独立图层,图例自动生成在真实内容正上方,全部标记经程序化核验(圈心与符号位置偏差 < 1mm)。

## 为什么不用插入点?(核心方法:炸开真值法)

天正/设计院图块常有 **~161000mm 的基点偏移**和旋转/镜像属性——直接按 `InsertionPoint` 定位会全部错位,`GetBoundingBox` 对旋转/镜像的 INSERT 也返回未变换几何。本工具把块参照在临时副本中逐个 `Explode()`,以炸出图元的总包围盒中心作为**符号真实显示位置**,再做多阈值一对一匹配(带数量闭环断言,杜绝漏配/重配)。

## 快速开始 Quick Start

```bash
pip install pywin32 python-docx

# 1. 对比 → truth.json(自动炸开取真值、匹配、断言闭环)
python scripts/dwg_compare.py --old 旧图.dwg --new 新图.dwg \
  --map "枪式摄像机=gun;A$C69077507=sph" --out truth.json
#   两版图纸基点/坐标系不一致时,按共识偏移向量修正后重跑(负值必须用 = 连接):
# python scripts/dwg_compare.py ... --old-offset=-207.5866,9.8875 --out truth.json

# 2. 在新图副本上画分色标记 + 图例
python scripts/dwg_mark.py --new-dwg 新图.dwg --mark-dwg 标记图.dwg \
  --truth truth.json --spec "gun=枪机;sph=球机" --old-label 20260325 --radius 600

# 3. 程序化核验(必须全过再交付)
python scripts/dwg_verify.py --mark-dwg 标记图.dwg --truth truth.json \
  --spec "gun=枪机;sph=球机" --old-label 20260325 --radius 600

# 4. 生成 Word 对比方案
python scripts/gen_report.py --truth truth.json --out 方案.docx \
  --spec "gun=枪式摄像机;sph=球机" --title 图纸名 \
  --old-dwg 旧图.dwg --new-dwg 新图.dwg --mark-file 标记图.dwg
```

## 四个脚本

| 脚本 | 作用 |
|---|---|
| `scripts/dwg_compare.py` | 炸开真值法取符号真实位置 + 一对一匹配(闭环断言、疑似移位对提示、`--old-offset` 两版基点/坐标系差异修正、AutoCAD 自动拉起) |
| `scripts/dwg_mark.py` | 新图副本上画分色标记圈 + 叉 + 数据驱动图例(不用 EXTMAX,避开杂散图元;画前自动清理底图同名标记图层残留) |
| `scripts/dwg_verify.py` | 逐图层核验圈数/半径/圈心配对(0.5mm 容差)/叉线数量 |
| `scripts/gen_report.py` | 生成带坐标表的 Word 对比方案(含基点修正自动注记) |

`references/pitfalls.md` 收录了 AutoCAD COM + 天正($TCHSYS$)环境的 12 类坑:gen_py 早绑定缓存污染、RPC 拒绝重试、ModelSpace.Item 不可用、acad.exe 崩溃恢复路径、scratch 文件锁、匹配闭环、EXTMAX 天边图例、两版图纸基点偏移修正、源图标记图层残留清理、COM 属性获取需整体 lambda 重试等——做 AutoCAD COM 自动化的人都用得上。

## 作为 Claude Code 技能使用

把本目录复制到 `~/.claude/skills/dwg-block-compare/`,对 Claude Code 说"对比这两个 DWG 里的摄像机图块并标记差异"即可自动触发全流程。详见 [SKILL.md](SKILL.md)。

## 依赖

- Windows + AutoCAD(已验证 2023;COM 自动化,无需额外 CAD 插件)
- Python 3.8+ 与 `pywin32`、`python-docx`

## 关键词 Keywords

DWG diff · DWG comparison · AutoCAD automation · CAD block compare · 图块对比 · 图纸版本对比 · 图纸差异标记 · 天正图纸 · 设计变更标记 · 设备点位核对 · 弱电深化 · pywin32 COM · Claude Code skill
