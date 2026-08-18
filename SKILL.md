---
name: dwg-block-compare
description: 对比新旧两版 AutoCAD DWG 图纸中指定图块(摄像机、音箱、插座等设备符号)的位置与数量差异,并在新版图纸副本上分色圈画标记(红=一致/黄=微调/绿=新增/蓝+叉=旧有新无),生成 Word 对比方案。当用户要求"对比 DWG 图块"、"图纸版本差异标记"、"新增/删除设备标记"、"CAD 图块分析标记"时使用。依赖 Windows + AutoCAD + pywin32,基于炸开真值法处理天正图纸的大偏移基点块。
---

# DWG 图块对比标记

对比两版 DWG 中指定图块的位置差异,在新图副本上画分色标记圈 + 图例,输出可核验的标记图和 Word 方案。管线已在光谷四十小项目(综合楼/地下室/室外,天正+AutoCAD 2023 环境)全量验证。

## 工作流(四步,顺序执行)

### 1. 对比 → truth JSON

```bash
python scripts/dwg_compare.py --old 旧图.dwg --new 新图.dwg \
  --map "枪式摄像机=gun;A$C69077507=sph" --out work/truth.json
```

- `--map` 为 块名=标签,分号分隔;中文块名与匿名块(`A$C…`)都支持。块名不确定时先在 CAD 里 `LI` 查,或问用户。
- 自动: 附着/启动 AutoCAD、清理残留 scratch 文档、临时副本炸开取真实位置、匹配、闭环断言、疑似移位对提示。
- 控制台会打印各类数量与"建议圈半径≈对角线中位/2+250"——**记下这个数,第 2 步要用**。

分类口径: <1mm 完全一致(exact) / 1~200mm 基本一致(near) / >200mm 移位或删除(old_only)与新增(new_only)。200~500mm 的对不消耗配额,两侧分别归类。

### 2. 标记 → 标记图 DWG

```bash
python scripts/dwg_mark.py --new-dwg 新图.dwg \
  --mark-dwg 新图名_对比标记.dwg --truth work/truth.json \
  --spec "gun=枪机;sph=球机" --old-label 20260325 --radius 600
```

- `--radius`: 综合/常规 600,符号大的(如地下室 diag≈947)用 750;参考第 1 步打印的建议值。
- `--old-label` 用旧版日期号,进图层名。
- 每设备 4 图层: 红(1)一致、黄(2)微调、绿(3)新增、蓝(5)+叉 旧有新无(画在旧图原位置);图例按内容坐标范围放在真实内容正上方。

### 3. 核验(必须跑,全绿才算完)

```bash
python scripts/dwg_verify.py --mark-dwg 新图名_对比标记.dwg \
  --truth work/truth.json --spec "gun=枪机;sph=球机" \
  --old-label 20260325 --radius 600
```

逐图层核对圈数、半径、圈心与 truth 配对(容差 0.5mm)、叉线数量。任何失败先查参数是否与 mark 一致。

### 4. 报告

```bash
python scripts/gen_report.py --truth work/truth.json --out xx_对比方案.docx \
  --spec "gun=枪式摄像机;sph=球机" --title 室外弱电总平面图 \
  --old-dwg 室外弱电总平面图20260325.dwg --new-dwg 室外弱电总平面图20260604.dwg \
  --mark-file 室外弱电总平面图20260604_对比标记.dwg --radius 600
```

## 关键约束

- **位置一律用炸开真值**(dwg_compare.py 已实现),绝不用 InsertionPoint——本项目块基点偏移 ~161000mm 且带旋转/镜像。
- **标记图必须核验通过后才能交付**。
- `--spec` 的 tag 必须与 `--map` 的 tag 对应;dwg_mark 与 dwg_verify 的 `--spec/--old-label/--radius` 必须完全一致。
- 疑似移位对(dwg_compare 打印 ⚠ 行)要在最终回复与文档里单独列出,供设计确认,不要擅自归类。
- 输出目录: 标记图与 docx 放用户指定目录(通常为新图所在目录),truth JSON 与 scratch 放工作目录(默认 `C:\Users\hwdem\dwg_work`)。

## 报错排障

报 com_error / AttributeError / CLASSNOTREG / "服务器运行失败" / PermissionError 时,**先读 [references/pitfalls.md](references/pitfalls.md)**——天正+AutoCAD COM 环境的 9 类坑与修法(gen_py 缓存删除、COM 重试、acad.exe 崩溃恢复路径、scratch 锁文件、选择集残留、匹配闭环、EXTMAX 天边图例等)都在里面,勿重新发明。
