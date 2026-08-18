# -*- coding: utf-8 -*-
"""在新版 DWG 副本上画对比标记: 每类设备 4 图层(红=一致 黄=微调 绿=新增 蓝+叉=旧有新无) + 图例

用法:
  python dwg_mark.py --new-dwg 新图.dwg --mark-dwg 输出标记图.dwg --truth truth.json \
      --spec "gun=枪机;sph=球机" --old-label 20260325 --radius 600
"""
import argparse
import json
import os
import shutil
import sys
import time

import win32com.client
import pythoncom

sys.stdout.reconfigure(encoding="utf-8")

COLORS = {"exact": 1, "near": 2, "new": 3, "gone": 5}  # 红黄绿蓝(ACI)


def com_retry(fn, *args, tries=15, wait=1.5, **kwargs):
    """天正环境瞬时代理故障既报 com_error 也报 AttributeError(名字解析失败),都重试"""
    for i in range(tries):
        try:
            return fn(*args, **kwargs)
        except (pythoncom.com_error, AttributeError):
            if i == tries - 1:
                raise
            time.sleep(wait)
    return None


def pt_var(x, y, z=0.0):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [x, y, z])


def open_doc(acad, path, readonly=False):
    for i in range(15):
        try:
            return acad.Documents.Open(path, readonly)
        except Exception as e:
            print(f"  open 重试 {i}: {type(e).__name__} {e}")
            time.sleep(3)
    raise RuntimeError(f"无法打开 {path}")


def main():
    ap = argparse.ArgumentParser(description="DWG 对比标记绘制")
    ap.add_argument("--new-dwg", required=True, help="新版 DWG(复制为标记底图)")
    ap.add_argument("--mark-dwg", required=True, help="标记图输出路径(*.dwg)")
    ap.add_argument("--truth", required=True, help="dwg_compare.py 输出的 JSON")
    ap.add_argument("--spec", required=True,
                    help="tag=中文名 映射,分号分隔,如 \"gun=枪机;sph=球机\"")
    ap.add_argument("--old-label", required=True, help="旧版标识(用于图层名,如 20260325)")
    ap.add_argument("--radius", type=float, default=600.0, help="圈半径 mm(综合楼/室外600 地下室750)")
    args = ap.parse_args()

    names = {}
    for item in args.spec.split(";"):
        tag, name = item.split("=", 1)
        names[tag.strip()] = name.strip()

    truth = json.load(open(args.truth, encoding="utf-8"))

    # 图层名: 每设备 4 层。gone 层名两种历史变体都试,保持与既有标记图兼容
    layers = {}
    for tag, cname in names.items():
        layers[f"{tag}_exact"] = (f"标记-{cname}与{args.old_label}版一致", COLORS["exact"])
        layers[f"{tag}_near"] = (f"标记-{cname}轻微偏移(50~200mm)", COLORS["near"])
        layers[f"{tag}_new"] = (f"标记-{cname}新增(旧图无)", COLORS["new"])
        layers[f"{tag}_gone"] = (f"标记-旧图有新图无-{cname}(移位或删除)", COLORS["gone"])
    # 单设备时 gone 层名不带后缀(历史兼容: 枪机独占时为"标记-旧图有新图无(移位或删除)")
    if len(names) == 1:
        only_tag = next(iter(names))
        layers[f"{only_tag}_gone"] = (f"标记-旧图有新图无(移位或删除)", COLORS["gone"])

    if os.path.exists(args.mark_dwg):
        os.remove(args.mark_dwg)
    shutil.copy(args.new_dwg, args.mark_dwg)
    print(f"已复制 -> {os.path.basename(args.mark_dwg)}")

    acad = win32com.client.GetActiveObject("AutoCAD.Application")
    doc = open_doc(acad, args.mark_dwg, False)
    time.sleep(3)
    print(f"打开: {doc.Name}  ReadOnly={doc.ReadOnly}")

    for key, (lname, color) in layers.items():
        lay = com_retry(doc.Layers.Add, lname)
        com_retry(lambda l=lay, c=color: setattr(l, "Color", c))

    all_pts = []
    for tag, cname in names.items():
        g = truth[tag]
        for cat, key in (("exact", "exact"), ("near", "near"), ("new_only", "new")):
            lname = layers[f"{tag}_{key}"][0]
            for (x, y) in g[cat]:
                c = com_retry(doc.ModelSpace.AddCircle, pt_var(x, y), args.radius)
                com_retry(lambda o=c, l=lname: setattr(o, "Layer", l))
        gone_l = layers[f"{tag}_gone"][0]
        for (x, y) in g["old_only"]:
            c = com_retry(doc.ModelSpace.AddCircle, pt_var(x, y), args.radius)
            com_retry(lambda o=c: setattr(o, "Layer", gone_l))
            r2 = args.radius * 0.70710678
            for dx1, dy1, dx2, dy2 in ((-r2, -r2, r2, r2), (-r2, r2, r2, -r2)):
                ln = com_retry(doc.ModelSpace.AddLine,
                               pt_var(x + dx1, y + dy1), pt_var(x + dx2, y + dy2))
                com_retry(lambda o=ln: setattr(o, "Layer", gone_l))
        all_pts += [p for k in ("exact", "near", "old_only", "new_only") for p in g[k]]
        print(f"[{cname}] 一致{len(g['exact'])} 微调{len(g['near'])} "
              f"新增{len(g['new_only'])} 移位/删除{len(g['old_only'])}")

    # 图例: 真实内容正上方(勿用 EXTMAX,杂散图元会到天边)
    top_y = max(y for _, y in all_pts) + args.radius + 3 * args.radius
    lx = min(x for x, _ in all_pts)
    h = args.radius * 1.2

    legend = [(f"图块对比标记·{'+'.join(names.values())}（对比基准：{args.old_label}版）",
               layers[f"{next(iter(names))}_exact"][0])]
    for tag, cname in names.items():
        g = truth[tag]
        legend.append((f"—— {cname}（旧版{len(g['old'])}台 / 新版{len(g['new'])}台）——",
                       layers[f"{tag}_exact"][0]))
        if g["exact"]:
            legend.append((f"红圈 {len(g['exact'])} 处：两版位置完全一致（<1mm）", layers[f"{tag}_exact"][0]))
        if g["near"]:
            legend.append((f"黄圈 {len(g['near'])} 处：位置基本一致（偏移50~200mm）", layers[f"{tag}_near"][0]))
        if g["new_only"]:
            legend.append((f"绿圈 {len(g['new_only'])} 处：本版新增（旧图无）", layers[f"{tag}_new"][0]))
        if g["old_only"]:
            legend.append((f"蓝圈+叉 {len(g['old_only'])} 处：旧图有、新图无（移位或删除，画在旧图原符号位置）",
                           layers[f"{tag}_gone"][0]))

    for k, (text, lname) in enumerate(legend):
        t = com_retry(doc.ModelSpace.AddText, text, pt_var(lx, top_y - k * h * 1.6), h)
        com_retry(lambda o=t, l=lname: setattr(o, "Layer", l))
    print(f"图例 {len(legend)} 行 @({lx:.0f},{top_y:.0f})")

    com_retry(doc.Save)
    com_retry(doc.Close, False)
    print(f"完成: {args.mark_dwg}")


if __name__ == "__main__":
    main()
