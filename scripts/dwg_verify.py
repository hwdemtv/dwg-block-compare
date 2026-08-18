# -*- coding: utf-8 -*-
"""核验标记图: 逐图层统计圈/叉数量、圈心与 truth JSON 配对(容差0.5mm)、半径检查

用法:
  python dwg_verify.py --mark-dwg 标记图.dwg --truth truth.json \
      --spec "gun=枪机;sph=球机" --old-label 20260325 --radius 600
"""
import argparse
import json
import math
import sys
import time

import win32com.client
import pythoncom

sys.stdout.reconfigure(encoding="utf-8")


def com_retry(fn, *args, tries=15, wait=1.5, **kwargs):
    for i in range(tries):
        try:
            return fn(*args, **kwargs)
        except (pythoncom.com_error, AttributeError):
            if i == tries - 1:
                raise
            time.sleep(wait)
    return None


def new_ss(doc, name):
    for _ in range(5):
        try:
            existing = com_retry(doc.SelectionSets.Item, name)
            com_retry(existing.Delete)
        except Exception:
            pass
        try:
            return com_retry(doc.SelectionSets.Add, name)
        except pythoncom.com_error:
            time.sleep(1)
    raise RuntimeError(f"无法创建选择集 {name}")


def match_sets(centers, targets, tol=0.5):
    if len(centers) != len(targets):
        return None
    used = set()
    worst = 0.0
    for (x, y) in centers:
        best, bd = None, 1e18
        for j, (tx, ty) in enumerate(targets):
            if j in used:
                continue
            d = math.hypot(tx - x, ty - y)
            if d < bd:
                best, bd = j, d
        if best is None or bd > tol:
            return None
        used.add(best)
        worst = max(worst, bd)
    return worst


def main():
    ap = argparse.ArgumentParser(description="DWG 标记图程序化核验")
    ap.add_argument("--mark-dwg", required=True)
    ap.add_argument("--truth", required=True)
    ap.add_argument("--spec", required=True, help="tag=中文名,分号分隔,与 dwg_mark.py 一致")
    ap.add_argument("--old-label", required=True)
    ap.add_argument("--radius", type=float, default=600.0)
    args = ap.parse_args()

    names = {}
    for item in args.spec.split(";"):
        tag, name = item.split("=", 1)
        names[tag.strip()] = name.strip()

    truth = json.load(open(args.truth, encoding="utf-8"))

    def layer_names(tag, cname):
        gone = (f"标记-旧图有新图无(移位或删除)" if len(names) == 1
                else f"标记-旧图有新图无-{cname}(移位或删除)")
        return {
            "exact": f"标记-{cname}与{args.old_label}版一致",
            "near": f"标记-{cname}轻微偏移(50~200mm)",
            "new_only": f"标记-{cname}新增(旧图无)",
            "old_only": gone,
        }

    acad = win32com.client.GetActiveObject("AutoCAD.Application")
    doc = com_retry(acad.Documents.Open, args.mark_dwg, True)
    time.sleep(3)
    print(f"打开(只读): {doc.Name}")

    issues = []
    for tag, cname in names.items():
        lmap = layer_names(tag, cname)
        for tkey, lname in lmap.items():
            ss = new_ss(doc, f"SS_V{abs(hash((tag, tkey))) % 100000}")
            fc = win32com.client.VARIANT(pythoncom.VT_I2 | pythoncom.VT_ARRAY, [8])
            fv = win32com.client.VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_ARRAY, [lname])
            com_retry(ss.Select, 5, None, None, fc, fv)
            circles, lines = [], 0
            for k in range(com_retry(lambda: ss.Count)):
                e = com_retry(ss.Item, k)
                en = com_retry(lambda x=e: x.EntityName)
                if en == "AcDbCircle":
                    c = list(com_retry(lambda x=e: x.Center))
                    r = com_retry(lambda x=e: x.Radius)
                    circles.append(((c[0], c[1]), r))
                elif en == "AcDbLine":
                    lines += 1
            com_retry(ss.Delete)

            tgt = [(p[0], p[1]) for p in truth[tag][tkey]]
            gname = f"{cname}-{tkey}"
            if len(circles) != len(tgt):
                issues.append(f"{gname}: 圈 {len(circles)} != 期望 {len(tgt)}")
            elif match_sets([c for c, _ in circles], tgt) is None:
                issues.append(f"{gname}: 圈心无法配对(容差0.5mm)")
            else:
                w = match_sets([c for c, _ in circles], tgt)
                print(f"  ✓ {gname}: {len(circles)} 圈全部配对, 最大偏差 {w:.6f}mm")
            bad_r = sum(1 for _, r in circles if abs(r - args.radius) > 0.01)
            if bad_r:
                issues.append(f"{gname}: {bad_r} 个圈半径≠{args.radius:.0f}")
            exp_lines = 2 * len(tgt) if tkey == "old_only" else 0
            if lines != exp_lines:
                issues.append(f"{gname}: 叉线 {lines} != 期望 {exp_lines}")

    com_retry(doc.Close, False)
    print("=" * 50)
    if issues:
        print("问题:")
        for s in issues:
            print(" -", s)
        sys.exit(1)
    print("标记核验全部通过。")


if __name__ == "__main__":
    main()
