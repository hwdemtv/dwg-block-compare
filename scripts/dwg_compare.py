# -*- coding: utf-8 -*-
"""DWG 双版本图块对比: 炸开取真实符号位置 + 修正版一对一匹配

用法:
  python dwg_compare.py --old 旧图.dwg --new 新图.dwg \
      --map "枪式摄像机=gun;A$C5BEE1828=dome" --out truth.json [--acad-exe 路径]

输出 JSON: {tag: {block, old/new:[{ip,disp,w,h}], exact/near/old_only/new_only}}
匹配阈值: <1mm 一致, 1~200mm 微调, 200~500mm 视为移位(不配对,两侧分别归类), >500mm 不配对
"""
import argparse
import json
import math
import os
import shutil
import sys
import time

import win32com.client
import pythoncom

sys.stdout.reconfigure(encoding="utf-8")

SCRATCH_TAG = "scratch_cmp"


def com_retry(fn, *args, tries=15, wait=1.5, **kwargs):
    """天正环境瞬时代理故障既报 com_error 也报 AttributeError(名字解析失败),都重试"""
    for i in range(tries):
        try:
            return fn(*args, **kwargs)
        except (pythoncom.com_error, AttributeError) as e:
            if i == tries - 1:
                raise
            time.sleep(wait)
    return None


def attach_acad(acad_exe):
    try:
        return win32com.client.GetActiveObject("AutoCAD.Application")
    except Exception:
        pass
    if not acad_exe:
        raise RuntimeError(
            "AutoCAD COM 不可用且未提供 --acad-exe。崩后 COM 常无法自动重启,"
            "需直接启动 acad.exe(本机为 D:\\Program Files\\AutoCAD 2023\\AutoCAD 2023\\acad.exe)")
    os.startfile(acad_exe)
    for i in range(30):
        try:
            return win32com.client.GetActiveObject("AutoCAD.Application")
        except Exception:
            time.sleep(5)
    raise RuntimeError("启动 AutoCAD 后 150s 内 COM 仍不可用")


def close_scratch_docs(acad):
    """清理上次崩溃残留的 scratch 文档(会锁住临时文件)"""
    for i in range(com_retry(lambda: acad.Documents.Count)):
        try:
            d = com_retry(acad.Documents.Item, i)
            if SCRATCH_TAG in com_retry(lambda x=d: x.FullName):
                print(f"  关闭残留: {os.path.basename(d.FullName)}")
                com_retry(d.Close, False)
                time.sleep(1)
        except Exception:
            pass


def open_doc(acad, path, readonly=False):
    for i in range(15):
        try:
            return acad.Documents.Open(path, readonly)
        except Exception as e:
            print(f"  open 重试 {i}: {type(e).__name__} {e}")
            time.sleep(3)
    raise RuntimeError(f"无法打开 {path}")


def new_ss(doc, name):
    for attempt in range(5):
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


def explode_positions(acad, src_dwg, work_dir, block, label):
    """临时副本中逐个 Explode, 炸出图元包围盒中心=符号真实显示位置"""
    scratch = os.path.join(work_dir, f"{SCRATCH_TAG}_{label}.dwg")
    shutil.copy(src_dwg, scratch)
    doc = open_doc(acad, scratch, False)
    time.sleep(3)

    ss = new_ss(doc, "SS_X")
    fc = win32com.client.VARIANT(pythoncom.VT_I2 | pythoncom.VT_ARRAY, [2])
    fv = win32com.client.VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_ARRAY, [block])
    ss.Select(5, None, None, fc, fv)
    refs = [com_retry(ss.Item, i) for i in range(com_retry(lambda: ss.Count))]
    ss.Delete()
    print(f"[{label}][{block}] {len(refs)} 个")

    results, n_fail = [], 0
    for i, ref in enumerate(refs):
        ref = win32com.client.Dispatch(ref._oleobj_)
        try:
            ip = list(com_retry(lambda r=ref: r.InsertionPoint))
            prims = com_retry(ref.Explode)
            xs, ys = [], []
            for p in prims:
                try:
                    p1, p2 = p.GetBoundingBox()
                    # 过滤远处杂散图元(>1e6mm 视为垃圾)
                    if (p2[0] - p1[0]) < 1e6 and (p2[1] - p1[1]) < 1e6:
                        xs.extend([p1[0], p2[0]])
                        ys.extend([p1[1], p2[1]])
                except Exception:
                    pass
                try:
                    p.Delete()
                except Exception:
                    pass
            if xs:
                results.append({
                    "ip": [ip[0], ip[1]],
                    "disp": [(min(xs) + max(xs)) / 2, (min(ys) + max(ys)) / 2],
                    "w": max(xs) - min(xs), "h": max(ys) - min(ys),
                })
            else:
                n_fail += 1
        except Exception:
            n_fail += 1
        if (i + 1) % 40 == 0:
            print(f"  ... {i+1}/{len(refs)}")
    print(f"[{label}][{block}] 成功 {len(results)}, 失败 {n_fail}")
    com_retry(doc.Close, False)
    time.sleep(2)
    try:
        os.remove(scratch)
    except Exception:
        pass
    return results


def match(old, new):
    """修正版一对一匹配: 200~500mm 不消耗配额(按移位口径两侧分别归类), 带闭环断言"""
    old_pos = [r["disp"] for r in old]
    new_pos = [r["disp"] for r in new]
    pairs = []
    for oi, o in enumerate(old_pos):
        for ni, n in enumerate(new_pos):
            d = math.hypot(n[0] - o[0], n[1] - o[1])
            if d < 500:
                pairs.append((d, oi, ni))
    pairs.sort()
    uo, un = set(), set()
    exact, near = [], []
    for d, oi, ni in pairs:
        if oi in uo or ni in un:
            continue
        if d < 1:
            exact.append(new_pos[ni]); uo.add(oi); un.add(ni)
        elif d < 200:
            near.append(new_pos[ni]); uo.add(oi); un.add(ni)
    old_only = [old_pos[i] for i in range(len(old_pos)) if i not in uo]
    new_only = [new_pos[i] for i in range(len(new_pos)) if i not in un]
    assert len(exact) + len(near) + len(old_only) == len(old_pos), "旧侧数量不闭环!"
    assert len(exact) + len(near) + len(new_only) == len(new_pos), "新侧数量不闭环!"
    return exact, near, old_only, new_only


def flag_moved_pairs(old_only, new_only, radius_hint=2000):
    """疑似移位对: old_only 与 new_only 相距 < radius_hint,提示设计确认"""
    out = []
    for (ox, oy) in old_only:
        for (nx, ny) in new_only:
            d = math.hypot(nx - ox, ny - oy)
            if d < radius_hint:
                out.append((d, (ox, oy), (nx, ny)))
    out.sort()
    return out


def main():
    ap = argparse.ArgumentParser(description="DWG 双版本图块对比(炸开真值法)")
    ap.add_argument("--old", required=True, help="旧版 DWG 路径")
    ap.add_argument("--new", required=True, help="新版 DWG 路径(对比基准)")
    ap.add_argument("--map", required=True,
                    help="块名=标签 映射,分号分隔,如 \"枪式摄像机=gun;A$C5BEE1828=dome\"")
    ap.add_argument("--out", required=True, help="输出 JSON 路径")
    ap.add_argument("--work", default=r"C:\Users\hwdem\dwg_work", help="临时目录")
    ap.add_argument("--acad-exe", default=r"D:\Program Files\AutoCAD 2023\AutoCAD 2023\acad.exe",
                    help="acad.exe 路径(COM 不可用时直接启动)")
    args = ap.parse_args()

    blocks = {}
    for item in args.map.split(";"):
        name, tag = item.split("=", 1)
        blocks[name.strip()] = tag.strip()

    os.makedirs(args.work, exist_ok=True)
    acad = attach_acad(args.acad_exe)
    close_scratch_docs(acad)

    result = {}
    for block, tag in blocks.items():
        old = explode_positions(acad, args.old, args.work, block, "old")
        new = explode_positions(acad, args.new, args.work, block, "new")
        exact, near, old_only, new_only = match(old, new)
        print(f"\n=== [{tag}] {block}: 旧{len(old)} 新{len(new)} ===")
        print(f"完全一致 {len(exact)} | 微调 {len(near)} | 移位/删除 {len(old_only)} | 新增 {len(new_only)}")
        for name, pts in (("微调", near), ("移位/删除", old_only), ("新增", new_only)):
            for p in pts:
                print(f"  {name}: ({p[0]:.0f},{p[1]:.0f})")
        sizes = sorted(math.hypot(r["w"], r["h"]) for r in old + new)
        if sizes:
            print(f"  符号对角线: min={sizes[0]:.0f} 中位={sizes[len(sizes)//2]:.0f} max={sizes[-1]:.0f}"
                  f"  (建议圈半径≈中位/2+250)")
        for d, o, n in flag_moved_pairs(old_only, new_only):
            print(f"  ⚠ 疑似移位对: 旧({o[0]:.0f},{o[1]:.0f}) -> 新({n[0]:.0f},{n[1]:.0f}) 相距 {d:.0f}mm")
        result[tag] = {
            "block": block,
            "old": old, "new": new,
            "exact": [list(p) for p in exact],
            "near": [list(p) for p in near],
            "old_only": [list(p) for p in old_only],
            "new_only": [list(p) for p in new_only],
        }

    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False)
    print(f"\n已保存 {args.out}")


if __name__ == "__main__":
    main()
