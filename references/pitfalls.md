# AutoCAD COM + 天正环境坑清单(本机实测 2026-08)

按遇到频率排序。遇到报错先对照本表,再写新代码。

## 1. 图块插入点不可信 → 必须炸开取真值
- 本项目图块(枪式摄像机等)块定义几何距基点 ~161000mm,插入点处是空白。
- `INSERT` 带旋转/镜像(XScale 为负)时,`GetBoundingBox` 对块参照返回未变换几何,同样不可信。
- **正解**: 复制 DWG 到临时副本 → 逐个 `Explode()` → 炸出图元 `GetBoundingBox` 的总包围盒中心 = 符号真实显示位置。单图元 extents > 1e6mm 视为杂散图元丢弃。顺带记录 w/h 供圈半径选型(建议半径 ≈ 对角线中位数/2 + 250)。

## 2. ModelSpace.Item 遍历报 CLASSNOTREG → 一律用 SelectionSets 过滤
- `ModelSpace.Item(i)` 在本环境不可用。
- **正解**: `SelectionSets.Add` + `Select(5, None, None, filter_codes, filter_values)`,DXF 组码: 8=图层, 2=块名, 0=实体类型。
- Window/Crossing 选择(Select 0/1)只命中当前视口已绘制实体,跨区域探测前必须 ZoomExtents——按块名/图层全局过滤则无此问题。

## 3. gen_py 早绑定缓存污染 → 删缓存目录
- 一旦有任何脚本用过 `gencache.EnsureDispatch`,`%TEMP%\gen_py\3.11\AA9A2205-*`(AutoCAD typelib)缓存会让后续 `GetActiveObject/Dispatch/Layers.Add` 返回缺属性的早绑定代理: `Layer.Color`、`InsertionPoint`、`Center` 全部 AttributeError。
- `win32com.client.Dispatch(e._oleobj_)` 绕不过(内部仍查 gencache)。
- **根治**: 关 AutoCAD → 删除整个 `%TEMP%\gen_py` → 重开。之后全程动态代理正常。切勿在脚本里 import gencache。

## 4. 天正($TCHSYS$)对象 → 瞬时 RPC 拒绝,双重异常形态
- 图纸含天正对象时 COM 调用频发 `RPC_E_CALL_REJECTED`。
- 瞬时代理故障既抛 `com_error` **也抛 `AttributeError`**(如 `<unknown>.Open`、`Add.Count`)——`com_retry` 必须两者都 catch。
- 参数化调用失败时改用 `lambda x=obj: x.Prop` 形式重试。

## 5. AutoCAD 崩溃与恢复
- 连续高压 COM 调用可能把 AutoCAD 搞崩(RPC 服务器不可用)。操作完及时 `doc.Close(False)`,不长期挂文档。
- 崩后 COM 服务器常无法自动重启(`Dispatch` 报"服务器运行失败")。
- **正解**: 直接启动 acad.exe(本机 `D:\Program Files\AutoCAD 2023\AutoCAD 2023\acad.exe`,注意 C 盘 Autodesk 目录下只有组件),起来后 `GetActiveObject` 即可连上。

## 6. 残留状态 → 重跑前清理
- 崩溃脚本残留的命名选择集: 下次 `SelectionSets.Add` 同名报"命名选择集已存在" → Add 前 `Item(name).Delete` 重试。
- 崩溃脚本残留的 scratch 临时文档: 仍开着会锁住文件(PermissionError) → 遍历 `Documents` 按文件名 Close。

## 7. 匹配闭环(必须断言)
- 贪心一对一匹配若配对阈值(500mm)大于分类阈值(200mm),200~500mm 的对会被消耗却不落进任何类别(实测踩过: 169+6+5≠181)。
- **正解**: `d >= 200mm` 的对 `continue` 不消耗配额,两侧分别归入"移位/删除"与"新增";匹配完断言 `exact+near+old_only == len(old)` 与 `exact+near+new_only == len(new)`。
- old_only 与 new_only 中相距 < 2000mm 的对是"疑似移位对",报告里要单独列出供设计确认。

## 8. EXTMAX 不可信 → 图例按内容数据范围定位
- 图纸常有远处杂散图元,EXTMAX 可达 8.8e10mm,按 EXTMAX 放图例会放到"天边"。
- **正解**: 图例放在全部标记点 min/max 之上(`max_y + radius + 3*radius`)。

## 9. 其他
- AutoCAD 2023 实际装在 `D:\Program Files\AutoCAD 2023\`,C 盘只有组件。
- `WBlock` 按扩展名 .dxf 写出的仍是 DWG;要真 DXF 需 Open+SaveAs 格式码 61(ac2013_DXF)。
- ezdxf 求包围盒用 `ezdxf.bbox.extents(entities)`,不是实体方法;ezdxf 处理不了天正对象较多的图时退回 COM 法。
- 截被遮挡窗口: `ctypes PrintWindow(hwnd, hdc, 2)`;BMP 需 PIL 转 PNG 才能被图像分析工具读。
- 打开文档后 `time.sleep(3)` 再操作;`Open` 本身也要 15 次重试。

## 10. 两版图纸基点/坐标系不一致 → --old-offset 修正
- 直接按绝对坐标对比会把全部设备误判成"移位+新增"(旧40新43 曾报 0/0/40/43)。
- **识别信号**: 疑似移位对里大量距离完全相同、方向一致(如 37 对全部 208mm)。
- **正解**: 对每个 old 点取最近 new 点的向量,取众数即共识偏移(残差应≈0);`dwg_compare.py --old-offset=dx,dy` 平移旧坐标后再匹配。**负值必须用 `=` 连接**(如 `--old-offset=-207.5866,9.8875`),否则 argparse 当选项。truth JSON 会记录 `old_offset`,gen_report.py 自动写基点修正注记。

## 11. 源图含"标记-*"图层残留图元 → 污染核验计数
- 用户手工编辑过的源 DWG 可能已在标记图层上画过圈/线,复制进标记图后 verify 报"圈 N≠期望 M、叉线 K≠期望 2M"。
- **正解**: dwg_mark.py 画标记前按图层 Select+Delete 清残留(已内置,会打印"清理图层残留")。排查: 探针脚本按图层列圈心与 truth 求差;注意 Python 连续两次 round 有银行家舍入假阳性,比对用一次舍入+容差。

## 12. com_retry(obj.Method, args) 属性获取在 try 之外
- `com_retry(ss.Select, ...)` 先求值 `ss.Select`,瞬时代理故障的 AttributeError 发生在进入 com_retry 前,照样崩。
- **正解**: `com_retry(lambda: ss.Select(...))`,属性获取+调用放进同一个 lambda 整体重试。
