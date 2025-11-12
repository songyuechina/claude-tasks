# 示例测试场景：123/ab 的打开与绘制（非规则）

本文件仅作为“一个具体测试场景”的参考说明，用于验证控制器与打开策略。它不是规范、也不是规则来源。正式规则请以 `docs/OPEN_POLICY.md` 为准。

## 提示
- 开始前建议将 CAD 归位到“单文件不确定状态”（single_unsaved_state）；
- 打开策略与顺序规则请参考 `docs/OPEN_POLICY.md`；本文件仅为示例步骤，便于复现与验证。

> 以上规则已在控制器内部实现，具体约束见 `docs/OPEN_POLICY.md`。

---

## 测试 1：新建 123.dwg 并画三角形
1) 通过控制器或 `cad_cli.py new` 新建文件 `123.dwg`（也可 `new` 后 `saveas` 为该路径）。
2) 调用 CAD 基本操作.py 中的绘图函数（或控制器包装）绘制一个三角形（示例：依次三条直线）：
   - 可等价为三次 `draw_line(p1, p2)`，或使用 `AddLightWeightPolyline` 封闭绘制。
3) 成功后回到“单文件不确定状态”，再进行下一步。

建议命令示例（仅供参考）：
```
python cad_cli.py new D:/codex-tasks/20251010-0143-CAD开发/artifacts/dwgs/123.dwg
python cad_cli.py line 0 0 0 1000 0 0
python cad_cli.py line 1000 0 0 500 800 0
python cad_cli.py line 500 800 0 0 0 0
python cad_cli.py save
```

---

## 测试 2：新建 ab.dwg 并画矩形与样条
1) 新建文件 `ab.dwg`。
2) 调用 CAD 基本操作.py 的相关函数绘制：
   - 一个矩形（可用 4 条直线或 LWPOLYLINE 封闭）；
   - 一根样条（SPLINE，若外部脚本提供对应方法也可直接调用）。
3) 成功后回到“单文件不确定状态”，再进行下一步。

---

## 测试 3：把 ab.dwg 全量复制到 123.dwg
1) 以 `123.dwg` 为激活文件。
2) 使用 `cad_cli.py` 的插入命令把 `ab.dwg` 完整复制到 `123.dwg`：
   - 推荐：`python cad_cli.py insert-all <ab.dwg> 0 0 0 --no-explode`（若外部实现支持“标准块插入”可以不炸开）。
3) 验证 `123.dwg` 中能看到 `ab.dwg` 的全部内容。
4) 成功后回到“单文件不确定状态”，再进行下一步。

---

## 测试 4：把 ab.dwg 条件性复制到 123.dwg 原点
1) 以 `123.dwg` 为激活文件。
2) 使用 `cad_cli.py insert-area` 按矩形区域“摘取” `ab.dwg` 的部分内容，并复制到原点（0,0,0）：
   - 示例：`python cad_cli.py insert-area <ab.dwg> x1 y1 x2 y2 0 0 [0] [--no-explode]`
   - 其中 `(x1,y1)-(x2,y2)` 为选择区域；目标位置使用原点（0,0,0）。
3) 验证 `123.dwg` 中仅出现期望的局部内容，位置正确。
4) 成功后回到“单文件不确定状态”。

---

## 通过/失败的判定
- 打开与插入操作遵循 `docs/OPEN_POLICY.md`：
  - 无“只读副本”与重复打开；多个文件按顺序逐一完成，且每次完成后 CAD 空闲。
- 每步完成后能明确看到对应几何；
- 过程被干扰（弹窗、无响应等）可关闭干扰窗口后重试该步；
- 若出现 CAD 进程数异常（>1），先回到“单文件不确定状态”再继续。
- 打开前检查：如果进程数>1，先 `single_unsaved_state()` 并等待 3 秒再打开；若目标路径或同名文档已打开，则不再重复打开。

---

## 相关参考
- 统一打开策略与协同：`docs/OPEN_POLICY.md`
- 状态收敛与 CLI 总览：`docs/STATE_CONTROL.md`
- 可能用到的 CLI 命令：`cad_cli.py new / save / insert-all / insert-area / line / standard`

> 本文档仅记录测试步骤与要求，不强制实现细节。绘图函数请依据“CAD 基本操作.py”中可用方法或控制器包装选择具体实现。
