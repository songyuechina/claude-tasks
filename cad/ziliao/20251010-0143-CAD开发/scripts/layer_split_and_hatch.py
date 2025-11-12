# -*- coding: utf-8 -*-
# @script id: layer_split_and_hatch
# @script name: Split By Layer And Hatch Groups
# @script description: 针对一个 DWG：按图层拆分为多个文件；在每个文件中，查找封闭多段线，基于“最多两层包含”分组后做实体填充（SOLID）。
# @script usage: python scripts/layer_split_and_hatch.py [<source_dwg>]

from __future__ import annotations

import os
import re
import sys
from pathlib import Path
from typing import Dict, List, Tuple
import time
import json

# --- Ensure project root on sys.path & fix output root to this task folder ---
# This lets the script work whether you run via absolute path or from another CWD.
ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from shapely.geometry import Polygon
from shapely.ops import unary_union

from cad_automation import CadController

# Logging for this script under task folder
LOG_ROOT = ROOT / "artifacts" / "logs" / "layer_split_and_hatch"
LOG_ROOT.mkdir(parents=True, exist_ok=True)

def _now_ts() -> str:
    return time.strftime("%Y-%m-%d %H:%M:%S")

class RunLogger:
    def __init__(self, name_prefix: str = "run"):
        t = time.strftime("%Y%m%d-%H%M%S")
        self.run_id = t
        self.path = LOG_ROOT / f"{name_prefix}-{t}.log"
        self.json_path = LOG_ROOT / f"{name_prefix}-{t}.json"
        self.path.write_text("", encoding="utf-8")
        self.errors: List[str] = []
    def log(self, msg: str) -> None:
        line = f"[{_now_ts()}] {msg}\n"
        try:
            with self.path.open("a", encoding="utf-8") as f:
                f.write(line)
        except Exception:
            pass
        print(msg)
    def error(self, msg: str) -> None:
        self.errors.append(msg)
        self.log(f"ERROR: {msg}")
    def write_summary(self, data: dict) -> None:
        try:
            self.json_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass


def safe_name(name: str) -> str:
    return re.sub(r"[^\w\-_.]+", "_", name)


def _is_server_busy_exc(e: Exception) -> bool:
    try:
        # pywintypes.com_error: hresult -2147417846 is RPC_E_SERVERCALL_RETRYLATER / "应用程序正在使用中"
        import pywintypes  # type: ignore
        if isinstance(e, pywintypes.com_error):
            hr = getattr(e, "hresult", None)
            if hr in (-2147417846, 0x8001010A, 2147483648 - 2147417846):
                return True
    except Exception:
        pass
    # Fallback: message contains typical text
    msg = str(e)
    return ("正在使用中" in msg) or ("server busy" in msg.lower()) or ("RPC_E_SERVERCALL_RETRYLATER" in msg)


def _safe_list_modelspace(ops, attempts: int = 20, wait: float = 0.6):
    import time
    last_err = None
    for _ in range(max(1, attempts)):
        try:
            ms = ops.doc.ModelSpace
            return list(ms)
        except Exception as e:
            last_err = e
            if _is_server_busy_exc(e):
                try:
                    # best effort to allow CAD to settle
                    time.sleep(wait)
                except Exception:
                    pass
                continue
            raise
    if last_err:
        raise last_err
    return []


def _short_path(p: str) -> str:
    try:
        import ctypes
        from ctypes import wintypes
        GetShortPathNameW = ctypes.windll.kernel32.GetShortPathNameW
        GetShortPathNameW.argtypes = [wintypes.LPCWSTR, wintypes.LPWSTR, wintypes.DWORD]
        GetShortPathNameW.restype = wintypes.DWORD
        buf = ctypes.create_unicode_buffer(260)
        ret = GetShortPathNameW(p, buf, len(buf))
        return buf.value if ret else p
    except Exception:
        return p


def safe_save_as(ctl: CadController, out_path: str, tries: int = 12) -> bool:
    import time
    last_err = None
    for _ in range(max(1, tries)):
        try:
            # Try to use existing doc handle; avoid heavy re-init
            doc = getattr(ctl.ops, "doc", None)
            if doc is None:
                try:
                    ctl.ops.get_acad_doc()
                    doc = getattr(ctl.ops, "doc", None)
                except Exception:
                    doc = None
            ctl.wait_quiescent(45)
            sp = _short_path(str(Path(out_path).resolve()))
            if doc is not None:
                doc.SaveAs(sp)
            else:
                # As last resort fallback to controller method
                ctl.save_as(sp)
            ctl.wait_quiescent(45)
            return True
        except Exception as e:
            last_err = e
            if _is_server_busy_exc(e):
                time.sleep(0.8)
                continue
            try:
                ctl.ops.get_acad_doc()
            except Exception:
                pass
            time.sleep(0.5)
    if last_err:
        print(f"[warn] save_as failed after retries: {last_err!r}")
    return False


def list_layers_with_objects(ops) -> List[str]:
    layers = set()
    try:
        for ent in _safe_list_modelspace(ops):
            try:
                layers.add(str(getattr(ent, "Layer", "0")))
            except Exception:
                continue
    except Exception:
        pass
    return [x for x in layers if x]


def delete_except_layer(ops, keep_layer: str) -> int:
    removed = 0
    # Take a snapshot first to avoid iterator invalidation and handle server-busy.
    for ent in _safe_list_modelspace(ops):
        try:
            if str(getattr(ent, "Layer", "")) != keep_layer:
                ent.Delete()
                removed += 1
        except Exception:
            continue
    return removed


def get_closed_polylines(ops) -> List:
    closed = []
    for ent in _safe_list_modelspace(ops):
        try:
            if getattr(ent, "ObjectName", "") == "AcDbPolyline" and bool(getattr(ent, "Closed", False)):
                closed.append(ent)
        except Exception:
            continue
    return closed


def polyline_to_polygon(ent) -> Polygon:
    coords = list(ent.Coordinates)
    pts = [(coords[i], coords[i + 1]) for i in range(0, len(coords), 2)]
    if pts and pts[0] != pts[-1]:
        pts.append(pts[0])
    return Polygon(pts)


def group_by_containment(polys: List[Polygon]) -> Dict[int, Dict[str, List[int]]]:
    """
    返回 {outer_idx: {"outers": [outer_idx], "inners": [idx...]}}
    对每个多边形，找面积最小且包含它的父多边形；形成最多两层（父->子）。
    """
    areas = [p.area for p in polys]
    parent = [-1] * len(polys)
    for i, p in enumerate(polys):
        best = -1
        best_area = float("inf")
        for j, q in enumerate(polys):
            if i == j:
                continue
            if q.contains(p):
                if areas[j] < best_area:
                    best_area = areas[j]
                    best = j
        parent[i] = best

    groups: Dict[int, Dict[str, List[int]]] = {}
    for i, par in enumerate(parent):
        if par == -1:
            # 外层
            groups.setdefault(i, {"outers": [i], "inners": []})
        else:
            # 子层，挂到最近父级
            groups.setdefault(par, {"outers": [par], "inners": []})
            groups[par]["inners"].append(i)
    return groups


def hatch_group(ops, ctl: CadController, layer_name: str, outer: Polygon, inners: List[Polygon]):
    """
    以“外 - 内（洞）”的关系生成填充。为避免 COM 对洞的处理差异，先做几何差集：outer - union(inners)。
    对结果每个面片，绘制 LWPOLYLINE 外边界并以 SOLID 填充。
    """
    ms = ops.doc.ModelSpace
    hatch_layer = f"HATCH_{layer_name}"

    geom = outer
    if inners:
        geom = outer.difference(unary_union(inners))

    def draw_and_hatch(poly: Polygon):
        xs, ys = poly.exterior.xy
        coords3d = [(float(x), float(y), 0.0) for x, y in zip(xs, ys)]
        # 画边界
        pl = ops.draw_lwpolyline(coords3d, layer_name=layer_name, closed=True)
        # 填充
        hatch = ms.AddHatch(0, "SOLID", False)
        hatch.AppendLoop(0, ops.vtobj(pl))
        hatch.Evaluate()
        hatch.Layer = hatch_layer

    if geom.is_empty:
        return
    if geom.geom_type == "Polygon":
        draw_and_hatch(geom)
    elif geom.geom_type == "MultiPolygon":
        for g in geom.geoms:
            draw_and_hatch(g)


def ensure_test_dwg(ctl: CadController, test_path: Path) -> str:
    ops = ctl.ops
    logger.mark("cad-ready")
    ctl.standardize_state()
    safe_save_as(ctl, str(test_path))
    ctl.wait_quiescent(30)

    # 画一些测试图形：两层包含 / 不包含 / 不同图层
    def rect(x1, y1, x2, y2):
        return [(x1, y1, 0.0), (x2, y1, 0.0), (x2, y2, 0.0), (x1, y2, 0.0), (x1, y1, 0.0)]

    def ensure_layer_quiet(name: str):
        layers = ops.doc.Layers
        try:
            lyr = layers.Item(name)
        except Exception:
            lyr = layers.Add(name)
        ops.doc.ActiveLayer = lyr

    def draw_lwpolyline_quiet(coords3d, layer_name: str, closed: bool = True):
        # 扁平 x,y
        raw = []
        for x, y, _ in coords3d:
            raw.extend((x, y))
        import pythoncom, win32com.client
        arr = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, raw)
        pl = ops.doc.ModelSpace.AddLightWeightPolyline(arr)
        pl.Layer = layer_name
        pl.Closed = bool(closed)
        return pl

    # L1
    ensure_layer_quiet("L1")
    draw_lwpolyline_quiet(rect(0, 0, 1000, 800), layer_name="L1", closed=True)           # 外
    draw_lwpolyline_quiet(rect(300, 200, 700, 600), layer_name="L1", closed=True)        # 内
    draw_lwpolyline_quiet(rect(1200, 0, 1900, 600), layer_name="L1", closed=True)        # 独立

    # L2
    ensure_layer_quiet("L2")
    draw_lwpolyline_quiet(rect(0, 1200, 1000, 2000), layer_name="L2", closed=True)       # 外
    draw_lwpolyline_quiet(rect(200, 1400, 500, 1700), layer_name="L2", closed=True)      # 内1
    draw_lwpolyline_quiet(rect(600, 1400, 900, 1700), layer_name="L2", closed=True)      # 内2

    ctl.save()
    return str(test_path)


def _is_under_root(p: Path, root: Path) -> bool:
    try:
        p.resolve().relative_to(root.resolve())
        return True
    except Exception:
        return False


def _resolve_out_root(src_path: str, *, out_dir: str | None, out_mode: str, allow_outside: bool) -> Path:
    """Choose output directory.
    - artifacts: ROOT/artifacts/dwgs/layer_split (default)
    - source: <src_dir>/<src_stem>__layer_split
    - explicit out_dir: use as-is
    If target dir is outside this task folder and allow_outside=False, fall back to artifacts.
    """
    artifacts_root = ROOT / "artifacts" / "dwgs" / "layer_split"
    if out_dir:
        candidate = Path(out_dir)
    elif out_mode == "source":
        sp = Path(src_path)
        candidate = sp.parent / f"{sp.stem}__layer_split"
    else:
        candidate = artifacts_root

    if not _is_under_root(candidate, ROOT) and not allow_outside:
        print(f"[info] Output directory outside task folder is not allowed by default.\n"
              f"       Requested: {candidate}\n"
              f"       Falling back to artifacts: {artifacts_root}\n"
              f"       To override, pass --out-mode source or --out-dir with --force-outside.")
        candidate = artifacts_root
    candidate.mkdir(parents=True, exist_ok=True)
    return candidate


def main() -> int:
    # CLI args
    import argparse
    parser = argparse.ArgumentParser(description="Split DWG by layer and hatch groups")
    parser.add_argument("source", nargs="?", help="Path to source DWG. If omitted, generate demo file")
    parser.add_argument("--out-mode", choices=["artifacts", "source"], default="source",
                        help="Output location: artifacts (default) or source (subfolder next to DWG)")
    parser.add_argument("--out-dir", default=None, help="Explicit output directory (overrides out-mode)")
    parser.add_argument("--force-outside", action="store_true",
                        help="Allow writing outside this task folder (use with --out-mode source/--out-dir)")
    args = parser.parse_args()

    logger = RunLogger("layer-split")
    logger.log("start layer_split_and_hatch")
    logger.mark("start")

    ctl = CadController()
    assert ctl.start_tarch_v9(), "无法启动/复用天正CAD"
    assert ctl.standardize_state(), "无法回退标准状态"
    ops = ctl.ops

    # 源文件：命令行指定或创建测试文件
    if args.source:
        src_path = args.source
    else:
        src_path = str((ROOT / "artifacts" / "dwgs" / "LAYER_HATCH_TEST.dwg").resolve())
        if not os.path.exists(src_path):
            src_path = ensure_test_dwg(ctl, Path(src_path))

    allow_outside = bool(args.force_outside or (args.out_mode == "source"))
    out_root = _resolve_out_root(src_path, out_dir=args.out_dir, out_mode=args.out_mode, allow_outside=allow_outside)
    logger.log(f"source: {src_path}")
    logger.log(f"output root: {out_root}")

    # 读取层与对象分布
    ctl.open_dwg(src_path)
    logger.mark("source-opened")
    ctl.wait_quiescent(45)
    layers = list_layers_with_objects(ops)
    base = Path(src_path).stem
    eff_layers = [x for x in layers if x and not x.startswith("$")]
    logger.log(f"detected layers with objects: {eff_layers}")
    logger.mark("layers-detected")

    produced = []
    file_stats: List[dict] = []
    for layer in eff_layers:
        if not layer or layer.startswith("$"):
            continue

        # 每个图层：重新打开源文件 -> 删除非目标图层 -> 另存为
        try:
            ctl.open_dwg(src_path)
            ctl.wait_quiescent(45)
            removed = delete_except_layer(ops, keep_layer=layer)
            logger.log(f"layer '{layer}': removed entities not on this layer: {removed}")
            ctl.wait_quiescent(30)
        except Exception as e:
            logger.error(f"prep for layer '{layer}' failed: {e!r}")
            continue

        dest = out_root / f"{safe_name(base)}__{safe_name(layer)}.dwg"
        ok_save = safe_save_as(ctl, str(dest))
        if not ok_save:
            # If saving failed due to transient issues, skip this layer gracefully
            logger.error(f"skip layer '{layer}' due to SaveAs failure")
            continue

        # 在该文件中：查找封闭多段线 -> 分组 -> 填充
        try:
            closed = get_closed_polylines(ops)
            polys = [polyline_to_polygon(ent) for ent in closed]
            groups = group_by_containment(polys)
            hatch_ok = 0
            hatch_err = 0
            for outer_idx, data in groups.items():
                outer = polys[outer_idx]
                inner_polys = [polys[i] for i in data.get("inners", [])]
                try:
                    hatch_group(ops, ctl, layer, outer, inner_polys)
                    hatch_ok += 1
                except Exception as ge:
                    hatch_err += 1
                    logger.error(f"hatch group failed on layer '{layer}': {ge!r}")
            file_stats.append({
                "layer": layer,
                "groups_count": len(groups),
                "hatch_ok": hatch_ok,
                "hatch_err": hatch_err,
                "dest": str(dest),
            })
        except Exception as e:
            logger.error(f"hatch for layer '{layer}' failed: {e!r}")
            # keep the saved file even if hatch failed

        ctl.save()
        produced.append(str(dest))
        ctl.close_current()

    # 回到标准状态并打印产物
    ctl.standardize_state()
    logger.mark("finished")
    required = ["start", "cad-ready", "source-opened", "layers-detected", "finished"]
    print("OUTPUT:")
    print(f" - OUTPUT ROOT: {out_root}")
    for p in produced:
        print(" -", p)
    general_ok = all(cp in logger.checkpoints for cp in required) and (len(logger.errors) == 0)
    specific_ok = (len(produced) == len(eff_layers)) and all((st.get("hatch_err", 0) == 0) for st in file_stats)
    verdict_ok = general_ok and specific_ok
    # summary and log
    summary = {
        "source": src_path,
        "output_root": str(out_root),
        "expected_layers": eff_layers,
        "expected_count": len(eff_layers),
        "produced": produced,
        "produced_count": len(produced),
        "errors": logger.errors,
        "checkpoints": logger.checkpoints,
        "required_checkpoints": required,
        "general_ok": general_ok,
        "specific_ok": specific_ok,
        "file_stats": file_stats,
        "verdict": "PASS" if verdict_ok else "FAIL",
        "run_id": logger.run_id,
        "log_file": str(logger.path),
    }
    logger.write_summary(summary)
    print(f"VERDICT: {'PASS' if verdict_ok else 'FAIL'}  (general {general_ok}, specific {specific_ok}, errors {len(logger.errors)})")
    print(f"LOG: {logger.path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
