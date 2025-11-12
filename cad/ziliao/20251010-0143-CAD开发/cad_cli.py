# -*- coding: utf-8 -*-
# @script id: cad_cli
# @script name: CAD CLI Entrypoint
# @script description: 启动/关闭CAD、DWG管理、基本操作（画线）、标准化状态
# @script commands: start, stop, open, list, close, save, saveas, new, insert-all, insert-area, line, start-open, standard, standard2
# @script usage: python cad_cli.py <command> [args]
from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

from cad_automation import CadController
# cad_operation removed per design


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="天正CAD自动化：启动/关闭、打开/关闭DWG、基本操作、保存"
    )
    p.add_argument("command", choices=[
        "start", "stop", "open", "list", "close", "save", "saveas", "new", "insert-all", "insert-area", "line", "start-open", "standard", "standard2"
    ], help="要执行的命令")

    p.add_argument("args", nargs="*", help="命令参数")
    p.add_argument("--tarch", dest="tarch", default=None, help="天正V9根目录，例如 C:/Tangent/TArchT20V9")
    p.add_argument("--ops", dest="ops", default=None, help="CAD基本操作.py 的路径 (可选)")
    return p


def _cmd_start(ctl: CadController, tarch: str | None = None) -> bool:
    return bool(ctl.start_tarch_v9(root=tarch))


def _cmd_stop(ctl: CadController) -> None:
    ctl.close_all_cad()


def _cmd_open(ctl: CadController, paths: list[str]) -> int:
    ctl.start_tarch_v9()
    return int(ctl.open_dwgs(paths))


def _cmd_list(ctl: CadController) -> list[str]:
    return ctl.list_open_documents()


def _cmd_close(ctl: CadController, name: str | None = None) -> None:
    if name:
        ctl.close_by_name(name)
    else:
        ctl.close_current()


def _cmd_save(ctl: CadController) -> None:
    ctl.save()


def _cmd_saveas(ctl: CadController, out_path: str) -> None:
    ctl.save_as(out_path)


def _cmd_line(ctl: CadController, x1: float, y1: float, z1: float, x2: float, y2: float, z2: float) -> None:
    ctl.start_tarch_v9()
    ctl.draw_line((x1, y1, z1), (x2, y2, z2))
    ctl.save()

def _cli_new(ctl: CadController, out_path: str | None = None) -> None:
    """Create a new DWG, or open if target path already exists.

    - 当指定了 out_path 且该路径已存在：不再新建，直接打开该文件；
    - 否则：新建一个空白文档并（可选）另存为 out_path。
    """
    ctl.start_tarch_v9()

    # If a target path is provided and already exists, just open it (no new doc)
    if out_path:
        try:
            p = str(Path(out_path).resolve())
        except Exception:
            p = out_path
        if os.path.exists(p):
            ctl.open_dwg(p)
            return

    # Otherwise, create a fresh unsaved DWG then optionally SaveAs
    ops = ctl.ops
    ops.get_acad_doc()
    ops.acad.Documents.Add()
    ctl.wait_quiescent(30)
    if out_path:
        ctl.save_as(out_path)


def _cli_insert_all(ctl: CadController, src: str, pos: tuple[float, float, float], no_explode: bool) -> None:
    ctl.start_tarch_v9()
    ops = ctl.ops
    if no_explode and hasattr(ops, "insert_standard_block"):
        ops.insert_standard_block(src, insertion_point=pos)
    elif hasattr(ops, "insert_and_explode_dwg"):
        ops.insert_and_explode_dwg(src, insertion_point=pos)
    else:
        ops.get_acad_doc()
        cmd = f"-INSERT\n\"{src}\"\n{pos[0]},{pos[1]},{pos[2]}\n1\n1\n0\n"
        ops.doc.SendCommand(cmd)


def _cli_insert_area(ctl: CadController, src: str, x1: float, y1: float, x2: float, y2: float, tx: float, ty: float, tz: float, no_explode: bool) -> None:
    blx, bly = (min(x1, x2), min(y1, y2))
    ins = (tx - blx, ty - bly, tz)
    ctl.start_tarch_v9()
    ops = ctl.ops
    if no_explode and hasattr(ops, "insert_standard_block"):
        ops.insert_standard_block(src, insertion_point=ins)
    elif hasattr(ops, "insert_and_explode_dwg"):
        ops.insert_and_explode_dwg(src, insertion_point=ins)
    else:
        ops.get_acad_doc()
        cmd = f"-INSERT\n\"{src}\"\n{ins[0]},{ins[1]},{ins[2]}\n1\n1\n0\n"
        ops.doc.SendCommand(cmd)


def _cmd_start_open(ctl: CadController, dwgs: list[str]) -> int:
    return int(ctl.start_and_open(dwgs))


def _cmd_standard(ctl: CadController, target: str | None = None):
    return ctl.standardize_state(target_doc_path=target)


def _cmd_standard2(ctl: CadController, a: str, b: str, active: str | None = None):
    return ctl.standardize_two_documents(a, b, active=active)


def _cmd_single_unsaved(ctl: CadController) -> str:
    return ctl.single_unsaved_state()


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    ctl = CadController(cad_ops_path=args.ops, tarch_root=args.tarch)

    # No global retry wrapper; invoke controller methods directly.

    cmd = args.command
    if cmd == "start":
        ok = ctl.start_tarch_v9()
        print("启动成功" if ok else "启动失败")
        return 0

    if cmd == "stop":
        ctl.close_all_cad()
        print("已请求关闭所有CAD进程")
        return 0

    if cmd == "open":
        if not args.args:
            print("用法: cad_cli.py open <dwg1> [dwg2 ...]")
            return 2
        ctl.start_tarch_v9()
        ok_count = ctl.open_dwgs(args.args)
        print(f"打开成功: {ok_count} / {len(args.args)}")
        return 0

    if cmd == "list":
        names = ctl.list_open_documents()
        for n in names:
            print(n)
        return 0

    if cmd == "close":
        # 可选传入文件名（含 .dwg），否则关闭当前
        if args.args:
            ctl.close_by_name(args.args[0])
        else:
            ctl.close_current()
        print("已请求关闭")
        return 0

    if cmd == "save":
        ctl.save()
        print("已保存当前文档")
        return 0

    if cmd == "saveas":
        if not args.args:
            print("用法: cad_cli.py saveas <输出路径.dwg>")
            return 2
        ctl.save_as(args.args[0])
        print(f"已另存为: {args.args[0]}")
        return 0

    if cmd == "new":
        out = args.args[0] if args.args else None
        _cli_new(ctl, out)
        print(f"已新建并另存为: {out}" if out else "已新建空白文档（未保存）")
        return 0

    if cmd == "insert-all":
        # 用法: cad_cli.py insert-all <dwg> [x y z] [--no-explode]
        if not args.args:
            print("用法: cad_cli.py insert-all <dwg> [x y z] [--no-explode]")
            return 2
        src = args.args[0]
        pos = (0.0, 0.0, 0.0)
        rest = args.args[1:]
        no_explode = False
        if len(rest) >= 3:
            try:
                pos = (float(rest[0]), float(rest[1]), float(rest[2]))
                rest = rest[3:]
            except Exception:
                pass
        if rest and rest[0] == "--no-explode":
            no_explode = True
        _cli_insert_all(ctl, src, pos, no_explode)
        print("已插入文件: {} @ {} ({})".format(src, pos, "no-explode" if no_explode else "explode"))
        return 0

    if cmd == "insert-area":
        # 用法: cad_cli.py insert-area <dwg> x1 y1 x2 y2 tx ty [tz] [--no-explode]
        if len(args.args) < 8:
            print("用法: cad_cli.py insert-area <dwg> x1 y1 x2 y2 tx ty [tz] [--no-explode]")
            return 2
        src = args.args[0]
        x1, y1, x2, y2 = map(float, args.args[1:5])
        tx, ty = map(float, args.args[5:7])
        tz = 0.0
        rest = args.args[7:]
        if rest and (rest[0] not in ("--no-explode",)):
            try:
                tz = float(rest[0])
                rest = rest[1:]
            except Exception:
                pass
        no_explode = (bool(rest) and rest[0] == "--no-explode")
        _cli_insert_area(ctl, src, x1, y1, x2, y2, tx, ty, tz, no_explode)
        print("已插入区域左下对齐: src={}, 区域=({},{})-({},{}) 目标=({},{},{})".format(src, x1, y1, x2, y2, tx, ty, tz))
        return 0

    if cmd == "line":
        # 坐标: x1 y1 z1 x2 y2 z2
        if len(args.args) != 6:
            print("用法: cad_cli.py line x1 y1 z1 x2 y2 z2")
            return 2
        x1, y1, z1, x2, y2, z2 = map(float, args.args)
        ctl.start_tarch_v9()
        ctl.draw_line((x1, y1, z1), (x2, y2, z2))
        ctl.save()
        print("已画线并保存")
        return 0

    if cmd == "start-open":
        if not args.args:
            print("用法: cad_cli.py start-open <dwg1> [dwg2 ...]")
            return 2
        ok = ctl.start_and_open(args.args)
        print(f"启动并打开成功数: {ok}")
        return 0

    if cmd == "standard":
        # 可选传入一个 dwg 路径，确保该文件为唯一打开文件
        target = args.args[0] if args.args else None
        name = ctl.standardize_state(target_doc_path=target)
        if name:
            print(f"标准单文件: {name}")
        else:
            print("标准状态设置失败")
        return 0

    if cmd == "standard2":
        # 用法: cad_cli.py standard2 <A.dwg> <B.dwg> [a|b]
        # 若未提供 A/B，则使用任务目录固定引用 artifacts/dwgs/00.dwg 与 01.dwg
        if len(args.args) < 2:
            root = Path(__file__).resolve().parent
            a = str((root / "artifacts" / "dwgs" / "00.dwg").resolve())
            b = str((root / "artifacts" / "dwgs" / "01.dwg").resolve())
            active = args.args[0] if args.args else None
        else:
            a = args.args[0]; b = args.args[1]
            active = args.args[2] if len(args.args) >= 3 else None
        names = ctl.standardize_two_documents(a, b, active=active)
        if names:
            print(f"标准双文件: {names[0]}, {names[1]}")
        else:
            print("双文件标准状态设置失败")
        return 0
    return 1


if __name__ == "__main__":
    raise SystemExit(main())





