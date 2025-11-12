# -*- coding: utf-8 -*-
# @script id: run_test_workflow
# @script name: Run Basic Open/Draw Tests
# @script description: 严格按步骤执行测试：每步前后都归位到单文件不确定状态；打开前检查进程数；避免重复打开；输出日志到 artifacts/logs。
# @script usage: python scripts/run_test_workflow.py

from __future__ import annotations

from pathlib import Path
import sys as _sys
import time

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in _sys.path:
    _sys.path.insert(0, str(ROOT))

from cad_automation import CadController
from utils.runlog import RunLogger


def new_and_save(ctl: CadController, path: Path) -> None:
    """Idempotent create/open helper.

    - If the DWG already exists, just open it (do not create a new one again).
    - If not exists, create a new blank document and SaveAs to that path.
    """
    if path.exists():
        ctl.open_dwg(str(path))
        ctl.wait_quiescent(30)
        return
    ops = ctl.ops
    ops.get_acad_doc()
    ops.acad.Documents.Add()
    ctl.wait_quiescent(30)
    ctl.save_as(str(path))
    ctl.wait_quiescent(30)


def draw_triangle(ctl: CadController) -> None:
    ctl.draw_line((0, 0, 0), (1000, 0, 0))
    ctl.draw_line((1000, 0, 0), (500, 800, 0))
    ctl.draw_line((500, 800, 0), (0, 0, 0))
    ctl.save()


def draw_rectangle(ctl: CadController, w: float = 1000.0, h: float = 600.0) -> None:
    ctl.draw_line((0, 0, 0), (w, 0, 0))
    ctl.draw_line((w, 0, 0), (w, h, 0))
    ctl.draw_line((w, h, 0), (0, h, 0))
    ctl.draw_line((0, h, 0), (0, 0, 0))
    ctl.save()


def insert_block_no_explode(ctl: CadController, src: Path, target: Path, ins=(0.0, 0.0, 0.0)) -> None:
    ctl.open_dwg(str(target))
    ctl.wait_quiescent(30)
    ops = ctl.ops
    ops.get_acad_doc()
    cmd = f"-INSERT\n\"{str(src)}\"\n{ins[0]},{ins[1]},{ins[2]}\n1\n1\n0\n"
    ops.doc.SendCommand(cmd)
    ctl.wait_quiescent(60)
    ctl.save()


def converge_single_unsaved(ctl: CadController, log: RunLogger) -> None:
    ctl.close_all_cad_and_wait(timeout=90.0)
    name = ctl.single_unsaved_state()
    log.log(f"single-unsaved: {name or '<none>'}")


def main() -> int:
    log = RunLogger('run_test_workflow')
    ctl = CadController()

    # 统一先归位
    log.mark('start')
    converge_single_unsaved(ctl, log)

    dwgs = ROOT / 'artifacts' / 'dwgs'
    path_123 = (dwgs / '123.dwg').resolve()
    path_ab = (dwgs / 'ab.dwg').resolve()

    # 测试1：新建123并画三角形
    log.mark('test1')
    new_and_save(ctl, path_123)
    ctl.open_dwg(str(path_123))
    draw_triangle(ctl)
    converge_single_unsaved(ctl, log)

    # 测试2：新建ab并画矩形
    log.mark('test2')
    new_and_save(ctl, path_ab)
    ctl.open_dwg(str(path_ab))
    draw_rectangle(ctl)
    converge_single_unsaved(ctl, log)

    # 测试3：把 ab 作为块整体插入到 123（原点）
    log.mark('test3')
    insert_block_no_explode(ctl, path_ab, path_123, ins=(0.0, 0.0, 0.0))
    converge_single_unsaved(ctl, log)

    # 测试4：按“区域左下角对齐”插入（不裁剪），仍然到原点
    log.mark('test4')
    # 区域左下角对齐到(0,0)
    x1, y1, x2, y2 = 0.0, 0.0, 500.0, 500.0
    blx, bly = (min(x1, x2), min(y1, y2))
    ins = (0.0 - blx, 0.0 - bly, 0.0)
    insert_block_no_explode(ctl, path_ab, path_123, ins=ins)
    converge_single_unsaved(ctl, log)

    # 结果输出
    log.mark('finished')
    names = ctl.list_open_documents()
    log.log('OPEN: ' + (', '.join(names) if names else '<none>'))
    log.write_summary({
        'dwg_123': str(path_123),
        'dwg_ab': str(path_ab),
        'open_list': names,
        'checkpoints': log.checkpoints,
        'errors': log.errors,
        'verdict': 'PASS' if not log.errors else 'FAIL',
    })
    print('DONE. See artifacts/logs/run_test_workflow/*.json for summary.')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
