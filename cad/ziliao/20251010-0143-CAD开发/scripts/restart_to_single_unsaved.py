# -*- coding: utf-8 -*-
# @script id: restart_to_single_unsaved
# @script name: Restart CAD (TArch V9 + 2021) To Single-Unsaved State
# @script description: 关闭全部 CAD 进程，调用 CAD 基本操作的 start_applicationV9(PTH, max_retries, retry_delay) 重启；随后收敛到“单文件未保存（不确定）”状态。
# @script usage: python scripts/restart_to_single_unsaved.py ["C:/Tangent/TArchT20V9"]

from __future__ import annotations

import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from cad_automation import CadController
from utils.runlog import RunLogger


def main() -> int:
    tarch_root = sys.argv[1] if len(sys.argv) >= 2 else r"C:\Tangent\TArchT20V9"

    log = RunLogger('restart_to_single_unsaved')
    log.mark('start')

    ctl = CadController(tarch_root=tarch_root)

    # 1) 关闭全部 CAD 进程
    log.log('close_all_cad()')
    ctl.close_all_cad()
    time.sleep(1.0)

    # 2) 通过外部模块显式重启（包含 retry_delay）
    log.log(f'start_applicationV9: PTH={tarch_root}, max_retries=3, retry_delay=2.0')
    try:
        ctl.ops.start_applicationV9(PTH=tarch_root, max_retries=3, retry_delay=2.0)  # type: ignore[arg-type]
    except TypeError:
        # 兼容旧签名（无 retry_delay）
        ctl.ops.start_applicationV9(PTH=tarch_root, max_retries=3)

    # 等待 COM 就绪
    t0 = time.time()
    while time.time() - t0 < 60:
        try:
            ctl.ops.get_acad_doc()
            break
        except Exception:
            time.sleep(0.5)

    # 3) 收敛到单文件未保存状态
    name = ctl.single_unsaved_state()
    names = ctl.list_open_documents()

    log.log(f'ACTIVE: {name or "<none>"}')
    log.log('OPEN: ' + (', '.join(names) if names else '<none>'))
    log.mark('finished')

    verdict = bool(name) and len(names) == 1
    log.write_summary({
        'tarch_root': tarch_root,
        'active': name,
        'open': names,
        'verdict': 'PASS' if verdict else 'FAIL',
        'errors': log.errors,
        'checkpoints': log.checkpoints,
    })
    print('VERDICT:', 'PASS' if verdict else 'FAIL')
    return 0 if verdict else 1


if __name__ == '__main__':
    raise SystemExit(main())

