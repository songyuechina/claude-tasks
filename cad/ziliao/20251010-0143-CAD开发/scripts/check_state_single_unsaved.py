# -*- coding: utf-8 -*-
# @script id: check_single_unsaved
# @script name: Check CAD State & Single-Unsaved
# @script description: 检查当前 CAD 界面文档状态，并测试 single_unsaved_state() 以收敛为“单文件未保存”状态。
# @script usage: python scripts/check_state_single_unsaved.py

from __future__ import annotations

from pathlib import Path
import sys as _sys

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in _sys.path:
    _sys.path.insert(0, str(ROOT))

from cad_automation import CadController
from utils.runlog import RunLogger


def main() -> int:
    log = RunLogger('check_single_unsaved')
    log.mark('start')

    ctl = CadController()
    ok = ctl.start_tarch_v9()
    log.log(f'start_tarch_v9: {ok}')

    before = ctl.list_open_documents()
    log.log('OPEN BEFORE: ' + (', '.join(before) if before else '<none>'))

    name = ctl.single_unsaved_state()
    log.log(f'single_unsaved_state -> {name or "<empty>"}')

    after = ctl.list_open_documents()
    log.log('OPEN AFTER: ' + (', '.join(after) if after else '<none>'))

    verdict = bool(after) and len(after) == 1 and (not name or (name in after or name == after[0]))
    log.mark('finished')
    log.write_summary({
        'open_before': before,
        'open_after': after,
        'returned_name': name,
        'verdict': 'PASS' if verdict else 'FAIL',
        'errors': log.errors,
        'checkpoints': log.checkpoints,
    })
    print('VERDICT:', 'PASS' if verdict else 'FAIL')
    return 0 if verdict else 1


if __name__ == '__main__':
    raise SystemExit(main())

