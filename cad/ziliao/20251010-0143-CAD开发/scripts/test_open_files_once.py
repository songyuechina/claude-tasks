# -*- coding: utf-8 -*-
# @script id: test_open_files_once
# @script name: Test Open Files (No Retry)
# @script description: 将 CAD 归为“单文件不确定状态”，然后逐个只打开一次指定 DWG，验证不会出现多只读副本。
# @script usage: python scripts/test_open_files_once.py

from __future__ import annotations

from pathlib import Path
import sys as _sys
import time

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in _sys.path:
    _sys.path.insert(0, str(ROOT))

from cad_automation import CadController


def _pick_targets() -> list[Path]:
    dwg_dir = ROOT / 'artifacts' / 'dwgs'
    cs = list(dwg_dir.glob('*.dwg'))
    # 期望三类：包含 'SS' 的，前缀 '4#' 的，前缀 '3#' 的
    one = next((p for p in cs if 'SS' in p.stem or '签' in p.stem), None)
    two = next((p for p in cs if p.stem.startswith('4#')), None)
    three = next((p for p in cs if p.stem.startswith('3#')), None)
    picks = [p for p in (one, two, three) if p is not None]
    # 如果缺失，就退而取任意 3 个
    if len(picks) < 3:
        for p in cs:
            if p not in picks:
                picks.append(p)
            if len(picks) >= 3:
                break
    return picks


def main() -> int:
    ctl = CadController()

    # 0) 若进程数不等于 1（含 >1 或 =0），统一收敛为单文件不确定状态
    cnt = ctl.cad_process_count(print_it=True)
    if cnt != 1:
        print('[info] process count != 1, converge to single-unsaved')
        name = ctl.single_unsaved_state()
        print('STATE:', 'OK' if name else 'FAIL', name)

    targets = _pick_targets()
    for idx, p in enumerate(targets, 1):
        print(f'\n=== [{idx}/{len(targets)}] OPEN-ONCE: {p} ===')

        # 每次打开前：确保通过 start_applicationV9 启动
        if ctl.cad_process_count(print_it=True) > 0:
            ctl.close_all_cad_and_wait(timeout=90.0)
        try:
            PTH = ctl.tarch_root
            try:
                ctl.ops.start_applicationV9(PTH=PTH, max_retries=3, retry_delay=2.0)
            except TypeError:
                ctl.ops.start_applicationV9(PTH=PTH, max_retries=3)
        except Exception as e:
            print('[ERR] start_applicationV9 failed:', repr(e))
            return 2
        t0 = time.time()
        while time.time() - t0 < 60:
            try:
                ctl.ops.get_acad_doc()
                break
            except Exception:
                time.sleep(0.5)

        # 打开一次（open_dwg 自身不做外层重试装饰）
        ok = ctl.open_dwg(str(p))
        print('open_dwg:', 'OK' if ok else 'NG')
        names = ctl.list_open_documents()
        print('OPEN LIST:', ', '.join(names) if names else '<none>')

        # 期望：此时只有 1 个文档（刚打开的目标）
        if len(names) != 1:
            print(f'[WARN] expected 1 open doc, got {len(names)}')

        # 关闭当前为下一个测试做准备
        ctl.close_current()
        time.sleep(0.3)

    print('\nDONE.')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
