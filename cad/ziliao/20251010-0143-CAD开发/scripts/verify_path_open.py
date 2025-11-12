# -*- coding: utf-8 -*-
# @script id: verify_path_open
# @script name: Verify Path And Open DWG
# @script description: 校验本地DWG路径（含中文/特殊字符）是否存在，并尝试用CadController打开
# @script usage: python scripts/verify_path_open.py <absolute_dwg_path>

from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Tuple

# Ensure project root on sys.path
ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from cad_automation import CadController
from utils.runlog import RunLogger


def get_short_path(p: str) -> str:
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


def exists_via_win32(p: str) -> bool:
    try:
        import ctypes
        INVALID_FILE_ATTRIBUTES = 0xFFFFFFFF
        GetFileAttributesW = ctypes.windll.kernel32.GetFileAttributesW
        GetFileAttributesW.argtypes = [ctypes.c_wchar_p]
        GetFileAttributesW.restype = ctypes.c_uint32
        attrs = GetFileAttributesW(p)
        return attrs != INVALID_FILE_ATTRIBUTES
    except Exception:
        return False


def main(argv: list[str] | None = None) -> int:
    argv = argv or sys.argv[1:]
    if not argv:
        print("Usage: python scripts/verify_path_open.py <absolute_dwg_path>")
        return 2
    in_path = argv[0]
    p_abs = os.path.abspath(in_path)
    p_short = get_short_path(p_abs)

    log = RunLogger('verify_path_open')
    log.mark('start')
    log.log(f"INPUT: {in_path}")
    log.log(f"ABS:   {p_abs}")
    log.log(f"SHORT: {p_short}")

    e1 = os.path.exists(p_abs)
    e2 = exists_via_win32(p_abs)
    verdict_exist = e1 or e2 or os.path.exists(p_short)
    log.log(f"exists(os): {e1}, exists(win32): {e2}, exists(short): {os.path.exists(p_short)}")
    if not verdict_exist:
        log.error("Path not found by available checks")
        log.write_summary({
            'input': in_path,
            'abs': p_abs,
            'short': p_short,
            'exists': False,
            'errors': log.errors,
            'verdict': 'FAIL',
        })
        return 3

    # Try opening via CadController
    ctl = CadController()
    ok = ctl.start_tarch_v9() and ctl.open_dwg(p_abs)
    log.log(f"open_dwg: {'OK' if ok else 'NG'}")
    names = ctl.list_open_documents()
    log.log("OPEN: " + ", ".join(names))
    log.mark('finished')

    log.write_summary({
        'input': in_path,
        'abs': p_abs,
        'short': p_short,
        'exists': True,
        'open_ok': bool(ok),
        'open_list': names,
        'errors': log.errors,
        'checkpoints': log.checkpoints,
        'verdict': 'PASS' if ok and not log.errors else 'FAIL',
    })
    return 0 if ok else 4


if __name__ == '__main__':
    raise SystemExit(main())

