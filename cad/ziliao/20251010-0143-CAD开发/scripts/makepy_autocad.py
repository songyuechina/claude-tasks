# -*- coding: utf-8 -*-
"""
Generate MakePy cache for AutoCAD COM (to avoid interactive makepy prompts).
Safe to run multiple times.
"""
from __future__ import annotations

import sys

try:
    import win32com.client.makepy as makepy  # type: ignore
except Exception as e:  # pragma: no cover
    print('import makepy failed:', repr(e))
    sys.exit(1)

specs = [
    'AutoCAD 2021 Type Library',
    'AutoCAD.Application',
]

ok = False
for s in specs:
    try:
        makepy.GenerateFromTypeLibSpec(s)
        print('makepy ok:', s)
        ok = True
    except Exception as e:  # pragma: no cover
        print('makepy fail for', s, ':', repr(e))

sys.exit(0 if ok else 1)

