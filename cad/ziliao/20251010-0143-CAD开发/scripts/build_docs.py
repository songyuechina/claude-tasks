# -*- coding: utf-8 -*-
"""
Build docs to combined Markdown/HTML/PDF.

Usage:
    python scripts/build_docs.py

Outputs under artifacts/docs/:
    - combined.md: 合并的规则文档（索引顺序）
    - combined.html: 若检测到 pandoc/markdown 库可用则生成
    - combined.pdf: 若检测到 pandoc 可用则生成
"""
from __future__ import annotations

import os
import shutil
import subprocess
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = ROOT / 'artifacts' / 'docs'
OUT_DIR.mkdir(parents=True, exist_ok=True)

ORDER = [
    'docs/INDEX.md',
    'docs/CONCEPTS.md',
    'docs/COORDINATION.md',
    'docs/OPEN_POLICY.md',
    'docs/OPERATIONS.md',
    'docs/TOOLS.md',
]

def read(p: Path) -> str:
    try:
        return p.read_text(encoding='utf-8')
    except Exception:
        return ''

def main() -> int:
    parts = []
    for rel in ORDER:
        fp = ROOT / rel
        if fp.exists():
            parts.append(read(fp))
    combined = '\n\n'.join(parts)
    (OUT_DIR / 'combined.md').write_text(combined, encoding='utf-8')

    # HTML via pandoc or python-markdown if available
    html_ok = False
    if shutil.which('pandoc'):
        try:
            subprocess.run([
                'pandoc', '-f', 'markdown', '-t', 'html',
                '-o', str(OUT_DIR / 'combined.html'),
                str(OUT_DIR / 'combined.md')
            ], check=False)
            html_ok = True
        except Exception:
            pass
    else:
        try:
            import markdown  # type: ignore
            html = markdown.markdown(combined)
            (OUT_DIR / 'combined.html').write_text(html, encoding='utf-8')
            html_ok = True
        except Exception:
            pass

    # PDF via pandoc if available
    if shutil.which('pandoc'):
        try:
            subprocess.run([
                'pandoc', str(OUT_DIR / 'combined.md'), '-o', str(OUT_DIR / 'combined.pdf')
            ], check=False)
        except Exception:
            pass

    print('Docs combined at:', OUT_DIR)
    print('HTML:', 'OK' if html_ok else 'SKIP')
    return 0

if __name__ == '__main__':
    raise SystemExit(main())

