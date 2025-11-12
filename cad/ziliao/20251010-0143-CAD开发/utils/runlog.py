# -*- coding: utf-8 -*-
"""
Lightweight run logger for scripts in this task folder.

Features:
- Write human-readable log lines to artifacts/logs/<script_name>/
- Maintain JSON summary with checkpoints, errors and verdict

Usage:
    from utils.runlog import RunLogger
    log = RunLogger('my_script')
    log.log('hello')
    log.mark('start')
    log.error('something bad')
    log.write_summary({...})
"""
from __future__ import annotations

import json
import os
import time
from pathlib import Path
from typing import List


ROOT = Path(__file__).resolve().parents[1]
LOGS = ROOT / 'artifacts' / 'logs'


def _now() -> str:
    return time.strftime('%Y-%m-%d %H:%M:%S')


class RunLogger:
    def __init__(self, script_name: str) -> None:
        self.script_name = script_name
        self.dir = LOGS / script_name
        self.dir.mkdir(parents=True, exist_ok=True)
        self.run_id = time.strftime('%Y%m%d-%H%M%S')
        self.path = self.dir / f'{script_name}-{self.run_id}.log'
        self.json_path = self.dir / f'{script_name}-{self.run_id}.json'
        self.path.write_text('', encoding='utf-8')
        self.errors: List[str] = []
        self.checkpoints: List[str] = []

    def log(self, msg: str) -> None:
        line = f'[{_now()}] {msg}\n'
        try:
            with self.path.open('a', encoding='utf-8') as f:
                f.write(line)
        except Exception:
            pass
        print(msg)

    def error(self, msg: str) -> None:
        self.errors.append(msg)
        self.log(f'ERROR: {msg}')

    def mark(self, name: str) -> None:
        self.checkpoints.append(name)
        self.log(f'CHECKPOINT: {name}')

    def write_summary(self, data: dict) -> None:
        try:
            self.json_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
        except Exception:
            pass

