#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import time
import win32gui, win32con, win32process
import psutil
from pathlib import Path
import logging

# —— 日志配置 ——  
LOG_FILE = Path(__file__).parent / "cad_dialog_killer.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)-5s %(message)s",
    filename=LOG_FILE,
    filemode="a"
)
logger = logging.getLogger("cad_killer")

# —— 配置区 ——  
CAD_PROCESS_NAMES = {"tarcht20v9.exe", "acad.exe"}
DIALOG_CLASS      = "#32770"
INTERVAL_SEC      = 15
CONTROL_FILE      = Path(
    r"D:/Myprogramsystem/XT/Basic_service_processing/取消窗口延迟控制.txt"
)

# 用来记录每个 hwnd 首次出现的时间
_first_seen: dict[int, float] = {}
# 用来记录上一次读取到的 delay 值
_last_delay: int | None = None

def read_delay() -> int:
    """从 CONTROL_FILE 读取延迟秒数，失败则返回 0"""
    try:
        txt = CONTROL_FILE.read_text(encoding="utf-8").strip()
        delay = max(0, int(txt))
        logger.debug(f"读到延迟 = {delay}s")
        return delay
    except Exception as e:
        logger.warning(f"读取延迟文件失败 ({e})，使用 0s")
        return 0

def get_cad_pids() -> set[int]:
    """获取所有正在运行的 CAD 进程 PID"""
    pids = set()
    for proc in psutil.process_iter(['pid','name']):
        name = proc.info.get('name') or ""
        if name.lower() in CAD_PROCESS_NAMES:
            pids.add(proc.info['pid'])
    return pids

def enum_and_maybe_close(hwnd, cad_pids, delay, now):
    """枚举窗口，视情况发送 ESC 关闭符合条件的 CAD 对话框"""
    # 不是可见窗口或不是标准对话框类，就重置记录
    if not win32gui.IsWindowVisible(hwnd) or win32gui.GetClassName(hwnd) != DIALOG_CLASS:
        _first_seen.pop(hwnd, None)
        return

    # 窗口不属于 CAD 进程，也清理
    _, pid = win32process.GetWindowThreadProcessId(hwnd)
    if pid not in cad_pids:
        _first_seen.pop(hwnd, None)
        return

    # 首次看到时记录时间
    t0 = _first_seen.get(hwnd)
    if t0 is None:
        _first_seen[hwnd] = now
        logger.info(f"发现对话框 hwnd={hwnd} pid={pid}，等待 {delay}s 后关闭")
        return

    # 已记录，检查是否超过 delay
    elapsed = now - t0
    if delay == 0 or elapsed >= delay:
        # 发送 Esc 键
        win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_ESCAPE, 0)
        win32gui.PostMessage(hwnd, win32con.WM_KEYUP,   win32con.VK_ESCAPE, 0)
        logger.info(f"关闭对话框 hwnd={hwnd} pid={pid}，存在了 {int(elapsed)}s（阈值 {delay}s）")
        _first_seen.pop(hwnd, None)

def main():
    global _last_delay
    logger.info(f"cad_dialog_killer 启动 (间隔 {INTERVAL_SEC}s)，控制文件：{CONTROL_FILE}")

    while True:
        now   = time.time()
        delay = read_delay()

        # 如果 delay 与上次不同，记录变更
        if delay != _last_delay:
            logger.info(f"【DELAY】从控制文件读取到 delay={delay}s (上次={_last_delay}s)")
            _last_delay = delay

        pids = get_cad_pids()
        if pids:
            win32gui.EnumWindows(
                lambda h, _: enum_and_maybe_close(h, pids, delay, now),
                None
            )
        else:
            if _first_seen:
                logger.debug("CAD 已退出，清空对话框首次出现记录")
            _first_seen.clear()

        time.sleep(INTERVAL_SEC)

if __name__ == "__main__":
    main()




    
