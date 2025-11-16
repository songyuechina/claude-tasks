# -*- coding: utf-8 -*-
import time, argparse, psutil, sys
from typing import Optional

# pywin32
import win32gui, win32con, win32api
# pywinauto 更稳的发键
from pywinauto.application import Application
from pywinauto.findwindows import find_windows

def find_hwnd_by_title_keyword(keyword: str) -> Optional[int]:
    hwnd_found = None
    def enum_handler(hwnd, _):
        nonlocal hwnd_found
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if keyword.lower() in title.lower():
                hwnd_found = hwnd
    win32gui.EnumWindows(enum_handler, None)
    return hwnd_found

def bring_to_front(hwnd: int):
    # 取消最小化并前置
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    win32gui.SetForegroundWindow(hwnd)

def send_keys_via_pywinauto(hwnd: int, text: str):
    # 控制台窗口用 pywinauto 发送按键更稳
    app = Application(backend="uia").connect(handle=hwnd)
    dlg = app.window(handle=hwnd)
    dlg.set_focus()
    # 直接发送文本与回车
    dlg.type_keys(text, with_spaces=True, pause=0.01)
    dlg.type_keys("{ENTER}")

def monitor_and_inject(pid: Optional[int], title_kw: Optional[str],
                       idle_seconds: float, message: str,
                       min_check_interval: float, cpu_threshold: float,
                       cooldown: float):
    # 锁定目标进程
    target_proc = None
    if pid:
        target_proc = psutil.Process(pid)
    else:
        # 通过标题找 hwnd，再从窗口取 PID
        hwnd = find_hwnd_by_title_keyword(title_kw)
        if not hwnd:
            print("找不到匹配的控制台窗口标题。")
            sys.exit(1)
        _, pid_from_hwnd = win32process.GetWindowThreadProcessId(hwnd)  # type: ignore
        target_proc = psutil.Process(pid_from_hwnd)

    # 找窗口句柄
    def resolve_hwnd():
        if title_kw:
            return find_hwnd_by_title_keyword(title_kw)
        try:
            # 从 PID 反查窗口（粗略做法）
            handles = find_windows(process=target_proc.pid)
            return handles[0] if handles else None
        except Exception:
            return None

    last_active = time.time()
    last_cpu_check = 0.0
    print(f"监控进程 PID={target_proc.pid}，阈值：空闲≥{idle_seconds}s 视为停滞。")

    # 先预热一次 CPU 统计
    target_proc.cpu_percent(None)

    while True:
        time.sleep(min_check_interval)
        if not target_proc.is_running():
            print("进程已退出，停止监控。")
            break

        now = time.time()
        if (now - last_cpu_check) >= min_check_interval:
            cpu = target_proc.cpu_percent(interval=None)  # 非阻塞取上次到现在的占用
            last_cpu_check = now
            # 简单判定：CPU 低于阈值即认为“无活动”
            if cpu >= cpu_threshold:
                last_active = now

        idle_for = now - last_active
        # 满足停滞
        if idle_for >= idle_seconds:
            hwnd = resolve_hwnd()
            if not hwnd:
                print("未找到控制台窗口句柄，跳过一次注入。")
                last_active = now  # 避免频繁尝试
                continue
            try:
                bring_to_front(hwnd)
                time.sleep(0.1)
                send_keys_via_pywinauto(hwnd, message)
                print(f"[{time.strftime('%H:%M:%S')}] 停滞 {int(idle_for)}s，已发送：{message}+Enter")
            except Exception as e:
                print("发送按键失败：", e)
            finally:
                # 注入后进入冷却期，避免狂刷
                last_active = time.time()
                time.sleep(cooldown)

def main():
    ap = argparse.ArgumentParser(description="监控已运行的控制台进程并在停滞时发送“继续+回车”（Windows）")
    ap.add_argument("--pid", type=int, help="目标进程 PID（可选，与 --title 二选一）")
    ap.add_argument("--title", type=str, help="控制台窗口标题关键字（可选，与 --pid 二选一）")
    ap.add_argument("--idle-seconds", type=float, default=60.0, help="CPU 低占用视为停滞的秒数")
    ap.add_argument("--message", type=str, default="继续", help="要发送的文本（默认“继续”）")
    ap.add_argument("--min-check-interval", type=float, default=1.0, help="最小检测周期秒")
    ap.add_argument("--cpu-threshold", type=float, default=1.0, help="CPU 活动阈值（％），低于它视为“无活动”")
    ap.add_argument("--cooldown", type=float, default=10.0, help="每次注入后的冷却秒数")
    args = ap.parse_args()

    if not args.pid and not args.title:
        ap.error("需要 --pid 或 --title 至少提供一个。")

    # 延迟导入以避免非 Windows 报错
    global win32process
    import win32process  # type: ignore

    monitor_and_inject(args.pid, args.title, args.idle_seconds, args.message,
                       args.min_check_interval, args.cpu_threshold, args.cooldown)

if __name__ == "__main__":
    if sys.platform != "win32":
        print("该脚本仅适用于 Windows（方案 B）。跨平台请用方案 A。")
        sys.exit(1)
    main()
