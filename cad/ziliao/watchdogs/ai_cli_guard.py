# -*- coding: utf-8 -*-
import os, sys, time, argparse, ctypes, io
from typing import Optional
import psutil

import win32gui, win32con
from pywinauto.application import Application
from pywinauto.findwindows import find_windows

"""
心跳为第 0 信号（最优先）：

- AI 被要求：所有任务执行过程中，每 60 秒必须写入一次时间戳
- 文件：D:\claude-tasks\cad\任务进度汇报.txt
- 格式：仅时间戳（如 2025-11-14 10:30:45）
- 守护逻辑：
  * 若心跳超过 --idle-seconds 未更新：注入一条状态检查消息（激活它）；
  * 若心跳超过 --restart-idle-window 未更新：认为卡死，重启；
  * CPU / stdout 只用于“辅助观察”，不再阻止注入/重启。

典型调用示例：

cd /d D:\Myprogramsystem\XT\Basic_service_processing\watchdogs
python ai_cli_guard.py ^
  --message "状态检查：是否遇到问题？" ^
  --idle-seconds 300 ^
  --restart-check-period 60 ^
  --restart-idle-window 900 ^
  --heartbeat-path D:\claude-tasks\cad\任务进度汇报.txt ^
  --stdout-log D:\claude-tasks\ai_console.log ^
  --start-bat D:\claude-tasks\启动Claude-CAD.bat ^
  --start-wait 10 ^
  --title-hint "Claude Code"

"""

# ===================== 窗口前置与发键 =====================
user32 = ctypes.windll.user32

def force_foreground(hwnd: int):
    """恢复并置前控制台窗口，处理最小化/前台锁/线程附着。"""
    try:
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            time.sleep(0.05)
        target_tid = user32.GetWindowThreadProcessId(hwnd, None)
        cur_tid = user32.GetCurrentThreadId()
        attached = user32.AttachThreadInput(cur_tid, target_tid, True)
        try:
            win32gui.BringWindowToTop(hwnd);   time.sleep(0.02)
            win32gui.SetForegroundWindow(hwnd);time.sleep(0.02)
            win32gui.SetActiveWindow(hwnd);    time.sleep(0.02)
            user32.ShowWindow(hwnd, win32con.SW_SHOW)
            user32.SetForegroundWindow(hwnd)
        finally:
            if attached:
                user32.AttachThreadInput(cur_tid, target_tid, False)
    except Exception:
        pass

def type_text_enter(hwnd: int, text: str):
    """优先 UIA，回退 win32，发送文本并回车。"""
    try:
        app = Application(backend="uia").connect(handle=hwnd, timeout=5)
        dlg = app.window(handle=hwnd); dlg.set_focus()
        dlg.type_keys(text, with_spaces=True, pause=0.01)
        dlg.type_keys("{ENTER}")
        return
    except Exception:
        pass
    try:
        app = Application(backend="win32").connect(handle=hwnd, timeout=5)
        dlg = app.window(handle=hwnd); dlg.set_focus()
        dlg.type_keys(text, with_spaces=True, pause=0.01)
        dlg.type_keys("{ENTER}")
    except Exception:
        pass

def hwnd_from_pid(pid: int) -> Optional[int]:
    """按进程 PID 找一个窗口句柄。"""
    try:
        handles = find_windows(process=pid)
        return handles[0] if handles else None
    except Exception:
        return None

def find_console_hwnd(title_hint: str, pid: Optional[int] = None) -> Optional[int]:
    """
    启动初期用 title_hint 帮助找一次 Claude 控制台窗口。
    之后主循环中仅在 hwnd 失效时用它兜底。

    注意：AI 后续可以随意改窗口标题，但 hwnd 不变，所以只在“刚启动”时依赖 title。
    """
    # 1) ConsoleWindowClass + 标题关键字
    if title_hint:
        try:
            handles = find_windows(title_re=title_hint, class_name="ConsoleWindowClass")
            if handles:
                return handles[0]
        except Exception:
            pass
        # 2) 任意窗口标题包含关键字
        try:
            handles = find_windows(title_re=title_hint)
            if handles:
                return handles[0]
        except Exception:
            pass

    # 3) 兜底：按 PID
    if pid:
        return hwnd_from_pid(pid)
    return None

# ===================== 文件 I/O 工具 =====================
def write_text(path: str, content: str):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with io.open(path, "w", encoding="utf-8") as f:
        f.write(content)

def read_text(path: str) -> Optional[str]:
    try:
        with io.open(path, "r", encoding="utf-8") as f:
            return f.read().strip()
    except Exception:
        return None

def safe_mtime(path: Optional[str]) -> Optional[float]:
    if not path:
        return None
    try:
        return os.path.getmtime(path)
    except Exception:
        return None

# ===================== 通过 .bat 启动新的 AI 控制台 =====================
def start_new_from_bat(bat_path: str, title_hint: str, start_wait: float) -> tuple[int, Optional[int]]:
    """
    直接新开控制台并运行 .bat（.bat 内再启动 ClaudeConsole.ps1 等）。
    返回: (启动进程 PID, Claude 控制台 HWND or None)

    说明：
    - 这里的 PID 是 cmd.exe 的 PID，用来检测“这整条链路是否还活着”；
    - 真正注入时使用的是 Claude 控制台的 hwnd，与标题无关。
    """
    if not os.path.exists(bat_path):
        raise FileNotFoundError(f"启动脚本不存在: {bat_path}")

    CREATE_NEW_CONSOLE = 0x00000010
    cmdline = ["cmd.exe", "/k", f'call "{bat_path}"']
    proc = psutil.Popen(cmdline, creationflags=CREATE_NEW_CONSOLE)

    # 等待 .bat 链路拉起来（包括 ClaudeConsole.ps1 / /cad 等）
    time.sleep(start_wait)

    # 尝试找到 Claude 控制台窗口（通常标题一开始是 title_hint，后面可能被 AI 改掉）
    hwnd_cli = None
    for _ in range(60):  # 最多再等 6 秒
        hwnd_cli = find_console_hwnd(title_hint, None)
        if hwnd_cli:
            break
        time.sleep(0.1)

    return proc.pid, hwnd_cli

# ===================== 守护主逻辑 =====================
def guard(pid: Optional[int],
          inject_message: str,
          cpu_idle_threshold: float,
          idle_seconds_for_inject: float,
          stop_flag_path: str,
          stop_check_period: float,
          restart_check_period: float,
          restart_idle_window: float,
          heartbeat_path: Optional[str],
          stdout_log_path: Optional[str],
          start_bat: str,
          start_wait: float,
          title_hint: str):

    # 若没提供 PID 或 PID 不存在，则用 .bat 自启动
    target_proc: psutil.Process
    if pid is not None:
        try:
            p = psutil.Process(pid)
            if p.is_running():
                target_proc = p
            else:
                target_proc = None  # type: ignore
        except psutil.NoSuchProcess:
            target_proc = None  # type: ignore
    else:
        target_proc = None  # type: ignore

    if target_proc is None:
        print("[INIT] 未检测到运行中的 AI 会话，使用 .bat 自启动……")
        new_pid, new_hwnd = start_new_from_bat(start_bat, title_hint, start_wait)
        print(f"[INIT] 已启动：PID={new_pid}")
        target_proc = psutil.Process(new_pid)
        hwnd = new_hwnd
    else:
        hwnd = find_console_hwnd(title_hint, target_proc.pid)

    print(f"[GUARD] 监控启动进程 PID={target_proc.pid}")
    print(f"[GUARD] 心跳文件: {heartbeat_path}")
    print(f"[GUARD] stdout 日志: {stdout_log_path}")
    print(f"[GUARD] 注入：心跳无更新 ≥ {int(idle_seconds_for_inject)}s")
    print(f"[GUARD] 重启：心跳无更新 ≥ {int(restart_idle_window)}s")

    # 启动即把“可结束标记”置 0
    if stop_flag_path:
        try:
            write_text(stop_flag_path, "0")
            print(f"[INIT] 已写入 {stop_flag_path} = 0")
        except Exception as e:
            print(f"[WARN] 写入 {stop_flag_path} 失败：{e}")

    # CPU 统计预热（仅供参考）
    try:
        target_proc.cpu_percent(None)
    except Exception:
        pass

    start_time = time.time()
    last_cpu_poll = 0.0
    last_file_check = 0.0
    last_restart_check = 0.0

    # 心跳相关
    last_hb_mtime = safe_mtime(heartbeat_path)
    last_hb_change: Optional[float] = None
    injected_since_last_hb: bool = False

    if last_hb_mtime:
        last_hb_change = start_time  # 启动前已有一次心跳，视为刚活过
        print(f"[INIT] 初始心跳 mtime = {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(last_hb_mtime))}")
    else:
        print("[INIT] 尚未发现心跳文件，将从启动时间开始计时。")

    # stdout 仅记录时间，不参与“是否卡住”的判定
    last_stdout_mtime = safe_mtime(stdout_log_path)

    while True:
        time.sleep(1.0)
        now = time.time()

        # 1) 启动进程是否活着（如果整个 cmd 被关掉，就重启整链路）
        if not target_proc.is_running():
            print("[INFO] 启动进程已退出，执行重启（.bat）。")
            new_pid, new_hwnd = start_new_from_bat(start_bat, title_hint, start_wait)
            print(f"[RESTART] 新会话 PID={new_pid}")
            target_proc = psutil.Process(new_pid)
            hwnd = new_hwnd
            try:
                target_proc.cpu_percent(None)
            except Exception:
                pass
            start_time = time.time()
            last_hb_mtime = safe_mtime(heartbeat_path)
            last_hb_change = None
            injected_since_last_hb = False
            last_stdout_mtime = safe_mtime(stdout_log_path)
            continue

        # 刷新 Claude 控制台窗口句柄（如果原来的 hwnd 失效了，用 title_hint 再找一次）
        if hwnd is None or not win32gui.IsWindow(hwnd):
            hwnd = find_console_hwnd(title_hint, None)

        # ===== CPU 仅供参考，不参与“卡住”决策 =====
        if (now - last_cpu_poll) >= 5.0:  # 慢一点采样
            try:
                cpu = target_proc.cpu_percent(interval=None)
                # 可以在这里打印出来调试：print(f"[CPU] {cpu:.1f}%")
            except Exception:
                cpu = 0.0
            last_cpu_poll = now

        # ===== stdout mtime 也只是参考 =====
        cur_stdout_mtime = safe_mtime(stdout_log_path)
        if cur_stdout_mtime and last_stdout_mtime and cur_stdout_mtime != last_stdout_mtime:
            last_stdout_mtime = cur_stdout_mtime

        # ===== 心跳：第 0 信号（真正决定“卡住/重启”的依据）=====
        cur_hb_mtime = safe_mtime(heartbeat_path)
        if cur_hb_mtime and cur_hb_mtime != last_hb_mtime:
            # 检测到新的进度汇报
            last_hb_mtime = cur_hb_mtime
            last_hb_change = now
            injected_since_last_hb = False
            # 可选：打印调试
            # print(f"[HB] {time.strftime('%H:%M:%S', time.localtime(cur_hb_mtime))}")

        # 计算“距离上次心跳过去多久”
        if last_hb_change is not None:
            hb_idle = now - last_hb_change
        else:
            # 还没有任何心跳，用“从守护启动到现在”的时间替代
            hb_idle = now - start_time

        # ===== 1) 心跳超时 → 注入状态检查消息 =====
        if hb_idle >= idle_seconds_for_inject and not injected_since_last_hb and hwnd:
            try:
                force_foreground(hwnd)
                time.sleep(0.1)
                type_text_enter(hwnd, inject_message)
                print(f"[INJECT {time.strftime('%H:%M:%S')}] {inject_message}+Enter (hb_idle={int(hb_idle)}s)")
            except Exception as e:
                print(f"[WARN] 注入失败：{e}")
            finally:
                injected_since_last_hb = True  # 本次“无心跳区间”只注入一次

        # ===== 2) 心跳长时间缺失 → 重启（不受 CPU / stdout 影响）=====
        # 只在 restart_check_period 的节奏上检查一次，避免过于频繁打印日志
        need_restart = False
        if (now - last_restart_check) >= restart_check_period:
            last_restart_check = now
            if hb_idle >= restart_idle_window:
                need_restart = True

        if need_restart:
            print(f"[RESTART] 心跳无更新 ≥ {int(hb_idle)}s，执行重启（.bat）。")
            if hwnd:
                try:
                    force_foreground(hwnd)
                    time.sleep(0.1)
                    type_text_enter(hwnd, "exit")
                    time.sleep(2)
                except Exception:
                    pass
            try:
                target_proc.terminate()
                target_proc.wait(timeout=5)
            except Exception:
                try:
                    target_proc.kill()
                except Exception:
                    pass

            new_pid, new_hwnd = start_new_from_bat(start_bat, title_hint, start_wait)
            print(f"[RESTART] 新会话 PID={new_pid}")
            target_proc = psutil.Process(new_pid)
            hwnd = new_hwnd
            try:
                target_proc.cpu_percent(None)
            except Exception:
                pass
            start_time = time.time()
            last_hb_mtime = safe_mtime(heartbeat_path)
            last_hb_change = None
            injected_since_last_hb = False
            last_stdout_mtime = safe_mtime(stdout_log_path)
            continue

        # ===== 3) 结束标记（外部写 1 让它优雅退出）=====
        if stop_flag_path and (now - last_file_check) >= stop_check_period:
            last_file_check = now
            val = read_text(stop_flag_path)
            if val == "1":
                print("[STOP] 检测到结束标记=1，尝试优雅退出。")
                if hwnd:
                    try:
                        force_foreground(hwnd)
                        time.sleep(0.1)
                        type_text_enter(hwnd, "exit")
                        time.sleep(2)
                    except Exception:
                        pass
                try:
                    target_proc.terminate()
                    target_proc.wait(timeout=5)
                except Exception:
                    try:
                        target_proc.kill()
                    except Exception:
                        pass
                print("[STOP] 已结束。")
                return

def main():
    ap = argparse.ArgumentParser(description="AI 命令行看门狗（心跳第 0 信号 + .bat 启动/重启）")
    ap.add_argument("--pid", type=int, help="已运行的启动进程 PID（可选；缺省时将用 .bat 自启动）")
    ap.add_argument("--message", default="状态检查：是否遇到问题？", help="停滞时注入的文本")
    ap.add_argument("--cpu-threshold", type=float, default=1.0, help="CPU 活动阈值（%）（仅打印参考）")
    ap.add_argument("--idle-seconds", type=float, default=300, help="心跳无更新达到该秒数就注入一次 message")
    ap.add_argument("--stop-flag", default=r"D:\claude-tasks\cad\可以结束对话.txt", help="结束标记文件路径")
    ap.add_argument("--stop-check-period", type=float, default=300, help="检查结束标记周期（秒）")
    ap.add_argument("--restart-check-period", type=float, default=60, help="重启检测周期（秒）")
    ap.add_argument("--restart-idle-window", type=float, default=900, help="触发重启需“心跳无更新”的时长（秒）")
    ap.add_argument("--heartbeat-path", default=r"D:\claude-tasks\cad\任务进度汇报.txt", help="任务进度汇报（心跳）文件路径")
    ap.add_argument("--stdout-log", default=r"D:\claude-tasks\ai_console.log", help="可选：控制台输出日志（仅参考 mtime）")
    ap.add_argument("--start-bat", default=r"D:\claude-tasks\启动Claude-CAD.bat", help="用于启动/重启 AI 的批处理文件路径")
    ap.add_argument("--start-wait", type=float, default=10.0, help="重启后等待 .bat 链路就绪的秒数")
    ap.add_argument("--title-hint", default="Claude Code", help="启动初期用于定位 Claude 控制台窗口的标题关键字")
    args = ap.parse_args()

    if sys.platform != "win32":
        print("仅支持 Windows。")
        sys.exit(1)

    guard(pid=args.pid,
          inject_message=args.message,
          cpu_idle_threshold=args.cpu_threshold,
          idle_seconds_for_inject=args.idle_seconds,
          stop_flag_path=args.stop_flag,
          stop_check_period=args.stop_check_period,
          restart_check_period=args.restart_check_period,
          restart_idle_window=args.restart_idle_window,
          heartbeat_path=args.heartbeat_path,
          stdout_log_path=args.stdout_log,
          start_bat=args.start_bat,
          start_wait=args.start_wait,
          title_hint=args.title_hint)

if __name__ == "__main__":
    main()
