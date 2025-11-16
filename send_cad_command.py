# -*- coding: utf-8 -*-
"""
从所有窗口中找到 Claude Code 控制台窗口，
把它激活到前台，然后发送 `/cad` + 两次回车。

依赖：
    pip install pyautogui pygetwindow
"""

import time

import pyautogui
import pygetwindow as gw


# 必须包含的关键字（不区分大小写）
TARGET_KEYWORDS = ["claude code", "claude"]
# 明确排除的关键字（例如看门狗自己的窗口）
EXCLUDE_KEYWORDS = ["ai_cli_guard.py"]


def find_claude_window():
    """在所有窗口中找到最像 Claude Code 的那个。"""
    wins = gw.getAllWindows()
    candidates = []

    for w in wins:
        title = (w.title or "").strip()
        t_low = title.lower()

        # 1) 必须包含任一目标关键字
        if not any(k in t_low for k in TARGET_KEYWORDS):
            continue

        # 2) 排除看门狗等无关窗口
        if any(k in t_low for k in EXCLUDE_KEYWORDS):
            continue

        # 3) 只要有尺寸的窗口（防止一些隐藏窗口）
        if w.width <= 0 or w.height <= 0:
            continue

        candidates.append(w)

    if not candidates:
        return None

    # 简单策略：面积最大的那个一般就是主控制台
    candidates.sort(key=lambda w: w.width * w.height, reverse=True)
    return candidates[0]


def activate_and_send_cad():
    """等待 Claude 窗口出现，然后激活并发送 /cad。"""
    # 最多等 20 秒，给 Claude 有时间启动
    deadline = time.time() + 20
    win = None
    while time.time() < deadline:
        win = find_claude_window()
        if win:
            break
        time.sleep(0.5)

    if not win:
        print("[ERROR] 未找到标题包含 'claude' 的控制台窗口。")
        return False

    # 激活 + 点击标题栏，确保焦点在这个窗口
    try:
        win.activate()
    except Exception:
        pass
    time.sleep(0.2)

    x = win.left + win.width // 2
    y = win.top + max(5, min(30, win.height // 10))
    pyautogui.click(x, y)
    time.sleep(0.1)

    print(f"[INFO] activated window: {win.title!r} at ({win.left},{win.top}) size {win.width}x{win.height}")

    # 发送 /cad + 两次回车（保险：第一次被 UI 吃掉时第二次还能触发）
    pyautogui.typewrite("/cad", interval=0.02)
    pyautogui.press("enter")
    time.sleep(0.1)
    pyautogui.press("enter")

    print("[INFO] sent '/cad' + 2x Enter")
    return True


if __name__ == "__main__":
    ok = activate_and_send_cad()
    if not ok:
        input("未成功找到窗口，按任意键退出...")
