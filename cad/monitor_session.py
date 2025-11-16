#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CAD任务会话监控脚本

功能：
1. 启动时将结束信号文件设置为0
2. 监控信号文件，检测到1时自动处理
3. 可选：自动关闭Claude对话窗口
"""

import time
from pathlib import Path
import sys

# 信号文件路径
SIGNAL_FILE = Path("D:/claude-tasks/cad/可以结束对话.txt")

def reset_signal():
    """重置信号文件为0"""
    try:
        with open(SIGNAL_FILE, "w", encoding="utf-8") as f:
            f.write("0")
        print(f"✅ 已重置信号文件为0: {SIGNAL_FILE}")
        return True
    except Exception as e:
        print(f"❌ 重置信号文件失败: {e}")
        return False

def check_signal():
    """检查信号文件状态"""
    try:
        if not SIGNAL_FILE.exists():
            print(f"⚠️ 信号文件不存在: {SIGNAL_FILE}")
            # 创建文件
            reset_signal()
            return False

        with open(SIGNAL_FILE, "r", encoding="utf-8") as f:
            content = f.read().strip()

        return content == "1"
    except Exception as e:
        print(f"❌ 读取信号文件失败: {e}")
        return False

def monitor_session(interval=5, auto_close=False):
    """
    监控会话状态

    Args:
        interval: 检查间隔（秒）
        auto_close: 是否自动关闭Claude窗口
    """
    print("="*60)
    print("CAD任务会话监控")
    print("="*60)

    # 启动时重置信号
    if not reset_signal():
        return 1

    print(f"\n🔍 开始监控（检查间隔: {interval}秒）")
    print(f"📁 监控文件: {SIGNAL_FILE}")
    print(f"🔔 自动关闭: {'是' if auto_close else '否'}")
    print("\n按 Ctrl+C 停止监控\n")

    try:
        check_count = 0
        while True:
            check_count += 1

            if check_signal():
                print(f"\n{'='*60}")
                print("✅ 检测到任务完成信号！")
                print(f"{'='*60}")

                # 重置信号文件
                reset_signal()

                if auto_close:
                    print("\n🔄 准备关闭Claude对话窗口...")
                    # 这里可以添加关闭窗口的逻辑
                    # 例如: subprocess.run(["taskkill", "/F", "/IM", "claude-code.exe"])
                    print("⚠️ 自动关闭功能需要手动实现")
                else:
                    print("\n📝 请手动关闭对话窗口")

                print("\n✅ 监控结束")
                return 0

            # 显示心跳
            if check_count % 10 == 0:
                elapsed = check_count * interval
                print(f"⏱ 已监控 {elapsed}秒 - 等待任务完成...")

            time.sleep(interval)

    except KeyboardInterrupt:
        print("\n\n⚠️ 监控已手动停止")
        return 0
    except Exception as e:
        print(f"\n\n❌ 监控异常: {e}")
        return 1

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="CAD任务会话监控脚本")
    parser.add_argument(
        "--interval",
        type=int,
        default=5,
        help="检查间隔（秒），默认5秒"
    )
    parser.add_argument(
        "--auto-close",
        action="store_true",
        help="检测到完成信号后自动关闭Claude窗口"
    )

    args = parser.parse_args()

    sys.exit(monitor_session(
        interval=args.interval,
        auto_close=args.auto_close
    ))
