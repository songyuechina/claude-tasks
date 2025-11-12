# -*- coding: utf-8 -*-
"""
CAD运行协同机制模块

实现CAD命令的同步等待、文档打开等待、空闲状态检测等协同机制
确保CAD操作命令的顺序执行和完整性
"""

import time
import os
import subprocess
from pathlib import Path

# Windows进程创建标志
DETACHED_PROCESS = 0x00000008

def wait_quiescent(min_quiet: float = 0.4, timeout: float = 30.0) -> bool:
    """
    等待CAD进入空闲状态

    Args:
        min_quiet: 最小安静时间(秒),在此时CAD保持空闲
        timeout: 超时时间(秒)

    Returns:
        bool: True表示CAD已空闲,False表示超时
    """
    try:
        # 导入CAD_basic模块中的get_acad_doc函数
        import sys
        sys.path.append(str(Path(__file__).parent))
        from CAD_basic import get_acad_doc

        start_time = time.time()
        quiet_start = None

        while time.time() - start_time < timeout:
            try:
                # 尝试获取CAD文档状态
                _, doc = get_acad_doc(max_wait=0.1)
                if doc:
                    # 检查文档是否处于空闲状态
                    # 通过尝试访问文档属性来判断是否空闲
                    try:
                        _ = doc.Name  # 简单的属性访问测试
                        if quiet_start is None:
                            quiet_start = time.time()
                        elif time.time() - quiet_start >= min_quiet:
                            print(f"[成功] CAD已进入空闲状态 (安静时间: {min_quiet}s)")
                            return True
                    except:
                        # 访问失败,重置静计时
                        quiet_start = None
                else:
                    quiet_start = None

            except:
                quiet_start = None

            time.sleep(0.1)

        print(f"[警告] CAD等待空闲超时 ({timeout}s)")
        return False

    except Exception as e:
        print(f"[错误] 等待CAD空闲时出错: {e}")
        return False

def wait_document_opened(path: str, timeout: float = 120.0) -> bool:
    """
    等待文档被CAD完全打开并加入Documents集合

    Args:
        path: 文档路径
        timeout: 超时时间(秒)

    Returns:
        bool: True表示文档已打开,False表示超时
    """
    try:
        import win32com.client

        start_time = time.time()
        target_path = os.path.abspath(path).lower()

        while time.time() - start_time < timeout:
            try:
                acad = win32com.client.GetActiveObject("AutoCAD.Application")
                documents = acad.Documents
                count = documents.Count

                for i in range(count):
                    doc = documents.Item(i)
                    doc_path = doc.FullName.lower() if doc.FullName else ""

                    # 检查路径匹配
                    if doc_path == target_path:
                        # 额外等待确保文档完全加载
                        time.sleep(0.5)
                        print(f"[成功] 文档已成功打开: {path}")
                        return True

                    # 检查文件名匹配(处理路径格式差异)
                    if os.path.basename(doc_path).lower() == os.path.basename(target_path).lower():
                        time.sleep(0.5)
                        print(f"[成功] 文档已成功打开(文件名匹配): {path}")
                        return True

            except Exception:
                pass

            time.sleep(0.3)

        print(f"[警告] 等待文档打开超时: {path}")
        return False

    except Exception as e:
        print(f"[错误] 等待文档打开时出错: {e}")
        return False

def send_cmd_with_sync(cmd: str, wait_after: float = 0.3, timeout: float = 30.0) -> bool:
    """
    发送CAD命令并同步等待执行完成

    Args:
        cmd: 要发送的CAD命令
        wait_after: 命令发送后的基础等待时间
        timeout: 等待CAD空闲的超时时间

    Returns:
        bool: True表示命令执行成功,False表示失败
    """
    try:
        # 导入CAD_basic模块中的get_acad_doc函数
        import sys
        sys.path.append(str(Path(__file__).parent))
        from CAD_basic import get_acad_doc

        # 发送命令
        _, doc = get_acad_doc()
        if not doc:
            print("[错误] 无法获取CAD文档")
            return False

        formatted_cmd = cmd if cmd.endswith("\n") else (cmd + "\n")
        doc.SendCommand(formatted_cmd)
        print(f"[成功] 发送CAD命令: {cmd.strip()}")

        # 基础等待
        time.sleep(wait_after)

        # 等待CAD空闲
        return wait_quiescent(timeout=timeout)

    except Exception as e:
        print(f"[错误] 发送CAD命令失败: {e}")
        return False

def start_cad_with_dialog_killer() -> bool:
    """
    启动CAD并运行弹窗治理脚本

    Returns:
        bool: True表示启动成功,False表示失败
    """
    try:
        # 导入CAD_basic模块中的start_applicationV9函数
        import sys
        sys.path.append(str(Path(__file__).parent))
        from CAD_basic import start_applicationV9

        # 启动CAD
        print("[启动] 正在启动天正CAD...")
        proc = start_applicationV9()
        if not proc:
            print("[错误] CAD启动失败")
            return False

        # 等待CAD启动完成
        time.sleep(3.0)

        # 启动弹窗治理脚本（使用DETACHED_PROCESS避免创建可见窗口）
        dialog_killer_path = Path(__file__).parent / "cad_dialog_killer.py"
        if dialog_killer_path.exists():
            # 检查是否已在运行
            import psutil
            already_running = False
            for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
                try:
                    cmdline = proc.info.get('cmdline') or []
                    if any("cad_dialog_killer" in str(arg) for arg in cmdline):
                        already_running = True
                        print("[信息] 弹窗治理脚本已在运行")
                        break
                except:
                    pass

            if not already_running:
                print("[保护] 启动弹窗治理脚本...")
                subprocess.Popen([
                    "python", str(dialog_killer_path)
                ], creationflags=DETACHED_PROCESS)
                time.sleep(1.0)
        else:
            print("[警告] 弹窗治理脚本不存在,跳过启动")

        print("[成功] CAD启动完成,弹窗治理已启用")
        return True

    except Exception as e:
        print(f"[错误] 启动CAD时出错: {e}")
        return False

def ensure_single_process() -> bool:
    """
    确保只有一个CAD进程运行

    Returns:
        bool: True表示成功确保单进程,False表示失败
    """
    try:
        import psutil

        CAD_PROCESS_NAME = "acad.exe"
        processes = []

        # 查找所有CAD进程
        for proc in psutil.process_iter(['pid', 'name']):
            if proc.info['name'] == CAD_PROCESS_NAME:
                processes.append(proc)

        if len(processes) == 0:
            print("[信息] 没有发现CAD进程")
            return True
        elif len(processes) == 1:
            print("[信息] 已是单CAD进程状态")
            return True
        else:
            print(f"[警告] 发现 {len(processes)} 个CAD进程,尝试关闭多余进程...")

            # 保留第一个进程,关闭其他进程
            for proc in processes[1:]:
                try:
                    proc.terminate()
                    print(f"[成功] 已终止CAD进程 PID: {proc.info['pid']}")
                except:
                    try:
                        proc.kill()
                        print(f"[成功] 已强制终止CAD进程 PID: {proc.info['pid']}")
                    except:
                        print(f"[错误] 无法终止CAD进程 PID: {proc.info['pid']}")

            time.sleep(1.0)
            print("[成功] 已确保单CAD进程状态")
            return True

    except Exception as e:
        print(f"[错误] 确保单进程时出错: {e}")
        return False

# 测试函数
if __name__ == "__main__":
    print("[测试] 测试CAD协同机制...")

    # 测试启动CAD和弹窗治理
    print("\n1. 测试启动CAD和弹窗治理:")
    start_cad_with_dialog_killer()

    # 测试确保单进程
    print("\n2. 测试确保单进程:")
    ensure_single_process()

    # 测试等待空闲
    print("\n3. 测试等待CAD空闲:")
    wait_quiescent()

    print("\n[成功] CAD协同机制测试完成")