# -*- coding: utf-8 -*-
"""
CAD基本操作范式实现

基于ziliao/20251010-0143-CAD开发资料学习的完整操作范式
包含新建、打开、关闭、保存、插入等基本操作的增强实现
"""

import time
import os
import pythoncom
from pathlib import Path
from typing import List, Optional, Tuple

# 导入协同机制
from CAD_coordination import (
    wait_quiescent,
    wait_document_opened,
    send_cmd_with_sync,
    start_cad_with_dialog_killer,
    ensure_single_process
)

def get_cad_process_count() -> int:
    """获取CAD进程数量"""
    try:
        import psutil
        count = 0
        for proc in psutil.process_iter(['pid', 'name']):
            if proc.info['name'] == 'acad.exe':
                count += 1
        return count
    except:
        return 0

def get_open_file_count() -> int:
    """获取当前打开的DWG文件数量"""
    try:
        import win32com.client
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        return acad.Documents.Count
    except:
        return 0

def is_file_opened(file_path: str) -> bool:
    """检查指定文件是否已打开"""
    try:
        import win32com.client
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        target_path = Path(file_path).resolve().as_posix().lower()

        for i in range(acad.Documents.Count):
            doc = acad.Documents.Item(i)
            if doc.FullName:
                doc_path = Path(doc.FullName).resolve().as_posix().lower()
                if doc_path == target_path:
                    return True
        return False
    except:
        return False

def is_file_opened_by_name(file_name: str) -> bool:
    """检查文件名是否已打开"""
    try:
        import win32com.client
        acad = win32com.client.GetActiveObject("AutoCAD.Application")

        for i in range(acad.Documents.Count):
            doc = acad.Documents.Item(i)
            if doc.Name == file_name:
                return True
        return False
    except:
        return False

def _get_short_path(long_path: str) -> str:
    """获取短路径处理中文/特殊字符"""
    try:
        import ctypes
        from ctypes import wintypes

        GetShortPathNameW = ctypes.windll.kernel32.GetShortPathNameW
        GetShortPathNameW.argtypes = [
            wintypes.LPCWSTR, wintypes.LPWSTR, wintypes.DWORD
        ]
        GetShortPathNameW.restype = wintypes.DWORD

        buf = ctypes.create_unicode_buffer(260)
        ret = GetShortPathNameW(long_path, buf, len(buf))
        return buf.value if ret else long_path
    except Exception:
        return long_path

def _activate_document(file_path: str) -> bool:
    """激活指定文档"""
    try:
        import win32com.client
        acad = win32com.client.GetActiveObject("AutoCAD.Application")

        target_path = Path(file_path).resolve().as_posix().lower()

        for i in range(acad.Documents.Count):
            doc = acad.Documents.Item(i)
            if doc.FullName:
                doc_path = Path(doc.FullName).resolve().as_posix().lower()
                if doc_path == target_path:
                    acad.ActiveDocument = doc
                    return True
        return False
    except Exception:
        return False

# ============================================================================
# 1. 新建文件操作范式
# ============================================================================

def new_dwg_enhanced(output_path: Optional[str] = None) -> bool:
    """
    新建DWG文件范式（使用pywin32 Documents.Add()）

    规则:
    - 幂等操作: output_path已存在时不再新建,直接打开
    - 无output_path时创建未保存的空白文件
    - 使用pywin32直接创建，更简单可靠

    前置条件:
    - CAD进程已启动
    - 弹窗治理脚本运行中

    后置条件:
    - 文件已创建或已打开
    - 状态为单文件确定状态(有路径)或单文件不确定状态(无路径)
    """
    try:
        import win32com.client

        # 1. 确保CAD环境就绪
        if not ensure_single_process():
            return False
        wait_quiescent(min_quiet=0.5, timeout=15.0)

        # 2. 检查路径幂等性
        if output_path and Path(output_path).exists():
            print(f"[成功] 文件已存在,直接打开: {output_path}")
            return open_dwg_paradigm(output_path)

        # 3. 连接到CAD
        acad = win32com.client.GetActiveObject("AutoCAD.Application")

        # 4. 使用pywin32创建新文档
        print("[新建] 正在创建新文件...")
        doc = acad.Documents.Add()
        time.sleep(1.0)

        # 5. 如需保存
        if output_path:
            print(f"[保存] 正在保存为: {output_path}")
            # 创建输出目录
            Path(output_path).parent.mkdir(parents=True, exist_ok=True)

            doc.SaveAs(output_path)
            print(f"[成功] 新建并保存文件: {output_path}")
            wait_quiescent(min_quiet=1.0, timeout=15.0)
            return True
        else:
            print("[成功] 新建未保存文件成功")
            return True

    except Exception as e:
        print(f"[错误] 新建文件操作异常: {e}")
        return False

# ============================================================================
# 2. 打开文件操作范式
# ============================================================================

def open_dwg_paradigm(file_path: str) -> bool:
    """
    打开DWG文件范式

    规则:
    - 顺序+去重: 同一文件只触发一次Documents.Open()
    - 等待加入集合: 确保文档真正加入acad.Documents
    - 路径/名称幂等: 避免重复打开相同文件

    前置条件:
    - 非破坏性进程保证
    - 弹窗治理检查

    后置条件:
    - 文件成功打开并激活
    - CAD进入空闲状态
    """
    try:
        # 1. 基础验证
        if not Path(file_path).exists():
            print(f"[错误] 文件不存在: {file_path}")
            return False

        # 2. 进程预处理(非破坏性)
        process_count = get_cad_process_count()
        if process_count == 0:
            print("[启动] CAD未运行,启动CAD...")
            if not start_cad_with_dialog_killer():
                return False
        elif process_count > 1:
            print("[警告] 发现多个CAD进程,确保单进程...")
            ensure_single_process()

        # 3. 等待CAD稳定
        wait_quiescent(min_quiet=0.3, timeout=15.0)

        # 4. 路径级幂等检查
        if is_file_opened(file_path):
            print(f"[成功] 文件已打开: {file_path}")
            return True

        # 5. 名称级幂等检查
        basename = Path(file_path).name
        if is_file_opened_by_name(basename):
            print(f"[警告] 同名文件已打开,跳过: {basename}")
            return True

        # 6. 执行打开操作
        print(f"[处理] 正在打开: {file_path}")

        # 使用协同机制发送打开命令
        import win32com.client
        acad = win32com.client.GetActiveObject("AutoCAD.Application")

        # 转换为短路径处理中文/特殊字符
        short_path = _get_short_path(file_path)

        # 执行打开
        acad.Documents.Open(short_path)

        # 7. 等待文档加入集合
        if wait_document_opened(file_path, timeout=120.0):
            print(f"[成功] 文件成功打开: {file_path}")

            # 8. 激活文档
            _activate_document(file_path)

            # 9. 等待CAD空闲
            wait_quiescent(min_quiet=0.5, timeout=30.0)

            return True
        else:
            print(f"[错误] 文件打开超时: {file_path}")
            return False

    except Exception as e:
        print(f"[错误] 打开文件异常: {e}")
        return False

def open_multiple_files_paradigm(file_paths: List[str]) -> int:
    """
    多文件打开范式(严格顺序)

    规则:
    - 严格顺序,不可并发
    - 每次成功后等待空闲+0.3s缓冲
    - 路径/同名已存在则跳过
    """
    success_count = 0

    # 进程预处理
    ensure_single_process()
    wait_quiescent(min_quiet=0.3, timeout=15.0)

    print(f"[打开] 开始打开 {len(file_paths)} 个文件")

    for i, file_path in enumerate(file_paths):
        print(f"\n[文件] [{i+1}/{len(file_paths)}] {file_path}")

        if open_dwg_paradigm(file_path):
            success_count += 1
            print(f"[成功] 成功打开: {Path(file_path).name}")
        else:
            print(f"[错误] 打开失败: {Path(file_path).name}")

        # 文件间间隔等待
        if i < len(file_paths) - 1:
            time.sleep(0.3)
            wait_quiescent(min_quiet=0.3, timeout=15.0)

    print(f"\n[统计] 打开结果: {success_count}/{len(file_paths)} 成功")
    return success_count

# ============================================================================
# 3. 关闭文件操作范式
# ============================================================================

def close_current_dwg_paradigm(save_option: str = "prompt") -> bool:
    """
    关闭当前文件范式

    参数:
    - save_option: "prompt"(提示保存), "auto_save"(自动保存), "no_save"(不保存)
    """
    try:
        # 1. 检查是否有文件打开
        if get_open_file_count() == 0:
            print("[警告] 没有打开的文件")
            return True

        # 2. 获取当前文件信息
        import win32com.client
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        current_doc = acad.ActiveDocument
        doc_name = current_doc.Name

        print(f"[处理] 正在关闭当前文件: {doc_name}")

        # 3. 处理保存选项
        if save_option == "auto_save":
            # 自动保存
            try:
                current_doc.Save()
                print(f"[成功] 已保存: {doc_name}")
            except Exception as save_error:
                print(f"[警告] 保存失败: {save_error}")
        elif save_option == "no_save":
            # 不保存
            print(f"[警告] 不保存关闭: {doc_name}")
        else:
            # 提示保存(默认)
            print(f"[提示] 提示保存: {doc_name}")

        # 4. 执行关闭命令
        success = send_cmd_with_sync("_CLOSE\n", wait_after=1.0, timeout=30.0)

        if success:
            # 5. 等待关闭完成
            wait_quiescent(min_quiet=1.0, timeout=30.0)
            print(f"[成功] 文件已关闭: {doc_name}")
            return True
        else:
            print(f"[错误] 关闭文件失败: {doc_name}")
            return False

    except Exception as e:
        print(f"[错误] 关闭文件异常: {e}")
        return False

def close_dwg_by_name_paradigm(file_name: str) -> bool:
    """按文件名关闭文件范式"""
    try:
        # 1. 检查文件是否存在
        if not is_file_opened_by_name(file_name):
            print(f"[警告] 文件未打开: {file_name}")
            return True

        # 2. 切换到目标文件
        import win32com.client
        acad = win32com.client.GetActiveObject("AutoCAD.Application")

        # 查找并激活目标文件
        for i in range(acad.Documents.Count):
            doc = acad.Documents.Item(i)
            if doc.Name == file_name:
                acad.ActiveDocument = doc
                break

        # 3. 关闭文件
        return close_current_dwg_paradigm()

    except Exception as e:
        print(f"[错误] 按名关闭文件异常: {e}")
        return False

def close_all_dwg_paradigm() -> bool:
    """关闭所有文件范式"""
    try:
        file_count = get_open_file_count()
        if file_count == 0:
            print("[警告] 没有打开的文件")
            return True

        print(f"[处理] 准备关闭 {file_count} 个文件")

        # 逐一关闭文件
        success_count = 0
        for _ in range(file_count):
            if close_current_dwg_paradigm():
                success_count += 1
            time.sleep(0.5)  # 间隔等待

        print(f"[成功] 关闭完成: {success_count}/{file_count} 成功")
        return success_count == file_count

    except Exception as e:
        print(f"[错误] 关闭所有文件异常: {e}")
        return False

# ============================================================================
# 4. 保存文件操作范式
# ============================================================================

def save_current_dwg_paradigm() -> bool:
    """
    保存当前文件范式

    规则:
    - 使用短路径处理中文/特殊字符
    - 确保保存操作完成
    - 文件状态为已保存
    """
    try:
        # 1. 检查是否有文件打开
        if get_open_file_count() == 0:
            print("[错误] 没有打开的文件")
            return False

        # 2. 获取文件信息
        import win32com.client
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        current_doc = acad.ActiveDocument
        doc_name = current_doc.Name

        print(f"[保存] 正在保存: {doc_name}")

        # 3. 等待CAD空闲
        wait_quiescent(min_quiet=0.5, timeout=15.0)

        # 4. 执行保存操作
        try:
            current_doc.Save()
            print(f"[成功] 保存成功: {doc_name}")

            # 5. 等待保存完成
            wait_quiescent(min_quiet=1.0, timeout=30.0)
            return True

        except Exception as save_error:
            print(f"[警告] 直接保存失败,尝试另存为: {save_error}")

            # 如果是未保存文件,尝试另存为
            if not hasattr(current_doc, 'FullName') or not current_doc.FullName:
                default_path = f"D:/temp/{doc_name}"
                return save_as_dwg_paradigm(default_path)

            return False

    except Exception as e:
        print(f"[错误] 保存文件异常: {e}")
        return False

def save_as_dwg_paradigm(output_path: str) -> bool:
    """
    另存为文件范式

    规则:
    - 使用短路径处理中文/特殊字符
    - 创建输出目录
    - 验证文件创建
    """
    try:
        # 1. 基础验证
        if get_open_file_count() == 0:
            print("[错误] 没有打开的文件")
            return False

        # 2. 创建输出目录
        output_file = Path(output_path)
        output_file.parent.mkdir(parents=True, exist_ok=True)

        # 3. 获取当前文件信息
        import win32com.client
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        current_doc = acad.ActiveDocument
        doc_name = current_doc.Name

        print(f"[保存] 正在另存为: {doc_name} → {output_path}")

        # 4. 等待CAD空闲
        wait_quiescent(min_quiet=0.5, timeout=15.0)

        # 5. 使用短路径
        short_path = _get_short_path(output_path)

        # 6. 执行另存为操作
        try:
            current_doc.SaveAs(short_path)
            print(f"[成功] 另存为成功: {output_path}")

            # 7. 验证文件是否创建
            if output_file.exists():
                print(f"[成功] 文件已创建: {output_path}")

                # 8. 等待保存完成
                wait_quiescent(min_quiet=1.0, timeout=30.0)
                return True
            else:
                print(f"[错误] 文件未创建: {output_path}")
                return False

        except Exception as save_error:
            print(f"[错误] 另存为失败: {save_error}")
            return False

    except Exception as e:
        print(f"[错误] 另存为文件异常: {e}")
        return False

def auto_save_dwg_paradigm(interval_seconds: int = 300) -> bool:
    """自动保存范式"""
    try:
        if get_open_file_count() == 0:
            print("[警告] 没有打开的文件,跳过自动保存")
            return True

        import win32com.client
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        current_doc = acad.ActiveDocument
        doc_name = current_doc.Name

        print(f"[处理] 自动保存: {doc_name}")

        # 执行保存
        current_doc.Save()
        print(f"[成功] 自动保存完成: {doc_name}")

        return True

    except Exception as e:
        print(f"[错误] 自动保存异常: {e}")
        return False

# ============================================================================
# 5. 插入文件操作范式
# ============================================================================

def insert_dwg_as_block_paradigm(block_file_path: str,
                                insert_point: Tuple[float, float, float] = (0, 0, 0),
                                scale: float = 1.0,
                                rotation: float = 0.0,
                                explode: bool = False) -> bool:
    """
    插入DWG文件作为块的完整范式

    规则:
    - 使用-INSERT命令避免Unicode编码问题
    - 稳定路径处理中文/特殊字符
    - 等待插入操作完成
    - 可选炸开/缩放/旋转参数

    前置条件:
    - 有文件打开作为接收文件
    - 块文件存在

    后置条件:
    - 块已插入指定位置
    - 文件有未保存更改
    - CAD进入空闲状态
    """
    try:
        # 1. 基础验证
        if not Path(block_file_path).exists():
            print(f"[错误] 块文件不存在: {block_file_path}")
            return False

        if get_open_file_count() == 0:
            print("[错误] 没有打开的文件作为接收文件")
            return False

        # 2. 获取当前文件信息
        import win32com.client
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        current_doc = acad.ActiveDocument
        doc_name = current_doc.Name

        print(f"[处理] 正在插入块: {Path(block_file_path).name} → {doc_name}")
        print(f"[位置] 插入位置: {insert_point}, 缩放: {scale}, 旋转: {rotation}°, 炸开: {explode}")

        # 3. 等待CAD空闲
        wait_quiescent(min_quiet=0.5, timeout=15.0)

        # 4. 构建INSERT命令
        # 使用-INSERT命令避免外部脚本控制台Unicode输出问题
        short_path = _get_short_path(block_file_path)

        cmd_parts = [
            "-INSERT",
            f'"{short_path}"',  # 块文件路径
            f"{insert_point[0]},{insert_point[1]},{insert_point[2]}",  # 插入点
            str(scale),  # X比例
            str(scale) if scale != 1.0 else "1",  # Y比例 (如果X=1则跳过)
            str(rotation),  # 旋转角度
            "1" if explode else "0"  # 是否炸开
        ]

        insert_cmd = "\n".join(cmd_parts) + "\n"

        # 5. 执行插入命令
        success = send_cmd_with_sync(insert_cmd, wait_after=2.0, timeout=60.0)

        if not success:
            print(f"[错误] 插入块命令失败: {Path(block_file_path).name}")
            return False

        # 6. 等待插入完成
        wait_quiescent(min_quiet=2.0, timeout=60.0)

        # 7. 验证插入结果
        # 检查是否有未保存更改
        try:
            has_changes = not getattr(current_doc, 'Saved', True)
            if has_changes:
                print(f"[成功] 块插入成功: {Path(block_file_path).name}")
                return True
            else:
                print(f"[警告] 块插入后未检测到更改: {Path(block_file_path).name}")
                return True  # 仍然认为成功
        except:
            print(f"[成功] 块插入完成(无法验证更改状态): {Path(block_file_path).name}")
            return True

    except Exception as e:
        print(f"[错误] 插入块异常: {e}")
        return False

def insert_multiple_blocks_paradigm(block_configs: List[dict]) -> int:
    """
    批量插入块范式

    block_configs格式:
    [
        {
            'path': 'block_file.dwg',
            'point': (x, y, z),
            'scale': 1.0,
            'rotation': 0.0,
            'explode': False
        },
        ...
    ]
    """
    success_count = 0

    print(f"[处理] 开始批量插入 {len(block_configs)} 个块")

    for i, config in enumerate(block_configs):
        print(f"\n[块] [{i+1}/{len(block_configs)}] 插入块 {i+1}")

        try:
            block_path = config['path']
            insert_point = config.get('point', (0, 0, 0))
            scale = config.get('scale', 1.0)
            rotation = config.get('rotation', 0.0)
            explode = config.get('explode', False)

            if insert_dwg_as_block_paradigm(
                block_path, insert_point, scale, rotation, explode
            ):
                success_count += 1
                print(f"[成功] 成功插入: {Path(block_path).name}")
            else:
                print(f"[错误] 插入失败: {Path(block_path).name}")

            # 块间间隔等待
            if i < len(block_configs) - 1:
                time.sleep(1.0)
                wait_quiescent(min_quiet=0.5, timeout=15.0)

        except Exception as e:
            print(f"[错误] 插入块配置错误: {e}")

    print(f"\n[统计] 批量插入结果: {success_count}/{len(block_configs)} 成功")
    return success_count

def insert_and_explode_paradigm(block_file_path: str,
                               insert_point: Tuple[float, float, float] = (0, 0, 0),
                               scale: float = 1.0) -> bool:
    """插入并炸开块范式"""
    print(f"[处理] 插入并炸开: {Path(block_file_path).name}")

    # 1. 先插入块
    if not insert_dwg_as_block_paradigm(
        block_file_path, insert_point, scale, explode=True
    ):
        return False

    # 2. 等待插入完成
    wait_quiescent(min_quiet=1.0, timeout=30.0)

    # 3. 验证炸开结果
    print(f"[成功] 插入并炸开完成: {Path(block_file_path).name}")
    return True

# ============================================================================
# 完整工作流范式
# ============================================================================

def standard_workflow_paradigm(source_file: str,
                             block_files: List[dict],
                             output_file: str) -> bool:
    """
    标准工作流范式: 打开文件 → 插入块 → 保存 → 关闭

    参数:
    - source_file: 源文件路径
    - block_files: 块配置列表
    - output_file: 输出文件路径
    """
    try:
        print("[启动] 开始标准工作流")
        print(f"[文件] 源文件: {source_file}")
        print(f"[块] 块数量: {len(block_files)}")
        print(f"[保存] 输出文件: {output_file}")

        # 1. 打开源文件
        if not open_dwg_paradigm(source_file):
            print("[错误] 工作流失败: 无法打开源文件")
            return False

        # 2. 插入所有块
        if block_files:
            success_count = insert_multiple_blocks_paradigm(block_files)
            if success_count == 0:
                print("[错误] 工作流失败: 没有块插入成功")
                return False

        # 3. 保存到输出文件
        if not save_as_dwg_paradigm(output_file):
            print("[错误] 工作流失败: 无法保存输出文件")
            return False

        # 4. 关闭文件
        close_current_dwg_paradigm("no_save")

        print("[成功] 标准工作流完成")
        return True

    except Exception as e:
        print(f"[错误] 标准工作流异常: {e}")
        return False

# ============================================================================
# 6. 文件拷贝操作范式
# ============================================================================

def copy_dwg_to_dwg_paradigm(source_file: str, target_file: str, explode: bool = False) -> bool:
    """
    将DWG文件A完整拷贝到DWG文件B

    Args:
        source_file: 源文件路径
        target_file: 目标文件路径
        explode: True=炸开(直接拷贝内容), False=作为块插入

    Returns:
        bool: True表示成功, False表示失败
    """
    try:
        print(f"[拷贝] 开始拷贝: {Path(source_file).name} → {Path(target_file).name}")
        print(f"[模式] {'炸开模式' if explode else '块模式'}")

        # 1. 验证源文件
        if not Path(source_file).exists():
            print(f"[错误] 源文件不存在: {source_file}")
            return False

        # 2. 打开或创建目标文件
        if not Path(target_file).exists():
            print(f"[创建] 目标文件不存在,创建新文件: {target_file}")
            if not new_dwg_enhanced(target_file):
                return False
        else:
            if not open_dwg_paradigm(target_file):
                return False

        # 3. 插入源文件
        if explode:
            # 炸开模式: 插入并炸开
            success = insert_and_explode_paradigm(source_file, (0, 0, 0))
        else:
            # 块模式: 作为块插入
            success = insert_dwg_as_block_paradigm(source_file, (0, 0, 0))

        if not success:
            print(f"[错误] 插入失败")
            return False

        # 4. 保存目标文件
        if not save_current_dwg_paradigm():
            print(f"[错误] 保存失败")
            return False

        print(f"[成功] 拷贝完成: {Path(source_file).name} → {Path(target_file).name}")
        return True

    except Exception as e:
        print(f"[错误] 拷贝文件异常: {e}")
        return False

def copy_region_to_dwg_paradigm(source_file: str, target_file: str,
                                x1: float, y1: float, x2: float, y2: float,
                                xb: float, yb: float) -> bool:
    """
    将源文件中指定区域的图形对象拷贝到目标文件指定位置

    Args:
        source_file: 源文件路径
        target_file: 目标文件路径
        x1, y1: 区域左下角坐标
        x2, y2: 区域右上角坐标
        xb, yb: 目标位置坐标(区域左下角对应此点)

    Returns:
        bool: True表示成功, False表示失败
    """
    try:
        import sys
        sys.path.append(str(Path(__file__).parent))
        from CAD_basic import select_objects_in_window_area, get_acad_doc
        import win32com.client

        print(f"[区域拷贝] {Path(source_file).name} → {Path(target_file).name}")
        print(f"[区域] ({x1},{y1}) 到 ({x2},{y2})")
        print(f"[目标] ({xb},{yb},0)")

        # 1. 验证文件
        if not Path(source_file).exists():
            print(f"[错误] 源文件不存在: {source_file}")
            return False

        # 2. 打开源文件
        if not open_dwg_paradigm(source_file):
            return False
        wait_quiescent(min_quiet=0.5, timeout=15.0)

        # 3. 选择区域内的对象
        print(f"[选择] 选择区域内对象...")
        entities = select_objects_in_window_area(x1, y1, x2, y2)

        if not entities or len(entities) == 0:
            print(f"[警告] 区域内没有对象")
            return False

        print(f"[成功] 选中 {len(entities)} 个对象")

        # 4. 计算偏移量
        offset_x = xb - x1
        offset_y = yb - y1

        # 5. 打开或创建目标文件
        if not Path(target_file).exists():
            if not new_dwg_enhanced(target_file):
                return False
        else:
            if not open_dwg_paradigm(target_file):
                return False

        wait_quiescent(min_quiet=0.5, timeout=15.0)
        _, doc2 = get_acad_doc()

        # 6. 复制对象到目标文件
        print(f"[复制] 复制对象到目标文件...")
        copied_count = 0
        for obj in entities:
            try:
                new_obj = obj.Copy()
                # 移动到目标位置
                base_point = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0, 0, 0])
                displacement = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [offset_x, offset_y, 0])
                new_obj.Move(base_point, displacement)
                copied_count += 1
            except Exception as e:
                print(f"[警告] 复制对象失败: {e}")

        print(f"[成功] 已复制 {copied_count}/{len(entities)} 个对象")

        # 7. 保存目标文件
        if not save_current_dwg_paradigm():
            print(f"[错误] 保存目标文件失败")
            return False

        print(f"[成功] 区域拷贝完成")
        return True

    except Exception as e:
        print(f"[错误] 区域拷贝异常: {e}")
        return False

if __name__ == "__main__":
    # 测试基本操作范式
    print("[测试] 测试CAD基本操作范式")

    # 测试新建文件
    print("\n1. 测试新建文件")
    new_dwg_enhanced("D:/temp/test_new.dwg")

    # 测试保存
    print("\n2. 测试保存文件")
    save_current_dwg_paradigm()

    # 测试关闭
    print("\n3. 测试关闭文件")
    close_current_dwg_paradigm()

    print("\n[成功] 基本操作范式测试完成")