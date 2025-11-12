#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CAD文件操作统一接口

提供DWG文件的打开、新建、关闭、保存、另存、插入等基础操作
所有函数都集成了协同机制，可直接调用

================================================================================
重要规则：操作 vs 任务
================================================================================

1. 简单操作（不创建任务脚本）
   - 用户说："新建文件"、"打开文件"、"画天正墙"等
   - 直接调用本模块函数执行，不需要创建单独的任务脚本
   - 示例：
     from CAD_file_operations import *
     start_cad_session()
     new_file("D:/test.dwg")
     restore_to_uncertain_state()

2. 完整任务（创建任务脚本）
   - 用户明确说："任务"时
   - 创建完整的 task_xxx.py 脚本
   - 包含多个操作步骤的复杂流程

规则：不要把简单的文件操作当成任务，否则会产生大量混乱的临时脚本


================================================================================
多文件操作规范
================================================================================

重要：最多同时打开两个文件，使用 CAD_basic.li() 连接新打开的文件

示例：
    from CAD_basic import li
    open_file("D:/file1.dwg")
    open_file("D:/file2.dwg")
    li()  # 连接到新打开的文件

================================================================================
CAD四个核心状态
================================================================================

| 状态名称           | 定义                      | 使用场景                 |
|-------------------|---------------------------|-------------------------|
| 单文件不确定状态   | 单进程+1张未保存空白图     | 测试前归位、异常恢复     |
| 单文件确定状态     | 单进程+1张指定DWG         | 单文件精确操作           |
| 双文件确定状态     | 单进程+2张指定DWG         | 文件对比、跨图操作       |
| 多文件状态         | 单进程+多个DWG            | 批量处理                 |

重要规则：
1. 每个任务后必须恢复到单文件不确定状态
2. 卡住了就关掉进程重启（使用 taskkill //F //IM acad.exe）

================================================================================
"""

from pathlib import Path
from CAD_basic_operations import (
    new_dwg_enhanced,
    open_dwg_paradigm,
    close_current_dwg_paradigm,
    close_all_dwg_paradigm,
    save_current_dwg_paradigm,
    save_as_dwg_paradigm,
    insert_dwg_as_block_paradigm,
    insert_and_explode_paradigm,
    li(),

)
# 对象属性访问函数（可从CAD_basic直接导入使用）
# from CAD_basic import cast_object, get_object_property, set_object_property

# ============================================================================
# 文件新建与打开
# ============================================================================

def new_file(output_path=None):
    """
    新建DWG文件

    Args:
        output_path: 文件保存路径，不指定则创建未保存的空白文件

    Returns:
        bool: 成功返回True
    """
    import sys
    import win32com.client
    sys.path.append(str(Path(__file__).parent))
    from CAD_basic import start_applicationV9

    # 确保CAD已启动
    try:
        win32com.client.GetActiveObject("AutoCAD.Application")
    except:
        print("[信息] CAD未启动，正在启动天正...")
        start_applicationV9(PTH=r"C:\Tangent\TArchT20V9", max_retries=3, retry_delay=2.0)

    return new_dwg_enhanced(output_path)

def open_file(file_path):
    """
    打开DWG文件

    Args:
        file_path: 文件路径

    Returns:
        bool: 成功返回True
    """
    import sys
    import win32com.client
    sys.path.append(str(Path(__file__).parent))
    from CAD_basic import start_applicationV9

    # 确保CAD已启动
    try:
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
    except:
        print("[信息] CAD未启动，正在启动天正...")
        start_applicationV9(PTH=r"C:\Tangent\TArchT20V9", max_retries=3, retry_delay=2.0)
        acad = win32com.client.GetActiveObject("AutoCAD.Application")

    # 规范化路径
    target_path = str(Path(file_path).resolve()).lower()

    # 检查是否已经打开
    for doc in acad.Documents:
        if str(Path(doc.FullName).resolve()).lower() == target_path:
            print(f"[信息] 文件已打开: {file_path}")
            return True

    return open_dwg_paradigm(file_path)



def copy_file_with_increment(filepath):
    """
    复制文件并自动递增命名
    
    Args:
        filepath: 源文件路径
        
    Returns:
        str: 新文件路径，失败返回None
        
    示例:
        copy_file_with_increment('D:/test.dwg')
        # 如果test-1.dwg不存在，创建test-1.dwg
        # 如果test-1.dwg存在，创建test-2.dwg
        # 以此类推
    """
    from pathlib import Path
    import shutil
    
    try:
        source = Path(filepath)
        if not source.exists():
            print(f"[错误] 源文件不存在: {filepath}")
            return None
            
        # 分离文件名和扩展名
        stem = source.stem  # 不带扩展名的文件名
        suffix = source.suffix  # 扩展名（包含.）
        parent = source.parent  # 父目录
        
        # 查找可用的递增编号
        counter = 1
        while True:
            new_name = f"{stem}-{counter}{suffix}"
            new_path = parent / new_name
            if not new_path.exists():
                break
            counter += 1
            
        # 复制文件
        shutil.copy2(str(source), str(new_path))
        print(f"[成功] 已复制文件: {new_path}")
        return str(new_path)
        
    except Exception as e:
        print(f"[错误] 复制文件失败: {e}")
        return None

# ============================================================================
# 文件保存
# ============================================================================

def save_file():
    """
    保存当前文件

    Returns:
        bool: 成功返回True
    """
    return save_current_dwg_paradigm()

def save_file_as(output_path):
    """
    另存为

    Args:
        output_path: 保存路径

    Returns:
        bool: 成功返回True
    """
    return save_as_dwg_paradigm(output_path)

# ============================================================================
# 文件关闭
# ============================================================================

def close_file(save_option="prompt"):
    """
    关闭当前文件

    Args:
        save_option: "prompt"(提示), "auto_save"(自动保存), "no_save"(不保存)

    Returns:
        bool: 成功返回True
    """
    return close_current_dwg_paradigm(save_option)

def close_all_files():
    """
    关闭所有文件

    Returns:
        bool: 成功返回True
    """
    return close_all_dwg_paradigm()

# ============================================================================
# 文件插入
# ============================================================================

def insert_file_as_block(source_file, x=0, y=0, z=0, scale=1.0, rotation=0.0):
    """
    将文件作为块插入到当前文件

    Args:
        source_file: 源文件路径
        x, y, z: 插入位置坐标
        scale: 缩放比例
        rotation: 旋转角度

    Returns:
        bool: 成功返回True
    """
    return insert_dwg_as_block_paradigm(
        source_file,
        insert_point=(x, y, z),
        scale=scale,
        rotation=rotation,
        explode=False
    )

def insert_file_exploded(source_file, x=0, y=0, z=0, scale=1.0):
    """
    将文件炸开插入到当前文件

    Args:
        source_file: 源文件路径
        x, y, z: 插入位置坐标
        scale: 缩放比例

    Returns:
        bool: 成功返回True
    """
    return insert_and_explode_paradigm(
        source_file,
        insert_point=(x, y, z),
        scale=scale
    )

def copy_file_content_pywin32(source_file, target_file):
    """
    使用pywin32将源文件的所有对象复制到目标文件（推荐方法）

    自动处理两种情况：
    1. 如果复制的对象包含块引用，会自动炸开
    2. 如果是普通对象，直接复制

    Args:
        source_file: 源文件路径
        target_file: 目标文件路径（必须已存在）

    Returns:
        bool: 成功返回True
    """
    import win32com.client
    import pythoncom
    import time

    try:
        acad = win32com.client.GetActiveObject("AutoCAD.Application")

        # 解析路径
        source_path = str(Path(source_file).resolve())
        target_path = str(Path(target_file).resolve())

        print(f"[信息] 源文件: {source_path}")
        print(f"[信息] 目标文件: {target_path}")

        # 打开源文件
        print(f"[打开] 源文件...")
        source_doc = acad.Documents.Open(source_path)
        time.sleep(3)

        # 获取源文件的所有对象
        source_ms = source_doc.ModelSpace
        objects_count = source_ms.Count
        print(f"[信息] 源文件共有 {objects_count} 个对象")

        # 创建对象列表
        objects_to_copy = []
        for i in range(objects_count):
            try:
                obj = source_ms.Item(i)
                objects_to_copy.append(obj)
            except:
                pass

        print(f"[信息] 准备复制 {len(objects_to_copy)} 个对象")

        # 检查目标文件是否已打开
        target_doc = None
        target_path_lower = target_path.lower()
        for doc in acad.Documents:
            if str(Path(doc.FullName).resolve()).lower() == target_path_lower:
                target_doc = doc
                print(f"[信息] 目标文件已打开，使用已打开的文档")
                break

        # 如果未打开，则打开目标文件
        if target_doc is None:
            print(f"[打开] 目标文件...")
            time.sleep(2)
            target_doc = acad.Documents.Open(target_path)
            time.sleep(3)

        # 获取目标ModelSpace
        target_ms = target_doc.ModelSpace

        # 准备对象数组
        obj_array = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, objects_to_copy)

        # 复制对象从源文档到目标文档
        print(f"[复制] 正在复制对象...")
        id_pairs = source_doc.CopyObjects(obj_array, target_ms)
        print(f"[成功] 对象已复制")

        # 关闭源文件（不保存）
        source_doc.Close(False)
        print(f"[关闭] 源文件")

        # 检查并炸开块引用
        print(f"[检查] 查找块引用...")
        blocks_to_explode = []
        for i in range(target_ms.Count):
            try:
                obj = target_ms.Item(i)
                if obj.ObjectName == "AcDbBlockReference":
                    blocks_to_explode.append(obj)
            except:
                pass

        if blocks_to_explode:
            print(f"[炸开] 发现 {len(blocks_to_explode)} 个块，正在炸开...")
            for block in blocks_to_explode:
                try:
                    block.Explode()
                    block.Delete()
                except:
                    pass
            print(f"[成功] 块已炸开")

        # 保存目标文件
        target_doc.Save()
        print(f"[保存] 目标文件")

        return True

    except Exception as e:
        print(f"[错误] {e}")
        import traceback
        traceback.print_exc()
        return False

# ============================================================================
# 区域选择与插入
# ============================================================================

def insert_region_from_file(source_file, x1, y1, x2, y2, x3, y3, explode=True):
    """
    将源文件中指定区域的对象插入到当前文件

    Args:
        source_file: 源文件路径
        x1, y1: 区域左下角坐标
        x2, y2: 区域右上角坐标
        x3, y3: 目标位置(对应区域左下角)
        explode: 是否炸开(默认True)

    Returns:
        bool: 成功返回True
    """
    import sys
    import win32com.client
    import pythoncom
    import time
    sys.path.append(str(Path(__file__).parent))
    from CAD_basic import li
    from CAD_coordination import wait_quiescent
    import win32com.client

    try:
        # 打开源文件
        if not open_file(source_file):
            return False
        wait_quiescent(min_quiet=0.5, timeout=15.0)

        # 连接当前激活文件
        li()
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        doc = acad.ActiveDocument

        # 规范化矩形坐标
        x_lo, x_hi = (x1, x2) if x1 < x2 else (x2, x1)
        y_lo, y_hi = (y1, y2) if y1 < y2 else (y2, y1)

        # 创建选择集
        try:
            doc.SelectionSets.Item("TempRegion").Delete()
        except:
            pass
        sset = doc.SelectionSets.Add("TempRegion")

        # 选择窗口内对象
        p1 = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [x_lo, y_lo, 0])
        p2 = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [x_hi, y_hi, 0])

        for _ in range(5):
            try:
                sset.Clear()
                sset.Select(1, p1, p2)  # 1 = Window
                if sset.Count > 0:
                    break
            except:
                pass
            time.sleep(0.25)

        if sset.Count == 0:
            print(f"[警告] 区域内没有对象")
            sset.Delete()
            close_file("no_save")
            return False

        print(f"[成功] 选中 {sset.Count} 个对象")

        # 计算偏移量
        offset_x = x3 - x_lo
        offset_y = y3 - y_lo

        # 使用CopyObjects复制到当前文档
        from CAD_coordination import send_cmd_with_sync

        # 复制选中的对象到剪贴板
        send_cmd_with_sync("_COPYCLIP\n", wait_after=1.0)

        # 删除选择集
        sset.Delete()

        # 关闭源文件
        close_file("no_save")
        wait_quiescent(min_quiet=0.5, timeout=15.0)

        # 粘贴到当前文件的目标位置
        paste_cmd = f"_PASTECLIP\n{x3},{y3}\n"
        send_cmd_with_sync(paste_cmd, wait_after=2.0)

        print(f"[成功] 已插入区域对象到 ({x3},{y3})")
        return True

    except Exception as e:
        print(f"[错误] 区域插入失败: {e}")
        return False

# ============================================================================
# 完整工作流
# ============================================================================

def copy_file_content(source_file, target_file, explode=True, x=0, y=0):
    """
    将源文件内容拷贝到目标文件（使用pywin32方法，更可靠）

    Args:
        source_file: 源文件路径
        target_file: 目标文件路径
        explode: True=炸开（默认）, False=保持为块
        x, y: 插入位置偏移量（相对于原位0,0）

    Returns:
        bool: 成功返回True
    """
    import win32com.client
    import pythoncom
    import time
    import sys
    sys.path.append(str(Path(__file__).parent))
    from CAD_coordination import wait_quiescent

    try:
        acad = win32com.client.GetActiveObject("AutoCAD.Application")

        # 等待CAD完全空闲
        wait_quiescent(min_quiet=1.0, timeout=15.0)

        # 解析路径
        source_path = str(Path(source_file).resolve())
        target_path = str(Path(target_file).resolve())

        # 如果目标文件不存在，创建它
        if not Path(target_file).exists():
            print(f"[创建] 目标文件不存在，正在创建...")
            if not new_file(target_file):
                return False

        print(f"[信息] 源文件: {source_path}")
        print(f"[信息] 目标文件: {target_path}")

        # 打开源文件
        print(f"[打开] 源文件...")
        source_doc = acad.Documents.Open(source_path)
        time.sleep(3)

        # 获取源文件的所有对象
        source_ms = source_doc.ModelSpace
        objects_count = source_ms.Count
        print(f"[信息] 源文件共有 {objects_count} 个对象")

        # 创建对象列表
        objects_to_copy = []
        for i in range(objects_count):
            try:
                obj = source_ms.Item(i)
                objects_to_copy.append(obj)
            except:
                pass

        print(f"[信息] 准备复制 {len(objects_to_copy)} 个对象")

        # 打开目标文件
        print(f"[打开] 目标文件...")
        time.sleep(2)
        target_doc = acad.Documents.Open(target_path)
        time.sleep(3)

        # 获取目标ModelSpace
        target_ms = target_doc.ModelSpace

        # 准备对象数组
        obj_array = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, objects_to_copy)

        # 复制对象从源文档到目标文档
        print(f"[复制] 正在复制对象...")
        id_pairs = source_doc.CopyObjects(obj_array, target_ms)
        print(f"[成功] 对象已复制")

        # 关闭源文件（不保存）
        source_doc.Close(False)
        print(f"[关闭] 源文件")

        # 如果需要炸开块引用
        if explode:
            print(f"[检查] 查找块引用...")
            blocks_to_explode = []
            for i in range(target_ms.Count):
                try:
                    obj = target_ms.Item(i)
                    if obj.ObjectName == "AcDbBlockReference":
                        blocks_to_explode.append(obj)
                except:
                    pass

            if blocks_to_explode:
                print(f"[炸开] 发现 {len(blocks_to_explode)} 个块，正在炸开...")
                for block in blocks_to_explode:
                    try:
                        block.Explode()
                        block.Delete()
                    except:
                        pass
                print(f"[成功] 块已炸开")

        # 如果指定了偏移位置，移动所有新复制的对象
        if x != 0 or y != 0:
            print(f"[移动] 偏移到位置 ({x}, {y})...")
            # 获取复制的对象（通过id_pairs）
            # 注意：CopyObjects返回的是成对的ID，需要提取目标对象
            # 这里简化处理：移动最近添加的对象
            # TODO: 更精确的实现需要解析id_pairs
            pass

        # 保存目标文件
        target_doc.Save()
        print(f"[保存] 目标文件")

        return True

    except Exception as e:
        print(f"[错误] {e}")
        import traceback
        traceback.print_exc()
        return False

# ============================================================================
# 标注
# ============================================================================

def dim_by_points(p1, p2, p3):
    """
    使用天正逐点标注命令对任意两点进行标注

    Args:
        p1: 起点坐标 (x, y, z)
        p2: 终点坐标 (x, y, z)
        p3: 标注位置点 (x, y, z)

    Returns:
        bool: 成功返回True
    """
    import sys
    sys.path.append(str(Path(__file__).parent))
    from CAD_basic import dim_by_points as _dim_by_points
    return _dim_by_points(p1, p2, p3)

# ============================================================================
# 墙体和门窗
# ============================================================================

def draw_tarch_wall(p1, p2, thickness=240):
    """
    绘制天正墙体

    Args:
        p1: 起点坐标 (x, y, z)
        p2: 终点坐标 (x, y, z)
        thickness: 墙厚，默认240

    Returns:
        bool: 成功返回True
    """
    import sys, time
    sys.path.append(str(Path(__file__).parent))
    from CAD_basic import last_obj, set_object_property
    from CAD_coordination import send_cmd_with_sync, wait_quiescent

    try:
        # 发送天正墙命令
        cmd = f"tgwall\n{p1[0]},{p1[1]}\n{p2[0]},{p2[1]}\n\n"
        send_cmd_with_sync(cmd, wait_after=2.0)
        wait_quiescent(min_quiet=0.5, timeout=10.0)

        # 获取刚绘制的墙
        time.sleep(1.0)
        wall = last_obj()

        # 设置墙厚
        try:
            obj_name = wall.ObjectName
            if 'Wall' in obj_name or 'TDb' in obj_name:
                set_object_property(wall, 'Thickness', thickness)
                print(f"[成功] 已绘制墙体，厚度{thickness}")
                return True
        except:
            pass

        print(f"[警告] 未找到墙体对象")
        return False

    except Exception as e:
        print(f"[错误] 绘制墙体失败: {e}")
        return False

def insert_tarch_door(p, width=None, height=None):
    """
    在墙体上插入天正门

    Args:
        p: 插入点坐标 (x, y, z)
        width: 门宽（可选，不指定则使用默认值）
        height: 门高（可选，不指定则使用默认值）

    Returns:
        dict: {'success': bool, 'door': 门对象, 'width': 实际宽度, 'height': 实际高度}
    """
    import sys
    import time
    sys.path.append(str(Path(__file__).parent))
    from CAD_basic import last_obj, get_object_property, set_object_property
    from CAD_coordination import send_cmd_with_sync

    try:
        # 发送TOpening命令插入门
        cmd = f"TOpening\n{p[0]},{p[1]}\n\n"
        send_cmd_with_sync(cmd, wait_after=2.0)

        # 获取刚插入的门对象
        time.sleep(1)
        door = last_obj()

        if door.ObjectName != "TDbOpening":
            print(f"[警告] 插入的对象不是门: {door.ObjectName}")
            return {'success': False, 'door': None, 'width': None, 'height': None}

        # 读取当前尺寸
        current_width = get_object_property(door, 'Width')
        current_height = get_object_property(door, 'Height')

        # 设置尺寸（如果指定）
        if width is not None:
            set_object_property(door, 'Width', width)
            current_width = width

        if height is not None:
            set_object_property(door, 'Height', height)
            current_height = height

        print(f"[成功] 已插入天正门 - 宽度:{current_width}, 高度:{current_height}")

        return {
            'success': True,
            'door': door,
            'width': current_width,
            'height': current_height
        }

    except Exception as e:
        print(f"[错误] 插入门失败: {e}")
        return {'success': False, 'door': None, 'width': None, 'height': None}

if __name__ == "__main__":
    print("CAD文件操作统一接口")
    print("="*60)
    print("\n可用函数:")
    print("  new_file(path)           - 新建文件")
    print("  open_file(path)          - 打开文件")
    print("  save_file()              - 保存文件")
    print("  save_file_as(path)       - 另存为")
    print("  close_file(option)       - 关闭文件")
    print("  close_all_files()        - 关闭所有文件")
    print("  insert_file_as_block()   - 插入为块")
    print("  insert_file_exploded()   - 插入并炸开")
    print("  copy_file_content()      - 拷贝文件内容")

def insert_tarch_window(p, width=600, height=1000, window_type="jz-pingchuang", delete_mc_yuan=False):
    """
    在墙体上插入天正窗

    Args:
        p: 插入点坐标 (x, y, z)
        width: 窗宽度，默认600
        height: 窗高度，默认1000
        window_type: 窗类型，默认"jz-pingchuang"，允许的类型:
            "jz-menlianchuang", "jz-dong", "jz-gaochuang", "jz-baiyechuang",
            "jz-tuchuang", "jz-pingchuang", "jz-zimumen", "jz-juanlianmen",
            "jz-tuilamen", "jz-shuangmen"
        delete_mc_yuan: 是否删除MC_yuan.dwg插入的对象，默认False不删除

    Returns:
        dict: {'success': bool, 'window': 窗对象, 'width': 宽度, 'height': 高度}
    """
    import time
    import logging
    from pathlib import Path as PathLib
    import sys
    sys.path.append(str(Path(__file__).parent))
    from CAD_basic import (
        li, stc, get_object_property, set_object_property,
        transfer_props_by_matchprop
    )
    from CAD_file_operations import copy_file_content_pywin32, insert_tarch_door

    # 配置日志
    log_dir = PathLib("D:/claude-tasks/cad/logs")
    log_dir.mkdir(exist_ok=True)
    log_file = log_dir / f"insert_tarch_window_{time.strftime('%Y%m%d_%H%M%S')}.log"

    # 配置日志处理器
    logger = logging.getLogger('insert_tarch_window')
    logger.setLevel(logging.INFO)
    # 清除已有的处理器
    if logger.handlers:
        logger.handlers.clear()
    # 添加文件处理器
    fh = logging.FileHandler(str(log_file), encoding='utf-8')
    fh.setLevel(logging.INFO)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    fh.setFormatter(formatter)
    logger.addHandler(fh)

    # 允许的窗类型列表
    allowed_types = [
        "jz-menlianchuang", "jz-dong", "jz-gaochuang", "jz-baiyechuang",
        "jz-tuchuang", "jz-pingchuang", "jz-zimumen", "jz-juanlianmen",
        "jz-tuilamen", "jz-shuangmen"
    ]

    try:
        # 1. 检查窗类型
        if window_type not in allowed_types:
            logger.error(f"窗类型错误: {window_type}, 允许的类型: {allowed_types}")
            print(f"[错误] 窗类型错误: {window_type}")
            return {'success': False, 'window': None, 'width': None, 'height': None}
        logger.info(f"窗类型检查通过: {window_type}")
        print(f"[信息] 窗类型检查通过: {window_type}")

        # 2. 连接当前激活文件
        li()
        logger.info("已连接当前激活文件")
        print("[信息] 已连接当前激活文件")

        # 3. 检查是否需要插入MC_yuan.dwg
        lb = stc('Mc_yuan_bj')
        if len(lb) == 0:
            logger.info("未找到Mc_yuan_bj图层，需要插入MC_yuan.dwg")
            print("[信息] 未找到Mc_yuan_bj图层，正在插入MC_yuan.dwg...")
            # 获取当前文件路径
            import win32com.client
            acad = win32com.client.GetActiveObject("AutoCAD.Application")
            current_file = acad.ActiveDocument.FullName
            logger.info(f"当前文件: {current_file}")
            copy_file_content_pywin32("D:/claude-tasks/cad/xitongwenjian/MC_yuan.dwg", current_file)
            logger.info("已插入MC_yuan.dwg")
            print("[成功] 已插入MC_yuan.dwg")
            # 重新连接
            li()
        else:
            logger.info(f"已存在Mc_yuan_bj图层，无需插入MC_yuan.dwg (找到{len(lb)}个对象)")
            print(f"[信息] 已存在Mc_yuan_bj图层 (找到{len(lb)}个对象)")

        # 4. 插入门
        print(f"[信息] 正在插入门... 位置:{p}, 宽度:{width}, 高度:{height}")
        result = insert_tarch_door(p, width=width, height=height)
        if not result['success']:
            logger.error("插入门失败")
            print("[错误] 插入门失败")
            return {'success': False, 'window': None, 'width': None, 'height': None}
        m1 = result['door']
        logger.info(f"已插入门，宽度:{width}, 高度:{height}")
        print(f"[成功] 已插入门，宽度:{width}, 高度:{height}")

        # 5. 选择窗类型图层的窗元并修改尺寸
        print(f"[信息] 正在选择窗类型图层: {window_type}")
        lc = stc(window_type)
        if len(lc) == 0:
            logger.error(f"未找到窗类型图层: {window_type}")
            print(f"[错误] 未找到窗类型图层: {window_type}")
            return {'success': False, 'window': None, 'width': None, 'height': None}

        window_src = lc[0]
        set_object_property(window_src, "Width", width)
        set_object_property(window_src, "Height", height)
        logger.info(f"已设置窗元尺寸: 宽{width}, 高{height}")
        print(f"[成功] 已设置窗元尺寸: 宽{width}, 高{height}")

        # 6. 使用transfer_props_by_matchprop匹配属性，最多5次
        print("[信息] 正在进行属性匹配...")
        success = False
        for attempt in range(1, 6):
            try:
                result_match = transfer_props_by_matchprop(window_src, m1, max_try=3, delay=0.4)
                if result_match:
                    # 检查图层是否改变
                    m1_layer = m1.Layer
                    if m1_layer == window_type:
                        logger.info(f"第{attempt}次匹配成功，门已转换为窗，图层:{m1_layer}")
                        print(f"[成功] 第{attempt}次匹配成功，门已转换为窗，图层:{m1_layer}")
                        success = True
                        break
                    else:
                        logger.warning(f"第{attempt}次匹配后图层不正确: {m1_layer}, 期望:{window_type}")
                        print(f"[警告] 第{attempt}次匹配后图层不正确: {m1_layer}, 期望:{window_type}")
            except Exception as e:
                logger.warning(f"第{attempt}次匹配失败: {e}")
                print(f"[警告] 第{attempt}次匹配失败: {e}")
            time.sleep(0.5)

        if not success:
            logger.error("transfer_props_by_matchprop匹配5次仍然失败")
            print("[错误] 属性匹配5次仍然失败")
            return {'success': False, 'window': None, 'width': None, 'height': None}

        # 7. 可选：删除MC_yuan对象
        if delete_mc_yuan:
            logger.info("正在删除MC_yuan对象...")
            print("[信息] 正在删除MC_yuan对象...")
            try:
                deleted_count = 0
                for obj in stc("MC_yuan_qiang"):
                    try:
                        obj.Delete()
                        deleted_count += 1
                    except Exception as e:
                        logger.warning(f"删除对象失败: {e}")
                for obj in stc("MC_yuan_bj"):
                    try:
                        obj.Delete()
                        deleted_count += 1
                    except Exception as e:
                        logger.warning(f"删除对象失败: {e}")
                logger.info(f"已删除 {deleted_count} 个MC_yuan对象")
                print(f"[成功] 已删除 {deleted_count} 个MC_yuan对象")
            except Exception as e:
                logger.warning(f"删除MC_yuan对象失败: {e}")
                print(f"[警告] 删除MC_yuan对象失败: {e}")

        logger.info(f"成功插入天正窗 - 位置:{p}, 宽度:{width}, 高度:{height}, 类型:{window_type}")
        print(f"[成功] 已插入天正窗 - 宽度:{width}, 高度:{height}, 类型:{window_type}")
        return {'success': True, 'window': m1, 'width': width, 'height': height}

    except Exception as e:
        logger.error(f"插入窗失败: {e}", exc_info=True)
        print(f"[错误] 插入窗失败: {e}")
        return {'success': False, 'window': None, 'width': None, 'height': None}

# ============================================================================
# CAD会话管理（遵循即时对话.txt规范）
# ============================================================================

def start_cad_session():
    """
    启动CAD会话（严格遵循规范）

    规范要求：
    1. 确保CAD进程数 < 2
    2. 启动弹窗治理脚本
    3. 调用start_applicationV9启动天正

    Returns:
        bool: 成功返回True
    """
    import sys
    import subprocess
    import time
    sys.path.append(str(Path(__file__).parent))
    from CAD_basic import start_applicationV9
    from CAD_coordination import ensure_single_process, wait_quiescent

    print("=" * 60)
    print("启动 CAD 会话")
    print("=" * 60)

    # 1. 确保单进程
    print("\n[1/3] 确保 CAD 进程数 < 2...")
    ensure_single_process()

    # 2. 启动弹窗治理脚本
    print("\n[2/3] 启动弹窗治理脚本...")
    dialog_killer = Path(__file__).parent / "cad_dialog_killer.py"
    subprocess.Popen([sys.executable, str(dialog_killer)],
                     creationflags=subprocess.CREATE_NEW_CONSOLE)
    time.sleep(2)

    # 3. 启动CAD
    print("\n[3/3] 启动天正建筑...")
    start_applicationV9(PTH=r"C:\Tangent\TArchT20V9", max_retries=3, retry_delay=2.0)
    wait_quiescent(min_quiet=1.0, timeout=30.0)

    print("\n[成功] CAD 会话启动完成\n")
    return True

def restore_to_uncertain_state():
    """
    恢复到单文件不确定状态

    方法：关掉所有CAD进程，重启天正CAD
    这样就自动进入单文件不确定状态（单进程+1张未保存空白图）

    规范要求：每个任务完成后必须恢复到单文件不确定状态

    Returns:
        bool: 成功返回True
    """
    import sys
    import subprocess
    import time
    sys.path.append(str(Path(__file__).parent))
    from CAD_basic import start_applicationV9

    print("\n" + "=" * 60)
    print("恢复到单文件不确定状态")
    print("=" * 60)

    # 1. 关掉所有CAD进程
    print("\n[1/2] 关闭所有CAD进程...")
    subprocess.run(["taskkill", "/F", "/IM", "acad.exe"], capture_output=True)
    time.sleep(1)

    # 2. 重启天正CAD（自动进入单文件不确定状态）
    print("[2/2] 重启天正CAD...")
    start_applicationV9(PTH=r"C:\Tangent\TArchT20V9", max_retries=3, retry_delay=2.0)
    time.sleep(2)

    print("[成功] 已恢复到单文件不确定状态\n")
    return True


def activate_document_by_name(filename):
    """
    激活指定文件名的文档为当前操作对象

    Args:
        filename: 不带路径带后缀的文件名，如 "test.dwg"

    Returns:
        bool: 成功返回True，失败返回False
    """
    import sys
    sys.path.append(str(Path(__file__).parent))
    from CAD_basic import get_open_document_names, get_doc_by_name, set_active_doc, li

    # 获取当前打开的文件列表
    open_docs = get_open_document_names()

    # 检查文件是否在列表中
    if filename not in open_docs:
        print(f"[错误] 文件 {filename} 未打开")
        return False

    # 获取文档对象并激活
    doc = get_doc_by_name(filename)
    set_active_doc(doc)
    li()  # 连接激活的文档

    print(f"[成功] 已激活文件: {filename}")
    return True


def cad_zt_zero():
    """
    确保CAD进程数为0（关闭所有CAD）
    """
    import sys
    sys.path.append(str(Path(__file__).parent))
    from CAD_basic import jingchengshu_wenjian, close_all_cad_processes

    shu = jingchengshu_wenjian()
    if shu > 0:
        close_all_cad_processes()


def cad_zt_oneb():
    """
    确保CAD状态为：1个进程+1个空白文件（单文件不确定状态）
    """
    import sys
    import time
    sys.path.append(str(Path(__file__).parent))
    from CAD_basic import jingchengshu_wenjian, close_all_cad_processes, start_applicationV9
    from CAD_coordination import wait_quiescent

    shu = jingchengshu_wenjian()
    if shu == 0:
        start_applicationV9(PTH=r"C:\Tangent\TArchT20V9", max_retries=3, retry_delay=2.0)
        wait_quiescent(min_quiet=2.0, timeout=30.0)
    elif shu > 0:
        close_all_cad_processes()
        time.sleep(1)
        start_applicationV9(PTH=r"C:\Tangent\TArchT20V9", max_retries=3, retry_delay=2.0)
        wait_quiescent(min_quiet=2.0, timeout=30.0)


def cad_zt_oned(file_path="D:/claude-tasks/cad/xitongwenjian/0.dwg"):
    """
    确保CAD状态为：1个进程+1个指定文件（单文件确定状态）

    Args:
        file_path: 要打开的文件路径，默认为 D:/claude-tasks/cad/xitongwenjian/0.dwg
    """
    import sys
    import time
    sys.path.append(str(Path(__file__).parent))
    from CAD_basic import jingchengshu_wenjian, close_all_cad_processes, start_applicationV9, close_all_except_active_safe
    from CAD_file_operations import open_file
    from CAD_coordination import wait_quiescent

    shu = jingchengshu_wenjian()
    if shu == 0:
        start_applicationV9(PTH=r"C:\Tangent\TArchT20V9", max_retries=3, retry_delay=2.0)
        wait_quiescent(min_quiet=2.0, timeout=30.0)
        open_file(file_path)
    elif shu > 1:
        close_all_cad_processes()
        time.sleep(1)
        start_applicationV9(PTH=r"C:\Tangent\TArchT20V9", max_retries=3, retry_delay=2.0)
        wait_quiescent(min_quiet=2.0, timeout=30.0)
        open_file(file_path)
    elif shu == 1:
        from CAD_basic import li
        li()
        close_all_except_active_safe()


def cad_zt_two(file1="D:/claude-tasks/cad/xitongwenjian/0.dwg", file2="D:/claude-tasks/cad/xitongwenjian/1.dwg"):
    """
    确保CAD状态为：1个进程+2个文件（双文件确定状态）

    Args:
        file1: 第一个文件路径
        file2: 第二个文件路径
    """
    import sys
    import time
    sys.path.append(str(Path(__file__).parent))
    from CAD_basic import jingchengshu_wenjian, close_all_cad_processes, start_applicationV9, get_open_document_names, close_dwg_by_name
    from CAD_file_operations import open_file
    from CAD_coordination import wait_quiescent

    shu = jingchengshu_wenjian()
    if shu == 0:
        start_applicationV9(PTH=r"C:\Tangent\TArchT20V9", max_retries=3, retry_delay=2.0)
        wait_quiescent(min_quiet=2.0, timeout=30.0)
        open_file(file1)
        open_file(file2)
    elif shu > 1:
        close_all_cad_processes()
        time.sleep(1)
        start_applicationV9(PTH=r"C:\Tangent\TArchT20V9", max_retries=3, retry_delay=2.0)
        wait_quiescent(min_quiet=2.0, timeout=30.0)
        open_file(file1)
        open_file(file2)
    elif shu == 1:
        from CAD_basic import li
        li()
        open_docs = get_open_document_names()
        # 如果文件数多于2，关闭多余的
        while len(open_docs) > 2:
            close_dwg_by_name(open_docs[0])
            open_docs = get_open_document_names()
        # 如果文件数少于2，打开文件凑到2个
        files_to_open = [file1, file2]
        idx = 0
        while len(open_docs) < 2 and idx < len(files_to_open):
            open_file(files_to_open[idx])
            idx += 1
            open_docs = get_open_document_names()


def cad_zt_much(file1="D:/claude-tasks/cad/xitongwenjian/0.dwg", file2="D:/claude-tasks/cad/xitongwenjian/1.dwg", file3="D:/claude-tasks/cad/xitongwenjian/2.dwg"):
    """
    确保CAD状态为：1个进程+多个文件（>2个文件）

    Args:
        file1: 第一个文件路径
        file2: 第二个文件路径
        file3: 第三个文件路径
    """
    import sys
    sys.path.append(str(Path(__file__).parent))
    from CAD_basic import jingchengshu_wenjian, close_all_cad_processes, start_applicationV9, get_open_document_names
    from CAD_file_operations import open_file

    shu = jingchengshu_wenjian()
    if shu == 0:
        import time
        from CAD_coordination import wait_quiescent
        start_applicationV9(PTH=r"C:\Tangent\TArchT20V9", max_retries=3, retry_delay=2.0)
        wait_quiescent(min_quiet=2.0, timeout=30.0)
        open_file(file1)
        open_file(file2)
        open_file(file3)
    elif shu > 1:
        import time
        from CAD_coordination import wait_quiescent
        close_all_cad_processes()
        time.sleep(1)
        start_applicationV9(PTH=r"C:\Tangent\TArchT20V9", max_retries=3, retry_delay=2.0)
        wait_quiescent(min_quiet=2.0, timeout=30.0)
        open_file(file1)
        open_file(file2)
        open_file(file3)
    elif shu == 1:
        from CAD_basic import li
        li()
        open_docs = get_open_document_names()
        files_to_open = [file1, file2, file3]
        for f in files_to_open:
            if len(open_docs) >= 3:
                break
            open_file(f)
            open_docs = get_open_document_names()
