#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
天正对象属性访问模块

天正对象（TDbOpening, TDbWall等）无法通过标准COM接口访问专有属性
需要使用IDispatch.Invoke方法通过DISPID访问属性

使用方法:
    from CAD_tarch_properties import get_tarch_property, set_tarch_property

    # 获取门窗宽度
    width = get_tarch_property(door_obj, 'Width')

    # 设置墙体厚度
    set_tarch_property(wall_obj, 'Thickness', 240)
"""

import pythoncom

# 天正对象属性DISPID映射
TARCH_PROPERTY_MAP = {
    'TDbOpening': {  # 天正门窗
        'Offset': 1,     # 偏移
        'Width': 2,      # 宽度
        'Type': 3,       # 类型（如"普通窗"）
        'Height': 10,    # 高度
        'Name': 11,      # 名称
        'Direction': 7,  # 开启方向
        'Angle': 8,      # 角度
    },
    'TDbWall': {  # 天正墙
        'Offset1': 1,    # 偏移1
        'Thickness': 2,  # 厚度
        'Thickness2': 3, # 厚度2
        'Length': 4,     # 长度
        'WallType': 11,  # 墙类型（如"直墙"）
        'Material': 13,  # 材质（如"砖墙"）
        'Hatch': 21,     # 填充图案
        'Surface': 22,   # 表面材质
    }
}

def get_tarch_property(obj, property_name):
    """
    获取天正对象的属性值

    Args:
        obj: 天正对象（TDbOpening, TDbWall等）
        property_name: 属性名称（如'Width', 'Height', 'Thickness'等）

    Returns:
        属性值，失败返回None

    示例:
        width = get_tarch_property(door_obj, 'Width')
        thickness = get_tarch_property(wall_obj, 'Thickness')
    """
    try:
        obj_type = obj.ObjectName
        if obj_type not in TARCH_PROPERTY_MAP:
            print(f"[警告] 不支持的天正对象类型: {obj_type}")
            return None

        dispid = TARCH_PROPERTY_MAP[obj_type].get(property_name)
        if dispid is None:
            print(f"[警告] {obj_type} 没有属性 {property_name}")
            return None

        oleobj = obj._oleobj_
        result = oleobj.Invoke(dispid, 0, pythoncom.DISPATCH_PROPERTYGET, True)
        return result
    except Exception as e:
        print(f"[错误] 获取属性失败: {e}")
        return None

def set_tarch_property(obj, property_name, value):
    """
    设置天正对象的属性值

    Args:
        obj: 天正对象
        property_name: 属性名称
        value: 属性值

    Returns:
        bool: 成功返回True

    示例:
        set_tarch_property(door_obj, 'Width', 1200)
        set_tarch_property(wall_obj, 'Thickness', 240)
    """
    try:
        obj_type = obj.ObjectName
        if obj_type not in TARCH_PROPERTY_MAP:
            print(f"[警告] 不支持的天正对象类型: {obj_type}")
            return False

        dispid = TARCH_PROPERTY_MAP[obj_type].get(property_name)
        if dispid is None:
            print(f"[警告] {obj_type} 没有属性 {property_name}")
            return False

        oleobj = obj._oleobj_
        oleobj.Invoke(dispid, 0, pythoncom.DISPATCH_PROPERTYPUT, True, value)
        return True
    except Exception as e:
        print(f"[错误] 设置属性失败: {e}")
        return False

def get_tarch_property_by_dispid(obj, dispid):
    """
    通过DISPID直接获取天正对象属性（用于访问未映射的属性）

    Args:
        obj: 天正对象
        dispid: 属性的DISPID编号

    Returns:
        属性值，失败返回None

    示例:
        # 访问DISPID 21的属性
        value = get_tarch_property_by_dispid(wall_obj, 21)
    """
    try:
        oleobj = obj._oleobj_
        result = oleobj.Invoke(dispid, 0, pythoncom.DISPATCH_PROPERTYGET, True)
        return result
    except:
        return None

def set_tarch_property_by_dispid(obj, dispid, value):
    """
    通过DISPID直接设置天正对象属性

    Args:
        obj: 天正对象
        dispid: 属性的DISPID编号
        value: 属性值

    Returns:
        bool: 成功返回True

    示例:
        set_tarch_property_by_dispid(wall_obj, 21, "新填充图案")
    """
    try:
        oleobj = obj._oleobj_
        oleobj.Invoke(dispid, 0, pythoncom.DISPATCH_PROPERTYPUT, True, value)
        return True
    except:
        return False

def get_all_tarch_properties(obj, max_dispid=100):
    """
    获取天正对象的所有属性（用于调试和探索）

    Args:
        obj: 天正对象
        max_dispid: 最大DISPID编号（默认100）

    Returns:
        dict: {dispid: value} 字典

    示例:
        props = get_all_tarch_properties(door_obj)
        for dispid, value in props.items():
            print(f"DISPID {dispid}: {value}")
    """
    oleobj = obj._oleobj_
    properties = {}

    for dispid in range(1, max_dispid + 1):
        try:
            result = oleobj.Invoke(dispid, 0, pythoncom.DISPATCH_PROPERTYGET, True)
            if result is not None:
                properties[dispid] = result
        except:
            pass

    return properties

if __name__ == "__main__":
    print("天正对象属性访问模块")
    print("="*60)
    print("\n支持的对象类型和属性:")
    for obj_type, props in TARCH_PROPERTY_MAP.items():
        print(f"\n{obj_type}:")
        for prop_name, dispid in sorted(props.items(), key=lambda x: x[1]):
            print(f"  {prop_name:15s} -> DISPID {dispid}")
