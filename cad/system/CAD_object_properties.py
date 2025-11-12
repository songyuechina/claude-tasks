#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CAD对象属性访问统一模块

支持：
1. CAD标准对象（直线、圆、多段线等）- 通过CastTo转换
2. 天正对象（门窗、墙等）- 通过IDispatch.Invoke访问

使用方法:
    from CAD_object_properties import get_object_property, set_object_property, cast_object

    # 自动识别对象类型并获取属性
    obj = cast_object(entity)  # 转换对象
    start_point = obj.StartPoint  # 访问属性
"""

import pythoncom
import time
import pywintypes
from win32com.client import CastTo

# 通用重试：吞"忙/拒绝/RPC down"
_RPC_BUSY = (-2147417846, -2147418111)
_RPC_DOWN = (-2147023174,)

def com_retry(fn, retries=30, delay=0.05):
    """COM调用重试机制"""
    for _ in range(retries):
        try:
            return fn()
        except pywintypes.com_error as e:
            code = e.args[0] if e.args else None
            if code in _RPC_BUSY + _RPC_DOWN:
                time.sleep(delay)
                continue
            raise
    return fn()

# CAD标准对象类型映射
_CAST_MAP = {
    # 基础几何
    "AcDbLine": "IAcadLine",
    "AcDbCircle": "IAcadCircle",
    "AcDbArc": "IAcadArc",
    "AcDbPoint": "IAcadPoint",
    "AcDbEllipse": "IAcadEllipse",
    "AcDbSpline": "IAcadSpline",
    # 多段线
    "AcDbPolyline": "IAcadLWPolyline",
    "AcDb2dPolyline": "IAcadPolyline",
    "AcDb3dPolyline": "IAcad3DPolyline",
    # 文字
    "AcDbText": "IAcadText",
    "AcDbMText": "IAcadMText",
    # 块/属性
    "AcDbBlockReference": "IAcadBlockReference",
    "AcDbAttribute": "IAcadAttributeReference",
    "AcDbAttributeDefinition": "IAcadAttribute",
    # 引线/标注
    "AcDbLeader": "IAcadLeader",
    "AcDbMLeader": "IAcadMLeader",
    "AcDbDimension": "IAcadDimension",
    "AcDbAlignedDimension": "IAcadDimAligned",
    "AcDbRotatedDimension": "IAcadDimRotated",
    "AcDbRadialDimension": "IAcadDimRadial",
    "AcDbDiametricDimension": "IAcadDimDiametric",
    "AcDbArcDimension": "IAcadDimArc",
    "AcDb3PointAngularDimension": "IAcadDim3PointAngular",
    "AcDb2LineAngularDimension": "IAcadDim2LineAngular",
    "AcDbOrdinateDimension": "IAcadDimOrdinate",
    # 其它
    "AcDbHatch": "IAcadHatch",
    "AcDbTable": "IAcadTable",
}

# 天正对象属性DISPID映射
_TARCH_PROPERTY_MAP = {
    'TDbOpening': {
        'Offset': 1, 'Width': 2, 'Type': 3, 'Direction': 7,
        'Angle': 8, 'Height': 10, 'Name': 11,
    },
    'TDbWall': {
        'Offset1': 1, 'Thickness': 2, 'Thickness2': 3, 'Length': 4,
        'WallType': 11, 'Material': 13, 'Hatch': 21, 'Surface': 22,
    }
}

def _maybe_cast(ent):
    """
    安全地转换CAD对象到专用接口（使用com_retry）

    Args:
        ent: CAD对象

    Returns:
        转换后的对象（如果是CAD标准对象）或原对象

    示例:
        obj = last_obj()
        obj = _maybe_cast(obj)
        coords = obj.Coordinates  # 访问多段线坐标
    """
    try:
        name = com_retry(lambda: ent.ObjectName)
        iface = _CAST_MAP.get(name)
        if iface:
            try:
                return CastTo(ent, iface)
            except Exception:
                return ent
        return ent
    except Exception:
        return ent

def cast_object(obj):
    """
    转换CAD对象到专用接口（兼容旧代码，内部调用_maybe_cast）

    Args:
        obj: CAD对象

    Returns:
        转换后的对象（如果是CAD标准对象）或原对象

    示例:
        line = cast_object(entity)
        start = line.StartPoint  # 访问直线起点
    """
    return _maybe_cast(obj)

def get_object_property(obj, property_name):
    """
    统一获取对象属性（自动识别CAD对象或天正对象）

    Args:
        obj: CAD对象或天正对象
        property_name: 属性名称

    Returns:
        属性值

    示例:
        # CAD对象
        start = get_object_property(line, 'StartPoint')

        # 天正对象
        width = get_object_property(door, 'Width')
    """
    try:
        obj_name = com_retry(lambda: obj.ObjectName)

        # 天正对象 - 使用DISPID访问
        if obj_name in _TARCH_PROPERTY_MAP:
            dispid = _TARCH_PROPERTY_MAP[obj_name].get(property_name)
            if dispid is None:
                print(f"[警告] {obj_name} 没有属性 {property_name}")
                return None

            oleobj = obj._oleobj_
            return oleobj.Invoke(dispid, 0, pythoncom.DISPATCH_PROPERTYGET, True)

        # CAD标准对象 - 先转换再访问
        obj = _maybe_cast(obj)
        return getattr(obj, property_name)

    except Exception as e:
        print(f"[错误] 获取属性失败: {e}")
        return None

def set_object_property(obj, property_name, value):
    """
    统一设置对象属性（自动识别CAD对象或天正对象）

    Args:
        obj: CAD对象或天正对象
        property_name: 属性名称
        value: 属性值

    Returns:
        bool: 成功返回True

    示例:
        set_object_property(door, 'Width', 1200)
    """
    try:
        obj_name = com_retry(lambda: obj.ObjectName)

        # 天正对象 - 使用DISPID设置
        if obj_name in _TARCH_PROPERTY_MAP:
            dispid = _TARCH_PROPERTY_MAP[obj_name].get(property_name)
            if dispid is None:
                return False

            oleobj = obj._oleobj_
            oleobj.Invoke(dispid, 0, pythoncom.DISPATCH_PROPERTYPUT, True, value)
            return True

        # CAD标准对象 - 先转换再设置
        obj = _maybe_cast(obj)
        setattr(obj, property_name, value)
        return True

    except Exception as e:
        print(f"[错误] 设置属性失败: {e}")
        return False

if __name__ == "__main__":
    print("CAD对象属性访问统一模块")
    print("="*60)
    print("\n支持的CAD标准对象:")
    for obj_type in sorted(_CAST_MAP.keys()):
        print(f"  {obj_type}")
    print("\n支持的天正对象:")
    for obj_type in _TARCH_PROPERTY_MAP.keys():
        print(f"  {obj_type}")
