#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
简化的删除图层测试
"""

import sys
from pathlib import Path

sys.path.append(str(Path(__file__).parent.parent / "scripts"))

from CAD_basic import (
    start_applicationV9,
    close_all_cad_processes,
    delete_layers_from_list,
    create_layers_from_list,
    get_acad_doc
)
from CAD_coordination import ensure_single_process, wait_quiescent

def main():
    print("="*60)
    print("删除图层功能测试")
    print("="*60)

    try:
        # 启动CAD
        print("\n[步骤1] 启动CAD...")
        proc = start_applicationV9(PTH=r"C:\Tangent\TArchT20V9")
        if not proc:
            print("[错误] CAD启动失败")
            return False

        ensure_single_process()
        wait_quiescent(min_quiet=2.0, timeout=30.0)

        # 创建测试图层
        test_layers = ["TestLayer1", "TestLayer2", "TestLayer3"]
        print(f"\n[步骤2] 创建测试图层: {test_layers}")
        create_layers_from_list(test_layers)

        # 验证图层已创建
        _, doc = get_acad_doc()
        layers = doc.Layers
        print("\n[验证] 检查图层是否创建成功...")
        for name in test_layers:
            try:
                layers.Item(name)
                print(f"  [成功] 图层 '{name}' 已创建")
            except:
                print(f"  [失败] 图层 '{name}' 未创建")

        # 删除部分图层
        delete_layers = ["TestLayer1", "TestLayer3"]
        print(f"\n[步骤3] 删除图层: {delete_layers}")
        result = delete_layers_from_list(delete_layers)

        # 验证删除结果
        print("\n[验证] 检查删除结果...")
        for name in delete_layers:
            try:
                layers.Item(name)
                print(f"  [失败] 图层 '{name}' 应该被删除但仍存在")
            except:
                print(f"  [成功] 图层 '{name}' 已被删除")

        # 检查未删除的图层
        remaining = "TestLayer2"
        try:
            layers.Item(remaining)
            print(f"  [成功] 图层 '{remaining}' 仍然存在（正确）")
        except:
            print(f"  [失败] 图层 '{remaining}' 不应该被删除")

        # 清理
        print(f"\n[清理] 删除剩余测试图层...")
        delete_layers_from_list([remaining])

        print("\n" + "="*60)
        print("测试完成!")
        print("="*60)
        print(f"\n结果: 成功删除 {len(result['deleted'])} 个图层")
        print(f"      失败 {len(result['failed'])} 个图层")

        return True

    except Exception as e:
        print(f"\n[错误] 测试异常: {e}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        print("\n[清理] 关闭所有CAD进程...")
        try:
            close_all_cad_processes()
            print("[成功] 环境已清理")
        except:
            pass

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
