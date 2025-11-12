#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
配置验证脚本
验证 .claude 目录的配置是否完整和正确
"""

import json
import sys
from pathlib import Path

# 设置标准输出编码为 UTF-8
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def verify_config():
    """验证配置完整性"""
    print("="*60)
    print("Claude 配置验证")
    print("="*60)

    base_dir = Path(__file__).parent
    workspace = base_dir.parent

    checks = []

    # 1. 检查目录结构
    print("\n📁 检查目录结构...")

    required_dirs = [
        base_dir / "commands",
    ]

    for dir_path in required_dirs:
        exists = dir_path.exists()
        checks.append(("目录", str(dir_path.relative_to(workspace)), exists))
        print(f"  {'✅' if exists else '❌'} {dir_path.relative_to(workspace)}")

    # 2. 检查配置文件
    print("\n📄 检查配置文件...")

    required_files = [
        base_dir / "README.md",
        base_dir / "CLAUDE_BEHAVIOR_RULES.md",
        base_dir / "claude_config.json",
        base_dir / "commands" / "init.md",
        base_dir / "commands" / "cad.md",
    ]

    for file_path in required_files:
        exists = file_path.exists()
        checks.append(("文件", str(file_path.relative_to(workspace)), exists))
        print(f"  {'✅' if exists else '❌'} {file_path.relative_to(workspace)}")

    # 3. 检查 JSON 格式
    print("\n🔍 检查 JSON 格式...")

    json_file = base_dir / "claude_config.json"
    if json_file.exists():
        try:
            with open(json_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
            print(f"  ✅ JSON 格式正确")
            checks.append(("JSON格式", "claude_config.json", True))

            # 显示配置信息
            print(f"\n  工作站名称: {config.get('name', 'N/A')}")
            print(f"  版本: {config.get('version', 'N/A')}")
            print(f"  模式: {config.get('behavior', {}).get('mode', 'N/A')}")

        except json.JSONDecodeError as e:
            print(f"  ❌ JSON 格式错误: {e}")
            checks.append(("JSON格式", "claude_config.json", False))
    else:
        print(f"  ❌ 文件不存在")
        checks.append(("JSON格式", "claude_config.json", False))

    # 4. 检查规范文档
    print("\n📚 检查规范文档...")

    required_docs = [
        workspace / "CAD_操作规范.md",
        workspace / "CAD_基本操作范式.md",
        workspace / "CAD_快速参考.md",
        workspace / "CAD_测试规范.md",
    ]

    for doc_path in required_docs:
        exists = doc_path.exists()
        checks.append(("规范文档", doc_path.name, exists))
        print(f"  {'✅' if exists else '❌'} {doc_path.name}")

    # 5. 检查核心模块
    print("\n🔧 检查核心模块...")

    scripts_dir = workspace / "scripts"
    required_modules = [
        "CAD_basic_operations.py",
        "CAD_coordination.py",
        "CAD_enhanced_functions.py",
        "CAD_test_framework.py",
        "cad_dialog_killer.py",
    ]

    for module in required_modules:
        module_path = scripts_dir / module
        exists = module_path.exists()
        checks.append(("核心模块", module, exists))
        print(f"  {'✅' if exists else '❌'} {module}")

    # 6. 统计结果
    print("\n" + "="*60)
    print("验证结果统计")
    print("="*60)

    total = len(checks)
    passed = sum(1 for _, _, result in checks if result)
    failed = total - passed

    print(f"\n总计: {total} 项")
    print(f"✅ 通过: {passed} 项")
    print(f"❌ 失败: {failed} 项")
    print(f"通过率: {passed/total*100:.1f}%")

    # 7. 失败项详情
    if failed > 0:
        print("\n❌ 失败项详情:")
        for check_type, name, result in checks:
            if not result:
                print(f"  - [{check_type}] {name}")

    # 8. 建议
    print("\n" + "="*60)
    print("使用建议")
    print("="*60)

    if passed == total:
        print("\n✅ 所有配置检查通过！")
        print("\n你可以：")
        print("  1. 在新会话中输入 /init 初始化环境")
        print("  2. 在 CAD 任务前输入 /cad 启动 CAD 模式")
        print("  3. 查看 .claude/README.md 了解详细使用方法")
    else:
        print("\n⚠️  部分配置缺失或错误")
        print("\n建议：")
        print("  1. 检查失败项是否存在")
        print("  2. 确认文件路径正确")
        print("  3. 验证 JSON 格式")
        print("  4. 查看 .claude/README.md 了解配置要求")

    print("\n" + "="*60)

    return passed == total

if __name__ == "__main__":
    try:
        success = verify_config()
        exit(0 if success else 1)
    except Exception as e:
        print(f"\n❌ 验证过程出错: {e}")
        import traceback
        traceback.print_exc()
        exit(1)
