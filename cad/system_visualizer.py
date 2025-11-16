#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
D:/claude-tasks 体系可视化和汇报工具

双击运行，生成以下内容：
1. 文件夹全景图（目录树形结构）
2. 关于Claude的能力汇报
3. Claude执行任务的流程汇报
4. 函数编写和测试情况汇报
5. 体系分析和建议

生成的汇报保存到：D:/claude-tasks/体系汇报.md
"""

import os
import sys
from pathlib import Path
from datetime import datetime
import json

def generate_directory_tree(root_path, output_file, indent="", max_depth=5, current_depth=0):
    """生成目录树形结构"""
    if current_depth > max_depth:
        return

    try:
        items = sorted(Path(root_path).iterdir(), key=lambda x: (not x.is_dir(), x.name.lower()))
    except PermissionError:
        return

    for i, item in enumerate(items):
        # 跳过隐藏文件和特定目录
        if item.name.startswith('.') and item.name not in ['.claude']:
            continue
        if item.name in ['__pycache__', 'node_modules', '.git']:
            continue

        is_last = i == len(items) - 1
        connector = "└── " if is_last else "├── "

        if item.is_dir():
            output_file.write(f"{indent}{connector}{item.name}/\n")
            extension = "    " if is_last else "│   "
            generate_directory_tree(item, output_file, indent + extension, max_depth, current_depth + 1)
        else:
            # 显示文件大小
            size = item.stat().st_size
            size_str = f"{size:,} bytes" if size < 1024 else f"{size/1024:.1f} KB"
            output_file.write(f"{indent}{connector}{item.name} ({size_str})\n")

def count_files_by_type(root_path):
    """统计各类型文件数量"""
    stats = {}
    for root, dirs, files in os.walk(root_path):
        # 跳过隐藏目录
        dirs[:] = [d for d in dirs if not d.startswith('.') and d not in ['__pycache__', 'node_modules']]

        for file in files:
            ext = Path(file).suffix.lower()
            if not ext:
                ext = "无扩展名"
            stats[ext] = stats.get(ext, 0) + 1

    return stats

def analyze_cad_functions():
    """分析CAD项目的函数情况"""
    cad_dir = Path("D:/claude-tasks/cad")

    # 分析CAD_file_operations.py
    file_ops = cad_dir / "scripts" / "CAD_file_operations.py"
    functions = []

    if file_ops.exists():
        with open(file_ops, 'r', encoding='utf-8') as f:
            content = f.read()
            # 简单提取函数定义
            import re
            func_pattern = r'def\s+(\w+)\s*\('
            functions = re.findall(func_pattern, content)

    return functions

def analyze_test_reports():
    """分析测试报告"""
    test_reports_dir = Path("D:/claude-tasks/cad/test_reports")
    reports = []

    if test_reports_dir.exists():
        for file in test_reports_dir.glob("*.xlsx"):
            reports.append({
                "name": file.name,
                "size": file.stat().st_size,
                "modified": datetime.fromtimestamp(file.stat().st_mtime).strftime('%Y-%m-%d %H:%M:%S')
            })

    return reports

def generate_report():
    """生成完整汇报"""
    root_path = Path("D:/claude-tasks")
    output_path = root_path / "体系汇报.md"

    with open(output_path, 'w', encoding='utf-8') as f:
        # 标题
        f.write("# D:/claude-tasks 体系汇报\n\n")
        f.write(f"**生成时间**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        f.write("---\n\n")

        # 一、目录全景图
        f.write("## 一、目录全景图\n\n")
        f.write("```\n")
        f.write("D:/claude-tasks/\n")
        generate_directory_tree(root_path, f, "", max_depth=3)
        f.write("```\n\n")

        # 文件统计
        f.write("### 文件统计\n\n")
        file_stats = count_files_by_type(root_path)
        f.write("| 文件类型 | 数量 |\n")
        f.write("|---------|------|\n")
        for ext, count in sorted(file_stats.items(), key=lambda x: x[1], reverse=True):
            f.write(f"| {ext} | {count} |\n")
        f.write("\n")

        # 二、关于Claude
        f.write("## 二、关于Claude\n\n")

        f.write("### 1. Claude能力汇报\n\n")

        f.write("#### 网络和软件能力\n\n")
        f.write("| 能力 | 状态 | 说明 |\n")
        f.write("|------|------|------|\n")
        f.write("| 上网搜索 | ✅ 支持 | 可以使用WebSearch工具搜索最新信息 |\n")
        f.write("| 网页抓取 | ✅ 支持 | 可以使用WebFetch工具获取网页内容 |\n")
        f.write("| 下载软件 | ❌ 不支持 | 无法下载和安装软件 |\n")
        f.write("| 软件升级 | ❌ 不支持 | 无法自动升级软件 |\n")
        f.write("| GitHub推送 | ❌ 受限 | 可以生成代码但无法直接推送到GitHub |\n")
        f.write("| 账号登录 | ❌ 不支持 | 无法使用用户密码登录网站 |\n")
        f.write("| 发布文章 | ❌ 不支持 | 无法自动发布文章到网站 |\n\n")

        f.write("#### MCP能力\n\n")
        f.write("当前未检测到MCP（Model Context Protocol）服务器连接。\n\n")

        f.write("### 2. 电脑操作权限\n\n")
        f.write("| 操作类型 | 权限 | 说明 |\n")
        f.write("|---------|------|------|\n")
        f.write("| 读取文件 | ✅ 全局 | 可以读取电脑上任意可访问的文件 |\n")
        f.write("| 写入文件 | ✅ 全局 | 可以在有权限的目录写入文件 |\n")
        f.write("| 编辑文件 | ✅ 全局 | 可以编辑现有文件 |\n")
        f.write("| 删除文件 | ❌ 受限 | 无直接删除工具，需通过Bash命令 |\n")
        f.write("| 移动文件 | ❌ 受限 | 无直接移动工具，需通过Bash命令 |\n")
        f.write("| 复制文件 | ❌ 受限 | 无直接复制工具，需通过Bash命令 |\n")
        f.write("| 执行命令 | ✅ 全局 | 可以通过Bash工具执行Shell命令 |\n")
        f.write("| Python执行 | ✅ 全局 | 可以执行Python脚本 |\n\n")

        f.write("### 3. Claude版本信息\n\n")
        f.write("| 项目 | 信息 |\n")
        f.write("|------|------|\n")
        f.write("| 模型 | Claude Sonnet 4.5 |\n")
        f.write("| 模型ID | claude-sonnet-4-5-20250929 |\n")
        f.write("| 知识截止 | 2025年1月 |\n")
        f.write("| 产品 | Claude Code (Anthropic官方CLI) |\n")
        f.write("| 平台 | Windows (win32) |\n\n")

        # 三、Claude执行任务
        f.write("## 三、Claude执行任务\n\n")

        f.write("### 1. 标准任务流程\n\n")
        f.write("#### 任务开始阶段\n\n")
        f.write("1. **读取规范文档**\n")
        f.write("   - `D:/claude-tasks/CLAUDE.md` - 通用规范\n")
        f.write("   - `D:/claude-tasks/cad/CLAUDE.md` - CAD项目规范\n")
        f.write("   - `cad/CAD_快速参考.md` - 协同机制\n")
        f.write("   - `cad/CAD_操作规范.md` - 核心规则\n\n")

        f.write("2. **读取任务需求**\n")
        f.write("   - `/cad` 命令：读取 `即时对话.txt`\n")
        f.write("   - `/jst` 命令：执行 `即时对话.txt` 的任务\n\n")

        f.write("3. **创建TodoList**\n")
        f.write("   - 分解任务为具体步骤\n")
        f.write("   - 标记任务状态（pending/in_progress/completed）\n\n")

        f.write("#### 任务执行阶段\n\n")
        f.write("1. **遵循强制规则**\n")
        f.write("   - 测试前调用 `cad_zt_zero()` 清理\n")
        f.write("   - 遵循1分钟原则（超时立即终止）\n")
        f.write("   - 使用 `CAD_file_operations.py` 进行文件操作\n")
        f.write("   - 使用规定函数进行天正操作\n\n")

        f.write("2. **日志记录**\n")
        f.write("   - 关键步骤打印日志\n")
        f.write("   - 记录执行时间\n")
        f.write("   - 错误详细traceback\n\n")

        f.write("3. **进度汇报**（条件执行）\n")
        f.write("   - **触发条件**：任务明确要求\"每隔X分钟汇报进度\"\n")
        f.write("   - **心跳监控**：每60秒写入时间戳到任务进度汇报.txt\n")
        f.write("   - **状态检查**：响应用户\"状态检查：是否遇到问题？\"\n\n")

        f.write("#### 任务结束阶段\n\n")
        f.write("1. **三次清理规则**（🔴 最重要）\n")
        f.write("   - 第1次：测试开始前调用 `cad_zt_zero()`\n")
        f.write("   - 第2次：测试脚本finally块调用 `cad_zt_zero()`\n")
        f.write("   - **第3次：任务完成汇报前调用 `cad_zt_zero()`（最容易遗忘）**\n\n")

        f.write("2. **清空即时对话.txt**\n")
        f.write("   - 每次任务结束后必须清空\n")
        f.write("   - 避免下次任务读取到旧内容\n\n")

        f.write("3. **条件性操作**（仅在任务明确要求时）\n")
        f.write("   - 🔔 **设置结束信号**：仅当任务有\"### 关机信号\"字段时，将`可以结束对话.txt`设为1\n")
        f.write("   - ⚠️ **警告**：设置为1会触发系统关机！\n\n")

        f.write("4. **完成汇报**\n")
        f.write("   - 生成测试报告（Excel + Markdown）\n")
        f.write("   - 归档测试文件\n")
        f.write("   - 发送完成消息\n\n")

        f.write("### 2. 任务格式模板\n\n")
        f.write("#### 简化版格式\n\n")
        f.write("```markdown\n")
        f.write("<任务>\n\n")
        f.write("**任务类型**：CAD操作和函数、脚本处理\n\n")
        f.write("**目标**：{任务简述}\n\n")
        f.write("**要求**：\n")
        f.write("1. {要求1}\n")
        f.write("2. {要求2}\n\n")
        f.write("**完成标准**：{如何判断任务完成}\n\n")
        f.write("**强制规范**：\n")
        f.write("- 🔴 任务完成前必须调用cad_zt_zero()清理CAD进程\n")
        f.write("- 📝 清空即时对话.txt\n\n")
        f.write("</任务>\n")
        f.write("```\n\n")

        # 四、函数编写和测试
        f.write("## 四、函数编写和测试\n\n")

        f.write("### 1. 函数编写规范\n\n")
        f.write("**核心原则**：\n")
        f.write("- 优先使用 `CAD_basic.py` 的数据类型转换函数\n")
        f.write("- CAD对象使用 `cast_object` 获取属性\n")
        f.write("- 天正对象使用 `get_object_property` 和 `set_object_property`\n")
        f.write("- 选择操作优先调用\"第四部分 选择方法\"\n")
        f.write("- 绘图操作优先调用\"第五部分 线面分析\"\n")
        f.write("- 图块操作优先调用\"第九部分 CAD图块\"\n\n")

        f.write("### 2. 函数清单\n\n")
        functions = analyze_cad_functions()
        f.write(f"**CAD_file_operations.py** 共有 {len(functions)} 个函数：\n\n")
        for i, func in enumerate(functions, 1):
            f.write(f"{i}. `{func}()`\n")
        f.write("\n")

        f.write("### 3. 测试报告\n\n")
        test_reports = analyze_test_reports()
        if test_reports:
            f.write("| 报告名称 | 大小 | 修改时间 |\n")
            f.write("|---------|------|----------|\n")
            for report in test_reports:
                size_kb = report['size'] / 1024
                f.write(f"| {report['name']} | {size_kb:.1f} KB | {report['modified']} |\n")
        else:
            f.write("暂无测试报告\n")
        f.write("\n")

        f.write("### 4. 测试文件位置\n\n")
        f.write("- **测试脚本**：`D:/claude-tasks/cad/tests/`\n")
        f.write("- **测试文件**：`D:/claude-tasks/cad/tests/test_files/`\n")
        f.write("- **测试报告**：`D:/claude-tasks/cad/test_reports/`\n")
        f.write("- **DWG存档**：`D:/claude-tasks/cad/test_reports/test_dwg_files/`\n\n")

        # 五、体系分析和建议
        f.write("## 五、体系分析和建议\n\n")

        f.write("### 1. 当前结构优势\n\n")
        f.write("✅ **规范体系完善**\n")
        f.write("- 有完整的CLAUDE.md规范文档\n")
        f.write("- 任务流程清晰明确\n")
        f.write("- 强制规则（三次清理）确保稳定性\n\n")

        f.write("✅ **测试体系健全**\n")
        f.write("- 测试覆盖率要求明确\n")
        f.write("- 有Excel和Markdown双重报告\n")
        f.write("- DWG文件归档便于复测\n\n")

        f.write("✅ **协同机制可靠**\n")
        f.write("- `wait_quiescent()` 确保CAD同步\n")
        f.write("- `send_cmd_with_sync()` 命令协同\n")
        f.write("- 1分钟原则避免卡死\n\n")

        f.write("### 2. 改进建议\n\n")
        f.write("#### 建议1：增强进度监控\n\n")
        f.write("**当前问题**：\n")
        f.write("- 心跳监控是条件执行，简单任务无监控\n")
        f.write("- 无法实时看到任务进度\n\n")

        f.write("**建议方案**：\n")
        f.write("- 所有任务默认启用轻量级心跳（每5分钟）\n")
        f.write("- 重要任务启用详细进度汇报（每1分钟）\n")
        f.write("- 增加进度百分比显示\n\n")

        f.write("#### 建议2：自动化测试流程\n\n")
        f.write("**当前问题**：\n")
        f.write("- 测试需要手动运行脚本\n")
        f.write("- 测试报告需要手动生成\n\n")

        f.write("**建议方案**：\n")
        f.write("- 开发 `run_all_tests.py` 一键运行所有测试\n")
        f.write("- 测试完成自动生成报告\n")
        f.write("- 集成到CI/CD流程\n\n")

        f.write("#### 建议3：文档版本管理\n\n")
        f.write("**当前问题**：\n")
        f.write("- 文档更新频繁\n")
        f.write("- 无历史版本记录\n\n")

        f.write("**建议方案**：\n")
        f.write("- 文档添加版本号和更新日志\n")
        f.write("- 保留最近3个版本\n")
        f.write("- 使用Git进行版本控制\n\n")

        f.write("### 3. 配置建议\n\n")
        f.write("#### Claude行为配置\n\n")
        f.write("建议在 `.claude/CLAUDE_BEHAVIOR_RULES.md` 增加：\n")
        f.write("- 自动错误恢复机制\n")
        f.write("- 智能重试策略\n")
        f.write("- 性能优化提示\n\n")

        f.write("#### 工作区配置\n\n")
        f.write("建议添加：\n")
        f.write("- `.editorconfig` 统一编辑器配置\n")
        f.write("- `.gitignore` 忽略临时文件\n")
        f.write("- `requirements.txt` 明确Python依赖\n\n")

        f.write("---\n\n")
        f.write("**汇报生成完毕**\n\n")
        f.write(f"生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"报告位置：{output_path}\n")

    return output_path

def main():
    """主函数"""
    print("="*60)
    print("D:/claude-tasks 体系可视化和汇报工具")
    print("="*60)
    print()

    try:
        print("[1/1] 正在生成体系汇报...")
        output_path = generate_report()
        print("\n[成功] 汇报生成成功！")
        print(f"\n报告位置：{output_path}")
        print("\n请打开文件查看完整汇报。")

        # 自动打开报告
        import subprocess
        subprocess.Popen(['notepad.exe', str(output_path)])

    except Exception as e:
        print(f"\n[错误] 生成汇报失败：{e}")
        import traceback
        traceback.print_exc()

    print("\n按任意键退出...")
    input()

if __name__ == "__main__":
    main()
