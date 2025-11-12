#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
批量运行所有测试任务
"""

import sys
import time
import json
from pathlib import Path
from datetime import datetime

# 添加脚本目录到路径
sys.path.append(str(Path(__file__).parent.parent / "scripts"))

# 导入测试任务
from TEST_TASK_001_open_file import TestTask001
from TEST_TASK_002_close_file import TestTask002

class TestRunner:
    """测试运行器"""

    def __init__(self):
        self.results = []
        self.start_time = datetime.now()

    def run_all_tests(self) -> bool:
        """运行所有测试任务"""
        print(" 开始运行所有CAD测试任务")
        print("=" * 60)

        # 测试任务列表
        test_tasks = [
            TestTask001(),
            TestTask002(),
        ]

        success_count = 0
        total_count = len(test_tasks)

        for i, test in enumerate(test_tasks, 1):
            print(f"\n📋 [{i}/{total_count}] 运行测试任务: {test.task_name}")
            print("-" * 40)

            try:
                # 运行测试
                result = test.run()
                self.results.append(result)

                if result.result.value == "成功":
                    success_count += 1
                    print(f"✅ 测试成功: {result.message}")
                else:
                    print(f"❌ 测试失败: {result.message}")

                # 任务间隔等待
                if i < total_count:
                    print("⏳ 等待3秒后继续下一个测试...")
                    time.sleep(3.0)

            except Exception as e:
                print(f"💥 测试任务异常: {e}")
                error_result = type('Result', (), {
                    'result': type('Result', (), {'value': '错误'})(),
                    'message': f"测试异常: {e}",
                    'task_id': test.task_id,
                    'task_name': test.task_name
                })()
                self.results.append(error_result)

        # 生成测试报告
        self.generate_report(success_count, total_count)

        return success_count == total_count

    def generate_report(self, success_count: int, total_count: int):
        """生成测试报告"""
        print("\n" + "=" * 60)
        print("📊 测试报告")
        print("=" * 60)

        end_time = datetime.now()
        duration = (end_time - self.start_time).total_seconds()

        print(f"总测试数: {total_count}")
        print(f"成功测试: {success_count}")
        print(f"失败测试: {total_count - success_count}")
        print(f"成功率: {success_count/total_count*100:.1f}%")
        print(f"总耗时: {duration:.2f}秒")

        print(f"\n📋 详细结果:")
        for result in self.results:
            status_icon = "✅" if result.result.value == "成功" else "❌"
            print(f"  {status_icon} {result.task_id}: {result.result.value}")

        # 保存详细报告
        report_data = {
            "summary": {
                "total_tests": total_count,
                "success_count": success_count,
                "failed_count": total_count - success_count,
                "success_rate": success_count/total_count*100,
                "duration_seconds": duration,
                "start_time": self.start_time.isoformat(),
                "end_time": end_time.isoformat()
            },
            "results": []
        }

        for result in self.results:
            result_data = {
                "task_id": result.task_id,
                "task_name": result.task_name,
                "result": result.result.value,
                "message": result.message,
                "start_time": result.start_time.isoformat() if hasattr(result, 'start_time') else None,
                "end_time": result.end_time.isoformat() if hasattr(result, 'end_time') else None,
                "initial_window_count": len(result.initial_windows) if hasattr(result, 'initial_windows') else 0,
                "final_window_count": len(result.final_windows) if hasattr(result, 'final_windows') else 0,
                "dialog_count": len(result.dialog_records) if hasattr(result, 'dialog_records') else 0
            }
            report_data["results"].append(result_data)

        # 保存报告文件
        report_file = Path(__file__).parent / f"test_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        with open(report_file, 'w', encoding='utf-8') as f:
            json.dump(report_data, f, ensure_ascii=False, indent=2)

        print(f"\n📄 详细报告已保存: {report_file}")

def main():
    """主函数"""
    runner = TestRunner()
    success = runner.run_all_tests()

    if success:
        print("\n🎉 所有测试通过!")
        return True
    else:
        print("\n⚠ 部分测试失败,请查看详细报告")
        return False

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)