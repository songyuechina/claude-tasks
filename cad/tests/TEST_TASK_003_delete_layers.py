#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试任务003: 删除图层功能测试
按照CAD测试规范编写
"""

import sys
from pathlib import Path

sys.path.append(str(Path(__file__).parent.parent / "scripts"))

from CAD_test_framework import TestTask, TestResult, TestStatus
from CAD_basic import delete_layers_from_list, create_layers_from_list, get_acad_doc

class TestTask003(TestTask):
    """测试任务003: 删除图层"""

    def __init__(self):
        super().__init__(
            task_id="TEST_TASK_003",
            task_name="删除图层功能测试",
            description="测试delete_layers_from_list函数删除指定图层"
        )

    def execute(self) -> TestStatus:
        """执行测试"""
        from datetime import datetime

        try:
            self.logger.info("开始执行删除图层测试")

            # 测试数据
            test_layers = ["TestLayer1", "TestLayer2", "TestLayer3"]
            delete_layers = ["TestLayer1", "TestLayer3"]

            # 步骤1: 创建测试图层
            self.logger.info(f"步骤1: 创建测试图层 {test_layers}")
            create_layers_from_list(test_layers)

            # 验证图层已创建
            _, doc = get_acad_doc()
            layers = doc.Layers
            created_count = 0
            for name in test_layers:
                try:
                    layers.Item(name)
                    created_count += 1
                except:
                    pass

            if created_count != len(test_layers):
                return TestStatus(
                    task_id=self.task_id,
                    task_name=self.task_name,
                    result=TestResult.FAILED,
                    message=f"图层创建失败，期望{len(test_layers)}个，实际{created_count}个",
                    start_time=datetime.now(),
                    end_time=datetime.now(),
                    initial_windows=[],
                    final_windows=[],
                    dialog_records=[]
                )

            self.logger.info(f"成功创建 {created_count} 个测试图层")

            # 步骤2: 删除指定图层
            self.logger.info(f"步骤2: 删除图层 {delete_layers}")
            result = delete_layers_from_list(delete_layers)

            # 步骤3: 验证删除结果
            self.logger.info("步骤3: 验证删除结果")

            # 检查已删除的图层
            for name in result['deleted']:
                try:
                    layers.Item(name)
                    return TestStatus(
                        task_id=self.task_id,
                        task_name=self.task_name,
                        result=TestResult.FAILED,
                        message=f"图层 '{name}' 应该被删除但仍然存在",
                        start_time=datetime.now(),
                        end_time=datetime.now(),
                        initial_windows=[],
                        final_windows=[],
                        dialog_records=[]
                    )
                except:
                    self.logger.info(f"验证通过: 图层 '{name}' 已被删除")

            # 检查未删除的图层
            remaining_layer = "TestLayer2"
            try:
                layers.Item(remaining_layer)
                self.logger.info(f"验证通过: 图层 '{remaining_layer}' 仍然存在")
            except:
                return TestStatus(
                    task_id=self.task_id,
                    task_name=self.task_name,
                    result=TestResult.FAILED,
                    message=f"图层 '{remaining_layer}' 不应该被删除但已不存在",
                    start_time=datetime.now(),
                    end_time=datetime.now(),
                    initial_windows=[],
                    final_windows=[],
                    dialog_records=[]
                )

            # 清理: 删除剩余测试图层
            self.logger.info("清理: 删除剩余测试图层")
            delete_layers_from_list([remaining_layer])

            # 测试成功
            return TestStatus(
                task_id=self.task_id,
                task_name=self.task_name,
                result=TestResult.SUCCESS,
                message=f"成功删除 {len(result['deleted'])} 个图层",
                start_time=datetime.now(),
                end_time=datetime.now(),
                initial_windows=[],
                final_windows=[],
                dialog_records=[]
            )

        except Exception as e:
            self.logger.error(f"测试执行异常: {e}")
            import traceback
            traceback.print_exc()

            return TestStatus(
                task_id=self.task_id,
                task_name=self.task_name,
                result=TestResult.ERROR,
                message=f"测试异常: {e}",
                start_time=datetime.now(),
                end_time=datetime.now(),
                initial_windows=[],
                final_windows=[],
                dialog_records=[]
            )

if __name__ == "__main__":
    import logging
    logging.basicConfig(level=logging.INFO)

    test = TestTask003()
    result = test.run()

    print(f"\n测试结果: {result.result.value}")
    print(f"消息: {result.message}")

    sys.exit(0 if result.result == TestResult.SUCCESS else 1)
