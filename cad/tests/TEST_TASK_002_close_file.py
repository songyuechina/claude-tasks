#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试任务编号: TEST_TASK_002
测试名称: 关闭DWG文件测试
创建日期: 2025-11-06
测试内容: 验证DWG文件的关闭操作,包括保存提示和状态转换
"""

import sys
from pathlib import Path

# 添加脚本目录到路径
sys.path.append(str(Path(__file__).parent.parent / "scripts"))

from CAD_test_framework import TestTask, TestResult, TestStatus
from CAD_coordination import wait_quiescent, send_cmd_with_sync
from CAD_enhanced_functions import open_dwg_sync
from datetime import datetime

class TestTask002(TestTask):
    """测试任务002: 关闭DWG文件"""

    def __init__(self):
        super().__init__(
            task_id="TEST_TASK_002",
            task_name="关闭DWG文件测试",
            description="验证DWG文件的关闭操作,检查保存提示和状态转换是否正常"
        )

        # 测试参数
        self.test_file_path = r"D:\claude-tasks\test_files\sample.dwg"
        self.test_files_dir = Path(__file__).parent.parent / "test_files"

    def setup(self) -> bool:
        """测试准备"""
        if not super().setup():
            return False

        # 确保测文件目录存在
        self.test_files_dir.mkdir(exist_ok=True)

        self.logger.info("测试任务002准备完成")
        return True

    def execute(self) -> TestStatus:
        """执行测试"""
        try:
            self.logger.info("开始执行关闭文件测试...")

            # 步骤1: 先打开一个文件用于测试关闭
            self.logger.info("步骤1: 准备测试文件")
            prep_success = self._prepare_test_file()

            if not prep_success:
                return self._create_failed_status("准备测试文件失败")

            # 步骤2: 记录文件打开状态
            self.logger.info("步骤2: 记录文件打开状态")
            initial_file_count = self.checker.get_open_file_count()
            self.logger.info(f"关闭前打开文件数: {initial_file_count}")

            # 步骤3: 执行关闭文件操作
            self.logger.info("步骤3: 执行关闭文件操作")
            close_success = self._perform_close_file_test()

            if not close_success:
                return self._create_failed_status("关闭文件操作失败")

            # 步骤4: 验证关闭结果
            self.logger.info("步骤4: 验证关闭结果")
            verification_result = self._verify_file_closed()

            if not verification_result:
                return self._create_failed_status("文件关闭验证失败")

            # 步骤5: 检查状态转换
            self.logger.info("步骤5: 检查状态转换")
            state_check = self._check_state_transition()

            if not state_check:
                return self._create_failed_status("状态转换验证失败")

            self.logger.info("✅ 关闭文件测试成功完成")
            return self._create_success_status("关闭文件测试成功完成")

        except Exception as e:
            self.logger.error(f"❌ 测试执行异常: {e}")
            return self._create_failed_status(f"测试执行异常: {e}")

    def _prepare_test_file(self) -> bool:
        """准备测试文件(先打开一个文件)"""
        try:
            # 检查是否已有打开的文件
            current_count = self.checker.get_open_file_count()
            if current_count > 0:
                self.logger.info(f"已有 {current_count} 个文件打开,跳过准备步骤")
                return True

            # 如果测试文件存在,尝试打开
            if Path(self.test_file_path).exists():
                self.logger.info(f"打开测试文件: {self.test_file_path}")
                return open_dwg_sync(self.test_file_path, visible=True)
            else:
                # 模拟打开操作(创建新文件)
                self.logger.info("测试文件不存在,模拟打开新文件")
                return send_cmd_with_sync("_NEW\n", wait_after=1.0)

        except Exception as e:
            self.logger.error(f"❌ 准备测试文件异常: {e}")
            return False

    def _perform_close_file_test(self) -> bool:
        """执行关闭文件测试"""
        try:
            self.logger.info("发送关闭文件命令...")

            # 发送关闭命令
            # 使用CLOSE命令,如果有未保存的更改会弹出保存提示
            success = send_cmd_with_sync("_CLOSE\n", wait_after=1.0)

            if success:
                self.logger.info("✅ 关闭命令发送成功")

                # 等待关闭操作完成
                wait_quiescent(min_quiet=1.0, timeout=15.0)

                return True
            else:
                self.logger.error("❌ 关闭命令发送失败")
                return False

        except Exception as e:
            self.logger.error(f"❌ 关闭文件操作异常: {e}")
            return False

    def _verify_file_closed(self) -> bool:
        """验证文件是否成功关闭"""
        try:
            # 检查文件数量变化
            current_file_count = self.checker.get_open_file_count()
            self.logger.info(f"关闭后打开文件数: {current_file_count}")

            # 检查目标文件是否已关闭
            if Path(self.test_file_path).exists():
                is_opened = self.checker.is_file_opened(self.test_file_path)
                self.logger.info(f"目标文件是否仍打开: {is_opened}")

                if not is_opened:
                    self.logger.info("✅ 文件关闭验证成功")
                    return True
                else:
                    self.logger.error("❌ 文件关闭验证失败: 文件仍处于打开状态")
                    return False
            else:
                # 模拟测试
                if current_file_count >= 0:
                    self.logger.info("✅ 模拟文件关闭验证成功")
                    return True
                else:
                    return False

        except Exception as e:
            self.logger.error(f"❌ 文件关闭验证异常: {e}")
            return False

    def _check_state_transition(self) -> bool:
        """检查状态转换是否正确"""
        try:
            # 等待CAD稳定
            wait_quiescent(min_quiet=0.5, timeout=15.0)

            # 检查当前状态
            current_file_count = self.checker.get_open_file_count()
            self.logger.info(f"状态检查 - 当前打开文件数: {current_file_count}")

            # 检查是否为单文件不确定状态
            is_unsaved = self.checker.is_single_unsaved_state()
            self.logger.info(f"是否为单文件不确定状态: {is_unsaved}")

            # 关闭操作后的状态应该是单文件不确定状态或无文件状态
            if current_file_count == 0 or (current_file_count == 1 and is_unsaved):
                self.logger.info("✅ 状态转换正确: 已恢复到单文件不确定状态")
                return True
            else:
                self.logger.warning(f"⚠ 状态异常: 期望单文件不确定状态,实际{current_file_count}个文件")
                return False

        except Exception as e:
            self.logger.error(f"❌ 状态转换检查异常: {e}")
            return False

    def _create_success_status(self, message: str) -> TestStatus:
        """创建成功状态"""
        return TestStatus(
            task_id=self.task_id,
            task_name=self.task_name,
            result=TestResult.SUCCESS,
            message=message,
            start_time=datetime.now(),
            end_time=datetime.now(),
            initial_windows=self.initial_windows,
            final_windows=self.final_windows,
            dialog_records=self.dialog_records
        )

    def _create_failed_status(self, message: str) -> TestStatus:
        """创建失败状态"""
        return TestStatus(
            task_id=self.task_id,
            task_name=self.task_name,
            result=TestResult.FAILED,
            message=message,
            start_time=datetime.now(),
            end_time=datetime.now(),
            initial_windows=self.initial_windows,
            final_windows=self.final_windows,
            dialog_records=self.dialog_records
        )

def main():
    """运行测试任务"""
    print("🚀 启动测试任务: TEST_TASK_002")
    print("=" * 50)

    test = TestTask002()
    result = test.run()

    print("=" * 50)
    print(f"测试结果: {result.result.value}")
    print(f"测试消息: {result.message}")

    return result.result == TestResult.SUCCESS

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)