# -*- coding: utf-8 -*-
"""
CAD测试框架

提供完整的测试任务执行环境,包括窗口监测、弹窗处理、状态检查等功能
"""

import time
import json
import win32gui
import win32process
import win32con
import psutil
import logging
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Optional, Any
from dataclasses import dataclass
from enum import Enum

# 导入CAD相关模块
from CAD_coordination import (
    wait_quiescent,
    wait_document_opened,
    ensure_single_process,
    start_cad_with_dialog_killer
)
from CAD_enhanced_functions import (
    start_cad_session_with_coordination,
    open_dwg_sync,
    open_dwg_enhanced
)

class TestResult(Enum):
    SUCCESS = "成功"
    FAILED = "失败"
    SKIPPED = "跳过"
    ERROR = "错误"

@dataclass
class WindowInfo:
    title: str
    class_name: str
    hwnd: int
    pid: int
    is_visible: bool
    timestamp: datetime

@dataclass
class TestStatus:
    task_id: str
    task_name: str
    result: TestResult
    message: str
    start_time: datetime
    end_time: datetime
    initial_windows: List[WindowInfo]
    final_windows: List[WindowInfo]
    dialog_records: List[Dict]

class CADWindowMonitor:
    """CAD窗口监测器"""

    def __init__(self):
        self.logger = logging.getLogger("CADWindowMonitor")
        self.dialog_records_file = Path(__file__).parent.parent / "tests" / "dialog_records.json"
        self.dialog_records_file.parent.mkdir(exist_ok=True)

    def capture_windows(self) -> List[WindowInfo]:
        """捕获当前所有窗口信息"""
        windows = []

        def enum_callback(hwnd, window_list):
            try:
                if win32gui.IsWindowVisible(hwnd):
                    window_title = win32gui.GetWindowText(hwnd)
                    if window_title:  # 只记录有标题的窗口
                        class_name = win32gui.GetClassName(hwnd)
                        _, pid = win32process.GetWindowThreadProcessId(hwnd)

                        window_info = WindowInfo(
                            title=window_title,
                            class_name=class_name,
                            hwnd=hwnd,
                            pid=pid,
                            is_visible=win32gui.IsWindowVisible(hwnd),
                            timestamp=datetime.now()
                        )
                        window_list.append(window_info)
            except Exception as e:
                self.logger.warning(f"获取窗口信息失败: {e}")
            return True

        win32gui.EnumWindows(enum_callback, windows)
        return windows

    def find_cad_windows(self) -> List[WindowInfo]:
        """查找CAD相关窗口"""
        windows = self.capture_windows()
        cad_windows = []

        for window in windows:
            # CAD相关窗口的特征
            if (any(keyword in window.title.lower() for keyword in ['autocad', '天正', 'tangent', 'cad']) or
                'acmd' in window.class_name.lower() or
                'acad' in window.class_name.lower()):
                cad_windows.append(window)

        return cad_windows

    def find_new_windows(self, initial: List[WindowInfo], current: List[WindowInfo]) -> List[WindowInfo]:
        """找出新增的窗口"""
        initial_hwnds = {w.hwnd for w in initial}
        new_windows = [w for w in current if w.hwnd not in initial_hwnds]
        return new_windows

    def record_dialog(self, dialog_info: Dict):
        """记录弹窗信息"""
        try:
            # 读取现有记录
            if self.dialog_records_file.exists():
                with open(self.dialog_records_file, 'r', encoding='utf-8') as f:
                    records = json.load(f)
            else:
                records = []

            # 添加新记录
            records.append(dialog_info)

            # 保存记录
            with open(self.dialog_records_file, 'w', encoding='utf-8') as f:
                json.dump(records, f, ensure_ascii=False, indent=2, default=str)

        except Exception as e:
            self.logger.error(f"记录弹窗信息失败: {e}")

    def close_window_by_title(self, title: str, exact_match: bool = False) -> bool:
        """根据标题关闭窗口"""
        try:
            windows = self.capture_windows()

            for window in windows:
                if exact_match:
                    if window.title == title:
                        win32gui.PostMessage(window.hwnd, win32con.WM_CLOSE, 0, 0)
                        return True
                else:
                    if title.lower() in window.title.lower():
                        win32gui.PostMessage(window.hwnd, win32con.WM_CLOSE, 0, 0)
                        return True

        except Exception as e:
            self.logger.error(f"关闭窗口失败: {e}")

        return False

    def force_close_error_windows(self) -> int:
        """强制关闭错误窗口"""
        error_keywords = ['错误', 'error', 'warning', '警告', '失败', 'failed']
        closed_count = 0

        windows = self.capture_windows()
        for window in windows:
            try:
                if any(keyword in window.title.lower() for keyword in error_keywords):
                    win32gui.PostMessage(window.hwnd, win32con.WM_CLOSE, 0, 0)
                    closed_count += 1

                    # 记录关闭的窗口
                    dialog_info = {
                        "timestamp": datetime.now().isoformat(),
                        "action": "强制关闭",
                        "window_title": window.title,
                        "window_class": window.class_name,
                        "handled_by": "test_framework"
                    }
                    self.record_dialog(dialog_info)

                    time.sleep(0.1)  # 等待窗口关闭

            except Exception as e:
                self.logger.warning(f"关闭错误窗口失败: {e}")

        return closed_count

class CADStateChecker:
    """CAD状态检查器"""

    def __init__(self):
        self.logger = logging.getLogger("CADStateChecker")

    def get_cad_process_count(self) -> int:
        """获取CAD进程数量"""
        count = 0
        for proc in psutil.process_iter(['pid', 'name']):
            if proc.info['name'] == 'acad.exe':
                count += 1
        return count

    def is_dialog_killer_running(self) -> bool:
        """检查弹窗治理脚本是否运行"""
        for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
            try:
                cmdline = proc.info.get('cmdline', [])
                if (proc.info['name'] == 'python.exe' and
                    any('cad_dialog_killer.py' in cmd for cmd in cmdline)):
                    return True
            except:
                continue
        return False

    def get_open_file_count(self) -> int:
        """获取当前打开的DWG文件数量"""
        try:
            import win32com.client
            acad = win32com.client.GetActiveObject("AutoCAD.Application")
            return acad.Documents.Count
        except:
            return 0

    def is_single_unsaved_state(self) -> bool:
        """检查是否单文件不确定状态"""
        try:
            file_count = self.get_open_file_count()
            if file_count != 1:
                return False

            import win32com.client
            acad = win32com.client.GetActiveObject("AutoCAD.Application")
            doc = acad.Documents.Item(0)

            # 检查是否为未保存状态(通常文件名为空或为DrawingX)
            name = doc.Name
            return not name or name.startswith('Drawing') or name == 'Drawing1'

        except:
            return False

    def is_file_opened(self, file_path: str) -> bool:
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

class TestTask:
    """测试任务基类"""

    def __init__(self, task_id: str, task_name: str, description: str = ""):
        self.task_id = task_id
        self.task_name = task_name
        self.description = description

        self.monitor = CADWindowMonitor()
        self.checker = CADStateChecker()
        self.logger = self._setup_logger()

        # 测试状态
        self.initial_windows: List[WindowInfo] = []
        self.final_windows: List[WindowInfo] = []
        self.dialog_records: List[Dict] = []

    def _setup_logger(self) -> logging.Logger:
        """设置日志"""
        log_dir = Path(__file__).parent.parent / "tests" / "logs" / self.task_id
        log_dir.mkdir(parents=True, exist_ok=True)

        logger = logging.getLogger(self.task_id)
        logger.setLevel(logging.INFO)

        # 清除现有处理器
        for handler in logger.handlers[:]:
            logger.removeHandler(handler)

        # 文件处理器
        file_handler = logging.FileHandler(
            log_dir / f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.log",
            encoding='utf-8'
        )
        file_handler.setFormatter(
            logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        )
        logger.addHandler(file_handler)

        # 控制台处理器
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(
            logging.Formatter(f'{self.task_id} - %(levelname)s - %(message)s')
        )
        logger.addHandler(console_handler)

        return logger

    def check_dialog_killer(self) -> bool:
        """检查弹窗治理脚本"""
        self.logger.info("检查弹窗治理脚本...")

        if self.checker.is_dialog_killer_running():
            self.logger.info("✅ 弹窗治理脚本正在运行")
            return True
        else:
            self.logger.error("❌ 弹窗治理脚本未运行")
            return False

    def capture_initial_windows(self):
        """记录初始窗口列表"""
        self.logger.info("记录初始窗口列表...")
        self.initial_windows = self.monitor.capture_windows()

        # 记录CAD窗口
        cad_windows = self.monitor.find_cad_windows()
        self.logger.info(f"发现 {len(cad_windows)} 个CAD相关窗口")

        # 保存窗口列表到日志
        window_list = [
            {
                "title": w.title,
                "class": w.class_name,
                "pid": w.pid,
                "timestamp": w.timestamp.isoformat()
            }
            for w in self.initial_windows
        ]

        with open(Path(__file__).parent.parent / "tests" / "logs" / self.task_id / "initial_windows.json",
                 'w', encoding='utf-8') as f:
            json.dump(window_list, f, ensure_ascii=False, indent=2)

    def check_cad_state(self) -> bool:
        """检查CAD状态"""
        self.logger.info("检查CAD状态...")

        # 检查进程数
        process_count = self.checker.get_cad_process_count()
        self.logger.info(f"CAD进程数: {process_count}")

        if process_count == 0:
            self.logger.error("❌ 未发现CAD进程")
            return False
        elif process_count > 1:
            self.logger.warning(f"⚠ 发现多个CAD进程 ({process_count} 个)")
            # 尝试确保单进程
            if not ensure_single_process():
                self.logger.error("❌ 无法确保单进程")
                return False

        # 检查打开文件数
        file_count = self.checker.get_open_file_count()
        self.logger.info(f"当前打开文件数: {file_count}")

        return True

    def set_initial_state(self) -> bool:
        """设置初始状态"""
        self.logger.info("设置初始状态...")

        try:
            # 启动CAD会话(如果需要)
            if self.checker.get_cad_process_count() == 0:
                if not start_cad_session_with_coordination():
                    self.logger.error("❌ CAD启动失败")
                    return False
                time.sleep(2.0)  # 等待启动完成

            # 等待CAD空闲
            wait_quiescent(min_quiet=1.0, timeout=30.0)

            # 关闭额外窗口
            closed_count = self.monitor.force_close_error_windows()
            if closed_count > 0:
                self.logger.info(f"✅ 关闭了 {closed_count} 个错误窗口")

            self.logger.info("✅ 初始状态设置完成")
            return True

        except Exception as e:
            self.logger.error(f"❌ 设置初始状态失败: {e}")
            return False

    def capture_final_windows(self):
        """记录最终窗口列表"""
        self.logger.info("记录最终窗口列表...")
        self.final_windows = self.monitor.capture_windows()

        # 找出新增窗口
        new_windows = self.monitor.find_new_windows(self.initial_windows, self.final_windows)
        if new_windows:
            self.logger.info(f"发现 {len(new_windows)} 个新增窗口")

            # 记录新增窗口
            for window in new_windows:
                dialog_info = {
                    "timestamp": window.timestamp.isoformat(),
                    "test_task": self.task_id,
                    "phase": "测试结束",
                    "window_title": window.title,
                    "window_class": window.class_name,
                    "hwnd": window.hwnd,
                    "pid": window.pid,
                    "action": "记录"
                }
                self.dialog_records.append(dialog_info)
                self.monitor.record_dialog(dialog_info)

        # 保存窗口列表
        window_list = [
            {
                "title": w.title,
                "class": w.class_name,
                "pid": w.pid,
                "timestamp": w.timestamp.isoformat()
            }
            for w in self.final_windows
        ]

        with open(Path(__file__).parent.parent / "tests" / "logs" / self.task_id / "final_windows.json",
                 'w', encoding='utf-8') as f:
            json.dump(window_list, f, ensure_ascii=False, indent=2)

    def restore_cad_state(self) -> bool:
        """恢复CAD状态"""
        self.logger.info("恢复CAD状态...")

        try:
            # 确保单进程
            ensure_single_process()

            # 关闭所有文件
            from CAD_coordination import send_cmd_with_sync
            try:
                send_cmd_with_sync("_CLOSE\n", wait_after=0.5)
            except:
                pass

            # 等待CAD空闲
            wait_quiescent(min_quiet=0.5, timeout=15.0)

            self.logger.info("✅ CAD状态已恢复")
            return True

        except Exception as e:
            self.logger.error(f"❌ 恢复CAD状态失败: {e}")
            return False

    def record_dialogs(self):
        """记录弹窗信息"""
        if self.dialog_records:
            self.logger.info(f"记录 {len(self.dialog_records)} 个弹窗信息")

            # 保存到测试日志
            with open(Path(__file__).parent.parent / "tests" / "logs" / self.task_id / "dialogs.json",
                     'w', encoding='utf-8') as f:
                json.dump(self.dialog_records, f, ensure_ascii=False, indent=2, default=str)

    def execute(self) -> TestStatus:
        """执行测试任务 - 子类实现"""
        raise NotImplementedError

    def run(self) -> TestStatus:
        """运行测试任务"""
        start_time = datetime.now()
        self.logger.info(f"开始执行测试任务: {self.task_name} ({self.task_id})")

        try:
            # 测试准备阶段
            if not self.setup():
                return TestStatus(
                    task_id=self.task_id,
                    task_name=self.task_name,
                    result=TestResult.FAILED,
                    message="测试准备失败",
                    start_time=start_time,
                    end_time=datetime.now(),
                    initial_windows=self.initial_windows,
                    final_windows=self.final_windows,
                    dialog_records=self.dialog_records
                )

            # 执行测试
            status = self.execute()
            status.start_time = start_time
            status.initial_windows = self.initial_windows
            status.final_windows = self.final_windows
            status.dialog_records = self.dialog_records

            # 测试清理阶段
            self.teardown()

            end_time = datetime.now()
            duration = (end_time - start_time).total_seconds()
            self.logger.info(f"测试任务完成: {status.result.value} (耗时: {duration:.2f}秒)")

            return status

        except Exception as e:
            end_time = datetime.now()
            self.logger.error(f"测试任务异常: {e}")

            return TestStatus(
                task_id=self.task_id,
                task_name=self.task_name,
                result=TestResult.ERROR,
                message=f"测试异常: {e}",
                start_time=start_time,
                end_time=end_time,
                initial_windows=self.initial_windows,
                final_windows=self.final_windows,
                dialog_records=self.dialog_records
            )

    def setup(self) -> bool:
        """测试准备"""
        # 1. 检查弹窗治理
        if not self.check_dialog_killer():
            return False

        # 2. 记录初始窗口
        self.capture_initial_windows()

        # 3. 检查CAD状态
        if not self.check_cad_state():
            return False

        # 4. 设置初始状态
        return self.set_initial_state()

    def teardown(self):
        """测试清理"""
        # 1. 记录最终窗口
        self.capture_final_windows()

        # 2. 恢复CAD状态
        self.restore_cad_state()

        # 3. 记录弹窗信息
        self.record_dialogs()

if __name__ == "__main__":
    # 测试框架示例
    logging.basicConfig(level=logging.INFO)

    # 测试窗口监测
    monitor = CADWindowMonitor()
    windows = monitor.capture_windows()
    print(f"发现 {len(windows)} 个窗口")

    # 测试状态检查
    checker = CADStateChecker()
    print(f"CAD进程数: {checker.get_cad_process_count()}")
    print(f"弹窗治理脚本运行: {checker.is_dialog_killer_running()}")