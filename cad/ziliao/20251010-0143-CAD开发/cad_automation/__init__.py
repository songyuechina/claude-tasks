# -*- coding: utf-8 -*-

"""
统一的CAD自动化本地封装入口。

其他脚本建议只从这里导入：
    from cad_automation.controller import CadController

保持该包内接口稳定，便于后续脚本复用。
"""

from .controller import CadController, load_cad_ops_module

__all__ = [
    "CadController",
    "load_cad_ops_module",
]

