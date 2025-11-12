# CAD 界面状态控制（快速总览）

本文汇总“收敛 CAD 界面状态”的核心函数、命令与脚本，便于在本目录被重新打开后快速上手控制本机 CAD 状态。

## 核心函数（CadController）
- 单文件不确定（未保存）状态：`single_unsaved_state()`
  - 位置：`cad_automation/controller.py:456`
  - 保障：单进程 + 单文档；必要时新建一张未保存空白图（不关心是哪张图）。
  - 成功打印标记：`[single-indeterminate] <Name>`；返回激活文件名。
- 单文件确定（稳定）状态：`standardize_state(target_doc_path=None)`
  - 位置：`cad_automation/controller.py:456`
  - 保障：单进程 + 单文档；目标可指定文件路径，或不传则保留当前激活图。
  - 成功打印标记：`[single] <Name>`；返回激活文件名。
- 双文件确定状态：`standardize_two_documents(a, b, active=None)`
  - 位置：`cad_automation/controller.py:532`
  - 保障：单进程 + 仅保留 A/B 两个目标；`active` 可选择激活 a 或 b。
- 双文件确定（默认文件）：`standardize_two_default(active=None)`
  - 位置：`cad_automation/controller.py:589`
  - 使用 `artifacts/dwgs/00.dwg` 与 `01.dwg`，便于快速演示/测试。

以上函数均在成功路径后等待 CAD 空闲（CLI 中还有统一“成功后等待 + 异常最多重试 5 次”的包装）。

## CLI 映射（开箱即用）
- 单文件不确定（未保存）：`python cad_cli.py new`
  - 传路径：`python cad_cli.py new D:/out/A.dwg`（新建后另存为）。
- 单文件确定：`python cad_cli.py standard [<dwg>]`
- 双文件确定：`python cad_cli.py standard2 <A.dwg> <B.dwg> [a|b]`
  - 无参时默认使用 `artifacts/dwgs/00.dwg` 与 `01.dwg`。

## 打开文件策略（摘要）
- 启动一律使用 `start_applicationV9`（天正 V9 + CAD2021）；
- `open_dwg` 只发起一次 `Documents.Open()`，等待加入集合；
- `open_dwgs` 严格顺序：路径/同名已存在则跳过；每次成功后等待空闲 + 0.3s；
- 进程约束默认非破坏：`cnt==0` 启动，`cnt>1` 收敛到单进程，`cnt==1` 仅等待空闲；
- 严格场景（如双文件确定）才使用 `single_unsaved_state()` 收敛；
- 详细规范见 `docs/OPEN_POLICY.md`（唯一规则来源）。
- 示例测试场景（非规则）：`docs/examples/TEST_SCENARIO_123_AB.md`。

## 辅助脚本
- 强制单文件未保存：`scripts/do_single_unsaved.py`
  - 打印 `ACTIVE` 与 `OPEN` 列表，便于人工确认。
- 检查并收敛至单未保存：`scripts/check_state_single_unsaved.py`
  - 产生日志与 JSON 摘要：`artifacts/logs/check_single_unsaved/`。
- 更新脚本清单与文档：`python scripts/index.py`
  - 生成 `artifacts/scripts_manifest.json` 与 `docs/SCRIPTS.md`。

## 日志与排障
- 控制器封装日志：`artifacts/logs/cad_automation.log`
- 脚本运行日志：`artifacts/logs/<脚本名>/`

若环境路径包含中文/空格等字符，请在命令行中加引号；打开容忍时间默认 120s，详见 `docs/RUN_SETTINGS.md`。
