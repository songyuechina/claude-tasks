# 打开 DWG 的协同与约束（Open Policy）

本文约束本项目中“打开 DWG 文件”的统一做法，确保顺序、幂等与稳定，避免出现多个只读副本。

## 总体目标
- 统一由天正 V9 + CAD2021 环境启动（CAD 基本操作.py 的 `start_applicationV9`）。
- 同一文件只触发一次 `Documents.Open()`；等待其真正加入 `acad.Documents` 之后再进行下一步。
- 打开多个文件时严格顺序、不可并发，逐个确认完成与空闲后再继续。
- 每个“独立测试”结束后，建议将 CAD 环境归位为“单文件不确定状态”（调用 `single_unsaved_state()`），再开始下一项测试，避免状态串扰。

## 预处理与进程约束
- 默认（非破坏性）
  - 进程数 `cnt == 0`：调用 `start_tarch_v9()` 启动；
  - `cnt > 1`：调用 `ensure_single_process()` 尽量收敛到单进程（不强制重启）；
  - `cnt == 1`：仅确保空闲并短暂等待（≈0.3s）。
- 严格场景（如进入“双文件确定状态”前）：
  - 使用 `single_unsaved_state()`：关闭全部进程 → `start_applicationV9(PTH, max_retries=3, retry_delay=2.0)` → 等待稳定（函数内自带 1–3s 稳定窗口）。

## 启动与弹窗
- 启动一律通过 `start_applicationV9(PTH=r"C:\\Tangent\\TArchT20V9", max_retries=3, retry_delay=2.0)`。
- 启动后检测天正进程是否存在（如 `tarcht20v9.exe`），未检测到则再补一次启动。
- 在关键节点调用 `_close_error_windows()` 尝试关闭“AutoCAD 错误报告”等干扰弹窗。

## 单个文件：`open_dwg(path)`
- 仅执行一次 `Documents.Open(short_path)`；随后调用 `wait_document_opened(path, timeout=self.open_timeout_s)` 等待文档加入集合；
- 超时回退：按“文件名”再等待一小段时间，兼容中文/特殊字符路径导致的 `FullName` 差异；
- 成功后 `wait_quiescent`，并激活该文档；
- 打开前自动进行“非破坏性的进程保证”。

## 多个文件：`open_dwgs(paths)`（严格顺序）
- 对每个目标：
  1) 若相同路径已打开 → 跳过；
  2) 若同名文档已存在 → 跳过（避免只读副本）；
  3) 调用 `open_dwg`，成功后 `wait_quiescent` 并插入 ≈0.3s 的缓冲；
- 打开前进行“非破坏性的进程保证”。

注意：打开前的约束（用于测试场景）
- 先检测当前 CAD 进程数是否不大于 1；若大于 1，先 `single_unsaved_state()` 并等待 3 秒再打开；
- 打开目标前检查是否已打开：
  - 路径级判断：目标绝对路径是否已在 `acad.Documents`；
  - 名称级判断：是否已存在同名文档（basename 相同）。
  满足任一条件都不再重复打开。

## 稳定等待（Intervals）
- `single_unsaved_state()` 结束后保留 1–3 秒的稳定窗口（默认 2s）。
- 连续打开多个文件时，每次成功后等待空闲 + 0.3s。

## 幂等与去重
- 路径级幂等：若目标绝对路径已在 `acad.Documents` 中，直接激活返回；
- 名称级保护：若存在同名文档（basename 一致），默认跳过以避免只读副本。

## 相关实现位置
- 控制器：`cad_automation/controller.py`
  - `single_unsaved_state()`、`open_dwg()`、`open_dwgs()`、`_ensure_one_process_before_open()`、`_close_error_windows()` 等。
- 脚本示例：
  - `scripts/open_three_fixed.py`（按顺序打开三个指定 DWG）。

> 以上约束为后续任务的统一规范：若发现某脚本/流程与此不符，请先调整到该策略以获得一致与稳定的行为。
