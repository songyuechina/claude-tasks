# 具体操作（OPERATIONS）

- 新建（幂等）
  - 入口：CLI `python cad_cli.py new <path>`；控制器：新建后可 `save_as(path)`。
  - 规则：当 `<path>` 已存在时不再新建，直接打开（已在 CLI 实现）。
  - 后置：`wait_quiescent`；必要时 `standardize_state`（仅保留 1 张）。

- 打开（顺序 + 去重）
  - 单个：`CadController.open_dwg(path)`（单次 Open + 等待加入集合）。
  - 多个：`CadController.open_dwgs(paths)`（严格顺序；路径/同名已存在则跳过；每次成功后等待空闲 + 0.3s）。
  - 前置：`_ensure_one_process_before_open()`（非破坏）。

- 关闭
  - 关闭当前：`cad_cli.py close` / `CadController.close_current()`；
  - 按名关闭：`cad_cli.py close <name>` / `close_by_name(name)`。

- 保存
  - `cad_cli.py save` / `CadController.save()`；另存：`save_as(path)`（内部使用短路径）。

- 插入
  - 稳定路径：使用 `SendCommand("-INSERT ...")` 完成块插入并 `wait_quiescent`；
  - 避免外部脚本在控制台输出 Unicode 导致编码异常；
  - 若需“炸开/裁剪区域”，建议实现“静默版 explode + 选择集”再启用。

- 状态收敛
  - 单不确定：`single_unsaved_state()`；
  - 单确定：`standardize_state(target)`；
  - 双确定：`standardize_two_documents(a,b,active)` / `standardize_two_default()`。

- 失败处理
  - 先 `_close_error_windows()`；再 `wait_quiescent`；仍不行时归位到单不确定后重试当前步。
