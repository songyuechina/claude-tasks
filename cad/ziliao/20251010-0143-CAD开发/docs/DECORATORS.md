# 协同与重试装饰器说明（codex 可读）

本文档描述本任务内用于 CAD 操作“协同机制 + 失败重试”的两个装饰器，以及在项目中的应用位置与用法建议。

- 位置：`utils/cad_decorators.py`
- 相关装饰器：
  - `cad_sync_retry`（新增，通用同步 + 重试）
  - `cad_operation`（原有，面向 COM 瞬时错误的重试 + 空闲等待）

## cad_sync_retry（推荐在 CLI/脚本层统一使用）
- 定义：`utils/cad_decorators.py:147`
- 语义：
  - 失败重试：当被装饰函数抛出异常时重试，默认最多 `5` 次、重试间隔 `1s`。
  - 协同机制：函数成功返回后，等待 CAD 空闲（quiescent）再返回，确保“上一操作完成后再执行下一指令”。
  - 控制器探测：优先从参数中寻找 `CadController` 实例（关键词参数名 `ctl` 或任意位置参数）；未找到则回退用 COM 方式检测空闲。
- 典型用法：
  ```python
  from utils.cad_decorators import cad_sync_retry

  @cad_sync_retry(max_attempts=5, delay=1.0, wait_idle=True)
  def my_op(ctl: CadController, ...):
      # 做一次 CAD 操作；异常将触发自动重试
      ...
  ```
- 关键特性：协同机制不依赖错误消息反馈，无论是否发生异常，只要成功返回都会等待 CAD 空闲。

## cad_operation（保留给底层 COM 操作函数）
- 定义：`utils/cad_decorators.py`（文件顶部已有说明）
- 语义：
  - 针对常见 COM 瞬时错误（如 Server Busy / RPC Call Rejected / RPC Unavailable）做指数回退式重试。
  - 成功后可选等待 CAD 空闲。
- 典型用法：
  ```python
  from utils.cad_decorators import cad_operation

  @cad_operation(retries=5, delay=1.0, wait_idle=True)
  def low_level_insert(...):
      ...  # 直接调用 win32com 接口
  ```

## 在项目中的应用
- CLI 层：`cad_cli.py`
  - 函数已直接加装饰器：`_cli_new`、`_cli_insert_all`、`_cli_insert_area`。
  - 在 `main()` 里对 `CadController` 实例的方法做统一包装，使所有 CLI 命令（`start/stop/open/list/close/save/saveas/new/insert-*/line/start-open/standard/standard2`）均具备“最多 5 次重试 + 成功后等待 CAD 空闲”的能力。
- 控制器层：`cad_automation/controller.py`
  - 控制器方法内部已包含空闲等待与多种稳健性处理；若在其它脚本中直接使用控制器实例，也可按需在调用处套上 `@cad_sync_retry`。

## 为新脚本启用（建议）
- 在脚本的关键步骤（如启动、打开、保存、插入、绘制等）使用 `@cad_sync_retry`。
- 若直接编写底层 COM 调用函数，可使用 `@cad_operation` 强化对瞬时错误的兼容性。

## 配置参数速览
- `max_attempts`：最大尝试次数（默认 5）。
- `delay`：重试前等待时间（默认 1.0 秒）。
- `wait_idle`：成功后是否等待 CAD 空闲（默认 True）。
- `idle_timeout`：等待空闲的最长时间（默认 30 秒，可按需调整）。

以上内容为整个开发文档配置的一部分；codex 可直接读取本文件与 `utils/cad_decorators.py` 获取实现与用法细节。
