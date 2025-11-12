# 概念与状态（CONCEPTS）

- 单文件不确定状态（Single Indeterminate）
  - 定义：单进程 + 仅 1 张未保存空白图；不关心是哪张图，必要时自动新建。
  - 入口：`CadController.single_unsaved_state()`；CLI/脚本已封装。
  - 使用场景：测试/任务前归位、恢复异常现场、保障后续顺序操作的“基线”。

- 单文件确定状态（Single Definite）
  - 定义：单进程 + 仅 1 张“指定/当前”DWG 保留在前台，其他全部关闭；
  - 入口：`CadController.standardize_state(target_doc_path=None)`。
  - 使用场景：只对一个确切目标图做操作。

- 双文件确定状态（Double Definite）
  - 定义：单进程 + 仅 2 张“指定”DWG 保留在前台；可选择激活 A/B；
  - 入口：`CadController.standardize_two_documents(a,b,active=None)` / `standardize_two_default()`。
  - 使用场景：需要 A/B 对照或跨图操作时。

- 不变量
  - 任意状态转换前，先确保进程约束、空闲与弹窗治理；
  - 所有“打开”均遵循：单次 Open + 等待加入集合 + 去重 + 顺序；
  - 路径/名称幂等：相同路径已打开或同名文档在集合中，则不再重复打开。
