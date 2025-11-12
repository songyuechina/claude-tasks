# CAD 自动化控制规范（索引）

本目录汇总“在本任务中编写/测试涉及 CAD 文件操作时”的统一规则、概念、工具与实践清单，供后续 codex 阅读交接使用。

- 概念与状态
  - [CONCEPTS.md](CONCEPTS.md)：单文件不确定状态、单文件确定状态、双文件确定状态等核心概念与不变量。
- 协同与时序
  - [COORDINATION.md](COORDINATION.md)：等待/空闲/进程收敛/弹窗治理/编码兼容/路径规范等通用协同机制。
- 具体操作规则
  - [OPEN_POLICY.md](OPEN_POLICY.md)：打开/顺序/去重/进程校验的唯一规则来源。
  - [OPERATIONS.md](OPERATIONS.md)：新建/打开/关闭/保存/插入等操作的前置条件、后置条件与 CLI/控制器映射。
- 实用工具与脚本
  - [TOOLS.md](TOOLS.md)：脚本与 CLI 清单、用法、输出位置与日志说明。
- 参考与示例
  - [STATE_CONTROL.md](STATE_CONTROL.md)：状态收敛与策略摘要（指向 OPEN_POLICY）。
  - examples/[TEST_SCENARIO_123_AB.md](examples/TEST_SCENARIO_123_AB.md)（示例场景，非规则）。

> 必读顺序建议：CONCEPTS → COORDINATION → OPEN_POLICY → OPERATIONS → TOOLS。
