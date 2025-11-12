# layer_split_and_hatch 专用验收标准

在满足《docs/TESTING_PRINCIPLES.md》的通用法则基础上，本脚本增加如下专用判定：

1) 层数-文件数一致性
- 以源 DWG 中“有对象的图层数”（忽略以 `$` 开头的系统层）作为期望值；
- 实际生成的 DWG 文件数应与期望一致。

2) 组计数（封闭多段线分组）
- 在每个分解后的文件中，统计“封闭多段线按包含关系分组”的组数；
- 每组由一条外多段线及其最近包含的内多段线集合构成。

3) 组填充结果
- 对每一组，执行填充（SOLID）；只要填充过程未反馈错误，即视为该组填充成功；
- 所有组的填充均无错误时，脚本的专用判定通过。

脚本在 `artifacts/logs/layer_split_and_hatch/` 生成 `*.json` 总结，包含：
- `expected_layers`、`produced`、`file_stats`（含每个文件的 `groups_count`、`hatch_ok`、`hatch_err`）等字段；
- `general_ok`（通用法则）、`specific_ok`（专用判定）与最终 `verdict`。
