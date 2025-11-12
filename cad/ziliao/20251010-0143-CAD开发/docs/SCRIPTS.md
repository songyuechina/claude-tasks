# 已登记脚本一览

- 本表由 `scripts/index.py` 扫描生成；建议每个脚本在顶部使用 # @script 元数据标注。

- `cad_cli.py` — CAD CLI Entrypoint
  - 描述: 启动/关闭CAD、DWG管理、基本操作（画线）、标准化状态
  - 命令: start, stop, open, list, close, save, saveas, line, start-open, standard
  - 用法: python cad_cli.py <command> [args]
- `scripts/layer_split_and_hatch.py` — Split By Layer And Hatch Groups
  - 描述: 针对一个 DWG：按图层拆分为多个文件；在每个文件中，查找封闭多段线，基于“最多两层包含”分组后做实体填充（SOLID）。
  - 用法: python scripts/layer_split_and_hatch.py [<source_dwg>]
- `scripts/test_flow.py` — CAD Smoke Test
  - 描述: 以标准化状态启动CAD，打开/保存/画线并列出DWG
  - 用法: python scripts/test_flow.py
- `scripts/verify_path_open.py` — Verify Path And Open DWG
  - 描述: 校验本地DWG路径（含中文/特殊字符）是否存在，并尝试用CadController打开
  - 用法: python scripts/verify_path_open.py <absolute_dwg_path>