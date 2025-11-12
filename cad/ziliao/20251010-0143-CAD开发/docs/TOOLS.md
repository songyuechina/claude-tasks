# 工具与脚本（TOOLS）

- CLI 命令（入口：`python cad_cli.py <cmd>`）
  - `new <path>`：幂等新建/打开；
  - `open <p1> [p2 ...]`：顺序打开；
  - `list`/`close`/`close <name>`/`save`/`saveas <path>`/`line x1 y1 z1 x2 y2 z2`；
  - `standard`/`standard2`/`start-open`。

- 脚本
  - `scripts/makepy_autocad.py`：生成 AutoCAD COM 缓存；
  - `scripts/do_single_unsaved.py`：归位到单不确定；
  - `scripts/insert_block_once.py`：以 123.dwg 为当前图把 ab.dwg 作为块插入原点（稳定、无炸开）；
  - `scripts/insert_area_align.py`：按矩形左下角对齐插入（对齐，不裁剪）；
  - `scripts/open_three_fixed.py`：顺序打开 3 个指定 DWG（示例）；
  - `scripts/run_test_workflow.py`：按规则串行跑“新建→绘制→插入→归位”的回归示例并输出日志。

- 日志
  - 统一位置：`artifacts/logs/<脚本名>/`；
  - 控制器封装日志：`artifacts/logs/cad_automation.log`。
