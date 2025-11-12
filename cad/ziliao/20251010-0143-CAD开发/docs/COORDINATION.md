# 协同与时序（COORDINATION）

- 进程约束
  - 默认（非破坏）：`cnt==0` 启动；`cnt>1` 收敛为单进程；`cnt==1` 等待空闲；
  - 严格场景：调用 `single_unsaved_state()`（关闭全部→`start_applicationV9`→稳定等待 1–3s）。
- 等待策略
  - `wait_quiescent(min_quiet≈0.4s, timeout)`：CAD 空闲再继续；
  - `wait_document_opened(path)`：文档加入 `acad.Documents` 后才认为打开成功；
  - 连续打开间插 `≈0.3s` 的微等待以消抖。
- 弹窗治理
  - 典型：AutoCAD 错误报告、缺少 SHX 等；
  - 处理：`_close_error_windows()`（激活标题→尝试点击关闭/Alt+F4）。
- 路径与编码
  - 路径含中文/空格/特殊字符请加引号；必要时使用 Win32 短路径；
  - 控制台建议 UTF-8；若外部脚本打印 Unicode 导致 GBK 报错，优先使用 SendCommand 的“静默”路径。
- MakePy
  - 首次使用 COM 前：`python scripts/makepy_autocad.py` 生成缓存，避免交互式 makepy。
