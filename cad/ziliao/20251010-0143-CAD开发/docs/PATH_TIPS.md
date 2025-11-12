# 路径与字符集提示

- 包含中文、空格、`#`、`[]`、`【】` 等字符的 DWG 路径，命令行请务必整体加引号：
  - 示例：`python cad_cli.py open "D:/路径/3#栋[电气]_t3.dwg"`
- 控制器在打开与另存时会自动尝试转换为 Windows 短路径，以提升与 COM 的兼容性（已内置于 `CadController.open_dwg/save_as`）。
- 如仍有“找不到路径/文件”的情况，可使用自检脚本定位问题：
  - `python scripts/verify_path_open.py <绝对DWG路径>`
  - 输出运行日志与 JSON 总结到 `artifacts/logs/verify_path_open/`。

