# 即时任务执行命令

按照 `D:/claude-tasks/cad/即时对话.txt` 的内容执行任务。

## 执行流程

1. **读取即时对话.txt**
   ```
   D:/claude-tasks/cad/即时对话.txt
   ```

2. **遵循CLAUDE.md规范**
   - 读取 `cad/CLAUDE.md` 了解项目规范
   - 遵循函数编写规范（第六章）
   - 遵循测试规范（第七章）

3. **执行任务**
   - 按即时对话.txt中的要求执行
   - 使用 CAD_file_operations.py 进行文件操作
   - 使用 CAD_basic.py 进行对象操作
   - 测试前后调用 cad_zt_zero()

4. **完成汇报**
   - 按规范整理测试结果
   - 生成Excel测试报告
   - 生成Markdown测试报告
   - 归档测试DWG文件
   - 将代码上传到GitHub

## 重要提醒

- **都是用中文回复**
- **遵循1分钟原则**
- **测试失败先检查CAD进程数和弹窗处理函数**
- **所有操作必须符合CLAUDE.md规范**
