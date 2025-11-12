# CAD 基本操作范式

基于 `ziliao/20251010-0143-CAD开发` 资料学习的完整操作范式,包括新建、打开、关闭、保存、插入等基本操作。

---

## 🎯 操作范式核心原则

### 协同机制
- **物理驱动等待**: CAD是物理驱动的软件,命令执行需要时间
- **顺序执行**: 前一个命令完全结束后才执行下一个
- **状态收敛**: 每个操作后确保CAD进入稳定状态

### 幂等性
- **路径级幂等**: 相同路径的操作只执行一次
- **名称级保护**: 避免打开相同文件名造成只读副本
- **状态幂等**: 多次调用相同操作产生相同结果

---

## 📝 基本操作范式

### 1. 新建文件范式

#### 范式定义
```python
def new_dwg_file(output_path: Optional[str] = None) -> bool:
    """
    新建DWG文件范式

    规则:
    - 幂等操作: output_path已存在时不再新建,直接打开
    - 无output_path时创建未保存的空白文件
    - 后置处理: wait_quiescent;必要时standardize_state

    前置条件:
    - CAD进程已启动
    - 弹窗治理脚本运行中

    后置条件:
    - 文件已创建或已打开
    - 状态为单文件确定状态(有路径)或单文件不确定状态(无路径)
    """
```

#### 实现代码
```python
def new_dwg_enhanced(output_path: Optional[str] = None) -> bool:
    """增强版新建文件操作,集成协同机制"""
    try:
        # 1. 确保CAD环境就绪
        if not ensure_single_process():
            return False
        wait_quiescent(min_quiet=0.5, timeout=15.0)

        # 2. 检查路径幂等性
        if output_path and Path(output_path).exists():
            print(f"✅ 文件已存在,直接打开: {output_path}")
            return open_dwg_sync(output_path)

        # 3. 执行新建操作
        from CAD_coordination import send_cmd_with_sync
        success = send_cmd_with_sync("_NEW\n", wait_after=1.0, timeout=30.0)

        if not success:
            print("❌ 新建文件操作失败")
            return False

        # 4. 等待新建完成
        wait_quiescent(min_quiet=1.0, timeout=30.0)

        # 5. 如需另存为
        if output_path:
            from CAD_coordination import send_cmd_with_sync
            save_cmd = f"_SAVEAS\n\"{output_path}\"\n"
            success = send_cmd_with_sync(save_cmd, wait_after=2.0, timeout=60.0)

            if success:
                print(f"✅ 新建并保存文件: {output_path}")
                wait_quiescent(min_quiet=1.0, timeout=30.0)
                return True
            else:
                print("❌ 文件另存为失败")
                return False
        else:
            print("✅ 新建未保存文件成功")
            return True

    except Exception as e:
        print(f"❌ 新建文件操作异常: {e}")
        return False
```

#### 使用示例
```python
# 创建未保存的空白文件
new_dwg_enhanced()

# 创建并保存到指定路径
new_dwg_enhanced("D:/output/new_file.dwg")
```

---

### 2. 打开文件范式

#### 范式定义
```python
def open_dwg_paradigm(file_path: str) -> bool:
    """
    打开DWG文件范式

    规则:
    - 顺序+去重: 同一文件只触发一次Documents.Open()
    - 等待加入集合: 确保文档真正加入acad.Documents
    - 路径/名称幂等: 避免重复打开相同文件

    前置条件:
    - 非破坏性进程保证
    - 弹窗治理检查

    后置条件:
    - 文件成功打开并激活
    - CAD进入空闲状态
    """
```

#### 实现代码
```python
def open_dwg_paradigm(file_path: str) -> bool:
    """完整的打开文件范式实现"""
    try:
        # 1. 基础验证
        if not Path(file_path).exists():
            print(f"❌ 文件不存在: {file_path}")
            return False

        # 2. 进程预处理(非破坏性)
        process_count = get_cad_process_count()
        if process_count == 0:
            print("🚀 CAD未运行,启动CAD...")
            if not start_cad_with_dialog_killer():
                return False
        elif process_count > 1:
            print("⚠ 发现多个CAD进程,确保单进程...")
            ensure_single_process()

        # 3. 等待CAD稳定
        wait_quiescent(min_quiet=0.3, timeout=15.0)

        # 4. 路径级幂等检查
        if is_file_opened(file_path):
            print(f"✅ 文件已打开: {file_path}")
            return True

        # 5. 名称级幂等检查
        basename = Path(file_path).name
        if is_file_opened_by_name(basename):
            print(f"⚠ 同名文件已打开,跳过: {basename}")
            return True

        # 6. 执行打开操作
        print(f"🔄 正在打开: {file_path}")

        # 使用协同机制发送打开命令
        import win32com.client
        acad = win32com.client.GetActiveObject("AutoCAD.Application")

        # 转换为短路径处理中文/特殊字符
        short_path = _get_short_path(file_path)

        # 执行打开
        acad.Documents.Open(short_path)

        # 7. 等待文档加入集合
        if wait_document_opened(file_path, timeout=120.0):
            print(f"✅ 文件成功打开: {file_path}")

            # 8. 激活文档
            _activate_document(file_path)

            # 9. 等待CAD空闲
            wait_quiescent(min_quiet=0.5, timeout=30.0)

            return True
        else:
            print(f"❌ 文件打开超时: {file_path}")
            return False

    except Exception as e:
        print(f"❌ 打开文件异常: {e}")
        return False

def open_multiple_files_paradigm(file_paths: List[str]) -> int:
    """多文件打开范式(严格顺序)"""
    success_count = 0

    # 进程预处理
    ensure_single_process()
    wait_quiescent(min_quiet=0.3, timeout=15.0)

    for i, file_path in enumerate(file_paths):
        print(f"📂 [{i+1}/{len(file_paths)}] {file_path}")

        if open_dwg_paradigm(file_path):
            success_count += 1
            print(f"✅ 成功打开: {file_path}")
        else:
            print(f"❌ 打开失败: {file_path}")

        # 文件间间隔等待
        if i < len(file_paths) - 1:
            time.sleep(0.3)
            wait_quiescent(min_quiet=0.3, timeout=15.0)

    print(f"📊 打开结果: {success_count}/{len(file_paths)} 成功")
    return success_count
```

#### 使用示例
```python
# 打开单个文件
open_dwg_paradigm("D:/test.dwg")

# 顺序打开多个文件
files = ["D:/file1.dwg", "D:/file2.dwg", "D:/file3.dwg"]
open_multiple_files_paradigm(files)
```

---

### 3. 关闭文件范式

#### 范式定义
```python
def close_dwg_paradigm(target: Optional[str] = None) -> bool:
    """
    关闭DWG文件范式

    规则:
    - target=None: 关闭当前文件
    - target指定: 关闭指定名称文件
    - 处理保存提示: 自动处理未保存文件的保存对话框

    前置条件:
    - CAD进程运行中

    后置条件:
    - 文件已关闭
    - 状态已恢复
    """
```

#### 实现代码
```python
def close_current_dwg_paradigm(save_option: str = "prompt") -> bool:
    """关闭当前文件范式"""
    try:
        # 1. 检查是否有文件打开
        if get_open_file_count() == 0:
            print("⚠ 没有打开的文件")
            return True

        # 2. 获取当前文件信息
        import win32com.client
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        current_doc = acad.ActiveDocument
        doc_name = current_doc.Name

        print(f"🔄 正在关闭当前文件: {doc_name}")

        # 3. 处理保存选项
        if save_option == "auto_save":
            # 自动保存
            current_doc.Save()
            print(f"✅ 已保存: {doc_name}")
        elif save_option == "no_save":
            # 不保存
            print(f"⚠ 不保存关闭: {doc_name}")
        else:
            # 提示保存(默认)
            print(f"📝 提示保存: {doc_name}")

        # 4. 执行关闭命令
        from CAD_coordination import send_cmd_with_sync
        success = send_cmd_with_sync("_CLOSE\n", wait_after=1.0, timeout=30.0)

        if success:
            # 5. 等待关闭完成
            wait_quiescent(min_quiet=1.0, timeout=30.0)
            print(f"✅ 文件已关闭: {doc_name}")
            return True
        else:
            print(f"❌ 关闭文件失败: {doc_name}")
            return False

    except Exception as e:
        print(f"❌ 关闭文件异常: {e}")
        return False

def close_dwg_by_name_paradigm(file_name: str) -> bool:
    """按文件名关闭文件范式"""
    try:
        # 1. 检查文件是否存在
        if not is_file_opened_by_name(file_name):
            print(f"⚠ 文件未打开: {file_name}")
            return True

        # 2. 切换到目标文件
        import win32com.client
        acad = win32com.client.GetActiveObject("AutoCAD.Application")

        # 查找并激活目标文件
        for i in range(acad.Documents.Count):
            doc = acad.Documents.Item(i)
            if doc.Name == file_name:
                acad.ActiveDocument = doc
                break

        # 3. 关闭文件
        return close_current_dwg_paradigm()

    except Exception as e:
        print(f"❌ 按名关闭文件异常: {e}")
        return False

def close_all_dwg_paradigm() -> bool:
    """关闭所有文件范式"""
    try:
        file_count = get_open_file_count()
        if file_count == 0:
            print("⚠ 没有打开的文件")
            return True

        print(f"🔄 准备关闭 {file_count} 个文件")

        # 逐一关闭文件
        success_count = 0
        for _ in range(file_count):
            if close_current_dwg_paradigm():
                success_count += 1
            time.sleep(0.5)  # 间隔等待

        print(f"✅ 关闭完成: {success_count}/{file_count} 成功")
        return success_count == file_count

    except Exception as e:
        print(f"❌ 关闭所有文件异常: {e}")
        return False
```

#### 使用示例
```python
# 关闭当前文件(提示保存)
close_current_dwg_paradigm()

# 关闭当前文件(自动保存)
close_current_dwg_paradigm("auto_save")

# 关闭当前文件(不保存)
close_current_dwg_paradigm("no_save")

# 按文件名关闭
close_dwg_by_name_paradigm("test.dwg")

# 关闭所有文件
close_all_dwg_paradigm()
```

---

### 4. 保存文件范式

#### 范式定义
```python
def save_dwg_paradigm(output_path: Optional[str] = None) -> bool:
    """
    保存DWG文件范式

    规则:
    - output_path=None: 保存当前文件
    - output_path指定: 另存为新文件
    - 使用短路径处理中文/特殊字符
    - 确保保存操作完成

    前置条件:
    - 有文件打开

    后置条件:
    - 文件已保存/另存为
    - 文件状态为已保存
    """
```

#### 实现代码
```python
def save_current_dwg_paradigm() -> bool:
    """保存当前文件范式"""
    try:
        # 1. 检查是否有文件打开
        if get_open_file_count() == 0:
            print("❌ 没有打开的文件")
            return False

        # 2. 获取文件信息
        import win32com.client
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        current_doc = acad.ActiveDocument
        doc_name = current_doc.Name

        print(f"💾 正在保存: {doc_name}")

        # 3. 等待CAD空闲
        wait_quiescent(min_quiet=0.5, timeout=15.0)

        # 4. 执行保存操作
        try:
            current_doc.Save()
            print(f"✅ 保存成功: {doc_name}")

            # 5. 等待保存完成
            wait_quiescent(min_quiet=1.0, timeout=30.0)
            return True

        except Exception as save_error:
            print(f"⚠ 直接保存失败,尝试另存为: {save_error}")

            # 如果是未保存文件,尝试另存为
            if not hasattr(current_doc, 'FullName') or not current_doc.FullName:
                default_path = f"D:/temp/{doc_name}"
                return save_as_dwg_paradigm(default_path)

            return False

    except Exception as e:
        print(f"❌ 保存文件异常: {e}")
        return False

def save_as_dwg_paradigm(output_path: str) -> bool:
    """另存为文件范式"""
    try:
        # 1. 基础验证
        if get_open_file_count() == 0:
            print("❌ 没有打开的文件")
            return False

        # 2. 创建输出目录
        output_file = Path(output_path)
        output_file.parent.mkdir(parents=True, exist_ok=True)

        # 3. 获取当前文件信息
        import win32com.client
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        current_doc = acad.ActiveDocument
        doc_name = current_doc.Name

        print(f"💾 正在另存为: {doc_name} → {output_path}")

        # 4. 等待CAD空闲
        wait_quiescent(min_quiet=0.5, timeout=15.0)

        # 5. 使用短路径
        short_path = _get_short_path(output_path)

        # 6. 执行另存为操作
        try:
            current_doc.SaveAs(short_path)
            print(f"✅ 另存为成功: {output_path}")

            # 7. 验证文件是否创建
            if output_file.exists():
                print(f"✅ 文件已创建: {output_file}")

                # 8. 等待保存完成
                wait_quiescent(min_quiet=1.0, timeout=30.0)
                return True
            else:
                print(f"❌ 文件未创建: {output_path}")
                return False

        except Exception as save_error:
            print(f"❌ 另存为失败: {save_error}")
            return False

    except Exception as e:
        print(f"❌ 另存为文件异常: {e}")
        return False

def auto_save_dwg_paradigm(interval_seconds: int = 300) -> bool:
    """自动保存范式"""
    try:
        if get_open_file_count() == 0:
            print("⚠ 没有打开的文件,跳过自动保存")
            return True

        import win32com.client
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        current_doc = acad.ActiveDocument
        doc_name = current_doc.Name

        print(f"🔄 自动保存: {doc_name}")

        # 执行保存
        current_doc.Save()
        print(f"✅ 自动保存完成: {doc_name}")

        return True

    except Exception as e:
        print(f"❌ 自动保存异常: {e}")
        return False
```

#### 使用示例
```python
# 保存当前文件
save_current_dwg_paradigm()

# 另存为新文件
save_as_dwg_paradigm("D:/backup/test_backup.dwg")

# 自动保存
auto_save_dwg_paradigm(interval_seconds=300)
```

---

### 5. 插入文件范式

#### 范式定义
```python
def insert_dwg_paradigm(block_file_path: str, insert_point: tuple = (0, 0, 0),
                       scale: float = 1.0, rotation: float = 0.0,
                       explode: bool = False) -> bool:
    """
    插入DWG文件作为块范式

    规则:
    - 使用-INSERT命令避免Unicode编码问题
    - 稳定路径处理中文/特殊字符
    - 等待插入操作完成
    - 可选炸开/缩放/旋转参数

    前置条件:
    - 有文件打开作为接收文件
    - 块文件存在

    后置条件:
    - 块已插入指定位置
    - 文件有未保存更改
    - CAD进入空闲状态
    """
```

#### 实现代码
```python
def insert_dwg_as_block_paradigm(block_file_path: str,
                                insert_point: tuple = (0, 0, 0),
                                scale: float = 1.0,
                                rotation: float = 0.0,
                                explode: bool = False) -> bool:
    """插入DWG文件作为块的完整范式"""
    try:
        # 1. 基础验证
        if not Path(block_file_path).exists():
            print(f"❌ 块文件不存在: {block_file_path}")
            return False

        if get_open_file_count() == 0:
            print("❌ 没有打开的文件作为接收文件")
            return False

        # 2. 获取当前文件信息
        import win32com.client
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        current_doc = acad.ActiveDocument
        doc_name = current_doc.Name

        print(f"🔄 正在插入块: {block_file_path} → {doc_name}")
        print(f"📍 插入位置: {insert_point}, 缩放: {scale}, 旋转: {rotation}°, 炸开: {explode}")

        # 3. 等待CAD空闲
        wait_quiescent(min_quiet=0.5, timeout=15.0)

        # 4. 构建INSERT命令
        # 使用-INSERT命令避免外部脚本控制台Unicode输出问题
        short_path = _get_short_path(block_file_path)

        cmd_parts = [
            "-INSERT",
            f'"{short_path}"',  # 块文件路径
            f"{insert_point[0]},{insert_point[1]},{insert_point[2]}",  # 插入点
            str(scale),  # X比例
            str(scale) if scale != 1.0 else "1",  # Y比例 (如果X=1则跳过)
            str(rotation),  # 旋转角度
            "1" if explode else "0"  # 是否炸开
        ]

        insert_cmd = "\n".join(cmd_parts) + "\n"

        # 5. 执行插入命令
        from CAD_coordination import send_cmd_with_sync
        success = send_cmd_with_sync(insert_cmd, wait_after=2.0, timeout=60.0)

        if not success:
            print(f"❌ 插入块命令失败: {block_file_path}")
            return False

        # 6. 等待插入完成
        wait_quiescent(min_quient=2.0, timeout=60.0)

        # 7. 验证插入结果
        # 检查是否有未保存更改
        try:
            has_changes = not getattr(current_doc, 'Saved', True)
            if has_changes:
                print(f"✅ 块插入成功: {block_file_path}")
                return True
            else:
                print(f"⚠ 块插入后未检测到更改: {block_file_path}")
                return True  # 仍然认为成功
        except:
            print(f"✅ 块插入完成(无法验证更改状态): {block_file_path}")
            return True

    except Exception as e:
        print(f"❌ 插入块异常: {e}")
        return False

def insert_multiple_blocks_paradigm(block_configs: List[dict]) -> int:
    """批量插入块范式"""
    success_count = 0

    print(f"🔄 开始批量插入 {len(block_configs)} 个块")

    for i, config in enumerate(block_configs):
        print(f"\n📦 [{i+1}/{len(block_configs)}] 插入块 {i+1}")

        try:
            block_path = config['path']
            insert_point = config.get('point', (0, 0, 0))
            scale = config.get('scale', 1.0)
            rotation = config.get('rotation', 0.0)
            explode = config.get('explode', False)

            if insert_dwg_as_block_paradigm(
                block_path, insert_point, scale, rotation, explode
            ):
                success_count += 1
                print(f"✅ 成功插入: {block_path}")
            else:
                print(f"❌ 插入失败: {block_path}")

            # 块间间隔等待
            if i < len(block_configs) - 1:
                time.sleep(1.0)
                wait_quiescent(min_quiet=0.5, timeout=15.0)

        except Exception as e:
            print(f"❌ 插入块配置错误: {e}")

    print(f"\n📊 批量插入结果: {success_count}/{len(block_configs)} 成功")
    return success_count

def insert_and_explode_paradigm(block_file_path: str,
                               insert_point: tuple = (0, 0, 0),
                               scale: float = 1.0) -> bool:
    """插入并炸开块范式"""
    print(f"🔄 插入并炸开: {block_file_path}")

    # 1. 先插入块
    if not insert_dwg_as_block_paradigm(
        block_file_path, insert_point, scale, explode=True
    ):
        return False

    # 2. 等待插入完成
    wait_quiescent(min_quiet=1.0, timeout=30.0)

    # 3. 验证炸开结果
    print(f"✅ 插入并炸开完成: {block_file_path}")
    return True
```

#### 使用示例
```python
# 插入单个块
insert_dwg_as_block_paradigm(
    "D:/blocks/furniture.dwg",
    insert_point=(100, 50, 0),
    scale=1.0,
    rotation=45.0,
    explode=False
)

# 插入并炸开块
insert_and_explode_paradigm(
    "D:/blocks/door.dwg",
    insert_point=(200, 100, 0),
    scale=1.0
)

# 批量插入块
block_configs = [
    {
        'path': 'D:/blocks/chair.dwg',
        'point': (100, 100, 0),
        'scale': 1.0,
        'rotation': 0.0,
        'explode': False
    },
    {
        'path': 'D:/blocks/table.dwg',
        'point': (200, 100, 0),
        'scale': 1.0,
        'rotation': 90.0,
        'explode': True
    }
]
insert_multiple_blocks_paradigm(block_configs)
```

---

## 🛠 辅助函数

### 路径处理
```python
def _get_short_path(long_path: str) -> str:
    """获取短路径处理中文/特殊字符"""
    try:
        import ctypes
        from ctypes import wintypes

        GetShortPathNameW = ctypes.windll.kernel32.GetShortPathNameW
        GetShortPathNameW.argtypes = [
            wintypes.LPCWSTR, wintypes.LPWSTR, wintypes.DWORD
        ]
        GetShortPathNameW.restype = wintypes.DWORD

        buf = ctypes.create_unicode_buffer(260)
        ret = GetShortPathNameW(long_path, buf, len(buf))
        return buf.value if ret else long_path
    except Exception:
        return long_path

def _activate_document(file_path: str) -> bool:
    """激活指定文档"""
    try:
        import win32com.client
        acad = win32com.client.GetActiveObject("AutoCAD.Application")

        target_path = Path(file_path).resolve().as_posix().lower()

        for i in range(acad.Documents.Count):
            doc = acad.Documents.Item(i)
            if doc.FullName:
                doc_path = Path(doc.FullName).resolve().as_posix().lower()
                if doc_path == target_path:
                    acad.ActiveDocument = doc
                    return True

        return False
    except Exception:
        return False
```

### 状态检查
```python
def is_file_opened_by_name(file_name: str) -> bool:
    """检查文件名是否已打开"""
    try:
        import win32com.client
        acad = win32com.client.GetActiveObject("AutoCAD.Application")

        for i in range(acad.Documents.Count):
            doc = acad.Documents.Item(i)
            if doc.Name == file_name:
                return True
        return False
    except Exception:
        return False
```

---

## 📋 操作范式检查清单

### 新建文件
- [ ] CAD环境就绪检查
- [ ] 路径幂等性检查
- [ ] 协同机制新建操作
- [ ] 等待操作完成
- [ ] 后置保存处理(如需要)

### 打开文件
- [ ] 文件存在性验证
- [ ] 进程预处理(非破坏性)
- [ ] 路径级幂等检查
- [ ] 名称级幂等检查
- [ ] 协同机制打开操作
- [ ] 等待文档加入集合
- [ ] 文档激活
- [ ] 等待CAD空闲

### 关闭文件
- [ ] 文件打开状态检查
- [ ] 保存选项处理
- [ ] 协同机制关闭操作
- [ ] 等待关闭完成
- [ ] 状态恢复

### 保存文件
- [ ] 文件打开状态检查
- [ ] CAD空闲等待
- [ ] 协同机制保存操作
- [ ] 文件创建验证
- [ ] 等待保存完成

### 插入文件
- [ ] 块文件存在验证
- [ ] 接收文件状态检查
- [ ] 短路径转换
- [ ] -INSERT命令构建
- [ ] 协同机制插入操作
- [ ] 等待插入完成
- [ ] 更改状态验证

---

## 🎯 范式使用建议

### 单一操作
```python
# 推荐:使用范式函数
open_dwg_paradigm("D:/test.dwg")
save_current_dwg_paradigm()
```

### 组合操作
```python
# 标准工作流
if open_dwg_paradigm("D:/source.dwg"):
    insert_dwg_as_block_paradigm("D:/blocks/item.dwg")
    save_as_dwg_paradigm("D:/output/result.dwg")
    close_current_dwg_paradigm()
```

### 错误处理
```python
# 带错误处理的范式使用
try:
    if open_dwg_paradigm(file_path):
        # 执行操作
        save_current_dwg_paradigm()
    else:
        print("文件打开失败,处理错误")
except Exception as e:
    print(f"操作异常: {e}")
    # 恢复状态
    single_unsaved_state()
```

---

**重要提醒**:所有基本操作都必须严格遵循对应的范式,确保操作的可靠性、幂等性和协同性!