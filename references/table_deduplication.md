# Word 表格合并单元格去重避坑指南

在使用 `python-docx` 库处理 Microsoft Word 中的表格（`Table`）时，**合并单元格（Merged Cells）** 是最经典也是最容易导致文本重复读取和重复写入的“地雷”。

本文档深度剖析该原理，并给出我们在这套工具链中采用的**全局单元格物理级去重方案**。

---

## 💣 经典痛点：为什么合并单元格会被读取/写入多次？

在 `python-docx` 的底层设计中：
1. `table.rows` 包含表格的所有行，`row.cells` 包含当前行的所有列单元格。
2. 当 Word 中存在**横向合并（colspan）** 或 **纵向合并（rowspan）** 的单元格时，合并区域内的所有物理格子（Grid）在 API 层面都会指向**同一个逻辑 `_Cell` 对象**。
3. 如果我们简单地遍历：
   ```python
   for row in table.rows:
       for cell in row.cells:
           # 提取或修改文字
           print(cell.text)
   ```
   *   对于一个跨 3 行 2 列合并的单元格，上述嵌套循环会将其**完全重复访问 6 次**！
   *   **后果**：提取文本时，同一个中文段落被提取 6 次，产生大量冗余；回写翻译时，英文翻译被在其末尾**重复追加 6 次**，导致文档排版彻底崩溃。

---

## 🛡️ 终极解决方案：基于 `_tc` 的表格级全局去重

为了完美解决横向跨列和纵向跨行合并的重复访问问题，我们设计了**表格级的物理格去重**算法。

### 1. 底层核心 `_tc` 属性
每个 `_Cell` 对象内部都有一个由 `python-docx` 维护的底层的 Oxml 元素属性：`_tc`。
*   `cell._tc` 代表该单元格在 OpenXML 标准中唯一的底层 XML 句柄对象（`CT_Tc`）。
*   即使由于合并单元格原因在循环中被多次访问，它们的 `_tc` 句柄也**绝对指向同一个内存地址**。

### 2. 算法实现（全局去重）
在提取（`extract.py`）和回填（`inject.py`）中，我们声明了一个表格级的 Set：`processed_cells`。
```python
for t_idx, table in enumerate(doc.tables):
    processed_cells = set() # 👈 每个表格声明一个去重集合
    for r_idx, row in enumerate(table.rows):
        for c_idx, cell in enumerate(row.cells):
            cell_id = cell._tc # 👈 提取物理单元格的底层句柄
            if cell_id in processed_cells:
                continue # 👈 若已处理过该物理单元格，直接跳过！
            processed_cells.add(cell_id)
            
            # 以下为安全的、仅执行一次的操作
            for p_idx, para in enumerate(cell.paragraphs):
                ...
```

### 3. 去重基准一致性
> [!IMPORTANT]
> 提取文本（生成待译 JSON）时的去重基准，必须与回填文本（读取翻译 JSON）时的**去重基准完全保持一致**！
> 
> *   如果在 `extract.py` 里由于去重只保留了物理上第一次遇到的坐标 `(t_idx, r_idx, c_idx)`，那么在 `inject.py` 回写时，也必须对表格进行一致的全局去重，才能在遇到该物理单元格时，精确匹配并完成单次注入。
