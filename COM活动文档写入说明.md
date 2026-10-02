# Windows 活动 Office 文档写入

`mcp-office-writer` 在 Windows 上可以连接已经打开的 Word/WPS 或 Excel 实例，直接修改它们内存中的活动文档并由该应用调用 `Save()`。这与常规 writer 的“创建临时文件再替换原文件”不同，因此适用于被 Word/Excel 自身锁定的文件。

## 工具

- `get_open_office_documents()`：列出当前可通过 COM 操作的 Word/WPS 与 Excel 文档；
- `edit_open_word_document(file_path, append_sections, default_font=None)`：显式编辑已打开的 Word/WPS 文档；
- `edit_open_excel_workbook(file_path, append_sheets)`：显式向已打开的 Excel 工作簿追加工作表。

原有的 `edit_word_document` 与 `edit_excel_workbook` 也会优先尝试活动 COM 文档；若目标文件未打开，则无行为变化，继续使用原本的文件写入器。

## 边界

- 仅连接既有应用实例，绝不启动、关闭或退出用户的 Word/WPS/Excel；
- Word 当前支持标题、段落、项目符号/编号列表和分页符；复杂表格、图片与分节请关闭文件后使用原 writer；
- Excel 当前支持追加工作表、表头和数据，包含基础的冻结窗格与筛选；
- 文件为只读、受保护、未出现在活动实例中，或当前 Python 没有 `pywin32` 时，工具会返回明确失败信息；普通 `edit_word_document` / `edit_excel_workbook` 会自动回退到原有文件 writer；
- 活动 Excel 在写入前会校验所有工作表名；运行中失败会删除本次新增工作表，避免半写入。活动 Word 会先校验全部内容块，并将一次写入置于单一撤销记录中；
- PowerPoint/WPS 表格组件尚未接入 COM 写入，仍保持原文件写入路径。
