# mcp-office-writer

MCP 服务器：创建、编辑、读取 Office 文档（Word / Excel / PowerPoint / PDF / 纯文本）。

## 功能速览

| 类别 | 工具 |
|------|------|
| 读取 | `read_word_document` / `read_excel_workbook` / `read_powerpoint` / `read_pdf_document` / `read_pdf_structured` / `read_txt_document` |
| 创建 | `create_word_document` / `create_excel_workbook` / `create_powerpoint` / `create_pdf_document` |
| Markdown 转换 | `create_word_from_markdown`（MCP 进程直接读源文件，不传全文 JSON） |
| 编辑 | `edit_word_document` / `edit_excel_workbook` / `edit_powerpoint` |
| 高保真转换 (dsoffice) | `convert_to_pdf` / `render_document_pages` / `recalculate_excel_formulas` |
| 活动文档 (Windows COM) | `get_open_office_documents` / `edit_open_word_document` / `edit_open_excel_workbook` / `edit_open_powerpoint` |
| PPT 优化 | `ppt_list_color_schemes` / `ppt_apply_color_scheme` / `ppt_optimize_slide_layout` / `ppt_clean_unused_placeholders` / `ppt_analyze_template` / `ppt_create_framework_diagram` |
| 活动 PPT 流程图 | `ppt_add_flowchart_slide`（COM 实时绘制阶段+节点+箭头） |
| 结构化编辑 | `inspect_office_document` / `batch_edit_office_document` / `merge_office_template` |
| 质量验证 | `validate_word_document` |

特性：场景模板（学术/报告/公文/合同等，支持国内字号）、模板继承、{{key}} 占位符合并、
原子写入（`atomic_save`，失败不损坏原文件）、输入文件安全校验（大小/解压炸弹）、
PDF 扫描件 OCR 回退（Tesseract）、COM 活动文档回滚、超时保护。

### 高保真转换（dsh-libreoffice-kit）

`convert_to_pdf` / `render_document_pages` / `recalculate_excel_formulas` 封装
DeepSeek 开源的 [dsh-libreoffice-kit](https://github.com/deepseek-ai/dsh-libreoffice-kit)
（`dsoffice` CLI，内置精简版 LibreOffice 引擎，无需本机安装 LibreOffice/Office）：

- `convert_to_pdf`：docx/xlsx/pptx 及 97-2003 二进制格式（.doc/.xls/.ppt）、.rtf → PDF，
  缺失字体按字形度量替换（版式最接近原样）并在 `missing_fonts` 中报告
- `render_document_pages`：Office/PDF → 逐页 PNG（指定页码、Excel 区域、dpi），适合预览图/验收截图
- `recalculate_excel_formulas`：重算 xlsx 公式写回缓存值（弥补 openpyxl 公式无缓存值，
  转换后 `read_excel_workbook` 的 data_only 模式可见数值）

环境要求：Node.js ≥ 22.19，`npm install -g @deepseek-ai/libreoffice-kit`；
Windows x64 需 MSVC v14 运行库。缺失时三个工具返回带安装指引的错误信息。

## 启动

```bat
start.bat
```

或直接：

```bat
.\venv\Scripts\python.exe -I -X utf8 -m mcp_office_writer
```

## 开发

```bat
rem 依赖安装（venv 已建好时）
.\venv\Scripts\python.exe -m pip install -e ".[dev]"

rem 运行测试
.\venv\Scripts\python.exe -m pytest tests -q
```

测试套件 83 个用例：单元（文档操作、PDF 回环、dsoffice 封装层、COM 失败分支、超时守卫、安全回归）+ MCP stdio 端到端（中文往返）。

## 配置（环境变量）

- `MCP_OFFICE_OCR=0`：关闭 PDF 扫描件 OCR
- `TESSERACT_CMD`：Tesseract 可执行文件路径（默认 `D:\cherry-workplace\tools\tesseract\tesseract.exe`）
- `MCP_OFFICE_OCR_LANG`：OCR 语言，默认 `chi_sim+eng`
- `MCP_OFFICE_DSOFFICE`：dsoffice 可执行文件路径（默认从 PATH 查找）
