# mcp-office-writer 文件夹结构

```
mcp-office-writer/
├── README.md                # 项目说明（功能、启动、开发、环境变量）
├── FOLDER_STRUCTURE.md      # 本文件
├── pyproject.toml           # 包定义（mcp<2 FastMCP v1、python-docx/openpyxl/python-pptx/pymupdf 等）
├── start.bat / start.ps1    # MCP 启动脚本（venv python -I -X utf8 -m mcp_office_writer）
├── mcp_office_writer/       # Python 包源码
│   ├── __init__.py          # FastMCP server 与全部 @mcp.tool() 注册
│   ├── __main__.py          # python -m mcp_office_writer 入口
│   ├── models.py            # Pydantic 请求/结果模型
│   ├── word_writer.py       # Word 创建/追加（场景模板、封面）
│   ├── excel_writer.py      # Excel 创建/追加（SheetDef）
│   ├── ppt_writer.py        # PPT 创建/追加/配色/布局优化/框架图
│   ├── pdf_writer.py        # PDF 创建（中文字体自动检测）
│   ├── markdown_writer.py   # Markdown → Word（磁盘直读）
│   ├── reader.py            # Office/PDF/TXT 读取（含 PDF OCR 回退）
│   ├── office_convert.py    # dsh-libreoffice-kit (dsoffice) 封装：→PDF/→PNG/公式重算
│   ├── document_ops.py      # 节点树检查 + 原子批量编辑
│   ├── template_merge.py    # {{key}} 占位符模板合并
│   ├── office_com.py        # Windows COM 活动文档写入（Word/WPS/Excel/PPT）
│   ├── covers.py / scenes.py # 封面配方与场景排版预设
│   ├── checker.py           # Word 文档质量验证
│   ├── guards.py            # with_timeout 超时保护基建
│   └── utils.py             # 路径校验/if_exists/原子保存/图片下载等公共工具
├── templates/               # 场景模板资源
├── image_cache/             # PPT 远程图片下载缓存
├── tests/                   # unittest/pytest 用例（83 个，test_convert.py 为 dsoffice 封装层）
└── venv/                    # 项目专属虚拟环境（Python 3.13）
```

## 说明

- `office_convert.py` 依赖外部 CLI `dsoffice`（npm 全局包 `@deepseek-ai/libreoffice-kit`，
  Node ≥ 22.19），不在 Python 依赖内；缺失时工具返回安装指引而非崩溃。
- `venv/` 与 `image_cache/` 为运行时产物，不纳入版本管理。
