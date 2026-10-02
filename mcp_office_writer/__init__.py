# -*- coding: utf-8 -*-

"""
mcp-office-writer — MCP server for creating/editing Office documents
Supports Word (.docx), Excel (.xlsx), and PowerPoint (.pptx)
"""
from __future__ import annotations

import sys


def _configure_stdio_utf8() -> None:
    """Keep all process text streams UTF-8 when launched on Windows.

    MCP's stdio transport already encodes protocol frames as UTF-8.  Windows,
    however, starts Python text streams using the active code page (commonly
    GBK).  Explicitly configuring the base streams prevents Chinese JSON from
    being corrupted by helpers or diagnostics that access ``sys.stdin`` or
    ``sys.stdout`` directly before/alongside the MCP transport wrapper.
    """
    for stream in (sys.stdin, sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(encoding="utf-8", errors="backslashreplace")


# 在 import 其他依赖之前先执行一次（幂等，main() 启动时仍会再次调用）
_configure_stdio_utf8()

from mcp.server.fastmcp import FastMCP
from mcp.server.fastmcp.server import ToolAnnotations

from .models import (
    WordContent, PageSetup, HeaderFooter, FontConfig, SectionItem,
    SheetDef, SlideDef,
    CreateResult, EditResult, OfficeOperation,
    FrameworkDiagram,
    PdfContent, PdfPageSetup, PdfFontConfig,
)
from .word_writer import create_word_document as _create_word, edit_word_document as _edit_word
from .excel_writer import create_excel_workbook as _create_excel, edit_excel_workbook as _edit_excel
from .ppt_writer import create_powerpoint as _create_ppt, edit_powerpoint as _edit_ppt
from .ppt_writer import list_color_schemes as _list_color_schemes, apply_color_to_slides as _apply_color
from .ppt_writer import optimize_slide_layout as _optimize_slides, clean_unused_placeholders as _clean_placeholders
from .ppt_writer import analyze_powerpoint_template as _analyze_template
from .ppt_writer import create_framework_diagram as _create_fw_diagram
from .reader import (
    read_word_document as _read_word,
    read_excel_workbook as _read_excel,
    read_powerpoint as _read_ppt,
    read_pdf_document as _read_pdf,
    read_pdf_structured as _read_pdf_structured,
    read_txt_document as _read_txt,
)
from .markdown_writer import create_word_from_markdown as _create_word_from_markdown
from .pdf_writer import create_pdf_document as _create_pdf
from .office_convert import (
    convert_to_pdf as _convert_to_pdf,
    render_document_pages as _render_document_pages,
    recalculate_excel_formulas as _recalculate_excel_formulas,
)
from .document_ops import (
    batch_edit_office_document as _batch_edit_office_document,
    inspect_office_document as _inspect_office_document,
)
from .template_merge import merge_office_template as _merge_office_template
from .guards import with_timeout

# 同步操作超时上限（秒）：COM 活动文档操作可能弹窗或未响应，
# Markdown 转换涉及磁盘 IO 与 Word 渲染，均加超时防止 MCP 调用永久挂起。
# dsoffice 转换需启动内置 LibreOffice 引擎，且首次运行可能下载引擎，放得更宽。
COM_OPERATION_TIMEOUT = 60
MARKDOWN_CONVERT_TIMEOUT = 120
DSOFFICE_OPERATION_TIMEOUT = 300

mcp = FastMCP("Office Writer")


def _read_result_to_markdown(result: dict) -> str:
    """把 read_* 的扁平 dict 转成干净的 markdown 文本。

    直接返回 dict 会被 FastMCP 整体序列化为嵌套 JSON，正文大字符串经历双重转义；
    这里把正文作为顶层文本返回，统计信息附在引用行，截断标记保留。
    """
    if not result.get("success"):
        return f"读取失败: {result.get('message') or '未知错误'}"
    parts = [str(result.get("content") or "")]
    stats = {
        k: v for k, v in result.items()
        if k not in {"success", "file_path", "content", "message", "truncated"} and v is not None
    }
    if stats:
        parts.append("> " + "；".join(f"{k}: {v}" for k, v in stats.items()))
    if result.get("truncated"):
        parts.append("> 内容过长已截断")
    return "\n\n".join(p for p in parts if p)


# ── Word ────────────────────────────────────────────────────────

@mcp.tool()
def create_word_document(
    output_path: str,
    content: WordContent,
    template: str | None = None,
    page_setup: PageSetup | None = None,
    header_footer: HeaderFooter | None = None,
    default_font: FontConfig | None = None,
    scene: str | None = None,
    cover_config: dict | None = None,
    if_exists: str = "overwrite",
) -> CreateResult:
    """
    创建 Word (.docx) 文档。

    支持模板模式：传入已有的 .docx/.dotx 文件作为 template，新文档将继承其样式/页面设置/页眉页脚；模板缺失或格式错误会明确报错，不会静默改用空白文档。
    内容通过 content.sections 描述，支持 heading/paragraph/table/image/page_break/bullet_list/numbered_list 等类型。

    场景模板 (scene): 预设排版规范，自动设置字体/行距/缩进/边距。
      可选值: academic(学术论文), report(报告/提案), official(公文 GB/T 9704),
             contract(合同), resume(简历), exam(试卷), journal(期刊论文)
      所有字号均支持国内标准（如"小四""五号""三号"）或直接 pt 数值。
      场景默认值可被 page_setup/header_footer/default_font 等显式参数覆盖。

    封面 (cover_config): 可选字典，包含 title/subtitle/author/date/organization/recipe_id。
      场景为 report 时自动生成封面。
    """
    return _create_word(output_path, content, template, page_setup, header_footer, default_font, scene, cover_config, if_exists)


@mcp.tool()
def edit_word_document(
    file_path: str,
    append_sections: list[SectionItem],
    default_font: FontConfig | None = None,
) -> EditResult:
    """
    向已有 Word 文档追加内容。

    在文档末尾追加段落/表格/图片等内容块。
    """
    return _edit_word(file_path, append_sections, default_font)


@mcp.tool()
def create_word_from_markdown(
    source_path: str,
    output_path: str,
    template: str | None = None,
    scene: str | None = "report",
    if_exists: str = "overwrite",
) -> CreateResult:
    """直接将磁盘上的 Markdown 长文转换为 Word (.docx)。

    该工具由 MCP 进程自行读取源文件，不会把全文塞进工具 JSON 参数，适合报告、制度、论文等长文本。
    支持标题、段落、列表、表格、代码块和常见行内 Markdown；独占一行的本地图片会嵌入文档，
    无法嵌入时会保留可见提示而不静默丢失。scene 默认为 report。
    """
    timeout_result = CreateResult(
        success=False, file_path=output_path,
        message=f"Markdown 转 Word 超时（>{MARKDOWN_CONVERT_TIMEOUT}s），请检查源文件是否过大",
    )
    return with_timeout(
        _create_word_from_markdown, MARKDOWN_CONVERT_TIMEOUT, timeout_result,
        source_path, output_path, template, scene, if_exists,
    )


# ── Excel ───────────────────────────────────────────────────────

@mcp.tool()
def create_excel_workbook(
    output_path: str,
    sheets: list[SheetDef],
    template: str | None = None,
    if_exists: str = "overwrite",
) -> CreateResult:
    """
    创建 Excel (.xlsx) 工作簿。

    支持多工作表，每个 sheet 可指定表头、数据、列宽、格式、合并单元格、冻结窗格和自动筛选。
    默认将文本原样写入；需要公式时，请在对应 sheet 中显式设置 allow_formulas=true。
    支持模板模式。
    """
    return _create_excel(output_path, sheets, template, if_exists)


@mcp.tool()
def edit_excel_workbook(
    file_path: str,
    append_sheets: list[SheetDef],
) -> EditResult:
    """
    向已有 Excel 文件追加新的工作表。
    """
    return _edit_excel(file_path, append_sheets)


# ── PowerPoint ──────────────────────────────────────────────────

@mcp.tool()
def create_powerpoint(
    output_path: str,
    slides: list[SlideDef],
    template: str | None = None,
    if_exists: str = "overwrite",
) -> CreateResult:
    """
    创建 PowerPoint (.pptx) 演示文稿。

    支持多种布局 (title/content/two_content/comparison/blank 等)。
    每张幻灯片可包含段落、表格、图片（本地/URL）、形状、图表、项目符号/编号列表元素。
    支持演讲者备注、背景色设置、自定义幻灯片尺寸(slide_width/slide_height)。支持模板模式。
    """
    return _create_ppt(output_path, slides, template, if_exists)


@mcp.tool()
def edit_powerpoint(
    file_path: str,
    append_slides: list[SlideDef],
) -> EditResult:
    """
    向已有 PPT 文件追加新的幻灯片。
    """
    return _edit_ppt(file_path, append_slides)


# ── PDF ───────────────────────────────────────────────────────

@mcp.tool()
def create_pdf_document(
    output_path: str,
    content: PdfContent,
    page_setup: PdfPageSetup | None = None,
    font_config: PdfFontConfig | None = None,
    if_exists: str = "overwrite",
) -> CreateResult:
    """
    创建 PDF (.pdf) 文档。

    支持段落 (paragraph)、标题 (heading)、表格 (table)、图片 (image)、
    分页 (page_break)、项目符号列表 (bullet_list)、编号列表 (numbered_list)。
    自动检测 Windows 系统的中文字体（黑体/宋体/微软雅黑等），
    也支持通过 font_config.cn_font_path 指定自定义字体文件。

    page_setup 可选配置：paper_size (A4/Letter/A3/B5)、orientation (portrait/landscape)、
    边距 (margin_top/bottom/left/right, 单位 mm)。

    font_config 可选配置：cn_font (中文字体名)、cn_font_path (自定义字体路径)、
    en_font (西文字体)、cn_font_size (正文字号如 '小四')、heading_font_size (标题字号如 '三号')。
    """
    return _create_pdf(output_path, content, page_setup, font_config, if_exists)


# ── Convert (dsh-libreoffice-kit / dsoffice) ────────────────────

@mcp.tool()
def convert_to_pdf(
    input_path: str,
    output_path: str | None = None,
    if_exists: str = "overwrite",
    sheet: str | None = None,
) -> dict:
    """将 Office 文档高保真转换为 PDF（内置 LibreOffice 引擎，无需本机安装 Office）。

    支持 .docx/.xlsx/.pptx，以及 97-2003 的 .doc/.xls/.ppt 和 .rtf。
    Excel 可用 sheet 参数只转换单个工作表（默认全部）。
    缺失字体按字形度量自动替换（版式最接近原样），并在 missing_fonts 中报告。
    output_path 缺省时输出到源文件同目录同名 .pdf。
    if_exists: overwrite(覆盖)/error(报错)/rename(自动改名)。
    """
    timeout_result = {
        "success": False, "file_path": output_path or input_path,
        "message": f"PDF 转换超时（>{DSOFFICE_OPERATION_TIMEOUT}s），首次运行可能正在下载引擎，可重试",
    }
    return with_timeout(_convert_to_pdf, DSOFFICE_OPERATION_TIMEOUT, timeout_result,
                        input_path, output_path, if_exists, sheet)


@mcp.tool()
def render_document_pages(
    input_path: str,
    output_dir: str | None = None,
    pages: str = "all",
    dpi: int = 144,
    sheet: str | None = None,
    cell_range: str | None = None,
    if_exists: str = "overwrite",
) -> dict:
    """将 Office/PDF 文档逐页渲染为 PNG 图片。

    支持 .docx/.xlsx/.pptx、97-2003 二进制格式（.doc/.xls/.ppt）和 .pdf。
    pages 为 'all' 或 '1,3,5' 页码列表；Excel 可用 sheet + cell_range（如 A1:D20）渲染指定区域；
    dpi 默认 144（72-600）。输出目录缺省时为源文件同目录的 <文件名>_pages。
    返回每页图片的路径与尺寸（images），适合生成文档预览图或封面缩略图。
    if_exists: overwrite(清空重建)/error(报错)。
    """
    timeout_result = {
        "success": False, "output_dir": output_dir or "",
        "message": f"页面渲染超时（>{DSOFFICE_OPERATION_TIMEOUT}s），首次运行可能正在下载引擎，可重试",
    }
    return with_timeout(_render_document_pages, DSOFFICE_OPERATION_TIMEOUT, timeout_result,
                        input_path, output_dir, pages, dpi, sheet, cell_range, if_exists)


@mcp.tool()
def recalculate_excel_formulas(
    input_path: str,
    output_path: str | None = None,
    if_exists: str = "overwrite",
) -> dict:
    """重算 Excel (.xlsx) 公式并把计算结果写为缓存值。

    create_excel_workbook 等工具写入的公式没有缓存值，其他读取工具看到的是空；
    本工具用 LibreOffice 引擎整体重算后写回，公式与数值并存。
    默认输出 <原名>_recalc.xlsx（也可指定 .xlsx 或 .ods 路径）。
    if_exists: overwrite(覆盖)/error(报错)/rename(自动改名)。
    """
    timeout_result = {
        "success": False, "file_path": output_path or input_path,
        "message": f"公式重算超时（>{DSOFFICE_OPERATION_TIMEOUT}s），可重试",
    }
    return with_timeout(_recalculate_excel_formulas, DSOFFICE_OPERATION_TIMEOUT, timeout_result,
                        input_path, output_path, if_exists)


# ── Active Office documents (Windows COM) ───────────────────────

@mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))
def get_open_office_documents() -> dict:
    """列出可由 Windows COM 直接编辑的活动 Word/WPS 与 Excel 文档。

    超时保护：60 秒内未返回视为应用未响应，返回空列表。
    """
    from .office_com import get_open_office_documents as _get_open
    return with_timeout(_get_open, COM_OPERATION_TIMEOUT, {"word": [], "excel": [], "powerpoint": []})


@mcp.tool()
def edit_open_word_document(
    file_path: str,
    append_sections: list[SectionItem],
    default_font: FontConfig | None = None,
) -> EditResult:
    """直接编辑已在 Word/WPS 中打开的文档并由该应用保存。

    仅连接现有窗口，不会启动 Word/WPS；文件未打开时返回明确错误。
    支持标题、段落、项目符号/编号列表和分页符。
    """
    from .office_com import com_is_available, try_edit_open_word_document as _edit_open
    if not com_is_available():
        return EditResult(success=False, file_path=file_path, message="Windows COM 不可用；请在安装 pywin32 的 Windows 环境中使用活动文档编辑")
    timeout_result = EditResult(success=False, file_path=file_path, message=f"活动 Word 文档操作超时（>{COM_OPERATION_TIMEOUT}s），可能的应用未响应")
    result = with_timeout(_edit_open, COM_OPERATION_TIMEOUT, timeout_result, file_path, append_sections, default_font)
    return result or EditResult(success=False, file_path=file_path, message="目标文件未在活动 Word/WPS 实例中打开")


@mcp.tool()
def edit_open_excel_workbook(file_path: str, append_sheets: list[SheetDef]) -> EditResult:
    """直接向已在 Excel 中打开的工作簿追加工作表并由该应用保存。

    超时保护：60 秒内未返回视为应用未响应。
    """
    from .office_com import com_is_available, try_edit_open_excel_workbook as _edit_open
    if not com_is_available():
        return EditResult(success=False, file_path=file_path, message="Windows COM 不可用；请在安装 pywin32 的 Windows 环境中使用活动工作簿编辑")
    timeout_result = EditResult(success=False, file_path=file_path, message=f"活动 Excel 工作簿操作超时（>{COM_OPERATION_TIMEOUT}s），可能的应用未响应")
    result = with_timeout(_edit_open, COM_OPERATION_TIMEOUT, timeout_result, file_path, append_sheets)
    return result or EditResult(success=False, file_path=file_path, message="目标文件未在活动 Excel 实例中打开")


@mcp.tool()
def edit_open_powerpoint(file_path: str, append_slides: list[SlideDef]) -> EditResult:
    """直接向已在 PowerPoint 中打开的演示文稿追加幻灯片并由该应用保存。

    仅连接现有窗口，不会启动 PowerPoint；文件未打开时返回明确错误。
    支持在已打开的 PPT 中实时追加幻灯片并查看效果。
    """
    from .office_com import com_is_available, try_edit_open_powerpoint as _edit_open
    if not com_is_available():
        return EditResult(success=False, file_path=file_path, message="Windows COM 不可用；请在安装 pywin32 的 Windows 环境中使用活动演示文稿编辑")
    timeout_result = EditResult(success=False, file_path=file_path, message=f"活动 PowerPoint 演示文稿操作超时（>{COM_OPERATION_TIMEOUT}s），可能的应用未响应")
    result = with_timeout(_edit_open, COM_OPERATION_TIMEOUT, timeout_result, file_path, append_slides)
    return result or EditResult(success=False, file_path=file_path, message="目标文件未在活动 PowerPoint 实例中打开")


# ── PPT 效果优化 ─────────────────────────────────────────────────

@mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))
def ppt_list_color_schemes() -> list[dict]:
    """列出所有可用的 PPT 预设配色方案（商务蓝、暗夜模式、暖色大地、极简优雅、活力缤纷、企业绿、科技紫、经典中国红、清新自然、海洋蓝、日落暖橙、高级灰 共 12 套）。"""
    return _list_color_schemes()


@mcp.tool()
def ppt_apply_color_scheme(file_path: str, scheme_name: str = "professional-blue") -> dict:
    """对已有 PPT 文件中的所有幻灯片应用指定配色方案。

    自动更新标题颜色、正文颜色、背景色、形状填充色、表格颜色和字体。
    scheme_name 可选值: professional-blue, modern-dark, warm-earth, elegant-minimal, vibrant, corporate-green
    """
    return _apply_color(file_path, scheme_name)


@mcp.tool()
def ppt_optimize_slide_layout(file_path: str) -> dict:
    """对已有 PPT 中的内容元素进行布局优化。

    自动调整多个元素之间的垂直间距使其均匀分布，优化文本框高度适配内容长度。
    适合在使用其他工具生成内容后，对排版进行一键美化。
    """
    return _optimize_slides(file_path)


@mcp.tool()
def ppt_clean_unused_placeholders(file_path: str) -> dict:
    """清理 PPT 文件中未使用的空白占位符（空标题框、空文本框等）。

    移除每张幻灯片上没有文字内容的占位符形状，使幻灯片更加整洁干净。
    """
    return _clean_placeholders(file_path)


@mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))
def ppt_analyze_template(file_path: str) -> dict:
    """分析已有 PPT/PPTX 模板，提取布局、配色和字体信息。

    返回幻灯片尺寸、可用布局名称列表、使用的字体和颜色列表，
    帮助了解模板结构以便后续创作。
    """
    return _analyze_template(file_path)


@mcp.tool()
def ppt_create_framework_diagram(
    output_path: str,
    diagram: FrameworkDiagram,
    if_exists: str = "overwrite",
) -> CreateResult:
    """自动生成 PPT 框架图（阶段+节点+箭头连接）。

    根据 FrameworkDiagram 定义自动布局并渲染为单个幻灯片：
    - 左侧圆柱形阶段标签，右侧阶段背景矩形容器
    - 节点支持 rect/rounded_rect/ellipse/diamond 四种形状
    - 相邻节点自动生成箭头，也可通过 edges 指定显式连接
    - 支持 12 套预设配色方案

    FrameworkDiagram 结构：
    { title, subtitle, stages: [{ label, color, nodes: [{ id, text, style, fill_color }] }],
      edges: [{ from_id, to_id, label }], color_scheme, node_width, node_height }
    """
    return _create_fw_diagram(output_path, diagram, if_exists)


@mcp.tool()
def ppt_add_flowchart_slide(
    file_path: str,
    diagram: FrameworkDiagram,
) -> EditResult:
    """在已打开的 PowerPoint 演示文稿末尾追加一张流程图幻灯片（COM 实时绘制）。

    与 ppt_create_framework_diagram 布局一致（阶段标签 + 节点 + 箭头），但直接绘制到
    当前打开的活动演示文稿中：仅连接现有窗口，不会启动 PowerPoint；文件未打开时返回明确错误。

    FrameworkDiagram 结构：
    { title, subtitle, stages: [{ label, color, nodes: [{ id, text, style, fill_color }] }],
      edges: [{ from_id, to_id, label, edge_style }], color_scheme, node_width, node_height }
    """
    from .office_com import com_is_available, try_add_flowchart_slide as _add_fw
    if not com_is_available():
        return EditResult(success=False, file_path=file_path, message="Windows COM 不可用；请在安装 pywin32 的 Windows 环境中使用活动演示文稿编辑")
    timeout_result = EditResult(success=False, file_path=file_path, message=f"活动 PowerPoint 流程图操作超时（>{COM_OPERATION_TIMEOUT}s），可能的应用未响应")
    result = with_timeout(_add_fw, COM_OPERATION_TIMEOUT, timeout_result, file_path, diagram)
    return result or EditResult(success=False, file_path=file_path, message="目标文件未在活动 PowerPoint 实例中打开")


# ── Read (binary file content extraction) ───────────────────────


@mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))
def read_word_document(file_path: str, max_chars: int = 50000) -> str:
    """
    读取 Word (.docx) 文档内容，提取文字、表格、列表等。
    返回 Markdown 文本，标题以 # 标注，表格以 Markdown 格式呈现。
    """
    return _read_result_to_markdown(_read_word(file_path, max_chars))


@mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))
def read_excel_workbook(file_path: str, max_rows_per_sheet: int = 1000, max_chars: int = 50000) -> str:
    """
    读取 Excel (.xlsx/.xlsm) 工作簿内容，提取所有工作表的数据。
    每个 sheet 以 Markdown 表格输出，支持 .xlsx 和 .xlsm 格式。
    """
    return _read_result_to_markdown(_read_excel(file_path, max_rows_per_sheet, max_chars))


@mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))
def read_powerpoint(file_path: str, max_chars: int = 50000) -> str:
    """
    读取 PowerPoint (.pptx) 演示文稿内容，提取幻灯片文字和表格。
    每张幻灯片以 ## Slide N 开头，包含布局名称、各形状文本、表格数据和演讲者备注。
    支持按幻灯片边界截断（max_chars）。
    """
    return _read_result_to_markdown(_read_ppt(file_path, max_chars))

@mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))
def read_pdf_document(file_path: str, max_chars: int = 50000) -> str:
    """
    读取 PDF 文档内容，逐页提取文字。
    每页以 ## Page N 开头，支持大文件截断保护。
    """
    return _read_result_to_markdown(_read_pdf(file_path, max_chars))


@mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))
def read_pdf_structured(
    file_path: str,
    max_chars: int = 50000,
    include_blocks: bool = False,
    extract_images: bool = False,
) -> dict:
    """
    读取 PDF 文档，返回结构化内容。

    相比 read_pdf_document 返回纯文本，此工具返回细粒度的结构化数据：
    - 文档元数据（标题/作者/主题/关键词/创建工具等）
    - 书签层级树（TOC）
    - 超链接列表
    - 逐页数据：纯文本 + 可选文本块信息（bbox/字体/字号/颜色）
    - 可选图片提取（保存到 PDF 所在目录的 _images/ 子目录）
    """
    return _read_pdf_structured(file_path, max_chars, include_blocks, extract_images)


@mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))
def read_txt_document(file_path: str, max_chars: int = 50000) -> str:
    """
    读取纯文本文件内容，自动检测编码。
    支持 .txt、.md、.markdown 等纯文本格式。
    支持 UTF-8/GBK/GB2312/Latin-1/UTF-16 等编码自动识别。
    """
    return _read_result_to_markdown(_read_txt(file_path, max_chars))


# ── Structured inspection and editing ───────────────────────────

@mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))
def inspect_office_document(file_path: str, max_nodes: int = 1000) -> dict:
    """读取 Office 文档节点树并返回稳定路径。

    支持 DOCX、XLSX/XLSM、PPTX。返回的路径可直接交给
    batch_edit_office_document，例如 /body/p[1]、/Sheet1/A1、
    /slide[1]/shape[2]。max_nodes 用于限制大型文档的响应体。
    """
    return _inspect_office_document(file_path, max_nodes)


@mcp.tool()
def batch_edit_office_document(
    file_path: str,
    operations: list[OfficeOperation],
    output_path: str | None = None,
) -> dict:
    """以原子事务批量编辑 DOCX、XLSX/XLSM 或 PPTX。

    支持 set/add/remove/move/replace。任一操作失败时不会保存任何修改，
    并返回 code/path/suggestion/valid_values 等结构化错误信息。
    output_path 为空时原地更新，否则写入同扩展名的新文件。
    """
    return _batch_edit_office_document(file_path, operations, output_path)


@mcp.tool()
def merge_office_template(
    template_path: str,
    output_path: str,
    data: dict,
    if_exists: str = "overwrite",
) -> dict:
    """使用 JSON 数据填充 Office 模板中的 {{key}} 占位符。

    支持 DOCX、XLSX/XLSM、PPTX，以及对象点路径和数组索引，
    如 {{user.name}}、{{items[0].title}}。返回未解析占位符清单。
    """
    return _merge_office_template(template_path, output_path, data, if_exists)


# ── Validate ────────────────────────────────────────────────────

@mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))
def validate_word_document(file_path: str) -> dict:
    """
    对已生成的 Word 文档进行质量验证。

    检查项包括：空白页检测、行间距一致性、表格完整性、标题层级连续性、图片嵌入、分节设置等。
    返回详细的验证报告，包含每个检查项的通过/警告/错误状态。
    """
    from .checker import validate_word_document as _validate
    return _validate(file_path).to_dict()




def main():
    """启动 MCP server。"""
    _configure_stdio_utf8()
    try:
        mcp.run(transport="stdio")
    except KeyboardInterrupt:
        raise
    except Exception as e:
        print(f"[mcp-office-writer] 启动失败: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
