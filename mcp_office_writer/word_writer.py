
"""
mcp-office-writer - Word (.docx) 写入/编辑逻辑
"""
from __future__ import annotations

import logging
import os
import re
from typing import TYPE_CHECKING

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

if TYPE_CHECKING:
    from .models import (
        WordContent, SectionItem, ParagraphStyle, PageSetup,
        HeaderFooter, FontConfig, CreateResult, EditResult,
    )

logger = logging.getLogger(__name__)

MAX_RUN_TEXT_CHARS = 16_000


def _add_text_runs(paragraph, text: str, font_name=None, font_size=None,
                   bold=None, italic=None, color=None) -> None:
    """Write long text in bounded runs for Word/OOXML compatibility."""
    for start in range(0, len(text), MAX_RUN_TEXT_CHARS):
        run = paragraph.add_run(text[start:start + MAX_RUN_TEXT_CHARS])
        _apply_font(run, font_name, font_size, bold, italic, color)

ALIGN_MAP = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

ORIENT_MAP = {
    "portrait": WD_ORIENT.PORTRAIT,
    "landscape": WD_ORIENT.LANDSCAPE,
}

PAPER_SIZE_MAP = {
    "A4": (Inches(8.27), Inches(11.69)),
    "Letter": (Inches(8.5), Inches(11.0)),
    "A3": (Inches(11.69), Inches(16.54)),
    "B5": (Inches(7.17), Inches(10.16)),
}


def _col_letter_to_index(cell_ref: str) -> int:
    """将 Excel 列字母转为 0-based 索引，支持 AA, AB 等多字母列

    >>> _col_letter_to_index("A1") -> 0
    >>> _col_letter_to_index("AB5") -> 27
    """
    col_str = re.sub(r"\d+", "", cell_ref).upper()
    result = 0
    for ch in col_str:
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result - 1


def _apply_font(run, font_name: str | None = None, font_size: int | str | None = None,
                bold: bool | None = None, italic: bool | None = None,
                color: str | None = None):
    """对 run 应用字体样式"""
    from .utils import cn_size_to_pt
    if font_name:
        run.font.name = font_name
        # 让东亚字体也生效（如 宋体）
        run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
    if font_size is not None:
        run.font.size = Pt(cn_size_to_pt(font_size))
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic
    if color:
        run.font.color.rgb = RGBColor(
            int(color[0:2], 16),
            int(color[2:4], 16),
            int(color[4:6], 16),
        )


def _apply_paragraph_format(paragraph, style: ParagraphStyle | None):
    """对段落应用样式"""
    if style is None:
        return
    pf = paragraph.paragraph_format

    if style.alignment:
        paragraph.alignment = ALIGN_MAP.get(style.alignment, WD_ALIGN_PARAGRAPH.LEFT)
    if style.line_spacing is not None:
        pf.line_spacing = style.line_spacing
    if style.first_line_indent is not None:
        pf.first_line_indent = Pt(style.first_line_indent)
    if style.space_before is not None:
        pf.space_before = Pt(style.space_before)
    if style.space_after is not None:
        pf.space_after = Pt(style.space_after)


def _set_page_setup(section, page_setup: PageSetup | None):
    """设置页面

    注意顺序：先按纸型设置宽高，再处理 orientation。若先交换宽高再按纸型赋值，
    paper_size 会把 landscape 交换后的宽高覆盖回纵向（paper_size 有默认值 "A4"，
    恒非 None）。
    """
    if page_setup is None:
        return
    if page_setup.paper_size:
        w, h = PAPER_SIZE_MAP.get(page_setup.paper_size, (Inches(8.27), Inches(11.69)))
        section.page_width = w
        section.page_height = h
    if page_setup.orientation:
        section.orientation = ORIENT_MAP.get(page_setup.orientation, WD_ORIENT.PORTRAIT)
        if page_setup.orientation == "landscape":
            section.page_width, section.page_height = section.page_height, section.page_width
    if page_setup.margin_top is not None:
        section.top_margin = Inches(page_setup.margin_top)
    if page_setup.margin_bottom is not None:
        section.bottom_margin = Inches(page_setup.margin_bottom)
    if page_setup.margin_left is not None:
        section.left_margin = Inches(page_setup.margin_left)
    if page_setup.margin_right is not None:
        section.right_margin = Inches(page_setup.margin_right)


def _set_header_footer(doc, hf: HeaderFooter | None):
    """设置所有 section 的页眉页脚"""
    if hf is None:
        return
    for section in doc.sections:
        if hf.header_text:
            header = section.header
            header.is_linked_to_previous = False
            p = header.paragraphs[0]
            p.text = hf.header_text
        if hf.footer_text:
            footer = section.footer
            footer.is_linked_to_previous = False
            p = footer.paragraphs[0]
            # 替换 {page} 为页码域
            if "{page}" in hf.footer_text:
                parts = hf.footer_text.split("{page}")
                # 清空默认段落
                p.clear()
                for i, part in enumerate(parts):
                    if part:
                        p.add_run(part)
                    if i < len(parts) - 1:
                        # 插入 PAGE 域
                        fld_char_begin = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="begin"/>')
                        p._element.append(fld_char_begin)
                        instr = parse_xml(f'<w:instrText {nsdecls("w")} xml:space="preserve"> PAGE </w:instrText>')
                        p._element.append(instr)
                        fld_char_end = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="end"/>')
                        p._element.append(fld_char_end)
            else:
                p.text = hf.footer_text


def _add_section_item(doc, item: SectionItem, default_font: FontConfig | None):
    """向文档添加一个内容块"""
    from .models import ParagraphStyle as PS

    t = item.type

    if t == "heading":
        level = item.level or 1
        p = doc.add_heading(item.text or "", level=level)
        # 应用段落样式
        if item.style:
            _apply_paragraph_format(p, item.style)
        # heading 也应用默认字体（重置颜色为黑色，避免 python-docx 内置 heading 的蓝色主题色）
        if default_font and p.runs:
            for run in p.runs:
                _apply_font(run, default_font.name, default_font.size, color="000000")

    elif t == "paragraph":
        p = doc.add_paragraph()
        style = item.style or PS()
        s_font = style.font_name or (default_font.name if default_font else None)
        s_size = style.font_size or (default_font.size if default_font else None)
        _add_text_runs(p, item.text or "", s_font, s_size, style.bold, style.italic, style.color)
        _apply_paragraph_format(p, style)

    elif t == "table":
        from .models import TableDef
        td = item.table or TableDef()
        rows = len(td.rows) + (1 if td.headers else 0)
        cols = max(len(td.headers) if td.headers else 0,
                   max((len(r) for r in td.rows), default=0))
        if rows == 0 or cols == 0:
            logger.warning("跳过空表格（rows=%d, cols=%d）", rows, cols)
            return

        table = doc.add_table(rows=rows, cols=cols)
        table.style = "Table Grid"

        # 表头
        if td.headers:
            for i, h in enumerate(td.headers):
                cell = table.rows[0].cells[i]
                cell.text = str(h)
                if td.header_style:
                    hs = td.header_style
                    for p in cell.paragraphs:
                        for run in p.runs:
                            _apply_font(run, hs.font_name, hs.font_size,
                                        True if hs.bold is None else hs.bold, False, hs.font_color)
                    if hs.bg_color:
                        shading = parse_xml(
                            f'<w:shd {nsdecls("w")} w:fill="{hs.bg_color}"/>'
                        )
                        cell._element.get_or_add_tcPr().append(shading)

        # 数据行
        for ri, row_data in enumerate(td.rows):
            row_idx = ri + (1 if td.headers else 0)
            for ci, val in enumerate(row_data):
                if ci < cols:
                    cell = table.rows[row_idx].cells[ci]
                    cell.text = str(val)

        # 列宽
        if td.column_widths:
            for row in table.rows:
                for ci, w in enumerate(td.column_widths):
                    if ci < len(row.cells):
                        row.cells[ci].width = Inches(w)

        # 合并单元格
        if td.merged_cells:
            for mc in td.merged_cells:
                try:
                    start_cell_str, end_cell_str = mc.range.split(":")
                    start_col = _col_letter_to_index(start_cell_str)
                    start_row = int(re.search(r"\d+", start_cell_str).group()) - 1
                    end_col = _col_letter_to_index(end_cell_str)
                    end_row = int(re.search(r"\d+", end_cell_str).group()) - 1
                    table.cell(start_row, start_col).merge(
                        table.cell(end_row, end_col)
                    )
                except (ValueError, IndexError, AttributeError):
                    logger.warning("无法解析合并范围: %s", mc.range)

    elif t == "image":
        from .models import ImageDef
        img = item.image or ImageDef(path="")
        if os.path.exists(img.path):
            w = img.width or 4.0
            h = img.height or None
            if h:
                doc.add_picture(img.path, width=Inches(w), height=Inches(h))
            else:
                doc.add_picture(img.path, width=Inches(w))
        else:
            logger.warning("图片路径不存在，已跳过: %s", img.path)

    elif t == "page_break":
        doc.add_page_break()

    elif t == "bullet_list":
        if item.items:
            for li in item.items:
                p = doc.add_paragraph(li, style="List Bullet")
                if default_font and p.runs:
                    _apply_font(p.runs[0], default_font.name, default_font.size)

    elif t == "numbered_list":
        if item.items:
            for li in item.items:
                p = doc.add_paragraph(li, style="List Number")
                if default_font and p.runs:
                    _apply_font(p.runs[0], default_font.name, default_font.size)

    elif t == "code":
        p = doc.add_paragraph()
        run = p.add_run(item.text or "")
        _apply_font(run, "Consolas", 10, False, False, "1E1E1E")
        # 应用段落样式
        if item.style:
            _apply_paragraph_format(p, item.style)
        # 灰色背景阴影 - 使用 nil 而非 clear 以确保渲染
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F5F5F5" w:val="nil"/>')
        p._element.get_or_add_pPr().append(shading)


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
    """创建 Word (.docx) 文档。

    支持模板模式：传入已有的 docx 文件作为 template，新文档将继承其样式/页面设置/页眉页脚。
    支持场景模板：scene 参数自动应用预设排版规范（academic/report/official/contract/resume/exam）。
    场景默认值可以被 page_setup/header_footer/default_font 等显式参数覆盖。
    支持封面生成：cover_config 传入封面配置（title/subtitle/author/date/organization/recipe_id）。
    内容通过 content.sections 描述，支持 heading/paragraph/table/image/page_break/bullet_list/numbered_list 等类型。
    """
    from .models import CreateResult
    from .utils import atomic_save, ensure_output_dir, resolve_if_exists, get_file_size, validate_input_file, validate_output_path
    from .scenes import get_scene
    from .covers import select_cover_recipe, generate_cover

    try:
        validate_output_path(output_path, (".docx",))
        ensure_output_dir(output_path)
        resolved_path = resolve_if_exists(output_path, if_exists)
    except (OSError, ValueError) as e:
        return CreateResult(success=False, file_path=output_path, message=str(e))

    # 加载场景配置
    scene_profile = get_scene(scene) if scene else None

    try:
        if template:
            validate_input_file(template, (".docx", ".dotx"))
            doc = Document(template)
            logger.info("使用模板文件: %s", template)
        else:
            doc = Document()

        # 合并场景配置与显式参数（显式参数优先）
        effective_page_setup = _merge_page_setup(scene_profile, page_setup)
        effective_hf = _merge_header_footer(scene_profile, header_footer, default_font)
        effective_font = _merge_font(scene_profile, default_font)

        # 封面生成（在内容之前）
        should_show_cover = scene_profile and scene_profile.show_cover
        has_cover_config = cover_config is not None

        if should_show_cover or has_cover_config:
            cover_data = cover_config or {}
            # 标题取自封面配置或文档第一个 heading（首个内容块可能无 text，如表格）
            first_text = content.sections[0].text if content.sections else None
            cover_title = cover_data.get("title", "") or first_text or "文档标题"
            from .covers import RECIPES
            recipe = select_cover_recipe(
                doc_type=cover_data.get("doc_type", "general"),
                scene=scene,
            )
            # 允许用户指定配方
            if cover_data.get("recipe_id") and cover_data["recipe_id"] in RECIPES:
                recipe = RECIPES[cover_data["recipe_id"]]
            generate_cover(
                doc, recipe,
                title=cover_title,
                subtitle=cover_data.get("subtitle"),
                author=cover_data.get("author"),
                date=cover_data.get("date"),
                organization=cover_data.get("organization"),
                extra_lines=cover_data.get("extra_lines"),
            )

        # 页面设置
        if effective_page_setup:
            for section in doc.sections:
                _set_page_setup(section, effective_page_setup)

        # 页眉页脚
        if effective_hf:
            _set_header_footer(doc, effective_hf)

        # 修复 Heading 样式颜色（去除 python-docx 默认蓝色主题色）
        _fix_heading_style_colors(doc)

        # 写入内容 - 传递场景信息给每个 item
        scene_data = {"profile": scene_profile} if scene_profile else {}
        for item in content.sections:
            _add_section_item_scene(doc, item, effective_font, scene_data)

        atomic_save(doc.save, resolved_path)
        size = get_file_size(resolved_path)
        scene_msg = f"（场景: {scene}）" if scene else ""
        cover_msg = "+封面" if (should_show_cover or has_cover_config) else ""
        logger.info("Word 文档创建成功: %s (%d 字节)%s%s", resolved_path, size, scene_msg, cover_msg)

        return CreateResult(
            success=True,
            file_path=resolved_path,
            file_size=size,
            message=f"Word 文档创建成功: {resolved_path}{scene_msg}{'（含封面）' if (should_show_cover or has_cover_config) else ''}",
        )
    except Exception as e:
        # 输入校验错误会直接返回给 MCP 调用方；无需把预期错误刷成完整堆栈，
        # 否则 Reasonix 的插件日志会被无关异常信息淹没。
        logger.error("Word 文档创建失败: %s", e, exc_info=not isinstance(e, ValueError))
        return CreateResult(
            success=False,
            file_path=resolved_path,
            message=f"Word 文档创建失败: {str(e)}",
        )


def edit_word_document(
    file_path: str,
    append_sections: list[SectionItem],
    default_font: FontConfig | None = None,
) -> EditResult:
    """向已有 Word 文档追加内容"""
    from .models import EditResult
    from .utils import atomic_save, validate_input_file

    # An Explorer-created .docx can remain a zero-byte placeholder until the
    # open Word/WPS instance performs its first save.  Try that active instance
    # before applying the disk OOXML validator.
    from .office_com import try_edit_open_word_document
    com_result = try_edit_open_word_document(file_path, append_sections, default_font)
    if com_result is not None:
        return com_result

    try:
        validate_input_file(file_path, (".docx",))
    except ValueError as e:
        return EditResult(
            success=False,
            file_path=file_path,
            message=str(e),
        )

    try:
        doc = Document(file_path)
        count = 0
        for item in append_sections:
            _add_section_item(doc, item, default_font)
            count += 1
        atomic_save(doc.save, file_path)
        logger.info("已追加 %d 个内容块到 Word 文档: %s", count, file_path)
        return EditResult(
            success=True,
            file_path=file_path,
            operation_count=count,
            message=f"已追加 {count} 个内容块到 Word 文档",
        )
    except Exception as e:
        logger.error("编辑 Word 文档失败: %s", e, exc_info=True)
        return EditResult(
            success=False,
            file_path=file_path,
            message=f"编辑 Word 文档失败: {str(e)}",
        )


def _fix_heading_style_colors(doc):
    """修复 Heading 样式定义，去除 python-docx 默认模板中的蓝色主题色

    python-docx 的默认模板中 Heading 1-4 使用了蓝色主题色
    (w:themeColor="accent1")。虽然在 _add_section_item_scene 中通过
    run-level 设置了黑色，但样式定义中的蓝色仍可能在某些渲染场景
    （如 Word 的样式检查器、PDF 导出）中回退生效。此函数从样式定义
    层面彻底消除蓝色。
    """
    from docx.shared import RGBColor
    for level in range(1, 9):
        try:
            style = doc.styles[f"Heading {level}"]
        except KeyError:
            continue
        # 强制设置样式字体颜色为黑色
        if style.font.color and style.font.color.rgb:
            style.font.color.rgb = RGBColor(0, 0, 0)
        # 同时清除 themeColor 引用（通过直接操作 XML）
        for el in style.element.iter(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color"
        ):
            if el.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}themeColor"):
                # 移除 themeColor/themeShade 属性，仅保留直接 val="000000"
                el.attrib.clear()
                el.set(
                    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val",
                    "000000",
                )


# ── 场景模板辅助函数 ────────────────────────────────────────────────


def _merge_page_setup(scene_profile, explicit_page_setup):
    """合并场景配置和显式页面设置，显式参数优先"""
    from .models import PageSetup as PS

    if scene_profile is None and explicit_page_setup is None:
        return None

    merged = PS()
    if scene_profile:
        merged.paper_size = scene_profile.paper_size  # type: ignore
        merged.margin_top = scene_profile.margin_top
        merged.margin_bottom = scene_profile.margin_bottom
        merged.margin_left = scene_profile.margin_left
        merged.margin_right = scene_profile.margin_right

    # 显式参数覆盖
    if explicit_page_setup:
        if explicit_page_setup.orientation is not None:
            merged.orientation = explicit_page_setup.orientation
        if explicit_page_setup.paper_size is not None:
            merged.paper_size = explicit_page_setup.paper_size
        if explicit_page_setup.margin_top is not None:
            merged.margin_top = explicit_page_setup.margin_top
        if explicit_page_setup.margin_bottom is not None:
            merged.margin_bottom = explicit_page_setup.margin_bottom
        if explicit_page_setup.margin_left is not None:
            merged.margin_left = explicit_page_setup.margin_left
        if explicit_page_setup.margin_right is not None:
            merged.margin_right = explicit_page_setup.margin_right

    return merged


def _merge_header_footer(scene_profile, explicit_hf, default_font):
    """合并场景配置和显式页眉页脚"""
    from .models import HeaderFooter as HF

    if scene_profile is None and explicit_hf is None:
        return None

    merged = HF()
    if scene_profile and scene_profile.page_number_style != "none":
        merged.footer_text = "— {page} —"

    if explicit_hf:
        if explicit_hf.header_text is not None:
            merged.header_text = explicit_hf.header_text
        if explicit_hf.footer_text is not None:
            merged.footer_text = explicit_hf.footer_text

    return merged


def _merge_font(scene_profile, explicit_font):
    """合并场景配置和显式字体"""
    from .models import FontConfig as FC

    if scene_profile is None and explicit_font is None:
        return None

    merged = FC()
    if scene_profile:
        merged.name = scene_profile.font_name
        merged.size = scene_profile.body_size

    if explicit_font:
        if explicit_font.name is not None:
            merged.name = explicit_font.name
        if explicit_font.size is not None:
            merged.size = explicit_font.size
        if explicit_font.bold is not None:
            merged.bold = explicit_font.bold
        if explicit_font.italic is not None:
            merged.italic = explicit_font.italic
        if explicit_font.color is not None:
            merged.color = explicit_font.color

    return merged


def _add_section_item_scene(doc, item: SectionItem, default_font, scene_data: dict):
    """向文档添加内容块（场景感知版本）"""
    from .models import ParagraphStyle as PS

    profile = scene_data.get("profile")
    t = item.type

    if t == "heading":
        level = item.level or 1
        p = doc.add_heading(item.text or "", level=level)
        # 场景 heading 字体（重置颜色为黑色，避免 python-docx 内置 heading 的蓝色主题色）
        hfont = (profile.heading_font if profile else None) or (default_font.name if default_font else None)
        hsize = (profile.heading_sizes.get(level) if profile else None) or (default_font.size if default_font else None)
        for run in p.runs:
            _apply_font(run, hfont, hsize, color="000000")
        # 段落格式
        if item.style:
            _apply_paragraph_format(p, item.style)
        elif profile:
            _apply_paragraph_format(p, PS(
                alignment=profile.alignment,
                line_spacing=profile.line_spacing,
                first_line_indent=profile.first_line_indent,
            ))

    elif t == "paragraph":
        p = doc.add_paragraph()
        style = item.style or PS()
        s_font = style.font_name or (profile.font_name if profile else None) or (default_font.name if default_font else None)
        s_size = style.font_size or (profile.body_size if profile else None) or (default_font.size if default_font else None)
        _add_text_runs(p, item.text or "", s_font, s_size, style.bold, style.italic, style.color)
        # 合并场景段落格式
        merged_style = PS(
            alignment=style.alignment or (profile.alignment if profile else None),
            line_spacing=style.line_spacing or (profile.line_spacing if profile else None),
            first_line_indent=style.first_line_indent or (profile.first_line_indent if profile else None),
            space_before=style.space_before,
            space_after=style.space_after,
        )
        _apply_paragraph_format(p, merged_style)

    elif t == "table":
        from .models import TableDef
        td = item.table or TableDef()
        rows = len(td.rows) + (1 if td.headers else 0)
        cols = max(len(td.headers) if td.headers else 0,
                   max((len(r) for r in td.rows), default=0))
        if rows == 0 or cols == 0:
            logger.warning("跳过空表格（rows=%d, cols=%d）", rows, cols)
            return
        table = doc.add_table(rows=rows, cols=cols)
        table.style = "Table Grid"
        # 表头
        if td.headers:
            for i, h in enumerate(td.headers):
                cell = table.rows[0].cells[i]
                cell.text = str(h)
                if td.header_style:
                    hs = td.header_style
                    for p in cell.paragraphs:
                        for run in p.runs:
                            _apply_font(run, hs.font_name, hs.font_size,
                                        True if hs.bold is None else hs.bold, False, hs.font_color)
                    if hs.bg_color:
                        shading = parse_xml(
                            f'<w:shd {nsdecls("w")} w:fill="{hs.bg_color}"/>'
                        )
                        cell._element.get_or_add_tcPr().append(shading)
        # 数据行
        for ri, row_data in enumerate(td.rows):
            row_idx = ri + (1 if td.headers else 0)
            for ci, val in enumerate(row_data):
                if ci < cols:
                    cell = table.rows[row_idx].cells[ci]
                    cell.text = str(val)
        # 列宽
        if td.column_widths:
            for row in table.rows:
                for ci, w in enumerate(td.column_widths):
                    if ci < len(row.cells):
                        row.cells[ci].width = Inches(w)
        # 合并单元格
        if td.merged_cells:
            for mc in td.merged_cells:
                try:
                    start_cell_str, end_cell_str = mc.range.split(":")
                    start_col = _col_letter_to_index(start_cell_str)
                    start_row = int(re.search(r"\d+", start_cell_str).group()) - 1
                    end_col = _col_letter_to_index(end_cell_str)
                    end_row = int(re.search(r"\d+", end_cell_str).group()) - 1
                    table.cell(start_row, start_col).merge(
                        table.cell(end_row, end_col)
                    )
                except (ValueError, IndexError, AttributeError):
                    logger.warning("无法解析合并范围: %s", mc.range)

    elif t == "image":
        from .models import ImageDef
        img = item.image or ImageDef(path="")
        if os.path.exists(img.path):
            w = img.width or 4.0
            h = img.height or None
            if h:
                doc.add_picture(img.path, width=Inches(w), height=Inches(h))
            else:
                doc.add_picture(img.path, width=Inches(w))
        else:
            logger.warning("图片路径不存在，已跳过: %s", img.path)

    elif t == "page_break":
        doc.add_page_break()

    elif t == "bullet_list":
        if item.items:
            for li in item.items:
                p = doc.add_paragraph(li, style="List Bullet")
                if profile:
                    _apply_paragraph_format(p, PS(
                        alignment=profile.alignment,
                        line_spacing=profile.line_spacing,
                    ))

    elif t == "numbered_list":
        if item.items:
            for li in item.items:
                p = doc.add_paragraph(li, style="List Number")
                if profile:
                    _apply_paragraph_format(p, PS(
                        alignment=profile.alignment,
                        line_spacing=profile.line_spacing,
                    ))

    elif t == "code":
        p = doc.add_paragraph()
        run = p.add_run(item.text or "")
        _apply_font(run, "Consolas", 10, False, False, "1E1E1E")
        if item.style:
            _apply_paragraph_format(p, item.style)
        elif profile:
            _apply_paragraph_format(p, PS(
                alignment=profile.alignment,
                line_spacing=profile.line_spacing,
            ))
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F5F5F5" w:val="nil"/>')
        p._element.get_or_add_pPr().append(shading)

    elif t == "toc":
        """生成目录 (Table of Contents) 字段"""
        toc_levels = item.toc_levels or "1-3"
        p = doc.add_paragraph()
        # 目录标题
        run = p.add_run(item.text or "目录")
        run.font.size = Pt(16)
        run.bold = True
        if profile:
            run.font.name = profile.font_name
            run._element.rPr.rFonts.set(qn("w:eastAsia"), profile.font_name)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        # TOC 字段（在下一段落）
        toc_p = doc.add_paragraph()
        # TOC 字段指令
        fld_begin = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="begin"/>')
        toc_p._element.append(fld_begin)
        instr_text = f' TOC \\o "{toc_levels}" \\h \\z \\u '
        instr = parse_xml(f'<w:instrText {nsdecls("w")} xml:space="preserve">{instr_text}</w:instrText>')
        toc_p._element.append(instr)
        fld_separate = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="separate"/>')
        toc_p._element.append(fld_separate)
        # 占位文本（Word 会在打开时更新）
        placeholder = doc.add_paragraph('（请在 Word 中右键点击目录区域，选择「更新域」以生成完整目录）')
        for r in placeholder.runs:
            r.font.size = Pt(9)
            r.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
            r.italic = True
        fld_end = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="end"/>')
        toc_p._element.append(fld_end)
        # TOC 后分页
        doc.add_page_break()

    elif t == "section_break":
        """添加分节符（新页开始）"""
        doc.add_section()
        logger.debug("添加分节符")
