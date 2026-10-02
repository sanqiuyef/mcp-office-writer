
"""
mcp-office-writer - PowerPoint (.pptx) 写入/编辑逻辑
"""
from __future__ import annotations

import os
from typing import TYPE_CHECKING

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE

if TYPE_CHECKING:
    from .models import SlideDef, SlideItem, SlideTableDef, SlideImageDef
    from .models import SlideShapeDef, TextStyle, CreateResult, EditResult
    from .models import ChartDef, FrameworkDiagram


# ── 预设配色方案 ──────────────────────────────────────────────────

_PRESET_COLOR_SCHEMES: dict[str, dict] = {
    "professional-blue": {
        "title_color": "000000", "body_color": "333333",
        "accent_color": "2563EB", "background_color": "F8FAFC",
        "shape_fill_color": "DBEAFE", "table_header_color": "1A3C6D",
        "title_font": "微软雅黑", "body_font": "微软雅黑",
        "title_size": 32, "body_size": 18,
        "display_name": "商务蓝", "description": "专业商务风格，深蓝主色配合浅灰背景",
    },
    "modern-dark": {
        "title_color": "FFFFFF", "body_color": "D1D5DB",
        "accent_color": "60A5FA", "background_color": "1F2937",
        "shape_fill_color": "374151", "table_header_color": "111827",
        "title_font": "微软雅黑", "body_font": "微软雅黑",
        "title_size": 32, "body_size": 18,
        "display_name": "暗夜模式", "description": "现代深色主题，白字深灰底，适合大屏演示",
    },
    "warm-earth": {
        "title_color": "7C2D12", "body_color": "44403C",
        "accent_color": "D97706", "background_color": "FFF7ED",
        "shape_fill_color": "FED7AA", "table_header_color": "7C2D12",
        "title_font": "微软雅黑", "body_font": "微软雅黑",
        "title_size": 32, "body_size": 18,
        "display_name": "暖色大地", "description": "温暖自然风格，橙色与棕色搭配",
    },
    "elegant-minimal": {
        "title_color": "18181B", "body_color": "52525B",
        "accent_color": "8B5CF6", "background_color": "FFFFFF",
        "shape_fill_color": "F4F4F5", "table_header_color": "18181B",
        "title_font": "微软雅黑", "body_font": "微软雅黑",
        "title_size": 28, "body_size": 16,
        "display_name": "极简优雅", "description": "黑白灰极简风格，紫色点缀，适合学术/设计",
    },
    "vibrant": {
        "title_color": "BE185D", "body_color": "27272A",
        "accent_color": "14B8A6", "background_color": "FFF1F2",
        "shape_fill_color": "FCE7F3", "table_header_color": "BE185D",
        "title_font": "微软雅黑", "body_font": "微软雅黑",
        "title_size": 34, "body_size": 20,
        "display_name": "活力缤纷", "description": "高对比度配色，玫红+青绿，适合创意/营销",
    },
    "corporate-green": {
        "title_color": "064E3B", "body_color": "374151",
        "accent_color": "10B981", "background_color": "ECFDF5",
        "shape_fill_color": "D1FAE5", "table_header_color": "064E3B",
        "title_font": "微软雅黑", "body_font": "微软雅黑",
        "title_size": 32, "body_size": 18,
        "display_name": "企业绿", "description": "稳重大气，深绿主色配合浅绿背景",
    },
    "tech-purple": {
        "title_color": "4C1D95", "body_color": "374151",
        "accent_color": "7C3AED", "background_color": "F5F3FF",
        "shape_fill_color": "EDE9FE", "table_header_color": "4C1D95",
        "title_font": "微软雅黑", "body_font": "微软雅黑",
        "title_size": 32, "body_size": 18,
        "display_name": "科技紫", "description": "科技感紫色渐变，适合互联网/AI/创新主题",
    },
    "classic-red": {
        "title_color": "991B1B", "body_color": "44403C",
        "accent_color": "DC2626", "background_color": "FEF2F2",
        "shape_fill_color": "FEE2E2", "table_header_color": "991B1B",
        "title_font": "微软雅黑", "body_font": "微软雅黑",
        "title_size": 34, "body_size": 20,
        "display_name": "经典中国红", "description": "庄重典雅的中国红，适合党政/党建/传统节日主题",
    },
    "fresh-nature": {
        "title_color": "166534", "body_color": "374151",
        "accent_color": "22C55E", "background_color": "F0FDF4",
        "shape_fill_color": "DCFCE7", "table_header_color": "166534",
        "title_font": "微软雅黑", "body_font": "微软雅黑",
        "title_size": 30, "body_size": 18,
        "display_name": "清新自然", "description": "清爽绿色系，适合环保/健康/农业/教育主题",
    },
    "ocean-blue": {
        "title_color": "000000", "body_color": "334155",
        "accent_color": "0EA5E9", "background_color": "F0F9FF",
        "shape_fill_color": "E0F2FE", "table_header_color": "0C4A6E",
        "title_font": "微软雅黑", "body_font": "微软雅黑",
        "title_size": 32, "body_size": 18,
        "display_name": "海洋蓝", "description": "清爽海洋蓝，适合医疗/教育/水利/科技主题",
    },
    "sunset-orange": {
        "title_color": "9A3412", "body_color": "44403C",
        "accent_color": "F97316", "background_color": "FFF7ED",
        "shape_fill_color": "FFEDD5", "table_header_color": "9A3412",
        "title_font": "微软雅黑", "body_font": "微软雅黑",
        "title_size": 32, "body_size": 20,
        "display_name": "日落暖橙", "description": "温暖活力的橙色系，适合餐饮/旅游/创意/生活主题",
    },
    "monochrome-gray": {
        "title_color": "1E293B", "body_color": "475569",
        "accent_color": "64748B", "background_color": "F8FAFC",
        "shape_fill_color": "F1F5F9", "table_header_color": "1E293B",
        "title_font": "微软雅黑", "body_font": "微软雅黑",
        "title_size": 28, "body_size": 16,
        "display_name": "高级灰", "description": "克制高级灰，适合设计/建筑/金融/高端品牌主题",
    },
}


ALIGN_MAP = {
    "left": PP_ALIGN.LEFT,
    "center": PP_ALIGN.CENTER,
    "right": PP_ALIGN.RIGHT,
}


def _hex_to_rgb(hex_color: str) -> RGBColor:
    h = hex_color.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _apply_text_style(run, style: TextStyle | None):
    if style is None:
        return
    if style.font_name:
        run.font.name = style.font_name
    if style.font_size:
        run.font.size = Pt(style.font_size)
    if style.bold is not None:
        run.font.bold = style.bold
    if style.italic is not None:
        run.font.italic = style.italic
    if style.color:
        run.font.color.rgb = _hex_to_rgb(style.color)


def _set_text_frame_alignment(tf, alignment_str: str | None):
    if alignment_str and alignment_str in ALIGN_MAP:
        tf.paragraphs[0].alignment = ALIGN_MAP[alignment_str]


def _add_paragraph_item(tf, item: SlideItem):
    """向文本框添加段落内容"""
    from .models import TextStyle

    if not item.text:
        return
    p = tf.add_paragraph()
    p.text = item.text
    p.level = item.level
    style = item.style or TextStyle()
    _apply_text_style(p.runs[0], style) if p.runs else None
    if style.alignment:
        p.alignment = ALIGN_MAP.get(style.alignment, PP_ALIGN.LEFT)


def _add_slide_table(slide, td: SlideTableDef, left, top, width, height):
    """在幻灯片上添加表格"""
    rows = len(td.rows) + (1 if td.headers else 0)
    cols = max(len(td.headers) if td.headers else 0,
               max((len(r) for r in td.rows), default=0))
    if rows == 0 or cols == 0:
        return

    table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
    table = table_shape.table

    if td.headers:
        for ci, h in enumerate(td.headers):
            cell = table.cell(0, ci)
            cell.text = str(h)

    for ri, row_data in enumerate(td.rows, 1 if td.headers else 0):
        for ci, val in enumerate(row_data):
            if ci < cols:
                table.cell(ri, ci).text = str(val)


def _add_slide_image(slide, img_def: SlideImageDef):
    """在幻灯片上添加图片（支持本地 path 或 URL 下载）"""
    if img_def.url:
        from .utils import download_image
        try:
            file_path = download_image(img_def.url)
        except Exception:
            # 下载失败时静默跳过（避免 PPT 创建中断）
            return
    elif img_def.path:
        file_path = img_def.path
    else:
        return

    if not os.path.exists(file_path):
        return
    left = Inches(img_def.x or 0.5)
    top = Inches(img_def.y or 1.5)
    width = Inches(img_def.width or 4.0)
    height = Inches(img_def.height) if img_def.height else None
    if height:
        slide.shapes.add_picture(file_path, left, top, width, height)
    else:
        slide.shapes.add_picture(file_path, left, top, width)


def _add_slide_shape(slide, shape_def: SlideShapeDef):
    """在幻灯片上添加形状"""
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.dml import MSO_LINE_DASH_STYLE

    shape_map = {
        "rect": MSO_SHAPE.RECTANGLE,
        "rounded_rect": MSO_SHAPE.ROUNDED_RECTANGLE,
        "oval": MSO_SHAPE.OVAL,
        "arrow": MSO_SHAPE.RIGHT_ARROW,
        "cylinder": MSO_SHAPE.CAN,
        "diamond": MSO_SHAPE.DIAMOND,
    }

    shape_type = shape_map.get(shape_def.shape_type, MSO_SHAPE.RECTANGLE)
    left = Inches(shape_def.x or 0.5)
    top = Inches(shape_def.y or 1.5)
    width = Inches(shape_def.width)
    height = Inches(shape_def.height)

    shape = slide.shapes.add_shape(shape_type, left, top, width, height)

    if shape_def.no_fill:
        shape.fill.background()
    elif shape_def.bg_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = _hex_to_rgb(shape_def.bg_color)

    if shape_def.line_color:
        shape.line.color.rgb = _hex_to_rgb(shape_def.line_color)
    if shape_def.line_width is not None:
        shape.line.width = Pt(shape_def.line_width)
    if shape_def.line_dash:
        shape.line.dash_style = MSO_LINE_DASH_STYLE.DASH

    if shape_def.text:
        tf = shape.text_frame
        if shape_def.vertical_text:
            # python-pptx 无竖排 API，直接注入 bodyPr vert 属性
            from pptx.oxml.ns import qn
            body_pr = tf._txBody.find(qn("a:bodyPr"))
            if body_pr is not None:
                body_pr.set("vert", "vert270")
        # 横排文字按框宽自动换行（避免长公式/长句溢出）；竖排不换行
        tf.word_wrap = not shape_def.vertical_text
        tf.text = shape_def.text
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        for run in tf.paragraphs[0].runs:
            if shape_def.font_color:
                run.font.color.rgb = _hex_to_rgb(shape_def.font_color)
            if shape_def.bold:
                run.font.bold = True
            run.font.size = Pt(shape_def.font_size or 12)


def _add_slide_connector(slide, connector_def):
    """在幻灯片上添加连接线/箭头"""
    from pptx.enum.shapes import MSO_CONNECTOR_TYPE
    from pptx.enum.dml import MSO_LINE_DASH_STYLE

    connector = slide.shapes.add_connector(
        MSO_CONNECTOR_TYPE.STRAIGHT,
        Inches(connector_def.x1), Inches(connector_def.y1),
        Inches(connector_def.x2), Inches(connector_def.y2),
    )
    if connector_def.color:
        connector.line.color.rgb = _hex_to_rgb(connector_def.color)
    connector.line.width = Pt(connector_def.width)
    if connector_def.dash:
        connector.line.dash_style = MSO_LINE_DASH_STYLE.DASH

    # 箭头挂在 a:ln 上，而 a:ln 位于 p:spPr 之下；直接 find(qn("a:ln")) 匹配不到，
    # 必须经 LineFormat._get_or_add_ln() 取（连接线颜色/宽度已创建该元素）。
    from pptx.oxml.ns import qn
    from lxml import etree
    line_elem = connector.line._get_or_add_ln()
    if connector_def.arrowhead:
        tail = line_elem.find(qn("a:tailEnd"))
        if tail is None:
            tail = etree.SubElement(line_elem, qn("a:tailEnd"))
        tail.set("type", {"triangle": "triangle", "open": "oval", "stealth": "stealth"}.get(
            connector_def.arrowhead_style, "triangle"))
        tail.set("w", "med")
        tail.set("len", "med")
    if connector_def.begin_arrowhead:
        head = line_elem.find(qn("a:headEnd"))
        if head is None:
            head = etree.SubElement(line_elem, qn("a:headEnd"))
        head.set("type", "triangle")
        head.set("w", "med")
        head.set("len", "med")


def _build_slide_body(slide, items: list[SlideItem] | None, default_left, default_top, default_width, default_height):
    """向幻灯片主体区域写入内容元素"""
    if not items:
        return

    for item in items:
        if item.type == "paragraph":
            # 将段落添加到现有的标题/正文占位符
            text_frame = None
            if slide.placeholders:
                # 找第一个正文占位符
                for ph in slide.placeholders:
                    if ph.placeholder_format.idx == 1:  # 正文
                        text_frame = ph.text_frame
                        break
                if text_frame is None and len(slide.placeholders) > 1:
                    text_frame = slide.placeholders[1].text_frame

            if text_frame:
                _add_paragraph_item(text_frame, item)
            else:
                # 没有占位符时手动加文本框
                txBox = slide.shapes.add_textbox(default_left, default_top, default_width, default_height)
                tf = txBox.text_frame
                tf.text = item.text or ""

        elif item.type == "table" and item.table:
            td = item.table
            _add_slide_table(slide, td, default_left, Inches(2.0), Inches(7.0), Inches(2.5))

        elif item.type == "image" and item.image:
            _add_slide_image(slide, item.image)

        elif item.type == "shape" and item.shape:
            _add_slide_shape(slide, item.shape)

        elif item.type == "connector" and item.connector:
            _add_slide_connector(slide, item.connector)

        elif item.type == "bullet_list" or item.type == "numbered_list":
            # 添加项目符号/编号列表
            txBox = slide.shapes.add_textbox(default_left, Inches(1.8), default_width, default_height)
            tf = txBox.text_frame
            tf.word_wrap = True
            for i, list_item in enumerate(item.items or []):
                if i == 0:
                    p = tf.paragraphs[0]
                else:
                    p = tf.add_paragraph()
                p.text = list_item
                p.level = item.level
                p.alignment = PP_ALIGN.LEFT
                if item.type == "bullet_list":
                    p.level = 0  # 项目符号由布局决定
                if item.style:
                    _apply_text_style(p.runs[0], item.style) if p.runs else None

        elif item.type == "chart" and item.chart:
            _add_slide_chart(slide, item.chart, default_left, Inches(1.5), Inches(8.0), Inches(4.5))


def _add_slide_chart(slide, chart_def: ChartDef, left, top, width, height):
    """在幻灯片上添加图表"""
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE

    chart_type_map = {
        "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "bar": XL_CHART_TYPE.BAR_CLUSTERED,
        "line": XL_CHART_TYPE.LINE_MARKERS,
        "pie": XL_CHART_TYPE.PIE,
    }
    ctype = chart_type_map.get(chart_def.chart_type, XL_CHART_TYPE.COLUMN_CLUSTERED)

    chart_data = CategoryChartData()
    chart_data.categories = chart_def.categories
    for series in chart_def.series:
        chart_data.add_series(series.name, series.values)

    chart_frame = slide.shapes.add_chart(ctype, left, top, width, height, chart_data)
    chart = chart_frame.chart

    if chart_def.title:
        chart.has_title = True
        chart.chart_title.text_frame.paragraphs[0].text = chart_def.title
        chart.chart_title.text_frame.paragraphs[0].font.size = Pt(12)

    if chart_def.has_legend:
        chart.has_legend = True
        from pptx.enum.chart import XL_LEGEND_POSITION
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM


def _set_slide_bg(slide, bg_color: str | None):
    """设置幻灯片背景色"""
    if bg_color is None:
        return
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = _hex_to_rgb(bg_color)


def _create_slide(prs, slide_def: SlideDef, layout_idx: int):
    """创建一张幻灯片"""

    layout = prs.slide_layouts[layout_idx]
    slide = prs.slides.add_slide(layout)

    # 设置背景
    _set_slide_bg(slide, slide_def.bg_color)

    # 标题
    if slide_def.title and slide.shapes.title:
        slide.shapes.title.text = slide_def.title

    # 副标题 (title/section 布局)
    if slide_def.subtitle:
        from pptx.enum.shapes import PP_PLACEHOLDER
        for ph in slide.placeholders:
            if ph.placeholder_format.type == PP_PLACEHOLDER.SUBTITLE:
                ph.text = slide_def.subtitle
                break

    # 主体内容
    if slide_def.body:
        _build_slide_body(slide, slide_def.body, Inches(0.5), Inches(1.5), Inches(8.0), Inches(4.5))

    # 双栏布局
    if slide_def.left or slide_def.right:
        _build_slide_body(slide, slide_def.left, Inches(0.5), Inches(1.5), Inches(4.0), Inches(4.5))
        _build_slide_body(slide, slide_def.right, Inches(5.0), Inches(1.5), Inches(4.0), Inches(4.5))

    # 演讲者备注
    if slide_def.notes:
        notes_slide = slide.notes_slide
        notes_slide.notes_text_frame.text = slide_def.notes

    return slide


def create_powerpoint(
    output_path: str,
    slides: list[SlideDef],
    template: str | None = None,
    if_exists: str = "overwrite",
) -> CreateResult:
    """创建 PowerPoint 演示文稿"""
    from .models import CreateResult
    from .utils import atomic_save, ensure_output_dir, get_file_size, get_ppt_layout_index, resolve_if_exists, validate_input_file, validate_output_path

    try:
        validate_output_path(output_path, (".pptx",))
        ensure_output_dir(output_path)
        resolved_path = resolve_if_exists(output_path, if_exists)
    except (OSError, ValueError) as e:
        return CreateResult(success=False, file_path=output_path, message=str(e))

    try:
        if template:
            validate_input_file(template, (".pptx", ".potx"))
            prs = Presentation(template)
        else:
            prs = Presentation()

        # 支持自定义幻灯片尺寸（取第一个指定宽高的 slide_def）
        for sd in slides:
            if sd.slide_width and sd.slide_height:
                prs.slide_width = Inches(sd.slide_width)
                prs.slide_height = Inches(sd.slide_height)
                break

        for slide_def in slides:
            layout_idx = get_ppt_layout_index(slide_def.layout)
            _create_slide(prs, slide_def, layout_idx)

        atomic_save(prs.save, resolved_path)
        size = get_file_size(resolved_path)

        return CreateResult(
            success=True,
            file_path=resolved_path,
            file_size=size,
            slides_count=len(slides),
            message=f"PowerPoint 演示文稿创建成功: {resolved_path}",
        )
    except Exception as e:
        return CreateResult(
            success=False,
            file_path=resolved_path,
            message=f"PowerPoint 演示文稿创建失败: {str(e)}",
        )


def edit_powerpoint(
    file_path: str,
    append_slides: list[SlideDef],
) -> EditResult:
    """向已有 PPT 追加幻灯片"""
    from .models import EditResult
    from .utils import atomic_save, get_ppt_layout_index, validate_input_file

    # 优先尝试 COM 活动文档编辑（实时交互查看效果）
    from .office_com import try_edit_open_powerpoint
    com_result = try_edit_open_powerpoint(file_path, append_slides)
    if com_result is not None:
        return com_result

    try:
        validate_input_file(file_path, (".pptx",))
    except ValueError as e:
        return EditResult(
            success=False,
            file_path=file_path,
            message=str(e),
        )

    try:
        prs = Presentation(file_path)
        count = 0
        for slide_def in append_slides:
            layout_idx = get_ppt_layout_index(slide_def.layout)
            _create_slide(prs, slide_def, layout_idx)
            count += 1
        atomic_save(prs.save, file_path)
        return EditResult(
            success=True,
            file_path=file_path,
            operation_count=count,
            message=f"已追加 {count} 张幻灯片到 PPT",
        )
    except Exception as e:
        return EditResult(
            success=False,
            file_path=file_path,
            message=f"编辑 PowerPoint 失败: {str(e)}",
        )


# ── 配色方案工具 ──────────────────────────────────────────────────

def list_color_schemes() -> list[dict]:
    """列出所有预设配色方案及其详情。"""
    results = []
    for name, config in _PRESET_COLOR_SCHEMES.items():
        results.append({"name": name, "display_name": config["display_name"], "description": config["description"],
                        "preview": {"title_color": config["title_color"], "body_color": config["body_color"],
                                     "accent_color": config["accent_color"], "background_color": config["background_color"]}})
    return results


def apply_color_to_slides(
    file_path: str,
    scheme_name: str = "professional-blue",
) -> dict:
    """对已有 PPT 文件中的所有幻灯片应用指定配色方案。

    将根据配色方案更新每张幻灯片的标题颜色、正文颜色、背景色、
    形状填充色和表格表头颜色。
    """
    from .utils import atomic_save, validate_input_file

    try:
        validate_input_file(file_path, (".pptx",))
    except ValueError as e:
        return {"success": False, "file_path": file_path, "message": str(e)}

    if scheme_name not in _PRESET_COLOR_SCHEMES:
        names = ", ".join(_PRESET_COLOR_SCHEMES)
        return {"success": False, "file_path": file_path, "message": f"未知配色方案: {scheme_name}。可用: {names}"}

    scheme = _PRESET_COLOR_SCHEMES[scheme_name]
    try:
        prs = Presentation(file_path)
        slides_updated = 0

        for slide in prs.slides:
            # 背景色
            if scheme["background_color"]:
                slide.background.fill.solid()
                slide.background.fill.fore_color.rgb = _hex_to_rgb(scheme["background_color"])

            for shape in slide.shapes:
                # 表格：表头行白字 + 表头背景色，数据行套正文样式
                # （GraphicFrame 没有 has_fill/has_text_frame 属性，先单独处理）
                if getattr(shape, "has_table", False):
                    table = shape.table
                    for row_index, row in enumerate(table.rows):
                        for cell in row.cells:
                            for paragraph in cell.text_frame.paragraphs:
                                for run in paragraph.runs:
                                    if row_index == 0:
                                        run.font.color.rgb = _hex_to_rgb("FFFFFF")
                                    else:
                                        run.font.color.rgb = _hex_to_rgb(scheme["body_color"])
                                        run.font.name = scheme["body_font"]
                                        run.font.size = Pt(scheme["body_size"])
                            if row_index == 0:
                                cell.fill.solid()
                                cell.fill.fore_color.rgb = _hex_to_rgb(scheme["table_header_color"])
                    continue

                # 标题
                if shape.has_text_frame and shape == slide.shapes.title:
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.font.color.rgb = _hex_to_rgb(scheme["title_color"])
                            run.font.name = scheme["title_font"]
                            run.font.size = Pt(scheme["title_size"])

                elif shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.font.color.rgb = _hex_to_rgb(scheme["body_color"])
                            run.font.name = scheme["body_font"]
                            run.font.size = Pt(scheme["body_size"])

                # 形状填充
                if getattr(shape, "has_fill", False) and not shape.has_text_frame:
                    try:
                        shape.fill.solid()
                        shape.fill.fore_color.rgb = _hex_to_rgb(scheme["shape_fill_color"])
                    except Exception:
                        pass

            slides_updated += 1

        atomic_save(prs.save, file_path)
        return {
            "success": True,
            "file_path": file_path,
            "scheme": scheme_name,
            "slides_updated": slides_updated,
            "message": f"已对 {slides_updated} 张幻灯片应用「{scheme['display_name']}」配色方案",
        }
    except Exception as e:
        return {"success": False, "file_path": file_path, "message": f"配色应用失败: {e}"}


# ── 布局优化 ──────────────────────────────────────────────────────

def optimize_slide_layout(
    file_path: str,
) -> dict:
    """对已有 PPT 中的幻灯片进行布局优化。

    自动调整内容元素的位置和间距：
    - 多个元素时均匀分布垂直间距
    - 调整文本框大小自适应内容
    - 对齐标题位置
    """
    from .utils import atomic_save, validate_input_file

    try:
        validate_input_file(file_path, (".pptx",))
    except ValueError as e:
        return {"success": False, "file_path": file_path, "message": str(e)}

    try:
        prs = Presentation(file_path)
        slides_optimized = 0

        for slide in prs.slides:
            shapes_with_text = []
            for shape in slide.shapes:
                if shape.has_text_frame and shape.text.strip():
                    shapes_with_text.append(shape)

            if not shapes_with_text:
                continue

            # 按 Y 坐标排序
            shapes_with_text.sort(key=lambda s: s.top)

            # 计算可用区域
            slide_height = prs.slide_height
            top_margin = Inches(0.8)
            bottom_margin = Inches(0.5)
            usable_height = slide_height - top_margin - bottom_margin

            n = len(shapes_with_text)
            if n <= 1:
                slides_optimized += 1
                continue

            # 均匀分布垂直间距
            segment_height = usable_height / n
            for i, shape in enumerate(shapes_with_text):
                new_top = top_margin + segment_height * i

                # 调整文本框高度
                if shape.has_text_frame:
                    text_lines = len(shape.text_frame.text.split('\n'))
                    line_height = Pt(24)
                    new_height = max(Inches(0.6), line_height * max(text_lines, 1) + Pt(8))
                    shape.height = new_height

                shape.top = new_top

            slides_optimized += 1

        atomic_save(prs.save, file_path)
        return {
            "success": True,
            "file_path": file_path,
            "slides_optimized": slides_optimized,
            "message": f"已优化 {slides_optimized} 张幻灯片的布局",
        }
    except Exception as e:
        return {"success": False, "file_path": file_path, "message": f"布局优化失败: {e}"}


# ── 占位符清理 ───────────────────────────────────────────────────

def clean_unused_placeholders(
    file_path: str,
) -> dict:
    """清理 PPT 文件中未使用的空白占位符（如空标题框、空文本框）。

    移除每张幻灯片上没有文字内容的占位符形状，使幻灯片更加整洁。
    """
    from .utils import atomic_save, validate_input_file

    try:
        validate_input_file(file_path, (".pptx",))
    except ValueError as e:
        return {"success": False, "file_path": file_path, "message": str(e)}

    try:
        prs = Presentation(file_path)
        total_removed = 0

        for slide in prs.slides:
            shapes_to_remove = []
            for shape in slide.shapes:
                if shape.is_placeholder:
                    if not shape.has_text_frame or not shape.text_frame.text.strip():
                        shapes_to_remove.append(shape)

            for shape in shapes_to_remove:
                sp = shape._element
                sp.getparent().remove(sp)
                total_removed += 1

        atomic_save(prs.save, file_path)
        return {
            "success": True,
            "file_path": file_path,
            "placeholders_removed": total_removed,
            "message": f"已清理 {total_removed} 个未使用的占位符",
        }
    except Exception as e:
        return {"success": False, "file_path": file_path, "message": f"占位符清理失败: {e}"}


# ── 模板分析 ──────────────────────────────────────────────────────

def analyze_powerpoint_template(
    file_path: str,
) -> dict:
    """分析已有 PPT 模板，提取布局、配色和字体信息。

    返回模板的幻灯片尺寸、可用布局名称、使用到的字体列表，
    以及可提取的配色方案建议。
    """
    from .utils import validate_input_file

    try:
        validate_input_file(file_path, (".pptx", ".potx"))
    except ValueError as e:
        return {"success": False, "file_path": file_path, "message": str(e)}

    try:
        prs = Presentation(file_path)
        layout_names = [layout.name for layout in prs.slide_layouts]
        fonts_found: set[str] = set()
        color_hexes: set[str] = set()

        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            if run.font.name:
                                fonts_found.add(run.font.name)
                            try:
                                color_hexes.add(str(run.font.color.rgb))
                            except Exception:
                                pass

        return {
            "success": True,
            "file_path": file_path,
            "slide_count": len(prs.slides),
            "slide_width": round(prs.slide_width / 914400, 2),  # EMU → 英寸
            "slide_height": round(prs.slide_height / 914400, 2),
            "layout_names": layout_names,
            "fonts_used": sorted(fonts_found),
            "color_hexes_found": sorted(color_hexes),
            "message": f"模板分析完成: {len(layout_names)} 种布局, {len(fonts_found)} 种字体, {len(color_hexes)} 种颜色",
        }
    except Exception as e:
        return {"success": False, "file_path": file_path, "message": f"模板分析失败: {e}"}


# ── 框架图生成 ────────────────────────────────────────────────────

def create_framework_diagram(
    output_path: str,
    diagram: FrameworkDiagram,
    if_exists: str = "overwrite",
) -> CreateResult:
    """自动生成 PPT 框架图（阶段标签 + 节点 + 箭头连接）。

    根据 FrameworkDiagram 定义自动布局并渲染为单个幻灯片：
    - 左侧圆柱形阶段标签
    - 阶段背景矩形容器
    - 节点按指定形状（rect/rounded_rect/ellipse/diamond）渲染
    - 相邻节点自动生成箭头，也可通过 edges 指定显式连接
    - 支持 12 套配色方案
    """
    from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR_TYPE

    from .models import CreateResult
    from .utils import atomic_save, ensure_output_dir, resolve_if_exists, validate_output_path

    try:
        validate_output_path(output_path, (".pptx",))
        ensure_output_dir(output_path)
        resolved_path = resolve_if_exists(output_path, if_exists)
    except (OSError, ValueError) as e:
        return CreateResult(success=False, file_path=output_path, message=str(e))

    try:
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        slide_layout = prs.slide_layouts[6]  # blank
        slide = prs.slides.add_slide(slide_layout)

        scheme = _PRESET_COLOR_SCHEMES.get(diagram.color_scheme, _PRESET_COLOR_SCHEMES["professional-blue"])
        scheme_bg = scheme["background_color"]
        scheme_title = scheme["title_color"]
        scheme_accent = scheme["accent_color"]
        scheme_body = scheme["body_color"]
        # 连接线用中性灰弱化（PptMcp 设计规范：箭头应比内容弱，中性灰优于主色）
        arrow_color = diagram.arrow_color or "94A3B8"

        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = _hex_to_rgb(scheme_bg)

        # ── 标题 ──────────────────────────────────────────────
        title_font = diagram.font_cn
        body_font_cn = diagram.font_cn
        body_font_en = diagram.font_en
        title_fz = diagram.title_font_size

        accent_bar_y = 0.62
        if diagram.title:
            title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.15), Inches(11.5), Inches(0.45))
            tf = title_box.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            _set_mixed_font_text(p, diagram.title, title_font, body_font_en, title_fz, True,
                               _hex_to_rgb(scheme_title))
            p.alignment = PP_ALIGN.CENTER
            title_y_offset = Inches(0.7)
        else:
            title_y_offset = Inches(0.2)

        if diagram.subtitle:
            sub_title = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.5), Inches(0.3))
            sub_tf = sub_title.text_frame
            p = sub_tf.paragraphs[0]
            p.text = diagram.subtitle
            p.alignment = PP_ALIGN.CENTER
            run = p.runs[0]
            run.font.size = Pt(12)
            run.font.color.rgb = _hex_to_rgb(scheme_body)
            run.font.name = scheme["body_font"]
            title_y_offset = Inches(0.95)
            accent_bar_y = 0.9

        # 标题下 accent 短横条（居中，60% 宽）
        if diagram.title:
            bar_w = 11.5 * 0.6
            bar = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Inches(0.8 + (11.5 - bar_w) / 2), Inches(accent_bar_y),
                Inches(bar_w), Inches(0.04),
            )
            bar.fill.solid()
            bar.fill.fore_color.rgb = _hex_to_rgb(scheme_accent)
            bar.line.fill.background()

        # ── 构建节点索引 ───────────────────────────────────────
        node_id_to_pos: dict[str, dict] = {}  # id → (shape, x, y, w, h)
        stage_count = len(diagram.stages)
        if stage_count == 0:
            atomic_save(prs.save, resolved_path)
            return CreateResult(success=True, file_path=resolved_path, file_size=0, slides_count=1, message="空框架图已创建")

        # 可用区域
        slide_width_in = 13.333
        content_left = 1.15  # 阶段标签左边
        content_right = slide_width_in - 0.3
        content_width = content_right - content_left

        # 阶段之间 gap
        usable_height = 7.5 - title_y_offset.inches - 0.3
        stage_unit_h = usable_height / stage_count - diagram.stage_gap

        shape_map = {
            # 基础形状
            "rect": MSO_SHAPE.RECTANGLE,
            "rounded_rect": MSO_SHAPE.ROUNDED_RECTANGLE,
            "ellipse": MSO_SHAPE.OVAL,
            "diamond": MSO_SHAPE.DIAMOND,
            # Flowchart 流程图标准形状
            "process": MSO_SHAPE.FLOWCHART_PROCESS,
            "decision": MSO_SHAPE.FLOWCHART_DECISION,
            "data": MSO_SHAPE.FLOWCHART_DATA,
            "document": MSO_SHAPE.FLOWCHART_DOCUMENT,
            "terminator": MSO_SHAPE.FLOWCHART_TERMINATOR,
            "predefined": MSO_SHAPE.FLOWCHART_PREDEFINED_PROCESS,
            # Basic Shapes
            "cylinder": MSO_SHAPE.CAN,
            "parallelogram": MSO_SHAPE.PARALLELOGRAM,
            # Block Arrows
            "chevron": MSO_SHAPE.CHEVRON,
            "right_arrow": MSO_SHAPE.RIGHT_ARROW,
            "left_arrow": MSO_SHAPE.LEFT_ARROW,
            # Callouts
            "callout": MSO_SHAPE.ROUNDED_RECTANGULAR_CALLOUT,
            "cloud": MSO_SHAPE.CLOUD,
        }

        for si, stage in enumerate(diagram.stages):
            stage_y = title_y_offset.inches + si * (stage_unit_h + diagram.stage_gap)
            stage_color = stage.color or scheme_accent

            # ── 阶段背景容器（极淡同色系底 + 同色系细边 + 柔和阴影） ──
            bg = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(content_left + 0.55), Inches(stage_y),
                Inches(content_width - 0.55), Inches(stage_unit_h),
            )
            bg.fill.solid()
            bg.fill.fore_color.rgb = _hex_to_rgb(_mix_with_white(stage_color, 0.96))
            bg.line.color.rgb = _hex_to_rgb(_mix_with_white(stage_color, 0.88))
            bg.line.width = Pt(0.75)
            _add_shape_shadow(bg, blur_pt=10, dist_pt=2, opacity_pct=25)

            # ── 阶段标签（实心主色胶囊 + 白字） ────────────────
            pill_w, pill_h = 1.25, 0.38
            pill = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(0.25), Inches(stage_y + stage_unit_h / 2 - pill_h / 2),
                Inches(pill_w), Inches(pill_h),
            )
            pill.adjustments[0] = 0.5  # 胶囊形
            pill.fill.solid()
            pill.fill.fore_color.rgb = _hex_to_rgb(stage_color)
            pill.line.color.rgb = _hex_to_rgb(_mix_with_black(stage_color, 0.2))
            pill.line.width = Pt(0.75)
            _add_shape_shadow(pill)
            pill_tf = pill.text_frame
            pill_tf.word_wrap = False
            pill_tf.margin_left = pill_tf.margin_right = Inches(0.02)
            pill_p = pill_tf.paragraphs[0]
            pill_p.alignment = PP_ALIGN.CENTER
            _set_mixed_font_text(pill_p, stage.label, body_font_cn, body_font_en,
                                 12, True, _hex_to_rgb("FFFFFF"))

            nodes = stage.nodes
            if not nodes:
                continue

            # ── 节点水平排列 ──────────────────────────────────
            node_count = len(nodes)
            node_w = diagram.node_width
            node_h = diagram.node_height
            total_nodes_w = node_count * node_w + (node_count - 1) * 0.3
            start_x = content_left + 0.6 + max(0, (content_width - 0.6 - total_nodes_w) / 2)
            node_cy = stage_y + stage_unit_h / 2 - node_h / 2

            prev_shape = None
            for ni, node in enumerate(nodes):
                nx = start_x + ni * (node_w + 0.3)
                n_color = node.fill_color or stage_color
                shape_type = shape_map.get(node.style, MSO_SHAPE.ROUNDED_RECTANGLE)

                shape = slide.shapes.add_shape(
                    shape_type,
                    Inches(nx), Inches(node_cy),
                    Inches(node_w), Inches(node_h),
                )
                # 节点：主色填充（不再提亮）+ 同色系深边框 + 白字
                shape.fill.solid()
                shape.fill.fore_color.rgb = _hex_to_rgb(n_color)
                shape.line.color.rgb = _hex_to_rgb(_mix_with_black(n_color, 0.25))
                shape.line.width = Pt(1.0)
                _add_shape_shadow(shape)

                tf = shape.text_frame
                tf.word_wrap = True
                tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE  # 放不下时自动缩小字号
                tf.margin_left = tf.margin_right = Inches(0.04)
                tf.margin_top = tf.margin_bottom = Inches(0.02)
                # 处理 \n 换行，每条行用中英文混排字体，深底统一白字
                # 兼容真实换行符（MCP JSON 转义）与字面反斜杠+n 两种写法
                lines = node.text.replace("\\n", "\n").split("\n")
                for li, line in enumerate(lines):
                    if not line.strip() and li != 0:
                        # 空行保留段落结构；首行空则跳过（无意义）
                        continue
                    if li == 0:
                        p = tf.paragraphs[0]
                    else:
                        p = tf.add_paragraph()
                    p.alignment = PP_ALIGN.CENTER
                    _set_mixed_font_text(
                        p, line, body_font_cn, body_font_en,
                        node.font_size or 14, node.bold,
                        _hex_to_rgb("FFFFFF"),
                    )

                # 步骤编号徽章（阶段内节点 ≥2 时）：白色小圆 + 主色 Consolas 数字
                if node_count >= 2 and node.style != "terminator":
                    badge_d = 0.24
                    badge = slide.shapes.add_shape(
                        MSO_SHAPE.OVAL,
                        Inches(nx - 0.08), Inches(node_cy - 0.08),
                        Inches(badge_d), Inches(badge_d),
                    )
                    badge.fill.solid()
                    badge.fill.fore_color.rgb = _hex_to_rgb("FFFFFF")
                    badge.line.color.rgb = _hex_to_rgb(n_color)
                    badge.line.width = Pt(1.0)
                    _add_shape_shadow(badge, blur_pt=4, dist_pt=1.5, opacity_pct=35)
                    btf = badge.text_frame
                    btf.margin_left = btf.margin_right = btf.margin_top = btf.margin_bottom = 0
                    bp = btf.paragraphs[0]
                    bp.alignment = PP_ALIGN.CENTER
                    br = bp.add_run()
                    br.text = str(ni + 1)
                    br.font.size = Pt(9)
                    br.font.bold = True
                    br.font.color.rgb = _hex_to_rgb(n_color)
                    br.font.name = "Consolas"

                node_id_to_pos[node.id] = {"shape": shape, "x": nx, "y": node_cy, "w": node_w, "h": node_h}

                # ── 节点间自动箭头（同阶段内，始终 STRAIGHT） ──
                if prev_shape is not None:
                    prev = node_id_to_pos[nodes[ni - 1].id]
                    _add_connector_arrow(
                        slide, arrow_color,
                        MSO_CONNECTOR_TYPE.STRAIGHT,
                        Inches(prev["x"] + prev["w"]), Inches(prev["y"] + prev["h"] / 2),
                        Inches(nx), Inches(node_cy + node_h / 2),
                    )

                prev_shape = shape

        # ── 显式 edges（跨阶段连接，支持 edge_style） ────
        connector_style = {
            "straight": MSO_CONNECTOR_TYPE.STRAIGHT,
            "elbow": MSO_CONNECTOR_TYPE.ELBOW,
            "curve": MSO_CONNECTOR_TYPE.CURVE,
        }
        for edge in diagram.edges:
            if edge.from_id not in node_id_to_pos or edge.to_id not in node_id_to_pos:
                continue
            f = node_id_to_pos[edge.from_id]
            t = node_id_to_pos[edge.to_id]
            fx = f["x"] + f["w"] / 2
            fy = f["y"] + f["h"]
            tx = t["x"] + t["w"] / 2
            ty = t["y"]
            ctype = connector_style.get(edge.edge_style, MSO_CONNECTOR_TYPE.STRAIGHT)
            _add_connector_arrow(slide, arrow_color, ctype,
                                 Inches(fx), Inches(fy), Inches(tx), Inches(ty))

            if edge.label:
                lx = (fx + tx) / 2 - 0.5
                ly = (fy + ty) / 2 - 0.15
                label_box = slide.shapes.add_textbox(Inches(lx), Inches(ly), Inches(1.0), Inches(0.3))
                ltf = label_box.text_frame
                p = ltf.paragraphs[0]
                p.text = edge.label
                p.alignment = PP_ALIGN.CENTER
                run = p.runs[0]
                run.font.size = Pt(9)
                run.font.italic = True
                run.font.color.rgb = _hex_to_rgb(arrow_color)

        atomic_save(prs.save, resolved_path)
        import os as _os
        size = _os.path.getsize(resolved_path) if _os.path.exists(resolved_path) else 0
        return CreateResult(
            success=True,
            file_path=resolved_path,
            file_size=size,
            slides_count=1,
            message=f"框架图生成成功: {len(diagram.stages)} 个阶段, {sum(len(s.nodes) for s in diagram.stages)} 个节点",
        )
    except Exception as e:
        return CreateResult(success=False, file_path=output_path, message=f"框架图生成失败: {e}")


def _set_mixed_font_text(paragraph, text: str, font_cn: str, font_en: str,
                         font_size: int, bold: bool = False,
                         color_rgb=None) -> None:
    """将文本按中英文拆分后写入段落，中文用 font_cn，西文用 font_en。

    使用正则 [⏎-￭] 块匹配中文字符，非中文段落使用西文字体。
    保留原始 run 用于首段，后续添加的 run 使用 add_run()。
    """
    import re
    from pptx.util import Pt

    # 清空段落默认文本
    paragraph.clear()

    # 按中文/非中文边界拆分
    # 中文字符 Unicode 范围：基本汉字 4E00-9FFF, 扩展A 3400-4DBF, 标点等
    pattern = re.compile(r'([\u4e00-\u9fff\u3400-\u4dbf\u3000-\u303f\uff00-\uffef]+)')
    segments = pattern.split(text)

    for seg in segments:
        if not seg:
            continue
        run = paragraph.add_run()
        run.text = seg
        run.font.size = Pt(font_size)
        run.font.bold = bold
        # 判断是否中文字段
        if pattern.match(seg):
            run.font.name = font_cn
        else:
            run.font.name = font_en
        if color_rgb is not None:
            run.font.color.rgb = color_rgb


def _add_connector_arrow(slide, arrow_color_hex: str, connector_type, x1, y1, x2, y2) -> None:
    """创建一条带箭头末端的连接线。"""
    from pptx.oxml.ns import qn
    from lxml import etree
    connector = slide.shapes.add_connector(connector_type, x1, y1, x2, y2)
    connector.line.color.rgb = _hex_to_rgb(arrow_color_hex)
    connector.line.width = Pt(1.5)
    # a:ln 位于 p:spPr 之下，直接 find(qn("a:ln")) 匹配不到，必须经 LineFormat 取
    tail_end = connector.line._get_or_add_ln()
    tail = tail_end.find(qn("a:tailEnd"))
    if tail is None:
        tail = etree.SubElement(tail_end, qn("a:tailEnd"))
    tail.set("type", "triangle")
    tail.set("w", "med")
    tail.set("len", "med")


def _lighten(hex_color: str, amount: int = 30) -> str:
    """将 HEX 颜色提亮指定量。"""
    h = hex_color.lstrip("#")
    r = min(255, int(h[0:2], 16) + amount)
    g = min(255, int(h[2:4], 16) + amount)
    b = min(255, int(h[4:6], 16) + amount)
    return f"{r:02X}{g:02X}{b:02X}"


def _mix_with_white(hex_color: str, weight: float = 0.9) -> str:
    """将 HEX 颜色向白色混合，weight=0.9 表示 90% 白 + 10% 原色（极淡同色系底）。"""
    h = hex_color.lstrip("#")
    r = int(h[0:2], 16)
    g = int(h[2:4], 16)
    b = int(h[4:6], 16)
    w = max(0.0, min(1.0, weight))
    return (
        f"{round(r + (255 - r) * w):02X}"
        f"{round(g + (255 - g) * w):02X}"
        f"{round(b + (255 - b) * w):02X}"
    )


def _mix_with_black(hex_color: str, weight: float = 0.25) -> str:
    """将 HEX 颜色向黑色混合，weight=0.25 生成同色系深边框色。"""
    h = hex_color.lstrip("#")
    r = int(h[0:2], 16)
    g = int(h[2:4], 16)
    b = int(h[4:6], 16)
    w = max(0.0, min(1.0, weight))
    return (
        f"{round(r * (1 - w)):02X}"
        f"{round(g * (1 - w)):02X}"
        f"{round(b * (1 - w)):02X}"
    )


def _add_shape_shadow(shape, blur_pt: float = 8.0, dist_pt: float = 3.0,
                      opacity_pct: int = 50) -> None:
    """给 python-pptx 形状加柔和外阴影（手写 OOXML a:effectLst/a:outerShdw）。

    python-pptx 无原生 shadow API（官方 scope 明确 out-of-scope），需在
    p:spPr 下手写 effectLst。参数对齐 PptxGenJS ShadowProps 的语义。
    """
    from lxml import etree
    from pptx.oxml.ns import qn

    sp_pr = shape._element.spPr
    # 移除可能已存在的 effectLst，避免重复阴影
    for old in sp_pr.findall(qn("a:effectLst")):
        sp_pr.remove(old)
    effect_lst = etree.SubElement(sp_pr, qn("a:effectLst"))
    outer = etree.SubElement(effect_lst, qn("a:outerShdw"))
    outer.set("blurRad", str(int(blur_pt * 12700)))
    outer.set("dist", str(int(dist_pt * 12700)))
    outer.set("dir", "5400000")  # 正下方向
    outer.set("rotWithShape", "0")
    srgb = etree.SubElement(outer, qn("a:srgbClr"))
    srgb.set("val", "000000")
    alpha = etree.SubElement(srgb, qn("a:alpha"))
    alpha.set("val", str(int(opacity_pct * 1000)))
