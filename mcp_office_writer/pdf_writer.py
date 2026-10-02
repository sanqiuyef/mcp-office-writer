"""
mcp-office-writer - PDF (.pdf) 创建模块
基于 fpdf2 库实现 PDF 文档生成，支持中文。
"""
from __future__ import annotations

import logging
import os
from pathlib import Path
from typing import TYPE_CHECKING

from fpdf import FPDF
from fpdf.enums import Align

from .models import CreateResult

if TYPE_CHECKING:
    from .models import (
        PdfContent, PdfSectionItem, PdfPageSetup,
        PdfFontConfig, ParagraphStyle,
    )

logger = logging.getLogger(__name__)

# ── 页面尺寸 (mm) ──
PAPER_SIZE_MM: dict[str, tuple[float, float]] = {
    "A4": (210, 297),
    "Letter": (215.9, 279.4),
    "A3": (297, 420),
    "B5": (176, 250),
}

# ── 中文字体探测 ──
_CANDIDATE_FONT_PATHS: dict[str, str] = {
    "simhei": r"C:\Windows\Fonts\simhei.ttf",
    "simsun": r"C:\Windows\Fonts\simsun.ttc",
    "msyh": r"C:\Windows\Fonts\msyh.ttc",
    "msyhbd": r"C:\Windows\Fonts\msyhbd.ttc",
    "simkai": r"C:\Windows\Fonts\simkai.ttf",
    "simfang": r"C:\Windows\Fonts\simfang.ttf",
    "fangsong": r"C:\Windows\Fonts\simfang.ttf",
    "kaiti": r"C:\Windows\Fonts\simkai.ttf",
}
_CANDIDATE_ORDER = ["simsun", "simhei", "msyh", "simkai", "simfang"]


def _find_chinese_font() -> tuple[str, str] | None:
    """自动探测可用的中文字体，返回 (字体族名, 路径)。"""
    for name in _CANDIDATE_ORDER:
        path = _CANDIDATE_FONT_PATHS.get(name, "")
        if path and os.path.isfile(path):
            return name, path
    # 遍历 Windows Fonts 目录搜索任意 .ttf/.ttc
    font_dir = r"C:\Windows\Fonts"
    if os.path.isdir(font_dir):
        for fn in sorted(os.listdir(font_dir)):
            if fn.lower().endswith((".ttf", ".ttc")):
                return fn.rsplit(".", 1)[0].lower(), os.path.join(font_dir, fn)
    return None


# ── fpdf2 对齐映射 ──
FPDF_ALIGN: dict[str, Align] = {
    "left": Align.L,
    "center": Align.C,
    "right": Align.R,
    "justify": Align.J,
}


class PdfWriter:
    """PDF 文档生成器"""

    def __init__(
        self,
        page_setup: PdfPageSetup | None = None,
        font_config: PdfFontConfig | None = None,
    ):
        from .models import PdfPageSetup, PdfFontConfig
        from .utils import cn_size_to_pt

        self.page_setup = page_setup or PdfPageSetup()
        self.font_config = font_config or PdfFontConfig()

        # 计算页面尺寸
        pw, ph = PAPER_SIZE_MM.get(
            self.page_setup.paper_size, PAPER_SIZE_MM["A4"]
        )
        if self.page_setup.orientation == "landscape":
            pw, ph = ph, pw

        self.pdf = FPDF(orientation="P", unit="mm", format=(pw, ph))
        self.pdf.set_auto_page_break(auto=True, margin=self.page_setup.margin_bottom)

        # 设置边距
        self.pdf.set_margin(self.page_setup.margin_left)

        # ── 注册中文字体 ──
        cn_font_name = self.font_config.cn_font
        cn_font_path = self.font_config.cn_font_path

        if cn_font_path and os.path.isfile(cn_font_path):
            family = cn_font_name or Path(cn_font_path).stem.lower()
            self.pdf.add_font(family, "", cn_font_path)
            self.pdf.add_font(family, "B", cn_font_path)
            self.cn_font = family
            logger.info("已注册自定义中文字体: %s (%s)", family, cn_font_path)
        elif cn_font_name:
            path = _CANDIDATE_FONT_PATHS.get(cn_font_name.lower(), cn_font_name)
            if os.path.isfile(path):
                self.pdf.add_font(cn_font_name, "", path)
                self.pdf.add_font(cn_font_name, "B", path)
                self.cn_font = cn_font_name
                logger.info("已注册中文字体: %s", cn_font_name)
            else:
                raise FileNotFoundError(
                    f"指定的中文字体文件不存在: {path}。"
                    f"可用候选: {', '.join(_CANDIDATE_FONT_PATHS)}"
                )
        else:
            found = _find_chinese_font()
            if found:
                name, path = found
                self.pdf.add_font(name, "", path)
                self.pdf.add_font(name, "B", path)
                self.cn_font = name
                logger.info("自动检测到中文字体: %s (%s)", name, path)
            else:
                raise FileNotFoundError(
                    "未找到可用的中文字体文件。请通过 font_config.cn_font_path 指定字体路径，"
                    "或在 C:\\Windows\\Fonts 安装中文字体（如 黑体 simhei.ttf）。"
                )

        # 字号
        self.body_size = cn_size_to_pt(self.font_config.cn_font_size)
        self.heading_size = cn_size_to_pt(self.font_config.heading_font_size)
        self.en_font = self.font_config.en_font

    # ── 字体设置辅助 ──

    def _set_font(
        self,
        bold: bool = False,
        size: float | None = None,
        style: ParagraphStyle | None = None,
    ):
        """根据样式设置当前字体。"""
        from .utils import cn_size_to_pt

        use_bold = bold
        use_size = size or self.body_size

        if style:
            if style.bold is not None:
                use_bold = style.bold
            if style.font_size is not None:
                use_size = cn_size_to_pt(style.font_size)

        style_str = "B" if use_bold else ""
        try:
            self.pdf.set_font(self.cn_font, style_str, use_size)
        except Exception:
            self.pdf.set_font(self.cn_font, "", use_size)

    def _parse_color(self, hex_color: str | None) -> tuple[int, int, int] | None:
        """HEX 颜色转为 RGB 三元组。"""
        if not hex_color:
            return None
        from .utils import hex_to_rgb
        return hex_to_rgb(hex_color)

    # ── 内容渲染 ──

    def _render_heading(self, item: PdfSectionItem):
        """渲染标题。"""
        level = max(1, min(item.level or 1, 6))
        text = item.text or ""

        # 标题字号：1 级最大，在 heading_size 与 body_size 之间线性递减，保证单调
        span = max(0.0, self.heading_size - self.body_size)
        step = span / 5
        sizes = {k: round(self.body_size + step * (6 - k), 1) for k in range(1, 7)}
        h_size = sizes.get(level, self.body_size)

        self._set_font(bold=True, size=h_size, style=item.style)
        self.pdf.ln(4)

        if item.style and item.style.alignment:
            align = FPDF_ALIGN.get(item.style.alignment, Align.L)
            self.pdf.cell(
                0, h_size * 0.5, text, new_x="LMARGIN", new_y="NEXT",
                align=align,
            )
        else:
            self.pdf.cell(
                0, h_size * 0.5, text, new_x="LMARGIN", new_y="NEXT",
                align=Align.L,
            )
        self.pdf.ln(2)

    def _render_paragraph(self, item: PdfSectionItem):
        """渲染段落。"""
        text = item.text or ""
        if not text.strip():
            self.pdf.ln(4)
            return

        self._set_font(style=item.style)
        align = Align.J
        if item.style and item.style.alignment:
            align = FPDF_ALIGN.get(item.style.alignment, Align.J)

        # 首行缩进
        indent = 0.0
        if item.style and item.style.first_line_indent:
            # first_line_indent 单位为 pt，转为 mm (1pt ≈ 0.3528 mm)
            indent = item.style.first_line_indent * 0.3528

        self.pdf.set_left_margin(self.page_setup.margin_left + indent)
        self.pdf.multi_cell(
            0, self.body_size * 0.6, text, align=align,
        )
        self.pdf.set_left_margin(self.page_setup.margin_left)

    def _render_table(self, item: PdfSectionItem):
        """渲染表格。"""
        if not item.table:
            return
        table = item.table
        headers = table.headers or []
        rows = table.rows or []

        if not headers and not rows:
            return

        self.pdf.ln(2)

        # 如果只有行没有表头，自动用行数填充
        if not headers and rows:
            headers = [f"Col{i + 1}" for i in range(len(rows[0]))]

        # 计算列宽
        page_w = self.pdf.w - self.page_setup.margin_left - self.page_setup.margin_right
        col_count = len(headers) if headers else (len(rows[0]) if rows else 1)
        if table.column_widths and len(table.column_widths) == col_count:
            col_widths = [page_w * w for w in table.column_widths]
        else:
            col_widths = [page_w / col_count] * col_count

        # 表头样式
        self._set_font(bold=True, size=self.body_size - 1)
        header_align = Align.C

        # 绘制表头背景和文字
        header_h = self.body_size * 0.7
        self.pdf.set_fill_color(230, 230, 230)
        for i, header in enumerate(headers):
            self.pdf.cell(
                col_widths[i], header_h, str(header),
                border=1, fill=True, align=header_align,
            )
        self.pdf.ln()

        # 数据行
        self._set_font(size=self.body_size - 1)
        row_h = self.body_size * 0.6
        for row in rows:
            for i, cell in enumerate(row[:col_count]):
                if i == 0:
                    # 第一列左对齐
                    self.pdf.cell(
                        col_widths[i], row_h, str(cell) if cell else "",
                        border=1, align=Align.L,
                    )
                else:
                    self.pdf.cell(
                        col_widths[i], row_h, str(cell) if cell else "",
                        border=1, align=Align.C,
                    )
            self.pdf.ln()

        self.pdf.ln(2)

    def _render_image(self, item: PdfSectionItem):
        """渲染图片。"""
        if not item.image or not item.image.path:
            return
        img_path = item.image.path
        if not os.path.isfile(img_path):
            logger.warning("图片文件不存在，跳过: %s", img_path)
            self.pdf.ln(4)
            self._set_font()
            self.pdf.cell(
                0, self.body_size * 0.5,
                f"_[图片未找到: {img_path}]_",
                new_x="LMARGIN", new_y="NEXT",
            )
            return

        self.pdf.ln(2)
        try:
            page_w = self.pdf.w - self.page_setup.margin_left - self.page_setup.margin_right
            img_w = item.image.width or page_w
            img_h = item.image.height or 0

            # 宽度换算：英寸 → mm (1 inch = 25.4 mm)
            w_mm = img_w * 25.4
            h_mm = img_h * 25.4 if img_h else 0

            if h_mm == 0:
                self.pdf.image(img_path, w=w_mm, x=self.page_setup.margin_left)
            else:
                self.pdf.image(img_path, w=w_mm, h=h_mm, x=self.page_setup.margin_left)
        except Exception as e:
            logger.warning("插入图片失败: %s, %s", img_path, e)
            self._set_font()
            self.pdf.cell(
                0, self.body_size * 0.5,
                f"_[图片加载失败: {img_path}]_",
                new_x="LMARGIN", new_y="NEXT",
            )
        self.pdf.ln(2)

    def _render_list(self, item: PdfSectionItem, ordered: bool = False):
        """渲染列表。"""
        items = item.items or []
        if not items:
            return

        self._set_font(style=item.style)
        self.pdf.ln(2)

        indent = 5.0  # 列表缩进 (mm)
        page_w = self.pdf.w - self.pdf.l_margin - self.pdf.r_margin

        for i, li_text in enumerate(items):
            if ordered:
                prefix = f"{i + 1}. "
            else:
                prefix = "- "  # CJK-safe bullet (• glyph not in SimHei)

            self.pdf.set_x(self.pdf.l_margin + indent)
            full_text = prefix + li_text
            self.pdf.multi_cell(
                page_w - indent, self.body_size * 0.55, full_text, align=Align.L,
            )
        self.pdf.ln(2)

    def _render_page_break(self):
        """分页。"""
        self.pdf.add_page()

    # ── 主入口 ──

    def write(self, content: PdfContent):
        """将 PdfContent 渲染为 PDF。"""
        self._set_font()
        self.pdf.add_page()

        for item in content.sections:
            if item.type == "heading":
                self._render_heading(item)
            elif item.type == "paragraph":
                self._render_paragraph(item)
            elif item.type == "table":
                self._render_table(item)
            elif item.type == "image":
                self._render_image(item)
            elif item.type == "page_break":
                self._render_page_break()
            elif item.type == "bullet_list":
                self._render_list(item, ordered=False)
            elif item.type == "numbered_list":
                self._render_list(item, ordered=True)

    def save(self, output_path: str):
        """保存 PDF 到文件。"""
        dir_name = os.path.dirname(output_path)
        if dir_name:
            os.makedirs(dir_name, exist_ok=True)
        self.pdf.output(output_path)


def create_pdf_document(
    output_path: str,
    content: PdfContent,
    page_setup: PdfPageSetup | None = None,
    font_config: PdfFontConfig | None = None,
    if_exists: str = "overwrite",
) -> CreateResult:
    """
    创建 PDF (.pdf) 文档。

    支持段落、标题、表格、图片、分页、项目符号列表、编号列表。
    自动检测并使用 Windows 系统中的中文字体（黑体/宋体/微软雅黑等）。
    支持自定义页面尺寸、边距、字体和字号。

    Args:
        output_path: .pdf 输出文件路径
        content: PdfContent，包含 sections 列表
        page_setup: 可选页面设置（纸张大小、方向、边距）
        font_config: 可选字体配置（中英文字体、字号）
        if_exists: "overwrite" | "error" | "rename"

    Returns:
        CreateResult: {success, file_path, file_size, message}
    """
    from .utils import (
        validate_output_path, resolve_if_exists, get_file_size,
    )

    try:
        validate_output_path(output_path, (".pdf",))
        output_path = resolve_if_exists(output_path, if_exists)
    except (ValueError, FileExistsError) as e:
        return CreateResult(
            success=False, file_path=output_path, file_size=0,
            message=str(e),
        )

    if not content.sections:
        return CreateResult(
            success=False, file_path=output_path, file_size=0,
            message="content.sections 不能为空",
        )

    try:
        writer = PdfWriter(page_setup=page_setup, font_config=font_config)
        writer.write(content)
        writer.save(output_path)
    except Exception as e:
        logger.error("创建 PDF 失败: %s", e, exc_info=True)
        return CreateResult(
            success=False, file_path=output_path, file_size=0,
            message=f"创建 PDF 失败: {e}",
        )

    file_size = get_file_size(output_path)
    sections_count = len(content.sections)
    return CreateResult(
        success=True,
        file_path=output_path,
        file_size=file_size,
        message=f"成功创建 PDF 文档，共 {sections_count} 个内容块",
    )
