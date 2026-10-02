"""
mcp-office-writer 封面配方系统
基于 document-skills 的封面设计规范，提供已验证的封面模板
"""
from __future__ import annotations

import logging
from dataclasses import dataclass
from typing import Literal

from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

logger = logging.getLogger(__name__)

# ── 常量 ─────────────────────────────────────────────────────────

COVER_TWIPS = 16838          # A4 高度 twips（保留，供外部参考）
COVER_SAFE_TWIPS = 15638     # 安全内容高度 (减去 1200 twips 安全边距)

# ── 封面配方定义 ──────────────────────────────────────────────────

CoverLayout = Literal["centered", "split", "modern", "minimal", "academic", "official", "brand"]


@dataclass
class CoverRecipe:
    """封面配方"""
    recipe_id: str
    display_name: str
    layout: CoverLayout
    title_font: str
    subtitle_font: str
    meta_font: str
    title_size_min: int = 22   # pt
    title_size_max: int = 40
    subtitle_size: int = 16
    meta_size: int = 12
    use_horizontal_line: bool = False
    title_color: str = "000000"
    subtitle_color: str = "333333"
    bg_color: str | None = None  # None = white


# ── 7 个预验证封面配方 ─────────────────────────────────────────────

RECIPES: dict[str, CoverRecipe] = {
    "R1": CoverRecipe(
        recipe_id="R1",
        display_name="经典居中",
        layout="centered",
        title_font="黑体",
        subtitle_font="宋体",
        meta_font="宋体",
        title_size_min=22,
        title_size_max=40,
        subtitle_size=16,
        meta_size=12,
        use_horizontal_line=True,
    ),
    "R2": CoverRecipe(
        recipe_id="R2",
        display_name="现代简约",
        layout="modern",
        title_font="微软雅黑",
        subtitle_font="微软雅黑",
        meta_font="宋体",
        title_size_min=24,
        title_size_max=38,
        subtitle_size=15,
        meta_size=11,
        use_horizontal_line=False,
    ),
    "R3": CoverRecipe(
        recipe_id="R3",
        display_name="上下分割",
        layout="split",
        title_font="黑体",
        subtitle_font="宋体",
        meta_font="宋体",
        title_size_min=22,
        title_size_max=36,
        subtitle_size=15,
        meta_size=11,
        bg_color="F5F5F5",
    ),
    "R4": CoverRecipe(
        recipe_id="R4",
        display_name="学术论文",
        layout="academic",
        title_font="黑体",
        subtitle_font="宋体",
        meta_font="宋体",
        title_size_min=26,
        title_size_max=42,
        subtitle_size=16,
        meta_size=12,
        use_horizontal_line=True,
    ),
    "R5": CoverRecipe(
        recipe_id="R5",
        display_name="公文红头",
        layout="official",
        title_font="宋体",
        subtitle_font="仿宋_GB2312",
        meta_font="仿宋_GB2312",
        title_size_min=22,
        title_size_max=30,
        subtitle_size=16,
        meta_size=14,
        title_color="CC0000",
    ),
    "R6": CoverRecipe(
        recipe_id="R6",
        display_name="极简",
        layout="minimal",
        title_font="微软雅黑",
        subtitle_font="微软雅黑",
        meta_font="微软雅黑",
        title_size_min=20,
        title_size_max=36,
        subtitle_size=14,
        meta_size=10,
        use_horizontal_line=False,
    ),
    "R7": CoverRecipe(
        recipe_id="R7",
        display_name="品牌展示",
        layout="brand",
        title_font="黑体",
        subtitle_font="微软雅黑",
        meta_font="宋体",
        title_size_min=24,
        title_size_max=40,
        subtitle_size=16,
        meta_size=11,
        use_horizontal_line=True,
        bg_color="1A1A2E",
        title_color="FFFFFF",
        subtitle_color="CCCCCC",
    ),
}


def select_cover_recipe(doc_type: str = "general", scene: str | None = None) -> CoverRecipe:
    """根据文档类型和场景选择最合适的封面配方"""
    if scene == "academic":
        return RECIPES["R4"]
    elif scene == "official":
        return RECIPES["R5"]
    elif scene == "report":
        return RECIPES["R1"]
    elif scene == "contract":
        return RECIPES["R1"]
    elif scene == "resume":
        return RECIPES["R6"]
    elif scene == "exam":
        return RECIPES["R6"]
    return RECIPES["R1"]


def _calc_title_size(title: str, max_width_chars: int = 18) -> int:
    """根据标题长度自动计算字体大小"""
    length = len(title)
    if length <= 6:
        return 40
    elif length <= 10:
        return 34
    elif length <= 14:
        return 28
    elif length <= 20:
        return 24
    else:
        return 22


def _create_cover_paragraph(target, text: str, font_name: str, font_size: int,
                            color: str, alignment: str, space_before: int = 0,
                            space_after: int = 0, bold: bool = False):
    """创建封面段落（target 为 Document 或容器表格单元格）"""
    p = target.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if alignment == "center" else WD_ALIGN_PARAGRAPH.LEFT
    pf = p.paragraph_format
    if space_before:
        pf.space_before = Pt(space_before)
    if space_after:
        pf.space_after = Pt(space_after)
    pf.line_spacing = 1.5

    run = p.add_run(text)
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
    run.font.size = Pt(font_size)
    run.font.color.rgb = RGBColor(
        int(color[0:2], 16),
        int(color[2:4], 16),
        int(color[4:6], 16),
    )
    if bold:
        run.bold = True
    return p


def _add_horizontal_line(target):
    """在封面中添加水平分割线（段落底部边框实现）"""
    p = target.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format
    pf.space_before = Pt(6)
    pf.space_after = Pt(6)
    # 使用段落底部边框作为水平线
    pPr = p._element.get_or_add_pPr()
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'  <w:bottom w:val="single" w:sz="12" w:space="1" w:color="000000"/>'
        f'</w:pBdr>'
    )
    pPr.append(pBdr)


def _make_cover_background_container(doc, cover_section, bg_color: str):
    """创建整页背景容器：单格无边框表格，行高占满封面可用高度。

    OOXML 的段落/表格是顺序块元素，无法层叠在文字之下充当背景；
    因此带背景色的封面配方把所有封面内容放进一个行高=可用页高的
    单格表格中，单元格底色即整页背景色。
    """
    container = doc.add_table(rows=1, cols=1)
    container.autofit = False
    # 去掉表格边框
    tbl_pr = container._tbl.tblPr
    tbl_pr.append(parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        + "".join(
            f'<w:{side} w:val="none" w:sz="0" w:space="0"/>'
            for side in ("top", "left", "bottom", "right", "insideH", "insideV")
        )
        + '</w:tblBorders>'
    ))
    # 单元格底色
    cell = container.cell(0, 0)
    cell._element.get_or_add_tcPr().append(
        parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg_color}" w:val="clear"/>')
    )
    # 行高 = 页面可用高度（twips）
    if cover_section is not None:
        usable_inches = (
            cover_section.page_height.inches
            - cover_section.top_margin.inches
            - cover_section.bottom_margin.inches
        )
        tr_pr = container.rows[0]._tr.get_or_add_trPr()
        tr_pr.append(parse_xml(
            f'<w:trHeight {nsdecls("w")} w:val="{int(max(1.0, usable_inches) * 1440)}" '
            f'w:hRule="atLeast"/>'
        ))
    return cell


def generate_cover(doc: Document, recipe: CoverRecipe,
                   title: str = "",
                   subtitle: str | None = None,
                   author: str | None = None,
                   date: str | None = None,
                   organization: str | None = None,
                   extra_lines: list[str] | None = None):
    """在文档中生成封面

    封面内容写入文档最前部（与正文同节，之后由调用方分页）。
    带背景色的配方（R3/R7）使用整页容器表格承载封面内容。
    """
    cover_section = doc.sections[0] if doc.sections else None

    # 带背景色时：封面内容全部放入整页背景容器，实现整页色块
    host = doc
    if recipe.bg_color:
        host = _make_cover_background_container(doc, cover_section, recipe.bg_color)

    # 计算标题大小
    title_size = _calc_title_size(title)

    # 顶部空白（垂直居中辅助）
    _create_cover_paragraph(host, "", recipe.title_font, 10, "000000", "center", space_before=0)

    # ── 标题 ──
    _create_cover_paragraph(host, title, recipe.title_font, title_size,
                            recipe.title_color, "center", space_before=60, space_after=12, bold=True)

    # 水平分割线
    if recipe.use_horizontal_line:
        _add_horizontal_line(host)

    # ── 副标题 ──
    if subtitle:
        _create_cover_paragraph(host, subtitle, recipe.subtitle_font, recipe.subtitle_size,
                                recipe.subtitle_color, "center", space_before=20, space_after=8)

    # ── 分隔间隙 ──
    _create_cover_paragraph(host, "", recipe.title_font, 10, "000000", "center", space_before=0)

    # ── 附加信息（底部对齐） ──
    meta_lines = []
    if author:
        meta_lines.append(author)
    if date:
        meta_lines.append(date)
    if organization:
        meta_lines.append(organization)
    if extra_lines:
        meta_lines.extend(extra_lines)

    before_space = 40
    for line in meta_lines:
        _create_cover_paragraph(host, line, recipe.meta_font, recipe.meta_size,
                                recipe.subtitle_color, "center", space_before=before_space, space_after=2)
        before_space = 0

    # 封面后的分页符 - 确保封面与正文分离
    doc.add_page_break()

    logger.info("已生成封面（配方: %s, 标题: %s）", recipe.recipe_id, title)
