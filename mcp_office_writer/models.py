
"""
mcp-office-writer 数据模型定义
所有 Pydantic 模型集中在此，供各 writer 模块共享
"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field, model_validator


# ── 通用类型 ──────────────────────────────────────────────────────

class FontConfig(BaseModel):
    """字体配置"""
    name: str | None = Field(None, description="字体名称，如 '宋体'、'微软雅黑'、'Arial'")
    size: int | str | None = Field(None, description="字号（国内标准如'小四''五号'，或直接 pt 数值）")
    bold: bool | None = Field(None, description="是否加粗")
    italic: bool | None = Field(None, description="是否斜体")
    color: str | None = Field(None, description="字体颜色 (HEX, 如 'FF0000')")


class CellStyle(BaseModel):
    """单元格样式"""
    bold: bool | None = None
    font_name: str | None = None
    font_size: int | str | None = None
    font_color: str | None = Field(None, description="字体颜色 HEX")
    bg_color: str | None = Field(None, description="背景颜色 HEX")
    number_format: str | None = Field(None, description="数字格式，如 '#,##0.00'、'0%'、'YYYY-MM-DD'")
    alignment: Literal["left", "center", "right"] | None = None
    wrap_text: bool | None = None


class PageSetup(BaseModel):
    """页面设置"""
    orientation: Literal["portrait", "landscape"] | None = None
    paper_size: Literal["A4", "Letter", "A3", "B5"] | None = "A4"
    margin_top: float | None = Field(None, description="上边距 (英寸)")
    margin_bottom: float | None = Field(None, description="下边距 (英寸)")
    margin_left: float | None = Field(None, description="左边距 (英寸)")
    margin_right: float | None = Field(None, description="右边距 (英寸)")


class HeaderFooter(BaseModel):
    """页眉页脚设置"""
    header_text: str | None = Field(None, description="页眉文字")
    footer_text: str | None = Field(None, description="页脚文字，用 {page} 代表页码")


PageNumberStyle = Literal["none", "roman", "arabic"]


class SceneProfile(BaseModel):
    """场景排版配置文件"""
    name: str = Field(..., description="场景名称标识")
    display_name: str = Field("", description="场景显示名称")
    description: str | None = Field(None, description="场景描述")

    # 字体
    font_name: str = "宋体"
    font_name_western: str = Field("Times New Roman", description="西文字体")
    heading_font: str | None = None  # None = 使用 font_name
    body_size: str = "小四"           # 正文字号（国内标准）
    heading_sizes: dict[int, str] = Field(default_factory=lambda: {1: "三号", 2: "小三", 3: "四号", 4: "小四"})

    # 段落
    line_spacing: float = 1.3
    first_line_indent: int | None = 28  # pt, ≈ 2 字符
    alignment: str = "justify"

    # 页面
    paper_size: str = "A4"
    margin_top: float = 1.0
    margin_bottom: float = 1.0
    margin_left: float = 1.18   # 3cm
    margin_right: float = 1.18

    # 页码格式
    page_number_style: PageNumberStyle = "arabic"
    show_cover: bool = False


class CoverConfig(BaseModel):
    """封面配置"""
    title: str = Field(..., description="封面主标题")
    subtitle: str | None = Field(None, description="封面副标题")
    author: str | None = Field(None, description="作者/创建者")
    date: str | None = Field(None, description="日期")
    organization: str | None = Field(None, description="机构/单位名称")
    extra_lines: list[str] | None = Field(None, description="附加信息行")
    recipe_id: str | None = Field(None, description="封面配方 ID (R1-R7)，自动选择时可留空")


class MergeRange(BaseModel):
    """表格合并单元格范围 (Excel A1 风格)"""
    range: str = Field(..., description="合并范围，如 'A1:B2'、'C1:C3'")


# ── Word 相关模型 ─────────────────────────────────────────────────

class ParagraphStyle(BaseModel):
    """段落样式"""
    font_name: str | None = None
    font_size: int | str | None = Field(None, description="字号（国内标准如'小四''五号'，或直接 pt 数值）")
    bold: bool | None = None
    italic: bool | None = None
    color: str | None = None
    alignment: Literal["left", "center", "right", "justify"] | None = None
    line_spacing: float | None = Field(None, description="行距倍数，如 1.0、1.5、2.0")
    first_line_indent: int | None = Field(None, description="首行缩进 (pt)，中文段落常用 28pt≈2字符")
    space_before: int | None = Field(None, description="段前距 (pt)")
    space_after: int | None = Field(None, description="段后距 (pt)")


class TableDef(BaseModel):
    """Word 表格定义"""
    headers: list[str] = Field(default_factory=list, description="表头行")
    rows: list[list[str]] = Field(default_factory=list, description="数据行")
    column_widths: list[float] | None = Field(None, description="列宽 (英寸)")
    header_style: CellStyle | None = Field(None, description="表头样式")
    merged_cells: list[MergeRange] | None = Field(None, description="合并单元格")


class ImageDef(BaseModel):
    """图片定义"""
    path: str = Field(..., description="图片文件路径")
    width: float | None = Field(None, description="宽度 (英寸)")
    height: float | None = Field(None, description="高度 (英寸)")


class SectionItem(BaseModel):
    """Word 文档的一个内容块"""
    type: Literal["heading", "paragraph", "table", "image", "page_break", "bullet_list", "numbered_list", "code", "toc", "section_break"]
    text: str | None = Field(None, description="文字内容")
    level: int | None = Field(None, description="标题层级 (1-9)，仅 heading 类型使用")
    language: str | None = Field(None, description="代码语言标签，仅 code 类型使用")
    style: ParagraphStyle | None = Field(None, description="段落样式")
    items: list[str] | None = Field(None, description="列表项，仅 bullet_list/numbered_list 使用")
    table: TableDef | None = Field(None, description="表格定义，仅 table 类型使用")
    image: ImageDef | None = Field(None, description="图片定义，仅 image 类型使用")
    toc_levels: str | None = Field(None, description="TOC 包含的标题层级，如 '1-3'，仅 toc 类型使用")


class WordContent(BaseModel):
    """Word 文档内容"""
    sections: list[SectionItem] = Field(..., description="文档内容块列表")


# ── Excel 相关模型 ───────────────────────────────────────────────

class ColumnFormat(BaseModel):
    """Excel 列格式"""
    number_format: str | None = None
    font_name: str | None = None
    font_color: str | None = None
    bg_color: str | None = None
    alignment: Literal["left", "center", "right"] | None = None


class SheetDef(BaseModel):
    """Excel 工作表定义"""
    name: str = Field(..., description="工作表名称")
    headers: list[str] | None = Field(None, description="表头行")
    data: list[list] = Field(default_factory=list, description="数据行（支持 str/int/float/None）")
    allow_formulas: bool = Field(False, description="是否将以 = 开头的字符串写为 Excel 公式；默认按纯文本写入")
    column_widths: list[float] | None = Field(None, description="列宽 (字符数)")
    header_style: CellStyle | None = Field(None, description="表头样式")
    column_formats: list[ColumnFormat] | None = Field(None, description="按列格式")
    merged_cells: list[MergeRange] | None = Field(None, description="合并单元格")
    freeze: str | None = Field(None, description="冻结窗格，如 'A2' 冻结首行")
    auto_filter: bool = Field(False, description="是否启用自动筛选")
    tab_color: str | None = Field(None, description="工作表标签颜色 HEX")


class ExcelContent(BaseModel):
    """Excel 工作簿内容"""
    sheets: list[SheetDef] = Field(..., description="工作表列表（至少 1 个）")


# ── PPT 相关模型 ──────────────────────────────────────────────────

LayoutType = Literal[
    "title", "section", "content", "two_content", "comparison",
    "title_only", "blank", "content_caption", "picture_caption"
]


class TextStyle(BaseModel):
    """PPT 文字样式"""
    font_name: str | None = None
    font_size: int | None = None
    bold: bool | None = None
    italic: bool | None = None
    color: str | None = None
    alignment: Literal["left", "center", "right"] | None = None


class SlideTableDef(BaseModel):
    """PPT 幻灯片表格定义"""
    headers: list[str] = Field(default_factory=list)
    rows: list[list[str]] = Field(default_factory=list)


class SlideImageDef(BaseModel):
    """PPT 幻灯片图片定义"""
    path: str | None = Field(None, description="图片文件路径（与 url 二选一）")
    url: str | None = Field(None, description="图片 URL（与 path 二选一），自动下载后插入")
    x: float | None = Field(None, description="左侧位置 (英寸)")
    y: float | None = Field(None, description="顶部位置 (英寸)")
    width: float | None = Field(None, description="宽度 (英寸)")
    height: float | None = Field(None, description="高度 (英寸)")

    @model_validator(mode="after")
    def _check_path_or_url(self):
        if not self.path and not self.url:
            raise ValueError("SlideImageDef 必须提供 path（本地文件）或 url（在线图片）之一")
        if self.path and self.url:
            raise ValueError("SlideImageDef 不能同时提供 path 和 url，请二选一")
        return self


class SlideShapeDef(BaseModel):
    """PPT 形状定义"""
    shape_type: Literal["rect", "rounded_rect", "oval", "arrow", "cylinder", "diamond"] = "rect"
    text: str | None = Field(None, description="形状内文字")
    x: float | None = Field(None, description="左侧位置 (英寸)")
    y: float | None = Field(None, description="顶部位置 (英寸)")
    width: float = Field(..., description="宽度 (英寸)")
    height: float = Field(..., description="高度 (英寸)")
    bg_color: str | None = Field(None, description="填充颜色 HEX，None 且 no_fill=False 时保持默认")
    font_color: str | None = Field(None, description="文字颜色 HEX")
    font_size: int | None = Field(None, description="字号 (pt)，默认 12")
    bold: bool | None = False
    no_fill: bool = Field(False, description="无填充（透明），用于虚线分组框等")
    line_color: str | None = Field(None, description="边框颜色 HEX，None 不设置边框")
    line_width: float | None = Field(None, description="边框宽度 (pt)")
    line_dash: bool = Field(False, description="边框是否虚线")
    vertical_text: bool = Field(False, description="文字竖排（圆柱体阶段标记用）")


class SlideConnectorDef(BaseModel):
    """PPT 连接线/箭头定义（坐标单位：英寸，相对幻灯片左上角）"""
    x1: float = Field(..., description="起点 X (英寸)")
    y1: float = Field(..., description="起点 Y (英寸)")
    x2: float = Field(..., description="终点 X (英寸)")
    y2: float = Field(..., description="终点 Y (英寸)")
    color: str | None = Field(None, description="线条颜色 HEX，留空用默认")
    width: float = Field(1.5, description="线宽 (pt)")
    dash: bool = Field(False, description="是否虚线（损失反向传播等特殊箭头）")
    arrowhead: bool = Field(True, description="终点是否带箭头")
    arrowhead_style: Literal["triangle", "open", "stealth"] = Field("triangle", description="箭头样式")
    begin_arrowhead: bool = Field(False, description="起点是否带箭头")


class SlideItem(BaseModel):
    """幻灯片中的一个内容元素"""
    type: Literal["paragraph", "table", "image", "shape", "connector", "bullet_list", "numbered_list", "chart"]
    text: str | None = None
    level: int = Field(0, description="缩进层级 (0=顶级)")
    style: TextStyle | None = None
    table: SlideTableDef | None = None
    image: SlideImageDef | None = None
    shape: SlideShapeDef | None = None
    connector: SlideConnectorDef | None = Field(None, description="连接线定义，仅 connector 类型使用")
    items: list[str] | None = Field(None, description="列表项，仅 bullet_list/numbered_list 使用")
    chart: ChartDef | None = Field(None, description="图表定义，仅 chart 类型使用")


class ChartDef(BaseModel):
    """PPT 图表定义"""
    chart_type: Literal["column", "bar", "line", "pie"] = Field("column", description="图表类型")
    categories: list[str] = Field(..., description="类别标签（X 轴）")
    series: list[ChartSeries] = Field(..., description="数据系列")
    has_legend: bool = Field(True, description="是否显示图例")
    title: str | None = Field(None, description="图表标题")


class ChartSeries(BaseModel):
    """图表数据系列"""
    name: str = Field(..., description="系列名称")
    values: list[float] = Field(..., description="数值列表")


class SlideDef(BaseModel):
    """幻灯片定义"""
    layout: LayoutType = "content"
    title: str | None = Field(None, description="幻灯片标题")
    subtitle: str | None = Field(None, description="副标题（仅 title/section 布局）")
    body: list[SlideItem] | None = Field(None, description="主体内容")
    left: list[SlideItem] | None = Field(None, description="左侧内容（仅 two_content/comparison 布局）")
    right: list[SlideItem] | None = Field(None, description="右侧内容（仅 two_content/comparison 布局）")
    notes: str | None = Field(None, description="演讲者备注")
    bg_color: str | None = Field(None, description="幻灯片背景色 HEX")
    slide_width: float | None = Field(None, description="幻灯片宽度 (英寸)，默认 10")
    slide_height: float | None = Field(None, description="幻灯片高度 (英寸)，默认 7.5")


class PptContent(BaseModel):
    """PPT 演示文稿内容"""
    slides: list[SlideDef] = Field(..., description="幻灯片列表（至少 1 个）")


# ── 框架图 ────────────────────────────────────────────────────────

class FwNode(BaseModel):
    """框架图节点"""
    id: str = Field(..., description="节点唯一标识")
    text: str = Field(..., description="节点显示文本，支持 \\n 换行")
    style: Literal[
        "rect", "rounded_rect", "ellipse", "diamond",
        "process", "decision", "data", "document", "terminator", "predefined",
        "cylinder", "parallelogram", "chevron", "right_arrow", "left_arrow",
        "callout", "cloud",
    ] = Field("rounded_rect", description="节点形状")
    fill_color: str | None = Field(None, description="填充色 HEX，留空使用阶段色")
    font_size: int = Field(12, description="字号")
    bold: bool = Field(True, description="是否加粗")


class FwEdge(BaseModel):
    """框架图连接线"""
    from_id: str = Field(..., description="起始节点 ID")
    to_id: str = Field(..., description="目标节点 ID")
    label: str | None = Field(None, description="箭头上的标签文字（可选）")
    edge_style: Literal["straight", "elbow", "curve"] = Field("straight", description="连接线样式：straight(直线)/elbow(正交)/curve(曲线)")


class FwStage(BaseModel):
    """框架图阶段分组"""
    label: str = Field(..., description="阶段标签，显示在左侧圆柱形上")
    color: str | None = Field(None, description="阶段背景色 HEX，留空使用配色方案默认")
    nodes: list[FwNode] = Field(default_factory=list, description="阶段内的节点列表")


class FrameworkDiagram(BaseModel):
    """框架图完整定义"""
    title: str = Field("", description="框架图总标题")
    subtitle: str | None = Field(None, description="框架图副标题")
    stages: list[FwStage] = Field(..., description="阶段列表，按从上到下顺序排列")
    edges: list[FwEdge] = Field(default_factory=list, description="节点间连接线，未指定的相邻节点自动生成箭头")
    color_scheme: str = Field("professional-blue", description="配色方案名称")
    node_width: float = Field(1.4, description="默认节点宽度 (英寸)")
    node_height: float = Field(0.55, description="默认节点高度 (英寸)")
    stage_gap: float = Field(0.15, description="阶段间垂直间距 (英寸)")
    arrow_color: str | None = Field(None, description="箭头颜色 HEX")
    font_cn: str = Field("宋体", description="中文字体")
    font_en: str = Field("Times New Roman", description="西文字体")
    font_size: int = Field(16, description="正文默认字号 (pt)，放不下时自动缩小")
    title_font_size: int = Field(24, description="标题字号 (pt)")


# ── 配色方案 ──────────────────────────────────────────────────────

class ColorScheme(BaseModel):
    """PPT 配色方案"""
    name: str = Field(..., description="方案名称标识")
    display_name: str = Field("", description="方案显示名称")
    description: str | None = Field(None, description="方案描述")
    title_color: str = Field(..., description="标题颜色 HEX")
    body_color: str = Field(..., description="正文颜色 HEX")
    accent_color: str = Field(..., description="强调色 HEX")
    background_color: str = Field(..., description="幻灯片背景色 HEX")
    shape_fill_color: str = Field(..., description="形状填充色 HEX")
    table_header_color: str = Field(..., description="表格表头背景色 HEX")
    title_font: str = Field("微软雅黑", description="标题字体")
    body_font: str = Field("微软雅黑", description="正文字体")
    title_size: int = Field(32, description="标题字号 (pt)")
    body_size: int = Field(18, description="正文字号 (pt)")


# ── 模板分析 ──────────────────────────────────────────────────────

class TemplateAnalysis(BaseModel):
    """PPT 模板分析结果"""
    file_path: str = Field(..., description="模板文件路径")
    slide_count: int = Field(0, description="幻灯片数量")
    slide_width: float = Field(0, description="幻灯片宽度 (英寸)")
    slide_height: float = Field(0, description="幻灯片高度 (英寸)")
    layout_names: list[str] = Field(default_factory=list, description="可用布局名称列表")
    color_schemes: list[ColorScheme] | None = Field(None, description="提取的配色方案")
    fonts_used: list[str] = Field(default_factory=list, description="模板中使用的字体")


# ── 返回结果 ──────────────────────────────────────────────────────

class CreateResult(BaseModel):
    """创建操作的返回结果"""
    success: bool
    file_path: str
    file_size: int = 0
    sheets_count: int | None = None
    slides_count: int | None = None
    message: str


class EditResult(BaseModel):
    """编辑操作的返回结果"""
    success: bool
    file_path: str
    operation_count: int = 0
    message: str


# ── PDF 相关模型 ─────────────────────────────────────────────────

class PdfTableDef(BaseModel):
    """PDF 表格定义"""
    headers: list[str] = Field(default_factory=list, description="表头行")
    rows: list[list[str]] = Field(default_factory=list, description="数据行")
    column_widths: list[float] | None = Field(None, description="列宽比例 (0~1)，如 [0.3, 0.4, 0.3]")


class PdfSectionItem(BaseModel):
    """PDF 文档的一个内容块"""
    type: Literal["heading", "paragraph", "table", "image", "page_break", "bullet_list", "numbered_list"]
    text: str | None = Field(None, description="文字内容")
    level: int | None = Field(None, description="标题层级 (1-6)，仅 heading 类型使用")
    style: ParagraphStyle | None = Field(None, description="段落样式")
    items: list[str] | None = Field(None, description="列表项，仅 bullet_list/numbered_list 使用")
    table: PdfTableDef | None = Field(None, description="表格定义，仅 table 类型使用")
    image: ImageDef | None = Field(None, description="图片定义，仅 image 类型使用")


class PdfContent(BaseModel):
    """PDF 文档内容"""
    sections: list[PdfSectionItem] = Field(..., description="文档内容块列表")


class PdfPageSetup(BaseModel):
    """PDF 页面设置"""
    paper_size: Literal["A4", "Letter", "A3", "B5"] = "A4"
    orientation: Literal["portrait", "landscape"] = "portrait"
    margin_top: float = Field(10, description="上边距 (mm)")
    margin_bottom: float = Field(10, description="下边距 (mm)")
    margin_left: float = Field(10, description="左边距 (mm)")
    margin_right: float = Field(10, description="右边距 (mm)")


class PdfFontConfig(BaseModel):
    """PDF 字体配置"""
    cn_font: str = Field("", description="中文字体名称，如 'simhei'、'simsun'，留空自动检测")
    cn_font_path: str | None = Field(None, description="自定义中文字体文件路径 (.ttf/.ttc)")
    en_font: str = Field("Helvetica", description="西文字体: Helvetica/Courier/Times")
    cn_font_size: int | str = Field("小四", description="正文字号")
    heading_font_size: int | str = Field("三号", description="标题字号")


class OfficeOperation(BaseModel):
    """A path-addressed Office document mutation used by batch editing."""

    op: Literal["set", "add", "remove", "move", "replace"]
    path: str = Field(..., description="目标路径，如 /body/p[1]、/Sheet1/A1 或 /slide[1]/shape[2]")
    element_type: str | None = Field(None, description="add 操作的元素类型，如 paragraph、sheet、slide、textbox")
    props: dict[str, Any] = Field(default_factory=dict, description="操作属性")
    target_path: str | None = Field(None, description="move 操作的目标路径")
    position: Literal["before", "after"] = Field("after", description="move 相对目标的位置")
