"""
mcp-office-writer 场景模板系统
预定义的文档排版场景配置文件
所有字号均采用国内标准（二号/小四/五号 等）
"""
from __future__ import annotations

from .models import SceneProfile


# ── 场景定义 ──────────────────────────────────────────────────────

ACADEMIC = SceneProfile(
    name="academic",
    display_name="学术论文",
    description="学术论文/学位论文排版，宋体正文+黑体标题，1.3倍行距，两字符缩进",
    font_name="宋体",
    font_name_western="Times New Roman",
    heading_font="黑体",
    body_size="小四",              # 12pt
    heading_sizes={1: "三号", 2: "四号", 3: "小四", 4: "五号"},
    line_spacing=1.3,
    first_line_indent=28,
    alignment="justify",
    paper_size="A4",
    margin_top=1.0,
    margin_bottom=1.0,
    margin_left=1.18,    # 3cm
    margin_right=1.18,
    page_number_style="arabic",
    show_cover=False,
)

REPORT = SceneProfile(
    name="report",
    display_name="报告/提案",
    description="正式报告排版，宋体正文+黑体标题，含封面和执行摘要结构",
    font_name="宋体",
    font_name_western="Times New Roman",
    heading_font="黑体",
    body_size="小四",              # 12pt
    heading_sizes={1: "三号", 2: "四号", 3: "小四", 4: "五号"},
    line_spacing=1.3,
    first_line_indent=28,
    alignment="justify",
    paper_size="A4",
    margin_top=1.0,
    margin_bottom=1.0,
    margin_left=1.18,
    margin_right=1.18,
    page_number_style="arabic",
    show_cover=True,
)

OFFICIAL = SceneProfile(
    name="official",
    display_name="公文",
    description="GB/T 9704 公文排版，仿宋_GB2312 正文，黑体标题，28pt 固定行距",
    font_name="仿宋_GB2312",
    font_name_western="Times New Roman",
    heading_font="黑体",
    body_size="三号",              # 16pt
    heading_sizes={1: "二号", 2: "三号", 3: "小三"},
    line_spacing=1.75,
    first_line_indent=40,
    alignment="justify",
    paper_size="A4",
    margin_top=1.46,        # 37mm
    margin_bottom=1.38,     # 35mm
    margin_left=1.1,        # 28mm
    margin_right=1.0,       # 26mm
    page_number_style="arabic",
    show_cover=False,
)

CONTRACT = SceneProfile(
    name="contract",
    display_name="合同/协议",
    description="合同排版，宋体正文，1.5倍行距，含签约信息和签名区",
    font_name="宋体",
    font_name_western="Times New Roman",
    heading_font="黑体",
    body_size="小四",              # 12pt
    heading_sizes={1: "三号", 2: "四号", 3: "小四"},
    line_spacing=1.5,
    first_line_indent=28,
    alignment="justify",
    paper_size="A4",
    margin_top=1.0,
    margin_bottom=1.0,
    margin_left=1.18,
    margin_right=1.18,
    page_number_style="arabic",
    show_cover=False,
)

RESUME = SceneProfile(
    name="resume",
    display_name="简历",
    description="简历排版，微软雅黑正文，1.15倍行距，紧凑排版",
    font_name="微软雅黑",
    font_name_western="Calibri",
    heading_font="微软雅黑",
    body_size="五号",              # 10.5pt
    heading_sizes={1: "三号", 2: "四号", 3: "五号"},
    line_spacing=1.15,
    first_line_indent=None,
    alignment="left",
    paper_size="A4",
    margin_top=0.7,
    margin_bottom=0.7,
    margin_left=0.8,
    margin_right=0.8,
    page_number_style="none",
    show_cover=False,
)

EXAM = SceneProfile(
    name="exam",
    display_name="试卷/测验",
    description="试卷排版，宋体正文，预留答题区域，分值标记",
    font_name="宋体",
    font_name_western="Times New Roman",
    heading_font="黑体",
    body_size="小四",              # 12pt
    heading_sizes={1: "三号", 2: "四号", 3: "小四"},
    line_spacing=1.5,
    first_line_indent=None,
    alignment="left",
    paper_size="A4",
    margin_top=0.8,
    margin_bottom=0.8,
    margin_left=0.8,
    margin_right=0.8,
    page_number_style="arabic",
    show_cover=False,
)

JOURNAL = SceneProfile(
    name="journal",
    display_name="期刊论文",
    description="期刊论文排版（测绘学报规范），宋体正文+黑体标题，五号正文，1.3倍行距。内置模板: templates/journal_template.docx",
    font_name="宋体",
    font_name_western="Times New Roman",
    heading_font="黑体",
    body_size="五号",              # 10.5pt
    heading_sizes={1: "小四", 2: "五号", 3: "五号"},
    line_spacing=1.3,
    first_line_indent=28,
    alignment="justify",
    paper_size="A4",
    margin_top=1.0,
    margin_bottom=1.0,
    margin_left=1.18,
    margin_right=1.18,
    page_number_style="arabic",
    show_cover=False,
)

# ── 注册表 ─────────────────────────────────────────────────────────

SCENE_REGISTRY: dict[str, SceneProfile] = {
    "academic": ACADEMIC,
    "report": REPORT,
    "official": OFFICIAL,
    "contract": CONTRACT,
    "resume": RESUME,
    "exam": EXAM,
    "journal": JOURNAL,
}


def get_scene(name: str) -> SceneProfile | None:
    """按名称获取场景配置"""
    return SCENE_REGISTRY.get(name)


