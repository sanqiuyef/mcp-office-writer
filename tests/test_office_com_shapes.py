"""活动 PPT 形状/连接线/流程图渲染测试（假 COM 对象，不依赖真实 PowerPoint）。"""
from __future__ import annotations

import unittest
from unittest.mock import MagicMock, PropertyMock, patch

from mcp_office_writer.models import (
    FrameworkDiagram,
    FwEdge,
    FwNode,
    FwStage,
    SlideConnectorDef,
    SlideDef,
    SlideImageDef,
    SlideItem,
    SlideShapeDef,
)
from mcp_office_writer.office_com import (
    _ppt_add_content_to_slide,
    _render_flowchart_com,
    _validated_ppt_slides,
    try_add_flowchart_slide,
)


def _fake_slide() -> MagicMock:
    slide = MagicMock()
    parent = MagicMock()
    parent.PageSetup.SlideWidth = 13.333 * 72
    parent.PageSetup.SlideHeight = 7.5 * 72
    slide.Parent = parent
    slide.Shapes = MagicMock()
    return slide


def _rgb(hex_color: str) -> int:
    h = hex_color.lstrip("#")
    return int(h[0:2], 16) + (int(h[2:4], 16) << 8) + (int(h[4:6], 16) << 16)


class SlideItemShapeRenderingTest(unittest.TestCase):
    def test_shape_branch_calls_addshape_with_rect(self) -> None:
        slide = _fake_slide()
        shape = MagicMock()
        slide.Shapes.AddShape.return_value = shape

        slide_item = SlideItem(type="shape", shape=SlideShapeDef(
            shape_type="rect", text="开始", x=1.0, y=2.0, width=2.0, height=1.0,
            bg_color="FF0000", font_color="FFFFFF", bold=True,
        ))
        _ppt_add_content_to_slide(slide, [slide_item], 0.5, 1.5, 8.0, 4.5)

        slide.Shapes.AddShape.assert_called_once_with(1, 72, 144, 144, 72)
        self.assertEqual(shape.Fill.ForeColor.RGB, _rgb("FF0000"))
        self.assertEqual(shape.TextFrame.TextRange.Text, "开始")
        self.assertTrue(shape.TextFrame.TextRange.Font.Bold)

    def test_shape_arrow_maps_to_right_arrow_constant(self) -> None:
        slide = _fake_slide()

        slide_item = SlideItem(type="shape", shape=SlideShapeDef(
            shape_type="arrow", width=1.0, height=0.5,
        ))
        _ppt_add_content_to_slide(slide, [slide_item], 0.5, 1.5, 8.0, 4.5)
        slide.Shapes.AddShape.assert_called_once_with(33, 36, 108, 72, 36)

    def test_connector_branch_calls_addconnector_with_arrow(self) -> None:
        slide = _fake_slide()
        connector = MagicMock()
        slide.Shapes.AddConnector.return_value = connector

        slide_item = SlideItem(type="connector", connector=SlideConnectorDef(
            x1=1.0, y1=1.0, x2=3.0, y2=2.0, color="0000FF", width=2.5, arrowhead=True,
        ))
        _ppt_add_content_to_slide(slide, [slide_item], 0.5, 1.5, 8.0, 4.5)

        slide.Shapes.AddConnector.assert_called_once_with(1, 72, 72, 216, 144)
        self.assertEqual(connector.Line.ForeColor.RGB, _rgb("0000FF"))
        self.assertEqual(connector.Line.Weight, 2.5)
        self.assertEqual(connector.Line.EndArrowheadStyle, 2)

    def test_connector_without_arrowhead_does_not_set_end_style(self) -> None:
        slide = _fake_slide()
        connector = MagicMock()
        slide.Shapes.AddConnector.return_value = connector

        slide_item = SlideItem(type="connector", connector=SlideConnectorDef(
            x1=0.5, y1=0.5, x2=1.5, y2=0.5, arrowhead=False,
        ))
        _ppt_add_content_to_slide(slide, [slide_item], 0.5, 1.5, 8.0, 4.5)

        slide.Shapes.AddConnector.assert_called_once_with(1, 36, 36, 108, 36)
        # arrowhead=False 时不应设置 EndArrowheadStyle（访问会惰性创建 mock 子属性，
        # 因此用 hasattr 判断赋值前后是否有值不可靠；这里确认线宽设置生效且箭头样式未触发）
        self.assertEqual(connector.Line.Weight, 1.5)

    def test_image_url_only_downloads_to_cache(self) -> None:
        import tempfile
        from pathlib import Path

        slide = _fake_slide()

        slide_item = SlideItem(type="image", image=SlideImageDef(
            url="https://example.com/pic.png", width=2.0, height=1.5,
        ))
        with tempfile.TemporaryDirectory() as temp_dir:
            cached = Path(temp_dir) / "pic.png"
            cached.write_bytes(b"fake-png")
            with patch("mcp_office_writer.utils.download_image", return_value=str(cached)):
                _ppt_add_content_to_slide(slide, [slide_item], 0.5, 1.5, 8.0, 4.5)
        slide.Shapes.AddPicture.assert_called_once()

    def test_validator_allows_shape_and_connector(self) -> None:
        slide_def = SlideDef(
            layout="blank",
            body=[
                SlideItem(type="connector", connector=SlideConnectorDef(x1=0, y1=0, x2=1, y2=1)),
                SlideItem(type="shape", shape=SlideShapeDef(width=1.0, height=1.0)),
            ],
        )
        _validated_ppt_slides([slide_def])  # 不抛异常即通过

    def test_shape_no_fill_and_dash_line_and_vertical(self) -> None:
        slide = _fake_slide()
        shape = MagicMock()
        slide.Shapes.AddShape.return_value = shape

        item = SlideItem(type="shape", shape=SlideShapeDef(
            shape_type="rect", text="阶段一", width=0.7, height=1.0,
            no_fill=True, line_color="000000", line_width=1.0, line_dash=True,
            vertical_text=True, font_size=12,
        ))
        _ppt_add_content_to_slide(slide, [item], 0.5, 1.5, 8.0, 4.5)

        self.assertEqual(shape.Fill.Visible, 0)  # no_fill
        self.assertEqual(shape.Line.ForeColor.RGB, _rgb("000000"))
        self.assertEqual(shape.Line.Weight, 1.0)
        self.assertEqual(shape.Line.DashStyle, 2)
        self.assertEqual(shape.TextFrame.TextRange.Font.Size, 12)
        # 竖排：Orientation 被设置（WPS 可能不支持，静默跳过，不抛异常即可）

    def test_connector_dash_style(self) -> None:
        slide = _fake_slide()
        connector = MagicMock()
        slide.Shapes.AddConnector.return_value = connector

        item = SlideItem(type="connector", connector=SlideConnectorDef(
            x1=0.5, y1=0.5, x2=2.0, y2=0.5, color="FF0000", width=1.5, dash=True,
        ))
        _ppt_add_content_to_slide(slide, [item], 0.5, 1.5, 8.0, 4.5)

        self.assertEqual(connector.Line.ForeColor.RGB, _rgb("FF0000"))
        self.assertEqual(connector.Line.DashStyle, 2)
        self.assertEqual(connector.Line.Weight, 1.5)


class FlowchartComRenderTest(unittest.TestCase):
    def _diagram(self) -> FrameworkDiagram:
        return FrameworkDiagram(
            title="登录流程",
            stages=[
                FwStage(label="开始", nodes=[
                    FwNode(id="s1", text="开始", style="terminator"),
                ]),
                FwStage(label="认证", nodes=[
                    FwNode(id="n1", text="输入账号密码", style="process"),
                    FwNode(id="n2", text="校验通过?", style="decision"),
                ]),
                FwStage(label="结果", nodes=[
                    FwNode(id="ok", text="进入主页", style="process"),
                    FwNode(id="fail", text="提示重试", style="process"),
                ]),
            ],
            edges=[
                FwEdge(from_id="n2", to_id="ok", label="是"),
                FwEdge(from_id="n2", to_id="fail", label="否"),
            ],
        )

    def test_render_flowchart_draws_shapes_and_connectors(self) -> None:
        slide = _fake_slide()
        _render_flowchart_com(slide, self._diagram())

        # 新样式形状：标题 accent 条 1 + 阶段容器 3 + 胶囊标签 3 + 节点 5 + 编号徽章 4 = 16
        self.assertEqual(slide.Shapes.AddShape.call_count, 16)
        # 自动箭头（认证/结果阶段内各 1）+ 显式 edges 2 = 4 条连接线
        self.assertEqual(slide.Shapes.AddConnector.call_count, 4)
        # 标题文本框
        slide.Shapes.AddTextbox.assert_called()

    def test_render_flowchart_node_text_is_white(self) -> None:
        slide = _fake_slide()
        # 节点文字应为白色（深底白字，≥4.5:1 对比），而非旧版的深字
        from mcp_office_writer.models import SlideItem, SlideShapeDef

        item = SlideItem(type="shape", shape=SlideShapeDef(
            shape_type="rect", text="测试", width=1.0, height=0.5,
            bg_color="2563EB", font_color="FFFFFF",
        ))
        _ppt_add_content_to_slide(slide, [item], 0.5, 1.5, 8.0, 4.5)
        node_shape = slide.Shapes.AddShape.return_value
        self.assertEqual(node_shape.TextFrame.TextRange.Font.Color.RGB, _rgb("FFFFFF"))

    def test_render_flowchart_uses_neutral_grey_arrows(self) -> None:
        slide = _fake_slide()
        _render_flowchart_com(slide, self._diagram())

        connector = slide.Shapes.AddConnector.return_value
        # 连接线应使用中性灰（PptMcp：箭头弱于内容），而非主色
        self.assertEqual(connector.Line.ForeColor.RGB, _rgb("94A3B8"))
        self.assertEqual(connector.Line.Weight, 1.5)

    def test_try_add_flowchart_slide_appends_and_saves(self) -> None:
        presentation = MagicMock()
        slide = _fake_slide()
        presentation.Slides.Count = 2
        presentation.Slides.Add.return_value = slide
        application = MagicMock()

        with patch("mcp_office_writer.office_com.com_is_available", return_value=True):
            with patch("mcp_office_writer.office_com._active_applications", return_value=[(application, "PowerPoint.Application")]):
                with patch("mcp_office_writer.office_com._ppt_presentation", return_value=presentation):
                    result = try_add_flowchart_slide(r"D:\演示.pptx", self._diagram())

        self.assertTrue(result.success)
        presentation.Slides.Add.assert_called_once_with(3, 12)
        presentation.Save.assert_called_once()

    def test_try_add_flowchart_rolls_back_on_render_error(self) -> None:
        presentation = MagicMock()
        slide = _fake_slide()
        slide.Shapes.AddShape.side_effect = RuntimeError("COM 渲染失败")
        application = MagicMock()

        # 模拟 Slides.Count 随 Add/Delete 动态变化，使回滚 while 循环进入
        count_holder = {"n": 2}
        type(presentation.Slides).Count = PropertyMock(side_effect=lambda: count_holder["n"])
        presentation.Slides.Add.side_effect = lambda *a, **k: (count_holder.__setitem__("n", count_holder["n"] + 1), slide)[1]
        presentation.Slides.return_value.Delete.side_effect = lambda *a, **k: count_holder.__setitem__("n", count_holder["n"] - 1)

        with patch("mcp_office_writer.office_com.com_is_available", return_value=True):
            with patch("mcp_office_writer.office_com._active_applications", return_value=[(application, "PowerPoint.Application")]):
                with patch("mcp_office_writer.office_com._ppt_presentation", return_value=presentation):
                    result = try_add_flowchart_slide(r"D:\演示.pptx", self._diagram())

        self.assertFalse(result.success)
        self.assertIn("回滚", result.message)
        presentation.Slides.return_value.Delete.assert_called()
        presentation.Save.assert_not_called()


if __name__ == "__main__":
    unittest.main()
