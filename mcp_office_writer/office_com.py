"""Windows COM bridge for Office documents that are already open."""

from __future__ import annotations

import os
from contextlib import contextmanager
from typing import Any, Iterator

from .models import EditResult, FrameworkDiagram, FontConfig, ParagraphStyle, SectionItem, SheetDef, SlideDef, SlideItem, TextStyle
from .utils import cn_size_to_pt


class ActiveOfficeDocumentError(RuntimeError):
    """The requested active-document operation cannot be completed safely."""


def com_is_available() -> bool:
    """Return whether this process can safely attempt Windows COM automation."""
    if os.name != "nt":
        return False
    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401
    except ImportError:
        return False
    return True


@contextmanager
def _com_apartment() -> Iterator[None]:
    import pythoncom

    pythoncom.CoInitialize()
    try:
        yield
    finally:
        pythoncom.CoUninitialize()


def _normalise(path: str) -> str:
    return os.path.normcase(os.path.abspath(path))


def _active_applications(prog_ids: tuple[str, ...]) -> list[tuple[Any, str]]:
    import win32com.client

    applications: list[tuple[Any, str]] = []
    for prog_id in prog_ids:
        try:
            applications.append((win32com.client.GetActiveObject(prog_id), prog_id))
        except Exception:
            continue
    return applications


def _word_document(application: Any, file_path: str) -> Any | None:
    target = _normalise(file_path)
    for index in range(1, application.Documents.Count + 1):
        document = application.Documents.Item(index)
        if _normalise(str(document.FullName)) == target:
            return document
    return None


def _excel_workbook(application: Any, file_path: str) -> Any | None:
    target = _normalise(file_path)
    for index in range(1, application.Workbooks.Count + 1):
        workbook = application.Workbooks.Item(index)
        if _normalise(str(workbook.FullName)) == target:
            return workbook
    return None


def get_open_office_documents() -> dict[str, object]:
    """List Word/WPS, Excel, and PowerPoint documents currently available through COM."""
    result: dict[str, object] = {"word": [], "excel": [], "powerpoint": []}
    if not com_is_available():
        return result
    try:
        with _com_apartment():
            for application, prog_id in _active_applications(("Word.Application", "kwps.Application")):
                for index in range(1, application.Documents.Count + 1):
                    document = application.Documents.Item(index)
                    result["word"].append({"file_path": str(document.FullName), "application": prog_id})
            for application, prog_id in _active_applications(("Excel.Application", "et.Application")):
                for index in range(1, application.Workbooks.Count + 1):
                    workbook = application.Workbooks.Item(index)
                    result["excel"].append({"file_path": str(workbook.FullName), "application": prog_id})
            for application, prog_id in _active_applications(("PowerPoint.Application",)):
                for index in range(1, application.Presentations.Count + 1):
                    presentation = application.Presentations.Item(index)
                    result["powerpoint"].append({"file_path": str(presentation.FullName), "application": prog_id})
    except Exception as exc:
        # COM 应用未响应/弹窗等异常：返回已收集的部分结果，不向调用方抛异常
        result["error"] = str(exc)
    return result


def _word_color(value: str) -> int:
    value = value.lstrip("#")
    if len(value) != 6:
        raise ActiveOfficeDocumentError(f"无效的颜色值: {value!r}")
    try:
        red, green, blue = int(value[:2], 16), int(value[2:4], 16), int(value[4:], 16)
    except ValueError as exc:
        raise ActiveOfficeDocumentError(f"无效的颜色值: {value!r}") from exc
    return red + (green << 8) + (blue << 16)


def _apply_word_style(insertion: Any, style: ParagraphStyle | None, default_font: FontConfig | None) -> None:
    if default_font:
        if default_font.name:
            insertion.Font.Name = default_font.name
        if default_font.size is not None:
            insertion.Font.Size = cn_size_to_pt(default_font.size)
        if default_font.bold is not None:
            insertion.Font.Bold = default_font.bold
        if default_font.italic is not None:
            insertion.Font.Italic = default_font.italic
        if default_font.color:
            insertion.Font.Color = _word_color(default_font.color)
    if not style:
        return
    if style.font_name:
        insertion.Font.Name = style.font_name
    if style.font_size is not None:
        insertion.Font.Size = cn_size_to_pt(style.font_size)
    if style.bold is not None:
        insertion.Font.Bold = style.bold
    if style.italic is not None:
        insertion.Font.Italic = style.italic
    if style.color:
        insertion.Font.Color = _word_color(style.color)

    paragraph = insertion.ParagraphFormat
    alignment = {"left": 0, "center": 1, "right": 2, "justify": 3}
    if style.alignment:
        paragraph.Alignment = alignment[style.alignment]
    if style.line_spacing is not None:
        paragraph.LineSpacingRule = 5  # wdLineSpaceMultiple
        paragraph.LineSpacing = 12 * style.line_spacing
    if style.first_line_indent is not None:
        paragraph.FirstLineIndent = style.first_line_indent
    if style.space_before is not None:
        paragraph.SpaceBefore = style.space_before
    if style.space_after is not None:
        paragraph.SpaceAfter = style.space_after


def _validate_word_sections(sections: list[SectionItem]) -> None:
    supported = {"heading", "paragraph", "bullet_list", "numbered_list", "page_break"}
    unsupported = [item.type for item in sections if item.type not in supported]
    if unsupported:
        raise ActiveOfficeDocumentError(
            f"活动 Word 文档暂不支持 {', '.join(repr(item) for item in unsupported)}；请关闭文件后使用原 writer"
        )


def _append_word_item(document: Any, item: SectionItem, default_font: FontConfig | None) -> None:
    content = document.Content

    def append_paragraph(text: str) -> Any:
        start = max(0, content.End - 1)
        content.InsertAfter("\r" + text)
        return document.Range(start + 1, max(start + 1, document.Content.End - 1))

    if item.type == "page_break":
        append_paragraph("").InsertBreak(7)  # wdPageBreak
        return
    if item.type in {"bullet_list", "numbered_list"}:
        for value in item.items or []:
            insertion = append_paragraph(str(value))
            _apply_word_style(insertion, item.style, default_font)
            if item.type == "bullet_list":
                insertion.ListFormat.ApplyBulletDefault()
            else:
                insertion.ListFormat.ApplyNumberDefault()
        return
    insertion = append_paragraph(item.text or "")
    if item.type == "heading":
        try:
            insertion.Style = f"Heading {item.level or 1}"
        except Exception:
            insertion.Font.Bold = True
    _apply_word_style(insertion, item.style, default_font)


def _undo_word_record(application: Any, undo_record: Any | None) -> None:
    if undo_record is None:
        return
    try:
        undo_record.EndCustomRecord()
        application.Undo()
    except Exception:
        # The primary operation remains failed.  The message below asks the user
        # not to save if an Office variant does not expose Word's undo record.
        pass


def try_edit_open_word_document(file_path: str, append_sections: list[SectionItem], default_font: FontConfig | None = None) -> EditResult | None:
    """Edit an active Word/WPS document, or return ``None`` when not applicable."""
    if not com_is_available():
        return None
    try:
        with _com_apartment():
            application = None
            prog_id = None
            document = None
            for candidate, candidate_prog_id in _active_applications(("Word.Application", "kwps.Application")):
                candidate_document = _word_document(candidate, file_path)
                if candidate_document is not None:
                    application, prog_id, document = candidate, candidate_prog_id, candidate_document
                    break
            if document is None or application is None or prog_id is None:
                return None
            try:
                if bool(document.ReadOnly):
                    raise ActiveOfficeDocumentError("活动 Word 文档为只读，无法保存")
                _validate_word_sections(append_sections)
                undo_record = getattr(application, "UndoRecord", None)
                if undo_record is not None:
                    undo_record.StartCustomRecord("mcp-office-writer")
                try:
                    for item in append_sections:
                        _append_word_item(document, item, default_font)
                    document.Save()
                except Exception:
                    _undo_word_record(application, undo_record)
                    raise
                if undo_record is not None:
                    undo_record.EndCustomRecord()
                return EditResult(success=True, file_path=file_path, operation_count=len(append_sections), message=f"已通过 {prog_id} 修改并保存活动 Word 文档")
            except Exception as exc:
                return EditResult(success=False, file_path=file_path, message=f"活动 Word 文档编辑失败: {exc}")
    except Exception as exc:
        # 前置 COM 查找（GetActiveObject/枚举文档）失败也返回错误，不向调用方裸抛
        return EditResult(success=False, file_path=file_path, message=f"活动 Word 文档操作失败: {exc}")


def _validate_excel_sheets(workbook: Any, sheets: list[SheetDef]) -> None:
    existing = {str(workbook.Worksheets.Item(index).Name).casefold() for index in range(1, workbook.Worksheets.Count + 1)}
    planned: set[str] = set()
    invalid_chars = set("[]:*?/\\")
    for sheet in sheets:
        name = sheet.name[:31]
        if not name or any(char in invalid_chars for char in name):
            raise ActiveOfficeDocumentError(f"无效的工作表名称: {sheet.name!r}")
        normalised = name.casefold()
        if normalised in existing or normalised in planned:
            raise ActiveOfficeDocumentError(f"工作表名称已存在: {name!r}")
        planned.add(normalised)


def _excel_color(value: str) -> int:
    return _word_color(value)


def _apply_excel_style(cell: Any, style: Any, default_bold: bool = False) -> None:
    if style is None:
        if default_bold:
            cell.Font.Bold = True
        return
    if style.font_name:
        cell.Font.Name = style.font_name
    if style.font_size is not None:
        cell.Font.Size = cn_size_to_pt(style.font_size)
    if style.bold is not None:
        cell.Font.Bold = style.bold
    elif default_bold:
        cell.Font.Bold = True
    if style.font_color:
        cell.Font.Color = _excel_color(style.font_color)
    if style.bg_color:
        cell.Interior.Color = _excel_color(style.bg_color)
    if style.number_format:
        cell.NumberFormat = style.number_format
    if style.alignment:
        cell.HorizontalAlignment = {"left": -4131, "center": -4108, "right": -4152}[style.alignment]
    if style.wrap_text is not None:
        cell.WrapText = style.wrap_text


def _write_sheet(workbook: Any, sheet: SheetDef) -> Any:
    worksheet = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
    worksheet.Name = sheet.name[:31]
    if sheet.headers:
        for column_index, value in enumerate(sheet.headers, start=1):
            cell = worksheet.Cells(1, column_index)
            cell.Value = value
            _apply_excel_style(cell, sheet.header_style, default_bold=True)
    start_row = 2 if sheet.headers else 1
    for row_index, row in enumerate(sheet.data, start=start_row):
        for column_index, value in enumerate(row, start=1):
            cell = worksheet.Cells(row_index, column_index)
            if isinstance(value, str) and value.startswith("=") and not sheet.allow_formulas:
                cell.Value = "'" + value
            else:
                cell.Value = value
            if sheet.column_formats and column_index <= len(sheet.column_formats):
                _apply_excel_style(cell, sheet.column_formats[column_index - 1])
    if sheet.column_widths:
        for column_index, width in enumerate(sheet.column_widths, start=1):
            worksheet.Columns(column_index).ColumnWidth = width
    if sheet.merged_cells:
        for merged in sheet.merged_cells:
            worksheet.Range(merged.range).Merge()
    if sheet.auto_filter and sheet.headers and sheet.data:
        last_column = max(len(sheet.headers), *(len(row) for row in sheet.data))
        worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(len(sheet.data) + 1, last_column)).AutoFilter()
    if sheet.tab_color:
        worksheet.Tab.Color = _excel_color(sheet.tab_color)
    if sheet.freeze:
        application = workbook.Application
        previous_sheet = application.ActiveSheet
        previous_selection = application.Selection
        try:
            worksheet.Activate()
            application.ActiveWindow.FreezePanes = False
            worksheet.Range(sheet.freeze).Select()
            application.ActiveWindow.FreezePanes = True
        finally:
            try:
                previous_sheet.Activate()
                previous_selection.Select()
            except Exception:
                pass
    return worksheet


def _rollback_excel_sheets(workbook: Any, worksheets: list[Any]) -> None:
    if not worksheets:
        return
    application = workbook.Application
    previous_alerts = application.DisplayAlerts
    try:
        application.DisplayAlerts = False
        for worksheet in reversed(worksheets):
            worksheet.Delete()
    finally:
        application.DisplayAlerts = previous_alerts


def try_edit_open_excel_workbook(file_path: str, append_sheets: list[SheetDef]) -> EditResult | None:
    """Append sheets to an active Excel workbook, or return ``None`` when not applicable."""
    if not com_is_available():
        return None
    try:
        with _com_apartment():
            application = None
            prog_id = None
            workbook = None
            for candidate, candidate_prog_id in _active_applications(("Excel.Application", "et.Application")):
                candidate_workbook = _excel_workbook(candidate, file_path)
                if candidate_workbook is not None:
                    application, prog_id, workbook = candidate, candidate_prog_id, candidate_workbook
                    break
            if workbook is None or application is None or prog_id is None:
                return None
            created: list[Any] = []
            try:
                if bool(workbook.ReadOnly):
                    raise ActiveOfficeDocumentError("活动 Excel 工作簿为只读，无法保存")
                _validate_excel_sheets(workbook, append_sheets)
                for sheet in append_sheets:
                    created.append(_write_sheet(workbook, sheet))
                workbook.Save()
                return EditResult(success=True, file_path=file_path, operation_count=len(append_sheets), message=f"已通过 {prog_id} 修改并保存活动 Excel 工作簿")
            except Exception as exc:
                try:
                    _rollback_excel_sheets(workbook, created)
                except Exception as rollback_exc:
                    return EditResult(success=False, file_path=file_path, message=f"活动 Excel 工作簿编辑失败: {exc}；回滚失败，请勿保存该工作簿: {rollback_exc}")
                return EditResult(success=False, file_path=file_path, message=f"活动 Excel 工作簿编辑失败，已回滚本次新增工作表: {exc}")
    except Exception as exc:
        # 前置 COM 查找（GetActiveObject/枚举工作簿）失败也返回错误，不向调用方裸抛
        return EditResult(success=False, file_path=file_path, message=f"活动 Excel 工作簿操作失败: {exc}")


# ── PowerPoint COM ──────────────────────────────────────────────

PPT_COM_LAYOUT_MAP: dict[str, int] = {
    "title": 1,               # ppLayoutTitle
    "text": 2,                # ppLayoutText
    "content": 2,             # ppLayoutText alias
    "section": 3,             # ppLayoutSectionHeader
    "two_content": 4,         # ppLayoutTwoContent
    "comparison": 5,          # ppLayoutComparison
    "title_only": 6,          # ppLayoutTitleOnly
    "blank": 12,              # ppLayoutBlank
    "content_caption": 7,     # ppLayoutContentCaption
    "picture_caption": 8,     # ppLayoutPictureCaption
}


def _ppt_presentation(application: Any, file_path: str) -> Any | None:
    """Find a Presentation by path in a PowerPoint.Application."""
    target = _normalise(file_path)
    for index in range(1, application.Presentations.Count + 1):
        presentation = application.Presentations.Item(index)
        if _normalise(str(presentation.FullName)) == target:
            return presentation
    return None


def _ppt_color(value: str) -> int:
    """Convert HEX color to Office COM COLORREF (same as _word_color)."""
    return _word_color(value)


# ── PPT 形状类型（MsoAutoShapeType 常量值，与 python-pptx MSO_SHAPE 数值一致） ──
_PPT_SHAPE_TYPE_MAP: dict[str, int] = {
    "rect": 1,                    # msoShapeRectangle
    "rounded_rect": 5,            # msoShapeRoundedRectangle
    "oval": 9,                    # msoShapeOval
    "diamond": 4,                 # msoShapeDiamond
    "arrow": 33,                  # msoShapeRightArrow
    "left_arrow": 34,             # msoShapeLeftArrow
    "chevron": 52,                # msoShapeChevron
    "parallelogram": 2,           # msoShapeParallelogram
    "process": 61,                # msoShapeFlowchartProcess
    "decision": 63,               # msoShapeFlowchartDecision
    "data": 62,                   # msoShapeFlowchartData
    "document": 60,               # msoShapeFlowchartDocument
    "terminator": 78,             # msoShapeFlowchartTerminator
    "predefined": 77,             # msoShapeFlowchartPredefinedProcess
    "cylinder": 98,               # msoShapeCan
    "callout": 105,               # msoShapeRoundedRectangularCallout
    "cloud": 179,                 # msoShapeCloud
}

# MsoArrowheadStyle 常量：msoArrowheadTriangle=2, msoArrowheadOpen=3, msoArrowheadStealth=4
_PPT_ARROWHEAD_STYLE: dict[str, int] = {
    "triangle": 2,
    "open": 3,
    "stealth": 4,
}

# MsoConnectorType 常量：msoConnectorStraight=1, msoConnectorElbow=2, msoConnectorCurve=3
_PPT_CONNECTOR_TYPE: dict[str, int] = {
    "straight": 1,
    "elbow": 2,
    "curve": 3,
}

# MsoAlignParagraph 常量：ppAlignLeft=1, ppAlignCenter=2, ppAlignRight=3
_PPT_ALIGN: dict[str, int] = {"left": 1, "center": 2, "right": 3}


def _ppt_apply_text_style(text_range: Any, style: TextStyle | None) -> None:
    """Apply TextStyle to a PowerPoint COM TextRange."""
    if style is None:
        return
    font = text_range.Font
    if style.font_name:
        font.Name = style.font_name
    if style.font_size is not None:
        font.Size = float(style.font_size)
    if style.bold is not None:
        font.Bold = style.bold
    if style.italic is not None:
        font.Italic = style.italic
    if style.color:
        font.Color.RGB = _ppt_color(style.color)


def _validated_ppt_slides(slides: list[SlideDef]) -> None:
    """Validate that all slide item types are supported by COM bridge."""
    supported = {"paragraph", "table", "image", "shape", "connector"}
    for idx, slide_def in enumerate(slides, start=1):
        for section_name in ("body", "left", "right"):
            items = getattr(slide_def, section_name, None)
            if not items:
                continue
            for item in items:
                if item.type not in supported:
                    raise ActiveOfficeDocumentError(
                        f"幻灯片 #{idx} 的 {item.type!r} 类型暂不支持活动 PPT 编辑"
                    )


def _ppt_add_content_to_slide(slide: Any, items: list[SlideItem] | None,
                               default_left: float, default_top: float,
                               default_width: float, default_height: float) -> None:
    """Add SlideItem contents to a COM Slide object."""
    if not items:
        return

    for item in items:
        if item.type == "paragraph" and item.text:
            # Always use a dedicated textbox for body paragraphs
            left_pt = int(default_left * 72)
            top_pt = int(default_top * 72)
            width_pt = int(default_width * 72)
            height_pt = int(default_height * 72)
            shape = slide.Shapes.AddTextbox(1, left_pt, top_pt, width_pt, height_pt)
            tf = shape.TextFrame
            tf.TextRange.Text = item.text
            _ppt_apply_text_style(tf.TextRange, item.style)
            if item.style and item.style.alignment:
                alignment = {"left": 1, "center": 2, "right": 3}
                tf.TextRange.ParagraphFormat.Alignment = alignment.get(item.style.alignment, 1)

        elif item.type == "table" and item.table:
            td = item.table
            rows = len(td.rows) + (1 if td.headers else 0)
            cols = max(len(td.headers) if td.headers else 0,
                       max((len(r) for r in td.rows), default=0))
            if rows == 0 or cols == 0:
                continue
            left_pt = int(default_left * 72)
            top_pt = int(default_top * 72)
            width_pt = int(default_width * 72)
            height_pt = int(default_height * 72)
            table_shape = slide.Shapes.AddTable(rows, cols, left_pt, top_pt, width_pt, height_pt)
            table = table_shape.Table

            if td.headers:
                for ci, h in enumerate(td.headers, start=1):
                    cell = table.Cell(1, ci)
                    cell.Shape.TextFrame.TextRange.Text = str(h)
            start_data_row = 2 if td.headers else 1
            for ri, row_data in enumerate(td.rows):
                for ci, val in enumerate(row_data, start=1):
                    if ci <= cols:
                        cell = table.Cell(start_data_row + ri, ci)
                        cell.Shape.TextFrame.TextRange.Text = str(val)

        elif item.type == "image" and item.image:
            img_def = item.image
            img_path = img_def.path
            if not img_path and img_def.url:
                # url-only 图片：先下载到缓存再插入，避免 os.path.exists(None) 抛 TypeError
                try:
                    from .utils import download_image
                    img_path = download_image(img_def.url)
                except Exception:
                    continue
            if not img_path or not os.path.exists(img_path):
                continue
            left_pt = int((img_def.x or default_left) * 72)
            top_pt = int((img_def.y or default_top) * 72)
            width_pt = int((img_def.width or 4.0) * 72)
            height_pt = int((img_def.height or 3.0) * 72)
            slide.Shapes.AddPicture(img_path, False, True, left_pt, top_pt, width_pt, height_pt)

        elif item.type == "shape" and item.shape:
            sd = item.shape
            shape_type = _PPT_SHAPE_TYPE_MAP.get(sd.shape_type, 1)
            left_pt = int((sd.x if sd.x is not None else default_left) * 72)
            top_pt = int((sd.y if sd.y is not None else default_top) * 72)
            width_pt = int(sd.width * 72)
            height_pt = int(sd.height * 72)
            shape = slide.Shapes.AddShape(shape_type, left_pt, top_pt, width_pt, height_pt)
            if sd.no_fill:
                shape.Fill.Visible = 0  # msoFalse
            elif sd.bg_color:
                shape.Fill.ForeColor.RGB = _ppt_color(sd.bg_color)
            if sd.line_color:
                shape.Line.ForeColor.RGB = _ppt_color(sd.line_color)
            if sd.line_width is not None:
                shape.Line.Weight = sd.line_width
            if sd.line_dash:
                shape.Line.DashStyle = 2  # msoLineDash
            if sd.text:
                text_range = shape.TextFrame.TextRange
                text_range.Text = sd.text
                text_range.ParagraphFormat.Alignment = 2  # ppAlignCenter
                if sd.vertical_text:
                    try:
                        shape.TextFrame.Orientation = 5  # msoTextOrientationVerticalFarEast
                    except Exception:
                        pass
                if sd.font_color:
                    text_range.Font.Color.RGB = _ppt_color(sd.font_color)
                if sd.font_size:
                    text_range.Font.Size = sd.font_size
                if sd.bold:
                    text_range.Font.Bold = True

        elif item.type == "connector" and item.connector:
            cd = item.connector
            connector = slide.Shapes.AddConnector(
                1,  # msoConnectorStraight
                int(cd.x1 * 72), int(cd.y1 * 72),
                int(cd.x2 * 72), int(cd.y2 * 72),
            )
            if cd.color:
                connector.Line.ForeColor.RGB = _ppt_color(cd.color)
            connector.Line.Weight = cd.width
            if cd.dash:
                connector.Line.DashStyle = 2  # msoLineDash
            # PowerPoint 的箭头样式挂在 Shape.Line 上（EndArrowheadStyle 不是 Shape 直接属性）
            if cd.arrowhead:
                connector.Line.EndArrowheadStyle = _PPT_ARROWHEAD_STYLE.get(cd.arrowhead_style, 2)
            if cd.begin_arrowhead:
                connector.Line.BeginArrowheadStyle = 2  # msoArrowheadTriangle


def try_edit_open_powerpoint(file_path: str, append_slides: list[SlideDef]) -> EditResult | None:
    """Append slides to an active PowerPoint presentation, or return ``None`` when not applicable."""
    if not com_is_available():
        return None
    try:
        with _com_apartment():
            application = None
            prog_id = None
            presentation = None
            for candidate, candidate_prog_id in _active_applications(("PowerPoint.Application",)):
                candidate_presentation = _ppt_presentation(candidate, file_path)
                if candidate_presentation is not None:
                    application, prog_id, presentation = candidate, candidate_prog_id, candidate_presentation
                    break
            if presentation is None or application is None or prog_id is None:
                return None

            try:
                _validated_ppt_slides(append_slides)

                created_count = 0
                initial_slide_count = presentation.Slides.Count
                for slide_def in append_slides:
                    layout_idx = PPT_COM_LAYOUT_MAP.get(slide_def.layout, 2)
                    position = presentation.Slides.Count + 1
                    slide = presentation.Slides.Add(position, layout_idx)

                    # 设置背景色
                    if slide_def.bg_color:
                        slide.FollowMasterBackground = False
                        slide.Background.Fill.ForeColor.RGB = _ppt_color(slide_def.bg_color)

                    # 设置标题
                    if slide_def.title:
                        if slide.Shapes.HasTitle:
                            slide.Shapes.Title.TextFrame.TextRange.Text = slide_def.title

                    # 设置副标题（仅 title/section 布局）
                    if slide_def.subtitle:
                        for shape in slide.Shapes:
                            if shape.HasTextFrame:
                                try:
                                    # ppPlaceholderSubtitle = 4
                                    if shape.PlaceholderFormat.Type == 4:
                                        shape.TextFrame.TextRange.Text = slide_def.subtitle
                                        break
                                except Exception:
                                    pass

                    # 主体内容
                    if slide_def.body:
                        _ppt_add_content_to_slide(slide, slide_def.body, 0.5, 1.5, 8.0, 4.5)

                    # 双栏布局
                    if slide_def.left:
                        _ppt_add_content_to_slide(slide, slide_def.left, 0.5, 1.5, 4.0, 4.5)
                    if slide_def.right:
                        _ppt_add_content_to_slide(slide, slide_def.right, 5.0, 1.5, 4.0, 4.5)

                    # 演讲者备注
                    if slide_def.notes:
                        # ppPlaceholderBody = 2, the notes text placeholder
                        for shape in slide.NotesPage.Shapes:
                            if shape.HasTextFrame:
                                try:
                                    if shape.PlaceholderFormat.Type == 2:  # ppPlaceholderBody
                                        shape.TextFrame.TextRange.Text = slide_def.notes
                                        break
                                except Exception:
                                    pass

                    created_count += 1

                presentation.Save()
                return EditResult(
                    success=True,
                    file_path=file_path,
                    operation_count=created_count,
                    message=f"已通过 {prog_id} 修改并保存活动 PowerPoint 演示文稿",
                )
            except Exception as exc:
                # 回滚：删除本次新增的所有幻灯片
                try:
                    while presentation.Slides.Count > initial_slide_count:
                        presentation.Slides(presentation.Slides.Count).Delete()
                except Exception as rollback_exc:
                    return EditResult(
                        success=False,
                        file_path=file_path,
                        message=f"活动 PowerPoint 演示文稿编辑失败: {exc}；回滚失败，请勿保存该演示文稿: {rollback_exc}",
                    )
                return EditResult(
                    success=False,
                    file_path=file_path,
                    message=f"活动 PowerPoint 演示文稿编辑失败，已回滚本次新增幻灯片: {exc}",
                )
    except Exception as exc:
        # 前置 COM 查找（GetActiveObject/枚举演示文稿）失败也返回错误，不向调用方裸抛
        return EditResult(success=False, file_path=file_path, message=f"活动 PowerPoint 演示文稿操作失败: {exc}")


# ── 活动 PPT 流程图（COM 渲染） ──────────────────────────────────


def _com_textbox(slide: Any, left_pt: int, top_pt: int, width_pt: int, height_pt: int) -> Any:
    """在 COM Slide 上添加文本框（msoTextOrientationHorizontal=1）。"""
    return slide.Shapes.AddTextbox(1, left_pt, top_pt, width_pt, height_pt)


def _com_shape(slide: Any, shape_type: int, left_pt: int, top_pt: int,
               width_pt: int, height_pt: int) -> Any:
    """在 COM Slide 上添加形状，返回 shape。"""
    return slide.Shapes.AddShape(shape_type, left_pt, top_pt, width_pt, height_pt)


def _com_connector(slide: Any, conn_type: int, x1: float, y1: float,
                   x2: float, y2: float, color_hex: str | None,
                   width_pt: float = 1.5, arrowhead: bool = True) -> None:
    """在 COM Slide 上添加连接线（坐标单位：英寸），可选箭头。"""
    connector = slide.Shapes.AddConnector(
        conn_type,
        int(x1 * 72), int(y1 * 72), int(x2 * 72), int(y2 * 72),
    )
    if color_hex:
        connector.Line.ForeColor.RGB = _ppt_color(color_hex)
    connector.Line.Weight = width_pt
    if arrowhead:
        connector.Line.EndArrowheadStyle = 2  # msoArrowheadTriangle


def _set_com_shape_text(shape: Any, text: str, font_cn: str, font_size: int,
                        bold: bool, color_hex: str) -> None:
    """设置 COM 形状的文字（\n 换行转 PowerPoint 垂直制表符 \v）。

    兼容真实换行符（MCP JSON 转义）与字面反斜杠+n 两种写法。
    """
    text_range = shape.TextFrame.TextRange
    text_range.Text = text.replace("\\n", "\n").replace("\n", "\v")
    text_range.ParagraphFormat.Alignment = 2  # ppAlignCenter
    text_range.Font.Name = font_cn
    text_range.Font.Size = font_size
    text_range.Font.Bold = bold
    text_range.Font.Color.RGB = _ppt_color(color_hex)


def _com_add_shadow(shape: Any, blur_pt: float = 8.0, offset_pt: float = 3.0,
                    transparency: float = 0.5) -> None:
    """给 COM 形状加柔和外阴影（Shape.Shadow 原生属性）。

    WPS 兼容层可能不支持部分属性，失败时静默跳过，不影响主流程。
    """
    try:
        shadow = shape.Shadow
        shadow.Visible = True
        shadow.Blur = blur_pt
        shadow.OffsetX = offset_pt
        shadow.OffsetY = offset_pt
        shadow.Transparency = transparency  # 0=不透明, 1=全透明
    except Exception:
        pass


def _render_flowchart_com(slide: Any, diagram: FrameworkDiagram) -> None:
    """在 COM Slide 上渲染 FrameworkDiagram（样式与 create_framework_diagram 一致）。"""
    from .ppt_writer import (
        _PRESET_COLOR_SCHEMES, _mix_with_black, _mix_with_white,
    )

    scheme = _PRESET_COLOR_SCHEMES.get(diagram.color_scheme, _PRESET_COLOR_SCHEMES["professional-blue"])
    # 连接线用中性灰弱化（PptMcp 设计规范：箭头应比内容弱）
    arrow_color = diagram.arrow_color or "94A3B8"
    body_font_cn = diagram.font_cn

    # 使用活动演示文稿的页面尺寸
    slide_width_in = float(slide.Parent.PageSetup.SlideWidth) / 72.0
    slide_height_in = float(slide.Parent.PageSetup.SlideHeight) / 72.0

    # 背景
    slide.FollowMasterBackground = False
    slide.Background.Fill.ForeColor.RGB = _ppt_color(scheme["background_color"])

    # ── 标题 ──────────────────────────────────────────────
    title_y_offset = 0.2
    accent_bar_y = 0.62
    if diagram.title:
        title_box = _com_textbox(slide, int(0.8 * 72), int(0.15 * 72), int(11.5 * 72), int(0.45 * 72))
        title_range = title_box.TextFrame.TextRange
        title_range.Text = diagram.title
        title_range.ParagraphFormat.Alignment = 2
        title_range.Font.Size = diagram.title_font_size
        title_range.Font.Bold = True
        title_range.Font.Name = body_font_cn
        title_range.Font.Color.RGB = _ppt_color(scheme["title_color"])
        title_y_offset = 0.7

    if diagram.subtitle:
        sub_box = _com_textbox(slide, int(0.8 * 72), int(0.5 * 72), int(11.5 * 72), int(0.3 * 72))
        sub_range = sub_box.TextFrame.TextRange
        sub_range.Text = diagram.subtitle
        sub_range.ParagraphFormat.Alignment = 2
        sub_range.Font.Size = 12
        sub_range.Font.Color.RGB = _ppt_color(scheme["body_color"])
        title_y_offset = 0.95
        accent_bar_y = 0.9

    # 标题下 accent 短横条
    if diagram.title:
        bar_w = 11.5 * 0.6
        bar = _com_shape(slide, 1, int((0.8 + (11.5 - bar_w) / 2) * 72), int(accent_bar_y * 72),
                         int(bar_w * 72), int(0.04 * 72))
        bar.Fill.ForeColor.RGB = _ppt_color(scheme["accent_color"])
        bar.Line.Visible = 0

    stage_count = len(diagram.stages)
    if stage_count == 0:
        return

    content_left = 1.15
    content_right = slide_width_in - 0.3
    content_width = content_right - content_left
    usable_height = slide_height_in - title_y_offset - 0.3
    stage_unit_h = usable_height / stage_count - diagram.stage_gap

    node_id_to_pos: dict[str, dict] = {}

    for si, stage in enumerate(diagram.stages):
        stage_y = title_y_offset + si * (stage_unit_h + diagram.stage_gap)
        stage_color = stage.color or scheme["accent_color"]

        # 阶段背景容器（极淡同色系底 + 同色系细边 + 阴影）
        bg = _com_shape(slide, 5, int((content_left + 0.55) * 72), int(stage_y * 72),
                        int((content_width - 0.55) * 72), int(stage_unit_h * 72))
        bg.Fill.ForeColor.RGB = _ppt_color(_mix_with_white(stage_color, 0.96))
        bg.Line.ForeColor.RGB = _ppt_color(_mix_with_white(stage_color, 0.88))
        bg.Line.Weight = 0.75
        _com_add_shadow(bg, blur_pt=10, offset_pt=2, transparency=0.75)

        # 阶段标签（实心主色胶囊 + 白字）
        pill_w, pill_h = 1.25, 0.38
        pill = _com_shape(slide, 5, int(0.25 * 72), int((stage_y + stage_unit_h / 2 - pill_h / 2) * 72),
                          int(pill_w * 72), int(pill_h * 72))
        pill.Fill.ForeColor.RGB = _ppt_color(stage_color)
        pill.Line.ForeColor.RGB = _ppt_color(_mix_with_black(stage_color, 0.2))
        pill.Line.Weight = 0.75
        _com_add_shadow(pill)
        _set_com_shape_text(pill, stage.label, body_font_cn, 12, True, "FFFFFF")

        nodes = stage.nodes
        if not nodes:
            continue

        # 节点水平排列
        node_count = len(nodes)
        node_w = diagram.node_width
        node_h = diagram.node_height
        total_nodes_w = node_count * node_w + (node_count - 1) * 0.3
        start_x = content_left + 0.6 + max(0, (content_width - 0.6 - total_nodes_w) / 2)
        node_cy = stage_y + stage_unit_h / 2 - node_h / 2

        prev_id: str | None = None
        for ni, node in enumerate(nodes):
            nx = start_x + [n.id for n in nodes].index(node.id) * (node_w + 0.3)
            n_color = node.fill_color or stage_color
            shape_type = _PPT_SHAPE_TYPE_MAP.get(node.style, 5)  # 默认 rounded_rect

            shape = _com_shape(slide, shape_type,
                               int(nx * 72), int(node_cy * 72),
                               int(node_w * 72), int(node_h * 72))
            # 节点：主色填充 + 同色系深边框 + 白字
            shape.Fill.ForeColor.RGB = _ppt_color(n_color)
            shape.Line.ForeColor.RGB = _ppt_color(_mix_with_black(n_color, 0.25))
            shape.Line.Weight = 1.0
            _com_add_shadow(shape)
            _set_com_shape_text(shape, node.text, body_font_cn, node.font_size or 14, node.bold, "FFFFFF")

            # 步骤编号徽章（阶段内节点 ≥2 时）：白色小圆 + 主色 Consolas 数字
            if node_count >= 2 and node.style != "terminator":
                badge_d = 0.24
                badge = _com_shape(slide, 9, int((nx - 0.08) * 72), int((node_cy - 0.08) * 72),
                                   int(badge_d * 72), int(badge_d * 72))
                badge.Fill.ForeColor.RGB = _ppt_color("FFFFFF")
                badge.Line.ForeColor.RGB = _ppt_color(n_color)
                badge.Line.Weight = 1.0
                try:
                    badge_range = badge.TextFrame.TextRange
                    badge_range.Text = str(ni + 1)
                    badge_range.ParagraphFormat.Alignment = 2
                    badge_range.Font.Size = 9
                    badge_range.Font.Bold = True
                    badge_range.Font.Name = "Consolas"
                    badge_range.Font.Color.RGB = _ppt_color(n_color)
                except Exception:
                    pass

            node_id_to_pos[node.id] = {"x": nx, "y": node_cy, "w": node_w, "h": node_h}

            # 同阶段相邻节点自动箭头（中性灰弱化）
            if prev_id is not None:
                prev = node_id_to_pos[prev_id]
                _com_connector(slide, 1,
                               prev["x"] + prev["w"], prev["y"] + prev["h"] / 2,
                               nx, node_cy + node_h / 2,
                               arrow_color)
            prev_id = node.id

    # ── 显式 edges（跨阶段连接） ───────────────────────────
    for edge in diagram.edges:
        if edge.from_id not in node_id_to_pos or edge.to_id not in node_id_to_pos:
            continue
        f = node_id_to_pos[edge.from_id]
        t = node_id_to_pos[edge.to_id]
        fx = f["x"] + f["w"] / 2
        fy = f["y"] + f["h"]
        tx = t["x"] + t["w"] / 2
        ty = t["y"]
        conn_type = _PPT_CONNECTOR_TYPE.get(edge.edge_style, 1)
        _com_connector(slide, conn_type, fx, fy, tx, ty, arrow_color)

        if edge.label:
            label_box = _com_textbox(slide, int(((fx + tx) / 2 - 0.5) * 72),
                                     int(((fy + ty) / 2 - 0.25) * 72),
                                     int(1.2 * 72), int(0.35 * 72))
            label_range = label_box.TextFrame.TextRange
            label_range.Text = edge.label
            label_range.ParagraphFormat.Alignment = 2
            label_range.Font.Size = 11
            label_range.Font.Italic = True
            label_range.Font.Color.RGB = _ppt_color(arrow_color)


def try_add_flowchart_slide(file_path: str, diagram: FrameworkDiagram) -> EditResult | None:
    """在活动 PowerPoint 演示文稿末尾追加一张流程图幻灯片，或返回 ``None`` 当不可用。

    布局与文件级 ``ppt_create_framework_diagram`` 一致（阶段标签 + 节点 + 箭头），
    但全程通过 COM 在活动演示文稿中渲染；失败时回滚删除新增的幻灯片。
    """
    if not com_is_available():
        return None
    try:
        with _com_apartment():
            application = None
            prog_id = None
            presentation = None
            for candidate, candidate_prog_id in _active_applications(("PowerPoint.Application",)):
                candidate_presentation = _ppt_presentation(candidate, file_path)
                if candidate_presentation is not None:
                    application, prog_id, presentation = candidate, candidate_prog_id, candidate_presentation
                    break
            if presentation is None or application is None or prog_id is None:
                return None

            initial_slide_count = presentation.Slides.Count
            try:
                slide = presentation.Slides.Add(initial_slide_count + 1, 12)  # ppLayoutBlank
                _render_flowchart_com(slide, diagram)
                presentation.Save()
                node_count = sum(len(s.nodes) for s in diagram.stages)
                return EditResult(
                    success=True,
                    file_path=file_path,
                    operation_count=node_count,
                    message=f"已通过 {prog_id} 追加流程图幻灯片（{len(diagram.stages)} 个阶段, {node_count} 个节点）",
                )
            except Exception as exc:
                try:
                    while presentation.Slides.Count > initial_slide_count:
                        presentation.Slides(presentation.Slides.Count).Delete()
                except Exception as rollback_exc:
                    return EditResult(
                        success=False, file_path=file_path,
                        message=f"流程图生成失败: {exc}；回滚失败，请勿保存该演示文稿: {rollback_exc}",
                    )
                return EditResult(
                    success=False, file_path=file_path,
                    message=f"流程图生成失败，已回滚新增幻灯片: {exc}",
                )
    except Exception as exc:
        # 前置 COM 查找失败也返回错误，不向调用方裸抛
        return EditResult(success=False, file_path=file_path, message=f"活动 PowerPoint 流程图操作失败: {exc}")
