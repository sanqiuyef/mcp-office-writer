
"""
mcp-office-writer - Excel (.xlsx) 写入/编辑逻辑
"""
from __future__ import annotations

from typing import TYPE_CHECKING

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

if TYPE_CHECKING:
    from .models import SheetDef, CellStyle, CreateResult, EditResult


def _close_workbook(workbook) -> None:
    """Release both the workbook and the optional macro archive it owns."""
    try:
        workbook.close()
    finally:
        vba_archive = getattr(workbook, "vba_archive", None)
        if vba_archive is not None:
            try:
                vba_archive.close()
            except (OSError, ValueError):
                # An archive may already have been closed by openpyxl internals.
                pass


def _build_font(style: CellStyle | None, default_size: int = 11) -> Font:
    if style is None:
        return Font(size=default_size)
    return Font(
        name=style.font_name,
        size=style.font_size or default_size,
        bold=style.bold,
        italic=False,
        color=style.font_color,
    )


def _build_fill(style: CellStyle | None) -> PatternFill | None:
    if style and style.bg_color:
        return PatternFill(start_color=style.bg_color, end_color=style.bg_color, fill_type="solid")
    return None


def _build_alignment(style: CellStyle | None) -> Alignment | None:
    if style is None:
        return None
    h = style.alignment
    if h or style.wrap_text:
        return Alignment(horizontal=h, wrap_text=style.wrap_text)
    return None


def _apply_cell_style(cell, style: CellStyle | None, default_size: int = 11):
    """对单元格应用样式"""
    if style is None:
        return
    if style.font_name or style.font_size or style.bold or style.font_color:
        cell.font = _build_font(style, default_size)
    if style.bg_color:
        cell.fill = _build_fill(style)
    if style.alignment or style.wrap_text:
        cell.alignment = _build_alignment(style)


def _get_column_style(sheet_def: SheetDef, col_idx: int) -> CellStyle:
    """获取某一列的列格式（如果有）"""
    from .models import CellStyle

    if sheet_def.column_formats and col_idx < len(sheet_def.column_formats):
        cf = sheet_def.column_formats[col_idx]
        cs = CellStyle(
            font_name=cf.font_name,
            font_color=cf.font_color,
            bg_color=cf.bg_color,
            number_format=cf.number_format,
            alignment=cf.alignment,
        )
        return cs
    return CellStyle()


def _write_sheet(ws, sheet_def: SheetDef):
    """写入一个工作表"""

    # 写入表头
    if sheet_def.headers:
        for ci, h in enumerate(sheet_def.headers, 1):
            cell = ws.cell(row=1, column=ci, value=h)
            _apply_cell_style(cell, sheet_def.header_style)
            # 如果没有特别指定表头样式，默认加粗
            if sheet_def.header_style is None:
                cell.font = Font(bold=True)

    start_row = 2 if sheet_def.headers else 1

    # 写入数据
    for ri, row_data in enumerate(sheet_def.data, start_row):
        for ci, val in enumerate(row_data, 1):
            cell = ws.cell(row=ri, column=ci)
            # 公式识别：字符串以 = 开头
            if sheet_def.allow_formulas and isinstance(val, str) and val.startswith("="):
                cell.value = val  # openpyxl 自动识别为公式
            elif isinstance(val, str) and val.startswith("="):
                cell.value = "'" + val
            else:
                cell.value = val

            # 应用列格式
            col_style = _get_column_style(sheet_def, ci - 1)
            if col_style:
                _apply_cell_style(cell, col_style)
                if col_style.number_format:
                    cell.number_format = col_style.number_format

    # 列宽
    if sheet_def.column_widths:
        for ci, w in enumerate(sheet_def.column_widths, 1):
            ws.column_dimensions[get_column_letter(ci)].width = w

    # 合并单元格
    if sheet_def.merged_cells:
        for mc in sheet_def.merged_cells:
            ws.merge_cells(mc.range)

    # 冻结窗格
    if sheet_def.freeze:
        ws.freeze_panes = sheet_def.freeze

    # 自动筛选
    if sheet_def.auto_filter and sheet_def.headers and sheet_def.data:
        last_col = len(sheet_def.headers)
        last_row = len(sheet_def.data) + (1 if sheet_def.headers else 0)
        ws.auto_filter.ref = f"A1:{get_column_letter(last_col)}{last_row}"

    # Sheet 标签颜色
    if sheet_def.tab_color:
        ws.sheet_properties.tabColor = sheet_def.tab_color


def create_excel_workbook(
    output_path: str,
    sheets: list[SheetDef],
    template: str | None = None,
    if_exists: str = "overwrite",
) -> CreateResult:
    """创建 Excel 工作簿"""
    from .models import CreateResult
    from .utils import atomic_save, ensure_output_dir, get_file_size, resolve_if_exists, validate_input_file, validate_output_path

    wb = None
    try:
        validate_output_path(output_path, (".xlsx", ".xlsm"))
        ensure_output_dir(output_path)
        resolved_path = resolve_if_exists(output_path, if_exists)
    except (OSError, ValueError) as e:
        return CreateResult(success=False, file_path=output_path, message=str(e))

    try:
        output_is_macro = resolved_path.lower().endswith(".xlsm")
        template_is_macro = bool(template and template.lower().endswith((".xlsm", ".xltm")))
        if output_is_macro and not template_is_macro:
            raise ValueError("生成 .xlsm 必须提供 .xlsm 或 .xltm 宏模板，以避免扩展名与内容不匹配")
        if template_is_macro and not output_is_macro:
            raise ValueError("使用宏模板时输出文件必须为 .xlsm，以避免丢失 VBA 内容")

        if template:
            validate_input_file(template, (".xlsx", ".xlsm", ".xltx", ".xltm"))
            wb = load_workbook(template, keep_vba=template.lower().endswith((".xlsm", ".xltm")))
            # 删除默认 sheet
            for name in list(wb.sheetnames):
                del wb[name]
        else:
            wb = Workbook()
            # 删除默认 sheet
            wb.remove(wb.active)

        for sheet_def in sheets:
            ws = wb.create_sheet(title=sheet_def.name[:31])  # Excel 限制 sheet 名 31 字符
            _write_sheet(ws, sheet_def)

        atomic_save(wb.save, resolved_path)
        size = get_file_size(resolved_path)

        return CreateResult(
            success=True,
            file_path=resolved_path,
            file_size=size,
            sheets_count=len(sheets),
            message=f"Excel 工作簿创建成功: {resolved_path}",
        )
    except Exception as e:
        return CreateResult(
            success=False,
            file_path=resolved_path,
            message=f"Excel 工作簿创建失败: {str(e)}",
        )
    finally:
        if wb is not None:
            _close_workbook(wb)


def edit_excel_workbook(
    file_path: str,
    append_sheets: list[SheetDef],
) -> EditResult:
    """向已有 Excel 追加工作表"""
    from .models import EditResult
    from .utils import atomic_save, validate_input_file

    # An open Excel instance can initialise an Explorer-created placeholder
    # before it becomes a valid OOXML archive on disk.
    from .office_com import try_edit_open_excel_workbook
    com_result = try_edit_open_excel_workbook(file_path, append_sheets)
    if com_result is not None:
        return com_result

    wb = None
    try:
        validate_input_file(file_path, (".xlsx", ".xlsm"))
    except ValueError as e:
        return EditResult(
            success=False,
            file_path=file_path,
            message=str(e),
        )

    try:
        # `.xlsm` 保存时必须保留 VBA archive，否则追加工作表会把宏静默剥离。
        wb = load_workbook(file_path, keep_vba=file_path.lower().endswith(".xlsm"))
        count = 0
        for sheet_def in append_sheets:
            ws = wb.create_sheet(title=sheet_def.name[:31])
            _write_sheet(ws, sheet_def)
            count += 1
        atomic_save(wb.save, file_path)
        return EditResult(
            success=True,
            file_path=file_path,
            operation_count=count,
            message=f"已追加 {count} 个工作表到 Excel 文件",
        )
    except Exception as e:
        return EditResult(
            success=False,
            file_path=file_path,
            message=f"编辑 Excel 工作簿失败: {str(e)}",
        )
    finally:
        if wb is not None:
            _close_workbook(wb)
