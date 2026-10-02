"""
mcp-office-writer 读取模块
从 Word (.docx)、Excel (.xlsx)、PowerPoint (.pptx) 文件中提取文本内容。
"""
import logging
import os
import re
from collections import Counter
from pathlib import Path

logger = logging.getLogger(__name__)
_TRUNCATION_MARKER = "_(内容已截断)_"

# ── OCR 配置（环境变量可覆盖） ──
_OCR_ENABLED = os.environ.get("MCP_OFFICE_OCR", "1") == "1"
_OCR_TESSERACT_CMD = os.environ.get(
    "TESSERACT_CMD",
    r"D:\cherry-workplace\tools\tesseract\tesseract.exe",
)
# Tesseract 可执行文件不可用时 OCR 自动降级（返回文本层并在结果中提示），
# 而不是抛 RuntimeError 导致整个 PDF 读取失败。
_OCR_TESSERACT_AVAILABLE = bool(_OCR_TESSERACT_CMD) and os.path.isfile(_OCR_TESSERACT_CMD)
_OCR_LANG = os.environ.get("MCP_OFFICE_OCR_LANG", "chi_sim+eng")
_OCR_DPI = int(os.environ.get("MCP_OFFICE_OCR_DPI", "300"))
_OCR_MAX_PAGES = int(os.environ.get("MCP_OFFICE_OCR_MAX_PAGES", "0"))  # 0=全部
_OCR_MIN_TEXT_CHARS = int(os.environ.get("MCP_OFFICE_OCR_MIN_TEXT_CHARS", "500"))
_OCR_QUICK_CHECK_PAGES = int(os.environ.get("MCP_OFFICE_OCR_QUICK_PAGES", "3"))


def _content_with_truncation_marker(parts: list[str], max_chars: int) -> str:
    """Return complete parts plus a marker without exceeding ``max_chars``."""
    content = "\n\n".join(parts)
    if not content:
        return _TRUNCATION_MARKER[:max_chars]
    candidate = f"{content}\n\n{_TRUNCATION_MARKER}"
    return candidate if len(candidate) <= max_chars else content


def _truncate_at_boundary(parts: list[str], max_chars: int) -> tuple[str, bool]:
    """按段落/表格边界截断，保证返回的 Markdown 结构完整"""
    content = "\n\n".join(parts)
    if len(content) <= max_chars:
        return content, False

    # 从第一个超过 max_chars 的 part 处截断
    length = 0
    for i, part in enumerate(parts):
        sep = "\n\n" if i > 0 else ""
        if length + len(sep) + len(part) > max_chars:
            # 在此 part 处截断：当前这个 part 不加入，之前的所有 part 保留
            return _content_with_truncation_marker(parts[:i], max_chars), True
        length += len(sep) + len(part)

    return content, False


# ── OCR 扫描件检测与识别 ──

def _pdf_text_needs_ocr(text: str) -> bool:
    """判断 PDF 文本层是否缺失或几乎只包含重复水印。"""
    if not _OCR_ENABLED:
        return False
    lines = [re.sub(r"\s+", "", line) for line in text.splitlines()]
    lines = [line for line in lines if len(line) >= 3]
    compact_text = "".join(lines)
    if len(compact_text) < _OCR_MIN_TEXT_CHARS:
        # 短文档：文本层存在可读内容（≥3 字符的行）即视为正常文本 PDF，
        # 不做 OCR——避免把正常短文档识别成乱码。
        # 纯扫描件在 get_text 下通常没有 ≥3 字符的文本行。
        return len(lines) == 0
    # 重复水印检测：大量重复行
    if len(lines) >= 8:
        repeated = Counter(lines).most_common(1)[0][1]
        if len(set(lines)) <= max(3, len(lines) // 8) or repeated / len(lines) >= 0.85:
            return True
    return False


def _ocr_pdf(filepath: str) -> str:
    """使用 PyMuPDF 渲染扫描件，再交给 Tesseract 逐页识别。"""
    import fitz
    from io import BytesIO
    from PIL import Image
    import pytesseract

    if not os.path.isfile(_OCR_TESSERACT_CMD):
        raise RuntimeError(f"未找到 Tesseract: {_OCR_TESSERACT_CMD}")

    pytesseract.pytesseract.tesseract_cmd = _OCR_TESSERACT_CMD
    document = fitz.open(filepath)
    page_count = len(document)
    limit = min(page_count, _OCR_MAX_PAGES) if _OCR_MAX_PAGES > 0 else page_count
    scale = _OCR_DPI / 72
    matrix = fitz.Matrix(scale, scale)
    pages = []

    try:
        for page_number in range(limit):
            page = document.load_page(page_number)
            pixmap = page.get_pixmap(matrix=matrix, alpha=False)
            with Image.open(BytesIO(pixmap.tobytes("png"))) as image:
                page_text = pytesseract.image_to_string(
                    image, lang=_OCR_LANG, config="--psm 6"
                )
            if page_text.strip():
                pages.append(f"## Page {page_number + 1}\n\n{page_text.strip()}")
            else:
                pages.append(f"## Page {page_number + 1}\n\n_(OCR 未识别到文字)_")
    finally:
        document.close()

    if limit < page_count:
        pages.append(
            f"\n\n_[OCR 仅处理前 {limit}/{page_count} 页，"
            f"受 MCP_OFFICE_OCR_MAX_PAGES 限制]_"
        )
    return "\n\n".join(pages)


def _safe_heading_level(style_name: str) -> int | None:
    """从样式名称中安全提取标题层级"""
    if not style_name:
        return None
    m = re.search(r"(?:Heading|heading)\s*(\d+)", style_name)
    if m:
        return int(m.group(1))
    return None


def read_word_document(file_path: str, max_chars: int = 50000) -> dict:
    """读取 Word (.docx) 文档内容，提取文字、表格、列表等。

    Args:
        file_path: .docx 文件路径
        max_chars: 返回文本的最大字符数（按段落/表格边界截断，保证结构完整）

    Returns:
        dict: {success, file_path, content, paragraphs_count, tables_count, truncated, message}
    """
    if max_chars < 1:
        return {
            "success": False, "file_path": file_path, "content": "",
            "paragraphs_count": 0, "tables_count": 0, "truncated": False,
            "message": "max_chars 必须大于 0",
        }

    if not os.path.isfile(file_path):
        return {
            "success": False,
            "file_path": file_path,
            "content": "",
            "paragraphs_count": 0,
            "tables_count": 0,
            "truncated": False,
            "message": f"文件不存在: {file_path}",
        }

    if not file_path.lower().endswith(".docx"):
        return {
            "success": False,
            "file_path": file_path,
            "content": "",
            "paragraphs_count": 0,
            "tables_count": 0,
            "truncated": False,
            "message": "不支持的文件格式，仅支持 .docx",
        }

    try:
        from .utils import validate_input_file
        validate_input_file(file_path, (".docx",))
    except ValueError as e:
        return {
            "success": False, "file_path": file_path, "content": "",
            "paragraphs_count": 0, "tables_count": 0, "truncated": False, "message": str(e),
        }

    try:
        from docx import Document

        doc = Document(file_path)
        parts = []
        para_count = 0
        total_para_count = 0
        total_table_count = len(doc.tables)

        # 读取正文段落
        for para in doc.paragraphs:
            total_para_count += 1
            style_name = para.style.name if para.style else ""
            text = para.text.strip()
            if not text:
                continue

            para_count += 1
            level = _safe_heading_level(style_name)
            if level is not None:
                prefix = "#" * min(level, 6)
                parts.append(f"{prefix} {text}")
            elif "List" in style_name or style_name.startswith("List"):
                prefix = "- " if "Bullet" in style_name else "1. "
                parts.append(f"{prefix}{text}")
            else:
                parts.append(text)

        # 读取表格
        table_count_read = 0
        for table in doc.tables:
            table_count_read += 1
            # 跳过空表
            if len(table.rows) == 0:
                logger.warning("跳过空表格 #%d", table_count_read)
                continue
            parts.append(f"\n[Table {table_count_read}]")
            try:
                header_cells = [cell.text.strip() for cell in table.rows[0].cells]
            except IndexError:
                logger.warning("表格 #%d 首行为空，跳过", table_count_read)
                continue
            if any(header_cells):
                parts.append("| " + " | ".join(header_cells) + " |")
                parts.append("| " + " | ".join("---" for _ in header_cells) + " |")
            for row in table.rows[1:]:
                cells = [cell.text.strip() for cell in row.cells]
                parts.append("| " + " | ".join(cells) + " |")

        # 按段落/表格边界截断
        content, truncated = _truncate_at_boundary(parts, max_chars)

        msg = f"成功读取文档，共 {para_count} 段落、{total_table_count} 个表格"
        if truncated:
            msg += f"（已截断至 {max_chars} 字符）"

        return {
            "success": True,
            "file_path": file_path,
            "content": content,
            "paragraphs_count": para_count,
            "tables_count": total_table_count,
            "truncated": truncated,
            "message": msg,
        }

    except Exception as e:
        logger.error("读取 Word 文档失败: %s", e, exc_info=True)
        return {
            "success": False,
            "file_path": file_path,
            "content": "",
            "paragraphs_count": 0,
            "tables_count": 0,
            "truncated": False,
            "message": f"读取 Word 文档失败: {e}",
        }


def read_excel_workbook(
    file_path: str, max_rows_per_sheet: int = 1000, max_chars: int = 50000
) -> dict:
    """读取 Excel (.xlsx/.xlsm) 工作簿内容，提取所有工作表的数据。

    Args:
        file_path: .xlsx 文件路径
        max_rows_per_sheet: 每个 sheet 最多读取的行数
        max_chars: 所有 sheet 合计最多返回的字符数

    Returns:
        dict: {success, file_path, content, sheets_count, total_rows, truncated, message}
    """
    from .utils import validate_input_file

    if max_rows_per_sheet < 1 or max_chars < 1:
        return {
            "success": False,
            "file_path": file_path,
            "content": "",
            "sheets_count": 0,
            "total_rows": 0,
            "truncated": False,
            "message": "max_rows_per_sheet 和 max_chars 必须大于 0",
        }

    try:
        validate_input_file(file_path, (".xlsx", ".xlsm"))
    except ValueError as e:
        return {
            "success": False,
            "file_path": file_path,
            "content": "",
            "sheets_count": 0,
            "total_rows": 0,
            "truncated": False,
            "message": str(e),
        }

    wb = None
    try:
        from openpyxl import load_workbook

        wb = load_workbook(file_path, read_only=True, data_only=True)
        parts = []
        sheet_count = 0
        total_rows = 0
        current_length = 0
        truncated = False

        def add_part(part: str) -> bool:
            nonlocal current_length, truncated
            separator_length = 1 if parts else 0
            if current_length + separator_length + len(part) > max_chars:
                truncated = True
                return False
            parts.append(part)
            current_length += separator_length + len(part)
            return True

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            sheet_count += 1
            if not add_part(f"## Sheet: {sheet_name}"):
                break

            row_count = 0
            for row_idx, row in enumerate(ws.iter_rows(values_only=True)):
                if row_idx >= max_rows_per_sheet:
                    add_part(f"  _(已截断，超过 {max_rows_per_sheet} 行限制)_")
                    truncated = True
                    break
                # 跳过完全空行
                values = [str(v) if v is not None else "" for v in row]
                if all(v == "" for v in values):
                    continue
                if not add_part("| " + " | ".join(values) + " |"):
                    break
                total_rows += 1
                row_count += 1

            if truncated:
                break

            if row_count == 0:
                add_part("  _(空工作表)_")

        content = "\n".join(parts)
        if truncated:
            marker = "\n_(内容已截断)_"
            content = (content[: max(0, max_chars - len(marker))] + marker)[:max_chars]

        return {
            "success": True,
            "file_path": file_path,
            "content": content,
            "sheets_count": sheet_count,
            "total_rows": total_rows,
            "truncated": truncated,
            "message": (
                f"成功读取工作簿，共 {sheet_count} 个工作表、{total_rows} 行数据"
                + (f"（已截断至 {max_chars} 字符）" if truncated else "")
            ),
        }

    except Exception as e:
        return {
            "success": False,
            "file_path": file_path,
            "content": "",
            "sheets_count": 0,
            "total_rows": 0,
            "truncated": False,
            "message": f"读取 Excel 工作簿失败: {e}",
        }
    finally:
        if wb is not None:
            wb.close()


def read_powerpoint(file_path: str, max_chars: int = 50000) -> dict:
    """读取 PowerPoint (.pptx) 演示文稿内容，提取幻灯片文字和表格。

    Args:
        file_path: .pptx 文件路径
        max_chars: 返回文本的最大字符数（按幻灯片边界截断，保证结构完整）

    Returns:
        dict: {success, file_path, content, slides_count, truncated, message}
    """
    if max_chars < 1:
        return {
            "success": False, "file_path": file_path, "content": "",
            "slides_count": 0, "truncated": False, "message": "max_chars 必须大于 0",
        }

    if not os.path.isfile(file_path):
        return {
            "success": False,
            "file_path": file_path,
            "content": "",
            "slides_count": 0,
            "truncated": False,
            "message": f"文件不存在: {file_path}",
        }

    if not file_path.lower().endswith(".pptx"):
        return {
            "success": False,
            "file_path": file_path,
            "content": "",
            "slides_count": 0,
            "truncated": False,
            "message": "不支持的文件格式，仅支持 .pptx",
        }

    try:
        from .utils import validate_input_file
        validate_input_file(file_path, (".pptx",))
    except ValueError as e:
        return {
            "success": False, "file_path": file_path, "content": "",
            "slides_count": 0, "truncated": False, "message": str(e),
        }

    try:
        from pptx import Presentation

        prs = Presentation(file_path)
        slide_parts = []
        slide_num = 0
        total_slides = len(prs.slides)
        truncated = False

        for slide in prs.slides:
            slide_num += 1
            lines = []
            lines.append(f"## Slide {slide_num}")

            # 获取 slide 布局名称
            if slide.slide_layout:
                lines.append(f"**Layout:** {slide.slide_layout.name}")

            # 遍历所有形状
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for para in shape.text_frame.paragraphs:
                        text = para.text.strip()
                        if not text:
                            continue
                        level = para.level
                        prefix = "  " * level + ("- " if level > 0 else "")
                        lines.append(f"{prefix}{text}")

                if shape.has_table:
                    table = shape.table
                    lines.append("\n[Table]")
                    for row_idx, row in enumerate(table.rows):
                        cells = [cell.text.strip() for cell in row.cells]
                        line = "| " + " | ".join(cells) + " |"
                        lines.append(line)
                        if row_idx == 0:
                            lines.append("| " + " | ".join("---" for _ in cells) + " |")

            # 演讲者备注
            if slide.has_notes_slide:
                notes = slide.notes_slide.notes_text_frame.text.strip()
                if notes:
                    lines.append(f"\n*Notes:* {notes}")

            slide_text = "\n".join(lines)

            # 按幻灯片边界截断
            current_content = "\n\n".join(slide_parts)
            sep = "\n\n" if slide_parts else ""
            if len(current_content) + len(sep) + len(slide_text) > max_chars:
                truncated = True
                break

            slide_parts.append(slide_text)

        content = (
            _content_with_truncation_marker(slide_parts, max_chars)
            if truncated else "\n\n".join(slide_parts)
        )
        actual_slides = slide_num if not truncated else (slide_num - 1)

        msg = f"成功读取演示文稿，共 {total_slides} 张幻灯片"
        if truncated:
            msg += f"（已截断至 {max_chars} 字符，显示前 {actual_slides} 张）"

        return {
            "success": True,
            "file_path": file_path,
            "content": content,
            "slides_count": actual_slides,
            "truncated": truncated,
            "message": msg,
        }

    except Exception as e:
        return {
            "success": False,
            "file_path": file_path,
            "content": "",
            "slides_count": 0,
            "truncated": False,
            "message": f"读取 PowerPoint 演示文稿失败: {e}",
        }


def read_pdf_document(file_path: str, max_chars: int = 50000) -> dict:
    """读取 PDF 文档内容，逐页提取文字（基于 PyMuPDF）。

    相比基于 pypdf 的旧版本，文本提取顺序和质量更好，
    且附加文档元数据（标题/作者等）和书签结构信息。

    Args:
        file_path: .pdf 文件路径
        max_chars: 返回文本的最大字符数（按页面边界截断，保证结构完整）

    Returns:
        dict: {success, file_path, content, pages_count, truncated, message}
    """
    if max_chars < 1:
        return {
            "success": False, "file_path": file_path, "content": "",
            "pages_count": 0, "truncated": False, "message": "max_chars 必须大于 0",
        }

    if not os.path.isfile(file_path):
        return {
            "success": False, "file_path": file_path, "content": "",
            "pages_count": 0, "truncated": False, "message": f"文件不存在: {file_path}",
        }

    if not file_path.lower().endswith(".pdf"):
        return {
            "success": False, "file_path": file_path, "content": "",
            "pages_count": 0, "truncated": False, "message": "不支持的文件格式，仅支持 .pdf",
        }

    try:
        from .utils import validate_input_file
        validate_input_file(file_path, (".pdf",))
    except ValueError as e:
        return {
            "success": False, "file_path": file_path, "content": "",
            "pages_count": 0, "truncated": False, "message": str(e),
        }

    try:
        import pymupdf

        doc = pymupdf.open(file_path)
        total_pages = doc.page_count

        # ── 构建元数据头部 ──
        header_parts = []
        meta = doc.metadata or {}
        meta_items = []
        for key, label in [
            ("title", "标题"), ("author", "作者"), ("subject", "主题"),
            ("keywords", "关键词"), ("creator", "创建工具"), ("producer", "生成工具"),
        ]:
            val = meta.get(key, "").strip()
            if val:
                meta_items.append(f"{label}: {val}")
        if meta_items:
            header_parts.append("--- 文档元数据 ---\n" + "\n".join(meta_items))

        # ── 书签 / TOC ──
        toc = doc.get_toc()
        if toc:
            toc_lines = ["--- 书签 ---"]
            for level, title_text, page_num in toc:
                indent = "  " * (level - 1)
                toc_lines.append(f"{indent}- [{title_text}](page {page_num})")
            header_parts.append("\n".join(toc_lines))

        header = ("\n\n".join(header_parts) + "\n\n") if header_parts else ""

        # ── 逐页提取文本 ──
        page_parts = []
        truncated = False

        for i in range(total_pages):
            page = doc[i]
            text = page.get_text("text")
            if text and text.strip():
                page_text = f"## Page {i + 1}\n\n{text.strip()}"
            else:
                page_text = f"## Page {i + 1}\n\n_(空页)_"

            # 按页面边界截断（计入 header 长度）
            current_content = "\n\n".join(page_parts)
            sep = "\n\n" if page_parts else ""
            if len(header) + len(current_content) + len(sep) + len(page_text) > max_chars:
                truncated = True
                break

            page_parts.append(page_text)

        body = "\n\n".join(page_parts)

        # ── 扫描件 OCR 回退 ──
        # 用页面原始 get_text 拼接结果判断（不含 "## Page N" 页眉等加工文本，
        # 且短文本只要文本层有内容就不会误触发 OCR）
        raw_text = "\n".join(doc[i].get_text("text") for i in range(total_pages))
        if not truncated and _pdf_text_needs_ocr(raw_text):
            logger.info("检测到扫描件 PDF，尝试 OCR: %s", file_path)
            doc.close()
            content = header + body

            if not _OCR_TESSERACT_AVAILABLE:
                # OCR 自动降级：返回文本层，并在结果中提示 OCR 不可用
                note = "OCR 不可用（未找到 Tesseract 可执行文件），已返回 PDF 文本层"
                logger.warning("%s: %s", note, file_path)
                return {
                    "success": True, "file_path": file_path, "content": content,
                    "pages_count": total_pages, "truncated": False,
                    "message": f"成功读取 PDF，共 {total_pages} 页（{note}）",
                }

            try:
                ocr_content = _ocr_pdf(file_path)
            except Exception as e:
                # OCR 执行失败同样降级，不因 OCR 问题导致整个 PDF 读取失败
                note = f"OCR 不可用（{e}），已返回 PDF 文本层"
                logger.warning("OCR 失败，降级返回文本层: %s", e)
                return {
                    "success": True, "file_path": file_path, "content": content,
                    "pages_count": total_pages, "truncated": False,
                    "message": f"成功读取 PDF，共 {total_pages} 页（{note}）",
                }

            ocr_msg = f"成功读取 PDF（OCR），共 {total_pages} 页"
            # OCR 结果也按 max_chars 截断（显式截断标志，与文本分支一致）
            ocr_truncated = False
            if len(ocr_content) > max_chars:
                ocr_truncated = True
                # 找到前 max_chars 内最后一个完整的 ## Page
                trunc_idx = ocr_content.rfind("\n## Page ", 0, max_chars)
                if trunc_idx > 0:
                    ocr_content = ocr_content[:trunc_idx] + f"\n\n{_TRUNCATION_MARKER}"
                else:
                    ocr_content = ocr_content[:max_chars] + f"\n\n{_TRUNCATION_MARKER}"
            if ocr_truncated:
                ocr_msg += f"（已截断至 {max_chars} 字符）"
            return {
                "success": True, "file_path": file_path, "content": ocr_content,
                "pages_count": total_pages, "truncated": ocr_truncated,
                "message": ocr_msg,
            }

        content = header + body

        if truncated:
            # 尝试在末尾追加截断标记，如果放不下则用 _content_with_truncation_marker 压缩
            marker_suffix = f"\n\n{_TRUNCATION_MARKER}"
            if len(content) + len(marker_suffix) <= max_chars:
                content += marker_suffix
            else:
                # 压缩主体部分（_content_with_truncation_marker 感知 remaining 限额）
                if page_parts:
                    trimmed = _content_with_truncation_marker(page_parts, max_chars - len(header))
                    content = header + trimmed
                else:
                    # 没有任何页面内容放入时，保证 content 不超过 max_chars
                    content = header[:max_chars]


        doc.close()

        msg = f"成功读取 PDF，共 {total_pages} 页"
        if truncated:
            msg += f"（已截断至 {max_chars} 字符）"

        return {
            "success": True, "file_path": file_path, "content": content,
            "pages_count": total_pages, "truncated": truncated, "message": msg,
        }
    except Exception as e:
        return {
            "success": False, "file_path": file_path, "content": "",
            "pages_count": 0, "truncated": False, "message": f"读取 PDF 失败: {e}",
        }


def read_pdf_structured(
    file_path: str,
    max_chars: int = 50000,
    include_blocks: bool = False,
    extract_images: bool = False,
) -> dict:
    """读取 PDF 文档，返回结构化内容。

    相比 read_pdf_document 返回纯文本，此工具返回细粒度的结构化数据：
    - 文档元数据（标题/作者/主题/关键词/创建工具等）
    - 书签层级树（TOC）
    - 超链接列表
    - 逐页数据：纯文本 + 文本块信息（bbox/字体/字号，仅 include_blocks=True）
    - 图片提取（仅 extract_images=True，保存到 PDF 所在目录的 _images/ 子目录）

    Args:
        file_path: .pdf 文件路径
        max_chars: 纯文本总长度的截断上限（按页边界截断）
        include_blocks: 是否返回每页的文本块详细信息（bbox、字体、字号、颜色）
        extract_images: 是否提取 PDF 中的图片并保存到磁盘

    Returns:
        dict: {success, file_path, pages_count, message,
               metadata, bookmarks, links, pages, truncated}
    """
    if max_chars < 1:
        return {
            "success": False, "file_path": file_path, "pages_count": 0,
            "message": "max_chars 必须大于 0",
            "metadata": {}, "bookmarks": [], "links": [], "pages": [],
            "truncated": False,
        }

    if not os.path.isfile(file_path):
        return {
            "success": False, "file_path": file_path, "pages_count": 0,
            "message": f"文件不存在: {file_path}",
            "metadata": {}, "bookmarks": [], "links": [], "pages": [],
            "truncated": False,
        }

    if not file_path.lower().endswith(".pdf"):
        return {
            "success": False, "file_path": file_path, "pages_count": 0,
            "message": "不支持的文件格式，仅支持 .pdf",
            "metadata": {}, "bookmarks": [], "links": [], "pages": [],
            "truncated": False,
        }

    try:
        from .utils import validate_input_file
        validate_input_file(file_path, (".pdf",))
    except ValueError as e:
        return {
            "success": False, "file_path": file_path, "pages_count": 0,
            "message": str(e),
            "metadata": {}, "bookmarks": [], "links": [], "pages": [],
            "truncated": False,
        }

    try:
        import pymupdf

        doc = pymupdf.open(file_path)
        total_pages = doc.page_count

        # ── 元数据 ──
        metadata = {}
        raw_meta = doc.metadata or {}
        for key in ("title", "author", "subject", "keywords",
                     "creator", "producer", "creationDate", "modDate"):
            val = raw_meta.get(key, "")
            if val and val.strip():
                metadata[key] = val.strip()

        # ── 书签 / TOC ──
        bookmarks = []
        raw_toc = doc.get_toc()
        for level, title_text, page_num in raw_toc:
            bookmarks.append({
                "level": level,
                "title": title_text,
                "page": page_num,
            })

        # ── 链接 ──
        links = []
        for i in range(total_pages):
            page = doc[i]
            for link_info in page.get_links():
                link_entry = {"page": i + 1, "kind": link_info.get("kind", "")}
                if "uri" in link_info:
                    link_entry["uri"] = link_info["uri"]
                if "page" in link_info:
                    link_entry["dest_page"] = link_info["page"]
                if "from" in link_info:
                    link_entry["bbox"] = link_info["from"]
                links.append(link_entry)

        # ── 图片输出目录 ──
        image_dir = None
        if extract_images:
            pdf_stem = Path(file_path).stem
            image_dir = Path(file_path).parent / f"{pdf_stem}_images"
            image_dir.mkdir(parents=True, exist_ok=True)

        # ── 逐页数据 ──
        pages = []
        truncated = False
        text_len = 0

        for i in range(total_pages):
            page = doc[i]
            page_w = page.rect.width
            page_h = page.rect.height

            # 纯文本（用于截断判断）
            page_text = page.get_text("text").strip()
            if not page_text:
                page_text = "_(空页)_"
            text_len += len(page_text)

            page_entry = {
                "page": i + 1,
                "width": round(page_w, 2),
                "height": round(page_h, 2),
                "text": page_text,
            }

            # 文本块详细信息
            if include_blocks:
                blocks_data = []
                raw_dict = page.get_text("dict")
                for block in raw_dict.get("blocks", []):
                    btype = block.get("type")
                    entry = {
                        "type": "text" if btype == 0 else "image",
                        "bbox": [round(v, 2) for v in block.get("bbox", [0, 0, 0, 0])],
                    }
                    if btype == 0:  # 文本块
                        lines = []
                        for line in block.get("lines", []):
                            spans = []
                            for span in line.get("spans", []):
                                spans.append({
                                    "text": span.get("text", ""),
                                    "font": span.get("font", ""),
                                    "size": round(span.get("size", 0), 1),
                                    "color": span.get("color", 0),
                                })
                            if spans:
                                lines.append({"spans": spans})
                        entry["lines"] = lines
                    elif btype == 1 and extract_images:  # 图片块
                        # 通过 xref 提取图片
                        img_info = block.get("image", {})
                        if isinstance(img_info, dict) and "xref" in img_info:
                            xref = img_info["xref"]
                            try:
                                img_data = doc.extract_image(xref)
                                ext = img_data.get("ext", "png")
                                img_bytes = img_data.get("image", b"")
                                if img_bytes and image_dir:
                                    img_path = image_dir / f"page{i+1}_xref{xref}.{ext}"
                                    with open(img_path, "wb") as f:
                                        f.write(img_bytes)
                                    entry["saved_path"] = str(img_path)
                                    entry["image_width"] = img_data.get("width", 0)
                                    entry["image_height"] = img_data.get("height", 0)
                            except Exception:
                                pass
                    blocks_data.append(entry)
                page_entry["blocks"] = blocks_data
            elif extract_images:
                # 即使不返回 blocks 细节，也可以提取图片
                img_list = page.get_images(full=True)
                blocks_data = []
                for img_idx, img in enumerate(img_list):
                    xref = img[0]
                    bbox_info = page.get_image_bbox(img)
                    try:
                        img_data = doc.extract_image(xref)
                        ext = img_data.get("ext", "png")
                        img_bytes = img_data.get("image", b"")
                        if img_bytes and image_dir:
                            img_path = image_dir / f"page{i+1}_img{img_idx}.{ext}"
                            with open(img_path, "wb") as f:
                                f.write(img_bytes)
                            entry = {
                                "type": "image",
                                "saved_path": str(img_path),
                                "image_width": img_data.get("width", 0),
                                "image_height": img_data.get("height", 0),
                            }
                            if bbox_info:
                                entry["bbox"] = [
                                    round(bbox_info.x0, 2), round(bbox_info.y0, 2),
                                    round(bbox_info.x1, 2), round(bbox_info.y1, 2),
                                ]
                            blocks_data.append(entry)
                    except Exception:
                        pass
                if blocks_data:
                    page_entry["blocks"] = blocks_data

            # 按 text 长度截断
            if text_len > max_chars:
                truncated = True
                # 当前页已包含，标记截断
                break

            pages.append(page_entry)

        # ── 扫描件 OCR 回退（结构化模式） ──
        if not truncated:
            all_text = "\n".join(p["text"] for p in pages)
            if _pdf_text_needs_ocr(all_text):
                logger.info("检测到扫描件 PDF（结构化），启用 OCR: %s", file_path)
                ocr_text = _ocr_pdf(file_path)
                # 按 ## Page N 拆分 OCR 结果
                ocr_pages = []
                for part in re.split(r"\n(?=## Page \d+)", ocr_text.strip()):
                    part = part.strip()
                    if not part:
                        continue
                    m = re.match(r"## Page (\d+)", part)
                    pnum = int(m.group(1)) if m else (len(ocr_pages) + 1)
                    # 去除 ## Page N 标题本身
                    text_body = re.sub(r"^## Page \d+\n\n", "", part, count=1)
                    ocr_pages.append({
                        "page": pnum,
                        "width": 0, "height": 0,
                        "text": text_body,
                    })
                if ocr_pages:
                    pages = ocr_pages

        doc.close()

        msg = f"成功读取 PDF，共 {total_pages} 页"
        extracted_count = sum(
            1 for p in pages
            for b in p.get("blocks", []) if b.get("type") == "image" and b.get("saved_path")
        )
        if extracted_count:
            msg += f"，提取 {extracted_count} 张图片到 {image_dir}"
        if truncated:
            msg += f"（已截断至 {max_chars} 字符）"
        if include_blocks:
            msg += "，含文本块详细信息"

        return {
            "success": True,
            "file_path": file_path,
            "pages_count": total_pages,
            "message": msg,
            "metadata": metadata,
            "bookmarks": bookmarks,
            "links": links,
            "pages": pages,
            "truncated": truncated,
        }
    except Exception as e:
        return {
            "success": False, "file_path": file_path, "pages_count": 0,
            "message": f"读取 PDF 失败: {e}",
            "metadata": {}, "bookmarks": [], "links": [], "pages": [],
            "truncated": False,
        }


def read_txt_document(file_path: str, max_chars: int = 50000) -> dict:
    """读取纯文本文件内容，支持多种编码自动检测。

    支持 .txt、.md、.markdown 等纯文本格式。

    Args:
        file_path: 文件路径（.txt / .md / .markdown 等纯文本格式）
        max_chars: 返回文本的最大字符数（按行边界截断）

    Returns:
        dict: {success, file_path, content, encoding, file_size, truncated, message}
    """
    if max_chars < 1:
        return {
            "success": False, "file_path": file_path, "content": "",
            "encoding": "", "file_size": 0, "truncated": False,
            "message": "max_chars 必须大于 0",
        }

    if not os.path.isfile(file_path):
        return {
            "success": False, "file_path": file_path, "content": "",
            "encoding": "", "file_size": 0,
            "truncated": False, "message": f"文件不存在: {file_path}",
        }

    try:
        from .utils import validate_input_file
        validate_input_file(file_path, ())
    except ValueError as e:
        return {
            "success": False, "file_path": file_path, "content": "",
            "encoding": "", "file_size": 0, "truncated": False, "message": str(e),
        }

    # Latin-1 永远可以解码，必须放在最后；否则 UTF-16 会被静默误判为 Latin-1。
    encodings = ["utf-8-sig", "utf-8", "utf-16", "gb18030", "latin-1"]
    for encoding in encodings:
        try:
            with open(file_path, "r", encoding=encoding) as f:
                full_content = f.read()

            # 按行边界截断
            truncated = len(full_content) > max_chars
            if truncated:
                lines = full_content.splitlines(keepends=True)
                content = ""
                for line in lines:
                    if len(content) + len(line) > max_chars:
                        if len(content) + len(_TRUNCATION_MARKER) <= max_chars:
                            content += _TRUNCATION_MARKER
                        elif not content:
                            content = _TRUNCATION_MARKER[:max_chars]
                        break
                    content += line
            else:
                content = full_content

            file_size = os.path.getsize(file_path)
            msg = f"成功读取文本文件（编码: {encoding}，{file_size} 字节）"
            if truncated:
                msg += f"（已截断至 {max_chars} 字符）"

            return {
                "success": True, "file_path": file_path, "content": content,
                "encoding": encoding, "file_size": file_size,
                "truncated": truncated, "message": msg,
            }
        except UnicodeDecodeError:
            continue
        except Exception as e:
            return {
                "success": False, "file_path": file_path, "content": "",
                "encoding": "", "file_size": 0,
                "truncated": False, "message": f"读取 TXT 文件失败: {e}",
            }

    return {
        "success": False, "file_path": file_path, "content": "",
        "encoding": "", "file_size": 0,
        "truncated": False,
        "message": "无法解码文件，尝试了 utf-8/gbk/gb2312/latin-1/utf-16 均失败",
    }
