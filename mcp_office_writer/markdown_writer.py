"""Markdown file to Word conversion without sending document text through MCP JSON."""

from __future__ import annotations

import re
from pathlib import Path

from .models import CreateResult, ImageDef, SectionItem, TableDef, WordContent
from .utils import validate_input_file

_HEADING = re.compile(r"^(#{1,6})\s+(.+?)\s*$")
_BULLET = re.compile(r"^\s*[-*+]\s+(.+?)\s*$")
_NUMBERED = re.compile(r"^\s*\d+[.)]\s+(.+?)\s*$")
_HORIZONTAL_RULE = re.compile(r"^\s*(?:---+|\*\*\*+|___+)\s*$")
_TABLE_DIVIDER = re.compile(r"^\s*\|?(?:\s*:?-{3,}:?\s*\|)+\s*$")
_IMAGE = re.compile(r'^\s*!\[([^\]]*)\]\(([^\s)]+)(?:\s+["\'][^"\']*["\'])?\)\s*$')
_SUPPORTED_IMAGE_SUFFIXES = (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff")


def _read_markdown(source_path: str) -> str:
    validate_input_file(source_path, (".md", ".markdown", ".txt"))
    data = Path(source_path).read_bytes()
    # UTF-16 必须在 GB18030 之前尝试；后者可以把带 NUL 的 UTF-16 字节
    # 错当成可解码文本，从而产生静默乱码。
    for encoding in ("utf-8-sig", "utf-8", "utf-16", "gb18030"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise ValueError("无法解码 Markdown 文件；仅支持 UTF-8、GB18030 或 UTF-16")


def _plain_text(value: str) -> str:
    """Keep text while removing the Markdown constructs this converter supports."""
    value = re.sub(r"!\[([^\]]*)\]\([^)]*\)", r"\1", value)
    value = re.sub(r"\[([^\]]+)\]\([^)]*\)", r"\1", value)
    value = re.sub(r"\\text(?:bf|it)\{([^{}]*)\}", r"\1", value)
    value = re.sub(r"(```|`|\*\*|__|~~)", "", value)
    return value.strip()


def _table_cells(line: str) -> list[str]:
    return [_plain_text(cell) for cell in line.strip().strip("|").split("|")]


def markdown_to_word_content(markdown: str, source_dir: Path | None = None) -> WordContent:
    """Convert common report Markdown structures into the existing Word schema."""
    sections: list[SectionItem] = []
    paragraph_lines: list[str] = []
    list_items: list[str] = []
    list_type: str | None = None
    table_lines: list[str] = []
    code_lines: list[str] = []
    in_code_block = False

    def flush_paragraph() -> None:
        nonlocal paragraph_lines
        text = " ".join(_plain_text(line) for line in paragraph_lines if _plain_text(line))
        if text:
            sections.append(SectionItem(type="paragraph", text=text))
        paragraph_lines = []

    def flush_list() -> None:
        nonlocal list_items, list_type
        if list_items and list_type:
            sections.append(SectionItem(type=list_type, items=list_items))
        list_items, list_type = [], None

    def flush_table() -> None:
        nonlocal table_lines
        rows = [_table_cells(line) for line in table_lines if not _TABLE_DIVIDER.match(line)]
        if rows:
            sections.append(SectionItem(
                type="table", table=TableDef(headers=rows[0], rows=rows[1:])
            ))
        table_lines = []

    def flush_code() -> None:
        nonlocal code_lines
        if code_lines:
            sections.append(SectionItem(type="code", text="\n".join(code_lines)))
        code_lines = []

    for raw_line in markdown.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        line = raw_line.rstrip()
        if line.startswith("```"):
            flush_paragraph()
            flush_list()
            flush_table()
            if in_code_block:
                flush_code()
            in_code_block = not in_code_block
            continue
        if in_code_block:
            code_lines.append(raw_line)
            continue
        if line.lstrip().startswith("|"):
            flush_paragraph()
            flush_list()
            table_lines.append(line)
            continue
        if table_lines:
            flush_table()
        image = _IMAGE.match(line)
        if image:
            flush_paragraph()
            flush_list()
            alt_text, image_target = image.groups()
            image_path = Path(image_target)
            if source_dir and not image_path.is_absolute():
                image_path = source_dir / image_path
            try:
                validate_input_file(str(image_path), _SUPPORTED_IMAGE_SUFFIXES)
            except ValueError:
                label = alt_text or image_target
                sections.append(SectionItem(type="paragraph", text=f"[图片未嵌入: {label}]"))
            else:
                sections.append(SectionItem(type="image", image=ImageDef(path=str(image_path))))
            continue
        heading = _HEADING.match(line)
        if heading:
            flush_paragraph()
            flush_list()
            sections.append(SectionItem(
                type="heading", level=min(len(heading.group(1)), 6), text=_plain_text(heading.group(2))
            ))
            continue
        if _HORIZONTAL_RULE.match(line):
            flush_paragraph()
            flush_list()
            continue
        quote = line.lstrip()
        if quote.startswith(">"):
            line = quote[1:].lstrip()
        bullet = _BULLET.match(line)
        numbered = _NUMBERED.match(line)
        if bullet or numbered:
            item_type = "bullet_list" if bullet else "numbered_list"
            if list_type and list_type != item_type:
                flush_list()
            flush_paragraph()
            list_type = item_type
            list_items.append(_plain_text((bullet or numbered).group(1)))
            continue
        if not line.strip():
            flush_paragraph()
            flush_list()
            continue
        flush_list()
        paragraph_lines.append(line)

    if in_code_block:
        flush_code()
    flush_table()
    flush_paragraph()
    flush_list()
    return WordContent(sections=sections)


def create_word_from_markdown(
    source_path: str,
    output_path: str,
    template: str | None = None,
    scene: str | None = "report",
    if_exists: str = "overwrite",
) -> CreateResult:
    """Create a Word document directly from a Markdown file on disk."""
    from .word_writer import create_word_document

    try:
        content = markdown_to_word_content(
            _read_markdown(source_path), source_dir=Path(source_path).parent
        )
        result = create_word_document(
            output_path=output_path,
            content=content,
            template=template,
            scene=scene,
            if_exists=if_exists,
        )
    except Exception as exc:
        return CreateResult(success=False, file_path=output_path, message=f"Markdown 转 Word 失败: {exc}")
    if result.success:
        result.message = f"Markdown 转 Word 成功，共转换 {len(content.sections)} 个内容块：{result.file_path}"
    return result
