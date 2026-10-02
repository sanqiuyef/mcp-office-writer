"""基于 DeepSeek dsh-libreoffice-kit（dsoffice CLI）的高保真 Office 转换。

dsoffice 随 npm 包 @deepseek-ai/libreoffice-kit 内置精简版 LibreOffice 引擎，
无需本机安装 LibreOffice 或 MS Office 即可工作，字体缺失时按字形度量替换并报告：

- convert:     docx/xlsx/pptx（含 97-2003 的 .doc/.xls/.ppt）→ PDF
- render:      Office/PDF → 逐页 PNG（可指定页码、工作表区域与 dpi）
- recalculate: 重算 xlsx 公式并写回缓存值（弥补 openpyxl 只写公式不存值）

CLI 输出为单行 JSON（失败时 {"code":"failed","error":...}，退出码非 0）。
dsoffice 拒绝覆盖已存在的输出文件/目录，本模块按 if_exists 语义先行处理。

环境要求：Node.js >= 22.19、`npm install -g @deepseek-ai/libreoffice-kit`；
Windows x64 另需 MSVC v14 x64 运行库。可用环境变量 MCP_OFFICE_DSOFFICE 指定 CLI 路径。
"""
from __future__ import annotations

import json
import logging
import os
import re
import shutil
import subprocess
from pathlib import Path

from .utils import (
    ensure_output_dir,
    resolve_if_exists,
    validate_input_file,
    validate_output_path,
)

logger = logging.getLogger(__name__)

# dsoffice convert/render 的输入格式（render 另支持直接渲染 PDF）
CONVERT_INPUT_SUFFIXES = (".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt", ".rtf")
RENDER_INPUT_SUFFIXES = CONVERT_INPUT_SUFFIXES + (".pdf",)
RECALC_INPUT_SUFFIXES = (".xlsx", ".ods")

DSOFFICE_ENV_VAR = "MCP_OFFICE_DSOFFICE"
DEFAULT_DPI = 144
MIN_DPI, MAX_DPI = 72, 600

DSOFFICE_INSTALL_HINT = (
    "未找到 dsoffice CLI。请先安装：npm install -g @deepseek-ai/libreoffice-kit"
    "（要求 Node.js >= 22.19；Windows x64 需 MSVC v14 运行库），"
    f"或通过环境变量 {DSOFFICE_ENV_VAR} 指定 dsoffice 可执行文件路径"
)


def find_dsoffice() -> str:
    """定位 dsoffice 可执行文件，找不到时抛出带安装指引的 RuntimeError。"""
    custom = os.environ.get(DSOFFICE_ENV_VAR)
    if custom:
        if Path(custom).is_file():
            return custom
        raise RuntimeError(f"{DSOFFICE_ENV_VAR} 指定的路径不存在: {custom}")
    resolved = shutil.which("dsoffice")
    if resolved:
        return resolved
    raise RuntimeError(DSOFFICE_INSTALL_HINT)


def _run_dsoffice(cli_path: str, args: list[str], timeout: float) -> dict:
    """执行 dsoffice 子命令并解析其单行 JSON 输出。

    Returns:
        CLI 返回的 JSON dict（成功时含 outputPath/outputDir 等字段）。

    Raises:
        RuntimeError: 进程失败、超时或输出不是合法 JSON。
    """
    cmd = [cli_path, *args]
    try:
        proc = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="backslashreplace",
            timeout=timeout,
            check=False,
        )
    except subprocess.TimeoutExpired as e:
        raise RuntimeError(f"dsoffice 执行超时（>{timeout}s）: {' '.join(cmd)}") from e
    except OSError as e:
        raise RuntimeError(f"dsoffice 无法启动: {e}") from e

    stdout = (proc.stdout or "").strip()
    try:
        payload = json.loads(stdout) if stdout else None
    except ValueError as e:
        stderr = (proc.stderr or "").strip()
        raise RuntimeError(
            f"dsoffice 输出无法解析（exit={proc.returncode}）: {stdout[:300] or stderr[:300]}"
        ) from e
    if not isinstance(payload, dict):
        # 刻意转成 RuntimeError：调用方只关心"CLI 输出异常"一种失败形态
        raise RuntimeError(  # noqa: TRY004
            f"dsoffice 输出不是 JSON 对象（exit={proc.returncode}）: {stdout[:300]}"
        )

    if proc.returncode != 0 or payload.get("code") == "failed":
        raise RuntimeError(f"dsoffice 转换失败: {payload.get('error') or payload}")
    return payload


def _delete_existing(path: Path, kind: str) -> None:
    """删除已存在的输出文件/目录（dsoffice 要求目标不存在）。"""
    if kind == "dir" and path.is_dir():
        shutil.rmtree(path)
    elif path.exists() or path.is_symlink():
        path.unlink()


def _prepare_target(path: Path, if_exists: str, kind: str = "file") -> Path:
    """按 if_exists 语义处理已存在的输出目标，返回实际应使用的目标路径。

    dsoffice 要求输出目标不存在：overwrite 删除已有目标后原路返回；error 抛
    FileExistsError；rename 返回一个必然不存在的新路径（如 "报告 (1).pdf"）。
    """
    resolved = resolve_if_exists(str(path), if_exists)
    target = Path(resolved)
    if target.exists() or target.is_symlink():
        _delete_existing(target, kind)
    return target


def _validate_pages(pages: str) -> str:
    """校验 --pages 参数：'all' 或 '1,3' 样式的页码列表。"""
    pages = (pages or "all").strip()
    if pages.lower() == "all":
        return "all"
    if not re.fullmatch(r"\d+(?:\s*,\s*\d+)*", pages):
        raise ValueError(f"pages 参数无效: {pages!r}，应为 'all' 或 '1,3,5' 样式")
    return ",".join(p.strip() for p in pages.split(","))


def convert_to_pdf(
    input_path: str,
    output_path: str | None = None,
    if_exists: str = "overwrite",
    sheet: str | None = None,
) -> dict:
    """将 Office 文档高保真转换为 PDF。

    支持 .docx/.xlsx/.pptx 及 97-2003 的 .doc/.xls/.ppt、.rtf。
    Excel 可用 sheet 指定单个工作表（默认全部）。
    返回 success/output_path/backend/missing_fonts（缺失字体按字形度量替换，不影响转换）。
    """
    validate_input_file(input_path, CONVERT_INPUT_SUFFIXES)
    src = Path(input_path)
    requested = Path(output_path) if output_path else src.with_suffix(".pdf")
    validate_output_path(str(requested), (".pdf",))
    target = _prepare_target(requested, if_exists, "file")
    ensure_output_dir(str(target))

    args = ["convert", "--input", str(src.resolve()), "--output", str(target.resolve())]
    if sheet:
        args += ["--sheet", sheet]
    payload = _run_dsoffice(find_dsoffice(), args, timeout=300)

    return {
        "success": True,
        "file_path": str(target),
        "backend": payload.get("backend"),
        "missing_fonts": payload.get("missingFonts") or [],
        "message": "转换成功" + (f"；缺失字体（已按字形度量替换）: {', '.join(payload.get('missingFonts') or [])}" if payload.get("missingFonts") else ""),
    }


def render_document_pages(
    input_path: str,
    output_dir: str | None = None,
    pages: str = "all",
    dpi: int = DEFAULT_DPI,
    sheet: str | None = None,
    cell_range: str | None = None,
    if_exists: str = "overwrite",
) -> dict:
    """将 Office/PDF 文档逐页渲染为 PNG 图片。

    支持 convert_to_pdf 的全部输入格式及 .pdf。pages 为 'all' 或 '1,3,5'；
    Excel 可用 sheet + cell_range（如 A1:D20）只渲染指定区域。
    dpi 默认 144（72-600）。输出目录必须为空/不存在：overwrite 先清空重建，error 直接报错。
    返回 success/output_dir/page_count/images（含每页路径与尺寸）。
    """
    validate_input_file(input_path, RENDER_INPUT_SUFFIXES)
    src = Path(input_path)
    requested = Path(output_dir) if output_dir else src.parent / f"{src.stem}_pages"
    target = _prepare_target(requested, if_exists, "dir")
    # dsoffice 要求输出目录不存在（fresh directory），这里只保证父目录可用
    target.parent.mkdir(parents=True, exist_ok=True)

    args = [
        "render",
        "--input", str(src.resolve()),
        "--output-dir", str(target.resolve()),
        "--pages", _validate_pages(pages),
        "--dpi", str(max(MIN_DPI, min(MAX_DPI, int(dpi)))),
    ]
    if sheet:
        args += ["--sheet", sheet]
    if cell_range:
        if not sheet:
            raise ValueError("cell_range 需要与 sheet 同时指定")
        args += ["--range", cell_range]
    payload = _run_dsoffice(find_dsoffice(), args, timeout=300)

    return {
        "success": True,
        "output_dir": str(target),
        "page_count": payload.get("pageCount"),
        "images": payload.get("images") or [],
        "missing_fonts": payload.get("missingFonts") or [],
        "message": f"渲染完成，共 {payload.get('pageCount')} 页",
    }


def recalculate_excel_formulas(
    input_path: str,
    output_path: str | None = None,
    if_exists: str = "overwrite",
) -> dict:
    """用 LibreOffice 引擎重算 xlsx 公式并把结果写为缓存值。

    openpyxl 写入的公式没有缓存值，多数读取工具（含 read_excel_workbook 的 data_only 模式）
    看到的都是空；转换后的文件公式与数值并存。默认输出 <原名>_recalc.xlsx，也可指定 .ods。
    """
    validate_input_file(input_path, RECALC_INPUT_SUFFIXES)
    src = Path(input_path)
    default_target = src.with_name(f"{src.stem}_recalc{src.suffix}")
    requested = Path(output_path) if output_path else default_target
    validate_output_path(str(requested), (".xlsx", ".ods"))
    target = _prepare_target(requested, if_exists, "file")
    ensure_output_dir(str(target))

    payload = _run_dsoffice(
        find_dsoffice(),
        ["recalculate", "--input", str(src.resolve()), "--output", str(target.resolve())],
        timeout=300,
    )

    return {
        "success": True,
        "file_path": str(target),
        "backend": payload.get("backend"),
        "missing_fonts": payload.get("missingFonts") or [],
        "message": "公式重算完成，缓存值已写入",
    }
