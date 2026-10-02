"""
mcp-office-writer 工具函数
"""
import os
import shutil
import tempfile
import logging
import zipfile
from collections.abc import Callable
from pathlib import Path

logger = logging.getLogger(__name__)

MAX_INPUT_FILE_BYTES = 100 * 1024 * 1024
MAX_OFFICE_UNCOMPRESSED_BYTES = 500 * 1024 * 1024

# ── 国内标准字号 → pt 映射 ────────────────────────────────────────
# 参考 GB/T 9704-2012 及中文排版惯例
CN_FONT_SIZE_MAP: dict[str, float] = {
    "初号": 42.0,
    "小初": 36.0,
    "一号": 26.0,
    "小一": 24.0,
    "二号": 22.0,
    "小二": 18.0,
    "三号": 16.0,
    "小三": 15.0,
    "四号": 14.0,
    "小四": 12.0,
    "五号": 10.5,
    "小五": 9.0,
    "六号": 7.5,
    "小六": 6.5,
    "七号": 5.5,
    "八号": 5.0,
}


def cn_size_to_pt(size: str | int | float) -> float:
    """将国内字号名转为 pt 值。数字则直接返回。"""
    if isinstance(size, (int, float)):
        return float(size)
    if size in CN_FONT_SIZE_MAP:
        return CN_FONT_SIZE_MAP[size]
    raise ValueError(f"未知的中文字号: {size}，可选值: {', '.join(CN_FONT_SIZE_MAP)}")


def ensure_output_dir(file_path: str) -> None:
    """确保输出目录存在"""
    parent = Path(file_path).parent
    if not parent.exists():
        parent.mkdir(parents=True, exist_ok=True)


def validate_output_path(file_path: str, allowed_suffixes: tuple[str, ...] | None = None) -> None:
    """验证输出路径是否合法

    Raises:
        ValueError: 路径无效
    """
    if not file_path or not file_path.strip():
        raise ValueError("输出路径不能为空")
    try:
        p = Path(file_path)
        if any(c in str(p) for c in '<>"|?*'):
            raise ValueError(f"路径包含非法字符: {file_path}")
        if allowed_suffixes and p.suffix.lower() not in allowed_suffixes:
            formats = "、".join(allowed_suffixes)
            raise ValueError(f"输出文件必须使用以下扩展名之一: {formats}")
    except (OSError, ValueError) as e:
        raise ValueError(f"路径无效: {file_path}") from e


def validate_input_file(file_path: str, allowed_suffixes: tuple[str, ...]) -> None:
    """Validate an Office/PDF input before a third-party parser opens it."""
    path = Path(file_path)
    if not path.is_file():
        raise ValueError(f"文件不存在: {file_path}")
    if allowed_suffixes and path.suffix.lower() not in allowed_suffixes:
        formats = "、".join(allowed_suffixes)
        raise ValueError(f"不支持的文件格式，仅支持: {formats}")
    if path.stat().st_size > MAX_INPUT_FILE_BYTES:
        raise ValueError(f"文件超过 {MAX_INPUT_FILE_BYTES // 1024 // 1024} MB 安全读取上限")
    if path.suffix.lower() in {
        ".docx", ".dotx", ".xlsx", ".xlsm", ".xltx", ".xltm", ".pptx", ".potx",
    }:
        try:
            with zipfile.ZipFile(path) as archive:
                if sum(item.file_size for item in archive.infolist()) > MAX_OFFICE_UNCOMPRESSED_BYTES:
                    raise ValueError("Office 文件解压后超过 500 MB 安全上限")
        except zipfile.BadZipFile as e:
            raise ValueError("Office 文件损坏或不是有效的 OOXML 压缩包") from e


def atomic_save(save: Callable[[str], None], file_path: str) -> None:
    """Save to a sibling temporary file, then atomically replace the target."""
    target = Path(file_path)
    fd, temp_path = tempfile.mkstemp(
        prefix=f".{target.stem}_", suffix=target.suffix, dir=target.parent
    )
    os.close(fd)
    try:
        save(temp_path)
        os.replace(temp_path, target)
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)


def _backup_file(file_path: str) -> str | None:
    """编辑前创建临时备份，返回备份路径；若文件不存在返回 None"""
    if not os.path.isfile(file_path):
        return None
    try:
        fd, backup_path = tempfile.mkstemp(
            suffix=Path(file_path).suffix + ".bak",
            prefix=Path(file_path).stem + "_",
            dir=Path(file_path).parent,
        )
        os.close(fd)
        shutil.copy2(file_path, backup_path)
        logger.debug("已创建备份: %s → %s", file_path, backup_path)
        return backup_path
    except OSError as e:
        logger.warning("创建备份失败: %s", e)
        return None


def resolve_if_exists(file_path: str, mode: str) -> str:
    """处理文件已存在的情况

    Args:
        file_path: 目标路径
        mode: "overwrite" | "error" | "rename"

    Returns:
        实际写入的路径
    """
    path = Path(file_path)
    if not path.exists():
        return file_path

    if mode == "overwrite":
        return file_path
    elif mode == "error":
        raise FileExistsError(f"文件已存在: {file_path}")
    elif mode == "rename":
        stem = path.stem
        suffix = path.suffix
        parent = path.parent
        counter = 1
        while True:
            new_name = f"{stem} ({counter}){suffix}"
            new_path = parent / new_name
            if not new_path.exists():
                return str(new_path)
            counter += 1
    else:
        raise ValueError(f"不支持的 if_exists 值: {mode}")


def get_file_size(file_path: str) -> int:
    """获取文件大小（字节）"""
    return os.path.getsize(file_path)


LAYOUT_MAP = {
    "title": 0,
    "text": 1,       # alias for text + content
    "section": 2,
    "two_content": 3,
    "comparison": 4,
    "title_only": 5,
    "blank": 6,
    "content_caption": 7,
    "picture_caption": 8,
}


def get_ppt_layout_index(layout_name: str) -> int:
    """将布局名称映射为 python-pptx 的布局索引"""
    idx = LAYOUT_MAP.get(layout_name, 1)
    return idx


def hex_to_rgb(hex_color: str) -> tuple[int, int, int]:
    """将 HEX 颜色字符串转为 RGB 元组"""
    hex_color = hex_color.lstrip("#")
    if len(hex_color) == 6:
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        return (r, g, b)
    return (0, 0, 0)


# ── 图片下载与缓存 ────────────────────────────────────────────────

_IMAGE_CACHE_DIR: str | None = None
_MAX_DOWNLOAD_BYTES = 20 * 1024 * 1024  # 单张图片下载上限 20 MB


def _get_image_cache_dir() -> str:
    """惰性初始化图片缓存目录。"""
    global _IMAGE_CACHE_DIR
    if _IMAGE_CACHE_DIR is None:
        base = Path(__file__).resolve().parent.parent
        cache = base / "image_cache"
        cache.mkdir(exist_ok=True)
        _IMAGE_CACHE_DIR = str(cache)
    return _IMAGE_CACHE_DIR


def download_image(url: str) -> str:
    """下载图片到缓存目录，返回本地路径。同 URL 自动复用缓存。

    仅允许 http/https，且带 20 MB 大小上限，防止 file:// 读取任意本地文件
    或超大响应占满内存/磁盘。
    """
    import hashlib
    import urllib.request
    from urllib.parse import urlparse

    parsed = urlparse(url)
    if parsed.scheme not in ("http", "https"):
        raise ValueError(f"不支持的图片 URL scheme: {parsed.scheme!r}（仅支持 http/https）")

    # 猜测扩展名
    path_part = parsed.path.lower()
    ext = ".png"
    for known in (".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp"):
        if path_part.endswith(known):
            ext = known
            break

    cache_key = hashlib.md5(url.encode()).hexdigest() + ext
    cache_path = os.path.join(_get_image_cache_dir(), cache_key)
    if os.path.exists(cache_path) and os.path.getsize(cache_path) > 0:
        return cache_path

    req = urllib.request.Request(url, headers={"User-Agent": "mcp-office-writer/1.0"})
    with urllib.request.urlopen(req, timeout=30) as resp:
        remaining = _MAX_DOWNLOAD_BYTES + 1
        chunks = []
        total = 0
        while remaining > 0:
            chunk = resp.read(min(65536, remaining))
            if not chunk:
                break
            chunks.append(chunk)
            total += len(chunk)
            remaining -= len(chunk)
    if total > _MAX_DOWNLOAD_BYTES:
        raise ValueError(f"图片超过 {_MAX_DOWNLOAD_BYTES // 1024 // 1024} MB 下载上限: {url}")
    data = b"".join(chunks)
    if not data:
        raise ValueError(f"图片下载结果为空: {url}")
    with open(cache_path, "wb") as f:
        f.write(data)
    return cache_path


def hex_to_ppt_color(hex_color: str) -> str:
    """将 HEX 颜色转为 python-pptx 所需的 RRGGBB 格式（无 #）"""
    return hex_color.lstrip("#")[:6]
