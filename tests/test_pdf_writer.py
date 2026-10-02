from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import pymupdf

from mcp_office_writer.models import PdfContent, PdfSectionItem
from mcp_office_writer.pdf_writer import create_pdf_document
from mcp_office_writer.reader import read_pdf_document


class PdfWriterTest(unittest.TestCase):
    def test_create_and_read_pdf_round_trip_with_chinese(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "报告.pdf"
            content = PdfContent(sections=[
                PdfSectionItem(type="heading", text="第一章 引言", level=1),
                PdfSectionItem(type="paragraph", text="这是一段中文内容，用于验证 PDF 生成与读取回环。"),
                PdfSectionItem(type="bullet_list", items=["第一点", "第二点"]),
            ])
            result = create_pdf_document(str(output), content)
            self.assertTrue(result.success, msg=result.message)
            self.assertTrue(output.exists())

            read_result = read_pdf_document(str(output))
            self.assertTrue(read_result["success"], msg=read_result["message"])
            text = read_result["content"]
            self.assertIn("第一章 引言", text)
            self.assertIn("中文内容", text)
            self.assertIn("第一点", text)

    def test_creates_pdf_pages(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "多页.pdf"
            content = PdfContent(sections=[
                PdfSectionItem(type="paragraph", text="第 1 页内容"),
                PdfSectionItem(type="page_break"),
                PdfSectionItem(type="paragraph", text="第 2 页内容"),
            ])
            result = create_pdf_document(str(output), content)
            self.assertTrue(result.success, msg=result.message)
            with pymupdf.open(str(output)) as doc:
                self.assertGreaterEqual(doc.page_count, 2)


if __name__ == "__main__":
    unittest.main()
