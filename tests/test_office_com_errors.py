from __future__ import annotations

import unittest
from unittest.mock import patch

from mcp_office_writer.models import EditResult, SectionItem
from mcp_office_writer.office_com import (
    get_open_office_documents,
    try_edit_open_excel_workbook,
    try_edit_open_powerpoint,
    try_edit_open_word_document,
)


def _com_unavailable():
    return False


def _com_available():
    return True


class OfficeComErrorBranchesTest(unittest.TestCase):
    def test_get_open_returns_empty_when_com_unavailable(self) -> None:
        with patch("mcp_office_writer.office_com.com_is_available", _com_unavailable):
            self.assertEqual(get_open_office_documents(), {"word": [], "excel": [], "powerpoint": []})

    def test_get_open_captures_com_error_without_raising(self) -> None:
        def raiser(*args, **kwargs):
            raise RuntimeError("应用未响应")

        with patch("mcp_office_writer.office_com.com_is_available", _com_available):
            with patch("mcp_office_writer.office_com._com_apartment", raiser):
                result = get_open_office_documents()
        self.assertEqual(result["word"], [])
        self.assertIn("error", result)

    def test_edit_open_word_returns_error_result_when_com_lookup_fails(self) -> None:
        def raiser(*args, **kwargs):
            raise RuntimeError("COM 初始化失败")

        with patch("mcp_office_writer.office_com.com_is_available", _com_available):
            with patch("mcp_office_writer.office_com._com_apartment", raiser):
                result = try_edit_open_word_document(
                    r"D:\不存在.docx", [SectionItem(type="paragraph", text="x")]
                )
        self.assertIsInstance(result, EditResult)
        self.assertFalse(result.success)
        self.assertIn("COM", result.message)

    def test_edit_open_excel_and_ppt_return_error_when_com_lookup_fails(self) -> None:
        def raiser(*args, **kwargs):
            raise RuntimeError("COM 初始化失败")

        with patch("mcp_office_writer.office_com.com_is_available", _com_available):
            with patch("mcp_office_writer.office_com._com_apartment", raiser):
                excel_result = try_edit_open_excel_workbook(r"D:\不存在.xlsx", [])
                ppt_result = try_edit_open_powerpoint(r"D:\不存在.pptx", [])
        self.assertFalse(excel_result.success)
        self.assertFalse(ppt_result.success)


if __name__ == "__main__":
    unittest.main()
