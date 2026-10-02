"""Regression test for Chinese JSON over the MCP stdio transport."""

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

import anyio
from mcp import ClientSession
from mcp.client.stdio import StdioServerParameters, stdio_client


class Utf8StdioTest(unittest.TestCase):
    def test_chinese_json_round_trip(self) -> None:
        source_text = "中文 JSON：北京，测试 ✅"

        async def check() -> None:
            with tempfile.TemporaryDirectory() as temp_dir:
                probe = Path(temp_dir) / "中文输入.txt"
                probe.write_text(source_text, encoding="utf-8")
                params = StdioServerParameters(
                    command=sys.executable,
                    args=["-m", "mcp_office_writer"],
                    cwd=Path(__file__).resolve().parents[1],
                    encoding="utf-8",
                )
                async with stdio_client(params) as (read_stream, write_stream):
                    async with ClientSession(read_stream, write_stream) as session:
                        await session.initialize()
                        result = await session.call_tool(
                            "read_txt_document", {"file_path": str(probe)}
                        )

            response = result.content[0].text
            self.assertIn(source_text, response)
            self.assertNotIn("\\u4e2d", response)

        anyio.run(check)


if __name__ == "__main__":
    unittest.main()
