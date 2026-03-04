#!/usr/bin/env python3
"""
MCP server for md2docx.

Exposes two tools:
  - convert_markdown_to_docx  : convert Markdown text → DOCX
  - convert_md_file_to_docx   : convert a .md file on disk → DOCX

Run via stdio (default) — compatible with Claude Code and Cline.
"""

import subprocess
import sys
import tempfile
from pathlib import Path

from mcp.server.fastmcp import FastMCP

mcp = FastMCP("md2docx")

_SCRIPT     = str(Path(__file__).parent / "md2docx.py")
_DEFAULT_STYLE = str(Path(__file__).parent / "style_default.json")


@mcp.tool()
def convert_markdown_to_docx(
    markdown_text: str,
    output_path: str,
    style_path: str = _DEFAULT_STYLE,
) -> str:
    """Convert Markdown text to a DOCX file.

    Args:
        markdown_text: Markdown content to convert.
        output_path: Absolute path where the .docx file will be saved.
        style_path: Path to a JSON style guide. Defaults to style_default.json.

    Returns:
        Confirmation message with the saved file path.
    """
    tmp = tempfile.NamedTemporaryFile(
        mode="w", suffix=".md", encoding="utf-8", delete=False
    )
    try:
        tmp.write(markdown_text)
        tmp.close()
        result = subprocess.run(
            [sys.executable, _SCRIPT, tmp.name, output_path, "--style", style_path],
            capture_output=True,
            text=True,
        )
        if result.returncode != 0:
            raise RuntimeError(result.stderr or result.stdout)
        return f"Saved → {output_path}"
    finally:
        Path(tmp.name).unlink(missing_ok=True)


@mcp.tool()
def convert_md_file_to_docx(
    input_path: str,
    output_path: str,
    style_path: str = _DEFAULT_STYLE,
) -> str:
    """Convert a Markdown file on disk to a DOCX file.

    Args:
        input_path: Absolute path to the input .md file.
        output_path: Absolute path where the .docx file will be saved.
        style_path: Path to a JSON style guide. Defaults to style_default.json.

    Returns:
        Confirmation message with the saved file path.
    """
    result = subprocess.run(
        [sys.executable, _SCRIPT, input_path, output_path, "--style", style_path],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        raise RuntimeError(result.stderr or result.stdout)
    return f"Saved → {output_path}"


if __name__ == "__main__":
    mcp.run()
