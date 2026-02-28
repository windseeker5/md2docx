#!/usr/bin/env python3
"""
MCP server for md2docx.

Exposes two tools:
  - convert_markdown_to_docx  : convert Markdown text → DOCX
  - convert_md_file_to_docx   : convert a .md file on disk → DOCX

Run via stdio (default) — compatible with Claude Code and Cline.
"""

import sys
from pathlib import Path

# Ensure md2docx.py is importable regardless of working directory
sys.path.insert(0, str(Path(__file__).parent))

import mistune
from mcp.server.fastmcp import FastMCP
import md2docx as converter

mcp = FastMCP("md2docx")

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
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)

    cfg = converter.load_style(style_path)
    converter.build_style_constants(cfg)

    tokens = mistune.create_markdown(renderer="ast", plugins=["table"])(markdown_text)
    doc = converter.setup_document()
    for token in tokens:
        converter.render_block(doc, token)

    doc.save(str(out))
    return f"Saved → {out}"


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
    md_text = Path(input_path).read_text(encoding="utf-8")
    return convert_markdown_to_docx(md_text, output_path, style_path)


if __name__ == "__main__":
    mcp.run()
