# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
# Install dependencies
pip install -r requirements.txt

# Basic conversion (uses style_default.json by default)
python md2docx.py input.md output.docx

# With a custom style guide
python md2docx.py input.md output.docx --style style_minipass.json
```

## Architecture

The entire converter is a single file: `md2docx.py`.

**Pipeline:** `mistune AST → token tree → python-docx Document`

1. **Style loading** (`load_style`, `build_style_constants`): reads a JSON style guide and overwrites module-level globals (`FONT_NAME`, `HEADING_STYLES`, `NORMAL`, `TABLE_*`, etc.). All styling flows through these globals.

2. **Document setup** (`setup_document`): creates a `python-docx` `Document`, applies margins and line spacing from globals, sets the Normal style font/size.

3. **Block dispatcher** (`render_block`): routes each mistune token by `type` to a dedicated `render_*` function (`render_heading`, `render_paragraph`, `render_blockquote`, `render_list`, `render_table`, `render_code_block`, `render_hr`).

4. **Inline renderer** (`render_inline`): recursively walks mistune inline AST nodes (text, strong, emphasis, codespan, link, image, etc.) and calls `doc.Paragraph.add_run()` with appropriate font settings via `_apply_run`.

**mistune usage:** `mistune.create_markdown(renderer='ast', plugins=['table'])` — produces a raw token list, not HTML. The `table` plugin is required for table support.

**Key discrepancy:** `main()` defaults to `style_default.json`, but only `style_minipass.json` is bundled. Pass `--style style_minipass.json` explicitly if that file doesn't exist. The cover page and footer keys in `style_minipass.json` are parsed by `build_style_constants` but **not rendered** — `md2docx.py` strips those features (they existed in an older `md_to_docx.py` version).

## Style Guide

Style guides are JSON files. `style_minipass.json` is the reference example. All color values are `"#RRGGBB"` hex strings. The `cover` and `footer` top-level keys are present in the JSON schema but ignored by the current script.
