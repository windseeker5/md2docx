# md2docx

Convert any Markdown file to a polished `.docx` — available as an **MCP tool** for Claude Code and Cline (VS Code), and as a standalone CLI.

---

## MCP Setup (Claude Code & Cline)

### Step 1 — Clone and install

```bash
git clone git@github.com:windseeker5/md2docx.git
cd md2docx
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

Note the full path to your `.venv` Python — you'll need it in the next step.

```bash
which python   # e.g. /home/you/Documents/DEV/md2docx/.venv/bin/python
```

---

### Step 2 — Claude Code

Run this one command (replace the path with wherever you cloned the repo):

```bash
claude mcp add --scope user md2docx \
  /path/to/md2docx/.venv/bin/python \
  /path/to/md2docx/server.py
```

Then restart Claude Code. Verify with `/mcp` — you should see **md2docx** listed.

---

### Step 3 — Cline (VS Code)

Open the Cline panel → **MCP Servers** → **Edit MCP Settings**, and add:

```json
{
  "md2docx": {
    "command": "/home/you/Documents/DEV/md2docx/.venv/bin/python",
    "args": ["/home/you/Documents/DEV/md2docx/server.py"],
    "disabled": false,
    "autoApprove": []
  }
}
```

Save and restart VS Code. The **md2docx** server will appear in Cline's MCP panel.

---

## Available MCP Tools

| Tool | Input | Description |
|------|-------|-------------|
| `convert_markdown_to_docx` | `markdown_text`, `output_path`, `style_path`* | Convert Markdown text to DOCX |
| `convert_md_file_to_docx` | `input_path`, `output_path`, `style_path`* | Convert a `.md` file on disk to DOCX |

`*` optional — defaults to `style_default.json`

**Example prompt to Claude:**
> Convert this markdown to a docx and save it to `/home/me/docs/report.docx`

---

## CLI Usage

```bash
# Default style (style_default.json — neutral Calibri)
python md2docx.py input.md output.docx

# Custom style
python md2docx.py input.md output.docx --style style_minipass.json
```

```
positional arguments:
  input          Input Markdown file (.md)
  output         Output Word document (.docx)

optional arguments:
  --style FILE, -s FILE
                 JSON style guide (default: style_default.json)
```

---

## Style Guides

| File | Description |
|------|-------------|
| `style_default.json` | Neutral Calibri — clean, professional, no cover page |
Copy `style_default.json`, rename it, and edit to create your own style. All colours are `"#RRGGBB"` hex strings.

### Style Guide Reference

#### `document`

| Key | Type | Description |
|-----|------|-------------|
| `page_size` | string | `"letter"` or `"A4"` |
| `line_spacing` | float | Line spacing multiplier (e.g. `1.15`) |
| `margins.top_cm` | float | Top margin in centimetres |
| `margins.bottom_cm` | float | Bottom margin in centimetres |
| `margins.left_cm` | float | Left margin in centimetres |
| `margins.right_cm` | float | Right margin in centimetres |

#### `fonts`

| Key | Description |
|-----|-------------|
| `body` | Font for body text, H2–H4, lists, tables |
| `h1_override` | Font for H1 only (`""` to use `body`) |

#### `headings` — one entry per level: `h1`, `h2`, `h3`, `h4`

| Key | Type | Description |
|-----|------|-------------|
| `size` | int | Font size in points |
| `color` | string | Text color |
| `bold` | bool | Bold weight |
| `italic` | bool | Italic style |

#### `body` / `blockquote` / `code_inline` / `code_block`

| Key | Description |
|-----|-------------|
| `size` | Font size in points |
| `color` | Text color |
| `border_color` | (`blockquote` only) Left border color |

#### `table`

| Key | Description |
|-----|-------------|
| `header_bg` | Header row background color |
| `header_text` | Header row text color |
| `body_text` | Body row text color |
| `header_size` | Header font size in points |
| `body_size` | Body font size in points |

---

## Requirements

Python 3.8+ — install all dependencies with:

```bash
pip install -r requirements.txt
```

Dependencies: `mistune`, `python-docx`, `lxml`, `mcp`
