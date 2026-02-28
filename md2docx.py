#!/usr/bin/env python3
"""
md2docx.py — Pure Markdown-to-DOCX converter (no cover page, no footer).

Converts standard Markdown (headings, paragraphs, lists, tables, blockquotes,
code) to a styled .docx document.  All styling is controlled by a JSON style
file; the bundled style_default.json uses neutral Calibri defaults.

Usage:
  python md2docx.py input.md output.docx
  python md2docx.py input.md output.docx --style my_style.json
"""

import argparse
import json
from pathlib import Path

import mistune
from lxml import etree
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ── style constants (defaults; overwritten by build_style_constants) ──────────

FONT_NAME       = 'Calibri'
FONT_NAME_SERIF = 'Calibri'   # H1 override (same unless style sets h1_override)

HEADING_STYLES = {
    1: {'size': 16, 'color': (0x00, 0x00, 0x00), 'bold': True,  'italic': False},
    2: {'size': 14, 'color': (0x1F, 0x38, 0x64), 'bold': True,  'italic': False},
    3: {'size': 12, 'color': (0x1F, 0x38, 0x64), 'bold': False, 'italic': True},
    4: {'size': 11, 'color': (0x40, 0x40, 0x40), 'bold': False, 'italic': True},
}
NORMAL      = {'size': 11, 'color': (0x21, 0x21, 0x21), 'bold': False, 'italic': False}
QUOTE       = {'size': 11, 'color': (0x55, 0x55, 0x55), 'bold': False, 'italic': True}
CODE_INLINE = {'size': 10, 'color': (0x33, 0x33, 0x33), 'bold': False, 'italic': False}
CODE_BLOCK  = {'size': 10, 'color': (0x33, 0x33, 0x33)}

QUOTE_BORDER_COLOR = 'AAAAAA'

TABLE_HEADER_BG   = '333333'
TABLE_HEADER_TEXT = (0xFF, 0xFF, 0xFF)
TABLE_HEADER_SIZE = 10
TABLE_BODY_TEXT   = (0x21, 0x21, 0x21)
TABLE_BODY_SIZE   = 10

DOC_LINE_SPACING = 1.15
DOC_MARGINS      = {'top': 2.5, 'bottom': 2.5, 'left': 2.5, 'right': 2.5}


# ── style loading ─────────────────────────────────────────────────────────────

def load_style(path: str) -> dict:
    """Load and return a JSON style guide file."""
    with open(path, encoding='utf-8') as f:
        return json.load(f)


def hex_to_rgb(h: str) -> tuple:
    """Convert '#RRGGBB' (or 'RRGGBB') hex string to (R, G, B) integer tuple."""
    h = h.lstrip('#')
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))


def build_style_constants(cfg: dict):
    """Populate module-level style constants from a loaded JSON style dict."""
    global FONT_NAME, FONT_NAME_SERIF, HEADING_STYLES, NORMAL, QUOTE, CODE_INLINE, CODE_BLOCK
    global QUOTE_BORDER_COLOR
    global TABLE_HEADER_BG, TABLE_HEADER_TEXT, TABLE_HEADER_SIZE
    global TABLE_BODY_TEXT, TABLE_BODY_SIZE
    global DOC_LINE_SPACING, DOC_MARGINS

    FONT_NAME       = cfg['fonts']['body']
    FONT_NAME_SERIF = cfg['fonts'].get('h1_override') or FONT_NAME

    heading_styles = {}
    for level, key in [(1, 'h1'), (2, 'h2'), (3, 'h3'), (4, 'h4')]:
        h = cfg['headings'][key]
        entry = {
            'size':   h['size'],
            'color':  hex_to_rgb(h['color']),
            'bold':   h['bold'],
            'italic': h['italic'],
        }
        if level == 1 and FONT_NAME_SERIF != FONT_NAME:
            entry['font'] = FONT_NAME_SERIF
        heading_styles[level] = entry
    HEADING_STYLES = heading_styles

    NORMAL = {
        'size':   cfg['body']['size'],
        'color':  hex_to_rgb(cfg['body']['color']),
        'bold':   False,
        'italic': False,
    }

    bq = cfg.get('blockquote', {})
    QUOTE = {
        'size':   bq.get('size', 11),
        'color':  hex_to_rgb(bq.get('color', '#555555')),
        'bold':   False,
        'italic': True,
    }
    QUOTE_BORDER_COLOR = bq.get('border_color', '#AAAAAA').lstrip('#')

    ci = cfg.get('code_inline', {})
    CODE_INLINE = {
        'size':   ci.get('size', 10),
        'color':  hex_to_rgb(ci.get('color', '#333333')),
        'bold':   False,
        'italic': False,
    }

    cb = cfg.get('code_block', {})
    CODE_BLOCK = {
        'size':  cb.get('size', 10),
        'color': hex_to_rgb(cb.get('color', '#333333')),
    }

    tbl = cfg.get('table', {})
    TABLE_HEADER_BG   = tbl.get('header_bg',   '#333333').lstrip('#')
    TABLE_HEADER_TEXT = hex_to_rgb(tbl.get('header_text', '#FFFFFF'))
    TABLE_HEADER_SIZE = tbl.get('header_size', 10)
    TABLE_BODY_TEXT   = hex_to_rgb(tbl.get('body_text',   '#212121'))
    TABLE_BODY_SIZE   = tbl.get('body_size',   10)

    doc_cfg = cfg.get('document', {})
    DOC_LINE_SPACING = doc_cfg.get('line_spacing', 1.15)
    margins = doc_cfg.get('margins', {})
    DOC_MARGINS = {
        'top':    margins.get('top_cm',    2.5),
        'bottom': margins.get('bottom_cm', 2.5),
        'left':   margins.get('left_cm',   2.5),
        'right':  margins.get('right_cm',  2.5),
    }


# ── helpers ──────────────────────────────────────────────────────────────────

def _apply_run(run, style, bold=None, italic=None):
    """Apply font settings to a docx Run."""
    run.font.name      = style.get('font', FONT_NAME)
    run.font.size      = Pt(style['size'])
    run.font.color.rgb = RGBColor(*style['color'])
    run.font.bold      = bold   if bold   is not None else style['bold']
    run.font.italic    = italic if italic is not None else style['italic']


def _set_cell_bg(cell, hex_color):
    """Set background colour of a table cell via raw XML."""
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement('w:shd')
    shd.set(qn('w:val'),   'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'),  hex_color)
    tcPr.append(shd)


def _add_para_border(para, side, color_hex, sz=6):
    """Add a single-line border on one side of a paragraph."""
    pPr  = para._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    edge = OxmlElement(f'w:{side}')
    edge.set(qn('w:val'),   'single')
    edge.set(qn('w:sz'),    str(sz))
    edge.set(qn('w:space'), '4')
    edge.set(qn('w:color'), color_hex)
    pBdr.append(edge)
    pPr.append(pBdr)


def _add_linebreak(para):
    """Insert a soft return (<w:br/>) inside a paragraph."""
    run = para.add_run()
    br  = OxmlElement('w:br')
    run._r.append(br)


# ── inline renderer ──────────────────────────────────────────────────────────

def render_inline(para, children, style, bold=False, italic=False):
    """Recursively walk mistune inline AST nodes and add formatted runs."""
    for token in (children or []):
        t = token.get('type', '')

        if t == 'text':
            raw = token.get('raw', '')
            if raw:
                run = para.add_run(raw)
                _apply_run(run, style,
                           bold=(style['bold'] or bold),
                           italic=(style['italic'] or italic))

        elif t == 'softline':
            run = para.add_run(' ')
            _apply_run(run, style,
                       bold=(style['bold'] or bold),
                       italic=(style['italic'] or italic))

        elif t == 'linebreak':
            _add_linebreak(para)

        elif t == 'strong':
            render_inline(para, token.get('children', []), style,
                          bold=True, italic=italic)

        elif t == 'emphasis':
            render_inline(para, token.get('children', []), style,
                          bold=bold, italic=True)

        elif t == 'link':
            # Render link text (no click-through needed)
            render_inline(para, token.get('children', []), style,
                          bold=bold, italic=italic)

        elif t == 'image':
            # Skip image references silently
            pass

        elif t == 'codespan':
            run = para.add_run(token.get('raw', ''))
            run.font.name      = 'Courier New'
            run.font.size      = Pt(CODE_INLINE['size'])
            run.font.color.rgb = RGBColor(*NORMAL['color'])
            run.font.bold      = False
            run.font.italic    = False

        elif t in ('raw_html', 'html_inline'):
            pass  # skip

        elif t == 'escape':
            raw = token.get('raw', '')
            if raw:
                run = para.add_run(raw)
                _apply_run(run, style,
                           bold=(style['bold'] or bold),
                           italic=(style['italic'] or italic))

        else:
            # Generic fallback: raw text or recurse into children
            raw = token.get('raw', '')
            if raw:
                run = para.add_run(raw)
                _apply_run(run, style,
                           bold=(style['bold'] or bold),
                           italic=(style['italic'] or italic))
            elif token.get('children'):
                render_inline(para, token['children'], style,
                              bold=bold, italic=italic)


# ── block renderers ──────────────────────────────────────────────────────────

def render_heading(doc, level, children):
    style = HEADING_STYLES.get(level, NORMAL)

    # Skip empty headings (image-only or whitespace-only)
    text_content = ''.join(
        c.get('raw', '') for c in (children or [])
        if c.get('type') in ('text', 'escape')
    ).strip()
    image_only = all(c.get('type') == 'image' for c in (children or [])) if children else False
    if not children or (not text_content and not image_only and
                        not any(c.get('children') for c in (children or []))):
        return  # truly empty heading — skip

    before = {1: 18, 2: 14, 3: 10, 4: 8}.get(level, 8)
    para = doc.add_paragraph()
    # Apply the built-in heading style so Google Docs / Word navigation
    # recognise the paragraph as a heading (structure), then our run-level
    # formatting (color, font, size) overrides the visual appearance.
    try:
        para.style = doc.styles[f'Heading {level}']
    except KeyError:
        pass
    para.paragraph_format.space_before = Pt(before)
    para.paragraph_format.space_after  = Pt(4)
    render_inline(para, children, style)


def render_paragraph(doc, children):
    # Skip pure-image paragraphs (e.g. ![][image1])
    if children and all(c.get('type') == 'image' for c in children):
        return

    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after  = Pt(6)
    render_inline(para, children, NORMAL)


def render_blockquote(doc, children):
    """Blockquote → indented italic paragraph with left border."""
    for child in (children or []):
        ct = child.get('type', '')
        if ct in ('paragraph', 'block_text'):
            para = doc.add_paragraph()
            para.paragraph_format.left_indent  = Inches(0.35)
            para.paragraph_format.right_indent = Inches(0.2)
            para.paragraph_format.space_before = Pt(3)
            para.paragraph_format.space_after  = Pt(3)
            _add_para_border(para, 'left', QUOTE_BORDER_COLOR, sz=12)
            render_inline(para, child.get('children', []), QUOTE)
        elif ct == 'list':
            render_list(doc, child)
        else:
            # Recurse (nested blockquotes, etc.)
            render_block(doc, child)


def render_list(doc, token):
    ordered = token.get('attrs', {}).get('ordered', False)
    depth   = token.get('attrs', {}).get('depth', 0)
    for item in (token.get('children') or []):
        if item.get('type') == 'list_item':
            render_list_item(doc, item, ordered, depth)


def render_list_item(doc, item_token, ordered, depth):
    """Render one list item, handling nested lists."""
    children = item_token.get('children') or []

    # Separate inline content (from block_text / paragraph) vs. nested lists
    inline_children = None
    nested_lists = []

    for child in children:
        ct = child.get('type', '')
        if ct == 'list':
            nested_lists.append(child)
        elif ct in ('paragraph', 'block_text'):
            if inline_children is None:
                inline_children = child.get('children', [])
        elif ct == 'blank_line':
            pass
        else:
            # Directly inline tokens (shouldn't happen, but be safe)
            if inline_children is None:
                inline_children = []
            inline_children.append(child)

    if inline_children is None:
        inline_children = []

    para = doc.add_paragraph()
    try:
        para.style = doc.styles['List Number' if ordered else 'List Bullet']
    except KeyError:
        pass

    para.paragraph_format.space_before = Pt(1)
    para.paragraph_format.space_after  = Pt(2)
    if depth > 0:
        para.paragraph_format.left_indent = Inches(0.25 + 0.25 * depth)

    render_inline(para, inline_children, NORMAL)

    # Ensure runs use the body font (list style might override font)
    for run in para.runs:
        run.font.name = FONT_NAME
        run.font.size = Pt(NORMAL['size'])

    # Recurse into nested lists
    for nested in nested_lists:
        nested_ordered = nested.get('attrs', {}).get('ordered', False)
        for nested_item in (nested.get('children') or []):
            if nested_item.get('type') == 'list_item':
                render_list_item(doc, nested_item, nested_ordered, depth + 1)


def render_table(doc, token):
    """Build a DOCX table from a mistune table token."""
    head_rows = []
    body_rows = []

    for section in (token.get('children') or []):
        st = section.get('type', '')
        if st == 'table_head':
            # mistune 3.x: table_head contains table_cell elements directly (no table_row wrapper)
            cells = [c for c in (section.get('children') or []) if c.get('type') == 'table_cell']
            if cells:
                head_rows.append(cells)
        elif st == 'table_body':
            for row in (section.get('children') or []):
                if row.get('type') == 'table_row':
                    body_rows.append(row.get('children') or [])

    all_rows = head_rows + body_rows
    if not all_rows:
        return

    num_cols = max(len(r) for r in all_rows)
    if num_cols == 0:
        return

    tbl = doc.add_table(rows=len(all_rows), cols=num_cols)
    tbl.style = 'Table Grid'

    for r_idx, row_cells in enumerate(all_rows):
        is_header = r_idx < len(head_rows)
        for c_idx, cell_token in enumerate(row_cells):
            if c_idx >= num_cols:
                break

            cell = tbl.cell(r_idx, c_idx)
            para = cell.paragraphs[0]
            para.clear()

            align = (cell_token.get('attrs') or {}).get('align')
            if align == 'center':
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif align == 'right':
                para.alignment = WD_ALIGN_PARAGRAPH.RIGHT

            cell_children = cell_token.get('children') or []

            if is_header:
                _set_cell_bg(cell, TABLE_HEADER_BG)
                h_style = {'size': TABLE_HEADER_SIZE, 'color': TABLE_HEADER_TEXT,
                           'bold': True, 'italic': False}
                render_inline(para, cell_children, h_style)
            else:
                t_style = {'size': TABLE_BODY_SIZE, 'color': TABLE_BODY_TEXT,
                           'bold': False, 'italic': False}
                render_inline(para, cell_children, t_style)

    # Small gap after table
    gap = doc.add_paragraph()
    gap.paragraph_format.space_after = Pt(6)


def render_hr(doc):
    """Horizontal rule → thin bottom-bordered paragraph."""
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(6)
    para.paragraph_format.space_after  = Pt(6)
    _add_para_border(para, 'bottom', 'C0C0C0', sz=4)


def render_code_block(doc, token):
    """Code block — rendered with body font + indent.
    Many 'code blocks' in practice are indented flow diagrams, not source code."""
    raw = token.get('raw', '').rstrip()
    if not raw:
        return
    for line in raw.split('\n'):
        para = doc.add_paragraph()
        run  = para.add_run(line)
        run.font.name      = FONT_NAME
        run.font.size      = Pt(CODE_BLOCK['size'])
        run.font.color.rgb = RGBColor(*CODE_BLOCK['color'])
        para.paragraph_format.left_indent  = Inches(0.4)
        para.paragraph_format.space_before = Pt(1)
        para.paragraph_format.space_after  = Pt(1)


# ── top-level dispatcher ─────────────────────────────────────────────────────

def render_block(doc, token):
    t        = token.get('type', '')
    children = token.get('children') or []
    attrs    = token.get('attrs')   or {}

    if t == 'heading':
        render_heading(doc, attrs.get('level', 1), children)

    elif t == 'paragraph':
        render_paragraph(doc, children)

    elif t == 'block_quote':
        render_blockquote(doc, children)

    elif t == 'list':
        render_list(doc, token)

    elif t == 'table':
        render_table(doc, token)

    elif t == 'thematic_break':
        render_hr(doc)

    elif t == 'block_code':
        render_code_block(doc, token)

    elif t == 'blank_line':
        pass  # paragraph spacing handles whitespace

    elif t == 'block_html':
        pass  # skip raw HTML blocks

    else:
        # Unknown block type — recurse into children if any
        for child in children:
            render_block(doc, child)


# ── document setup ────────────────────────────────────────────────────────────

def setup_document():
    doc = Document()

    # Remove the default empty paragraph Document() creates (if present)
    if doc.paragraphs:
        doc.paragraphs[0]._element.getparent().remove(doc.paragraphs[0]._element)

    sec = doc.sections[0]
    sec.top_margin    = Cm(DOC_MARGINS['top'])
    sec.bottom_margin = Cm(DOC_MARGINS['bottom'])
    sec.left_margin   = Cm(DOC_MARGINS['left'])
    sec.right_margin  = Cm(DOC_MARGINS['right'])

    normal_style = doc.styles['Normal']
    normal_style.font.name      = FONT_NAME
    normal_style.font.size      = Pt(NORMAL['size'])
    normal_style.font.color.rgb = RGBColor(*NORMAL['color'])
    normal_style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    normal_style.paragraph_format.line_spacing      = DOC_LINE_SPACING

    return doc


# ── main ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description='Convert Markdown to DOCX')
    parser.add_argument('input',  help='Input .md file')
    parser.add_argument('output', help='Output .docx file')
    parser.add_argument('--style', '-s',
                        default=str(Path(__file__).parent / 'style_default.json'),
                        help='JSON style guide (default: style_default.json)')
    args = parser.parse_args()

    cfg    = load_style(args.style)
    build_style_constants(cfg)
    tokens = mistune.create_markdown(renderer='ast', plugins=['table'])(
                 Path(args.input).read_text(encoding='utf-8'))
    doc = setup_document()
    for token in tokens:
        render_block(doc, token)
    doc.save(args.output)
    print(f'✓  Saved → {args.output}')


if __name__ == '__main__':
    main()
